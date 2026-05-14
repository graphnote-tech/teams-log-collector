// content.js — runs inside the Microsoft Teams tab.
//
// What this script does:
//   1. Reads the MSAL access-token cache that the Teams web client leaves in
//      localStorage.
//   2. Talks to either Microsoft Graph (if a Chat.Read-bearing Graph token is
//      cached) or the Teams internal messaging API
//      (*.ng.msg.teams.microsoft.com) to retrieve chat messages.
//   3. Streams progress back to popup.js over a long-lived port, then triggers
//      a JSON file download from inside the Teams page.
//
// This script never modifies the page or the user's Teams data; it only reads
// the auth cache and makes the same kind of GET requests the Teams web client
// itself makes.

(() => {
  "use strict";

  // Bump on every behavioural change. popup.js compares against this via the
  // ping handshake and triggers a tab reload if the values don't match — that
  // is the only reliable way to evict a previously-injected stale listener.
  const TLC_VERSION = "0.0.1";

  // If a previous run registered a listener under this key, drop it before we
  // register the new one. This keeps re-injections idempotent.
  if (window.__tlcListener) {
    try {
      chrome.runtime.onConnect.removeListener(window.__tlcListener);
    } catch {
      /* ignore */
    }
  }
  window.__tlcVersion = TLC_VERSION;

  // ----- Constants -----

  const GRAPH = "https://graph.microsoft.com/v1.0";

  // Audiences that identify a Microsoft Graph access token (URL form + GUID form).
  const GRAPH_AUDIENCES = new Set([
    "https://graph.microsoft.com",
    "00000003-0000-0000-c000-000000000000",
  ]);

  // Scopes that grant read access to chat messages via Graph.
  const CHAT_READ_SCOPES = new Set([
    "chat.read",
    "chat.readbasic",
    "chat.readwrite",
    "chat.readwrite.all",
    "chatmessage.read",
    "chatmessage.read.all",
  ]);

  // Fallback ordering when the skypeToken response does not name a region.
  const TEAMS_REGIONS = ["apac", "amer", "emea", "india", "br"];

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ----- Token discovery -----

  function decodeJwt(token) {
    if (!token || typeof token !== "string") return null;
    const part = token.split(".")[1];
    if (!part) return null;
    try {
      const padded = part + "=".repeat((4 - (part.length % 4)) % 4);
      const b64 = padded.replace(/-/g, "+").replace(/_/g, "/");
      return JSON.parse(atob(b64));
    } catch {
      return null;
    }
  }

  function enumerateAccessTokens() {
    const out = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      const raw = localStorage.getItem(key);
      if (!raw || raw[0] !== "{") continue;
      let obj;
      try {
        obj = JSON.parse(raw);
      } catch {
        continue;
      }
      if (obj.credentialType !== "AccessToken" || !obj.secret) continue;
      const jwt = decodeJwt(obj.secret);
      if (!jwt) continue;
      out.push({
        secret: obj.secret,
        target: obj.target || "",
        aud: jwt.aud,
        scp: jwt.scp || "",
        expMs: (jwt.exp || 0) * 1000,
        upn: jwt.upn || jwt.preferred_username,
      });
    }
    return out;
  }

  function scopesOf(t) {
    return t.scp.toLowerCase().split(/\s+/).filter(Boolean);
  }

  function findTokenByAudience(audCandidates, requireChatRead = false) {
    const auds = new Set(
      Array.isArray(audCandidates) ? audCandidates : [audCandidates]
    );
    const now = Date.now();
    let matches = enumerateAccessTokens().filter(
      (t) => t.expMs > now + 60_000 && auds.has(t.aud)
    );
    if (requireChatRead) {
      matches = matches.filter((t) =>
        scopesOf(t).some((s) => CHAT_READ_SCOPES.has(s))
      );
    }
    matches.sort((a, b) => b.expMs - a.expMs);
    return matches[0] || null;
  }

  function diagnose() {
    return enumerateAccessTokens()
      .sort((a, b) => b.expMs - a.expMs)
      .map((t) => ({
        aud: t.aud,
        scopes: t.scp,
        expiresInSec: Math.max(0, Math.round((t.expMs - Date.now()) / 1000)),
        upn: t.upn,
        hasChatRead: scopesOf(t).some((s) => CHAT_READ_SCOPES.has(s)),
      }));
  }

  // ----- Chat ID parsing -----

  // Accepts either a raw thread id, an encoded URL, or anything in between, and
  // returns the canonical "19:xxx@thread.v2" form (or @unq.gbl.spaces for 1:1).
  function parseChatIdFromUrl(input) {
    if (!input) return null;
    let s = String(input).trim();
    try {
      s = decodeURIComponent(s);
    } catch {
      /* keep raw */
    }
    const m = s.match(
      /19:[A-Za-z0-9_\-+=/]+@(?:thread\.v2|thread\.tacv2|unq\.gbl\.spaces|thread\.skype)/
    );
    return m ? m[0] : null;
  }

  // ----- Misc helpers -----

  function stripHtml(html) {
    if (!html) return "";
    const div = document.createElement("div");
    div.innerHTML = html;
    return (div.textContent || "").trim();
  }

  function graphSender(m) {
    return (
      m.from?.user?.displayName ||
      m.from?.application?.displayName ||
      m.from?.device?.displayName ||
      (m.messageType === "systemEventMessage" ? "system" : "unknown")
    );
  }

  function safeLabel(s) {
    return String(s || "chat")
      .replace(/[^\w぀-ヿ一-龯-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 60);
  }

  function fmtDate(d) {
    return d.toISOString().slice(0, 10);
  }

  function buildFilename(label, fromDate, toDate) {
    const safe = safeLabel(label);
    let range;
    if (fromDate && toDate) {
      range = `${fmtDate(fromDate)}_to_${fmtDate(toDate)}`;
    } else if (fromDate) {
      range = `from_${fmtDate(fromDate)}`;
    } else if (toDate) {
      range = `until_${fmtDate(toDate)}`;
    } else {
      range = `at_${fmtDate(new Date())}`;
    }
    return `teams-log_${safe}_${range}.json`;
  }

  function parseDate(d, endOfDay = false) {
    if (!d) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      return new Date(d + (endOfDay ? "T23:59:59.999Z" : "T00:00:00.000Z"));
    }
    const parsed = new Date(d);
    if (Number.isNaN(parsed.getTime())) throw new Error(`Invalid date: ${d}`);
    return parsed;
  }

  function downloadJSON(payload, filename) {
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ----- Microsoft Graph path -----

  async function graphFetch(url, token) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.status === 429) {
      const retry = Number(res.headers.get("Retry-After") || 5);
      await sleep(retry * 1000);
      return graphFetch(url, token);
    }
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      const err = new Error(`Graph ${res.status}`);
      err.status = res.status;
      err.body = body;
      throw err;
    }
    return res.json();
  }

  function chatLabel(chat) {
    if (chat.topic) return chat.topic;
    const names = (chat.members || [])
      .map((m) => m.displayName)
      .filter(Boolean)
      .slice(0, 5);
    return names.join(", ") || "(untitled)";
  }

  async function graphGetChat(token, chatId) {
    const url = `${GRAPH}/chats/${encodeURIComponent(chatId)}?$expand=members`;
    const data = await graphFetch(url, token);
    return { id: data.id, type: data.chatType, label: chatLabel(data) };
  }

  async function graphListChats(token) {
    const all = [];
    let url = `${GRAPH}/me/chats?$expand=members&$top=50`;
    while (url) {
      const data = await graphFetch(url, token);
      all.push(...(data.value || []));
      url = data["@odata.nextLink"];
    }
    return all;
  }

  async function graphCollectMessages(token, chatId, fromDate, toDate, onProgress) {
    const kept = [];
    let url = `${GRAPH}/chats/${encodeURIComponent(chatId)}/messages?$top=50`;
    let page = 0;
    while (url) {
      const data = await graphFetch(url, token);
      page++;
      // Graph returns chat messages ordered by lastModifiedDateTime DESC.
      // Because lastModified >= created, once the page's oldest
      // lastModifiedDateTime is below `fromDate`, every subsequent message
      // will also have createdDateTime below `fromDate`, so we can stop.
      let pageOldestLastMod = null;
      for (const m of data.value || []) {
        const created = new Date(m.createdDateTime);
        const lastMod = new Date(m.lastModifiedDateTime || m.createdDateTime);
        if (!pageOldestLastMod || lastMod < pageOldestLastMod) {
          pageOldestLastMod = lastMod;
        }
        if (toDate && created > toDate) continue;
        if (fromDate && created < fromDate) continue;
        kept.push({
          id: m.id,
          createdDateTime: m.createdDateTime,
          from: graphSender(m),
          text: stripHtml(m.body?.content),
        });
      }
      onProgress({
        page,
        kept: kept.length,
        oldest: pageOldestLastMod?.toISOString(),
      });
      if (fromDate && pageOldestLastMod && pageOldestLastMod < fromDate) break;
      url = data["@odata.nextLink"];
      await sleep(150);
    }
    return kept.sort((a, b) =>
      a.createdDateTime.localeCompare(b.createdDateTime)
    );
  }

  // ----- Teams internal messaging API path -----
  //
  // High-level flow:
  //   POST https://authsvc.teams.microsoft.com/v1.0/authz
  //        with Authorization: Bearer <AAD token, aud=api.spaces.skype.com>
  //   → response.tokens.skypeToken is what we authenticate with from here on.
  //
  //   GET https://{region}.ng.msg.teams.microsoft.com/v1/users/ME/
  //          conversations/{threadId}/messages?...
  //        with Authentication: skypetoken=<skypeToken>
  //   ↑ note: the header is "Authentication" (not "Authorization"), and the
  //     value uses the "skypetoken=" prefix. This is the same shape Teams web
  //     uses internally.

  /** Cached skypeToken so we don't hit /authz on every page. */
  let cachedSkypeToken = null;

  async function getSkypeToken() {
    if (cachedSkypeToken && cachedSkypeToken.expMs > Date.now() + 60_000) {
      return cachedSkypeToken.token;
    }
    const aad = findTokenByAudience("https://api.spaces.skype.com");
    if (!aad) {
      throw new Error(
        "No cached AAD token with aud=api.spaces.skype.com. Reopen Teams and retry."
      );
    }
    const res = await fetch("https://authsvc.teams.microsoft.com/v1.0/authz", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${aad.secret}`,
        "Content-Type": "application/json",
      },
      body: "{}",
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`authsvc/authz ${res.status}: ${body.slice(0, 250)}`);
    }
    const data = await res.json();
    const token = data.tokens?.skypeToken || data.skypeToken;
    if (!token) {
      throw new Error("authz response did not contain a skypeToken");
    }
    const expiresIn = Number(data.tokens?.expiresIn || data.expiresIn || 3600);

    // `regionGtms.msg` from authz can be either of:
    //   • "https://apac.ng.msg.teams.microsoft.com"          ← what we want
    //   • "https://teams.cloud.microsoft/api/mt/apac"         ← middleTier proxy
    // The proxy uses a different URL shape (404 if we naively append our path),
    // so extract the region letters and build the direct URL ourselves.
    const rawMsg = data.regionGtms?.msg || data.regionGtms?.middleTier;
    const region =
      data.tokens?.region ||
      data.region ||
      extractRegionFromUrl(rawMsg) ||
      null;

    cachedSkypeToken = {
      token,
      expMs: Date.now() + expiresIn * 1000,
      msgEndpoint: region ? `https://${region}.ng.msg.teams.microsoft.com` : null,
      region,
    };
    return token;
  }

  function extractRegionFromUrl(url) {
    if (!url) return null;
    let m = url.match(/\/api\/mt\/([a-z]+)/); // proxy form
    if (m) return m[1];
    m = url.match(/https?:\/\/([a-z]+)\.ng\.msg\./); // direct form
    if (m) return m[1];
    return null;
  }

  // Retry-aware GET against Teams messaging service. Honors Retry-After on
  // 429/503 and falls back to capped exponential backoff. Bounded to 6 tries
  // so a persistently-throttled chat surfaces as a clear error instead of
  // looping forever.
  async function teamsFetchJson(url, skypeToken) {
    let attempt = 0;
    while (true) {
      const res = await fetch(url, {
        headers: {
          // NOT "Authorization". Teams' messaging service rejects Bearer here.
          Authentication: `skypetoken=${skypeToken}`,
          BehaviorOverride: "redirectAs404",
        },
      });
      if (res.status === 429 || res.status === 503) {
        attempt++;
        if (attempt > 6) {
          const body = await res.text().catch(() => "");
          const err = new Error(`Teams ${res.status} after ${attempt} retries`);
          err.status = res.status;
          err.body = body;
          throw err;
        }
        const retryAfter = Number(res.headers.get("Retry-After"));
        const wait =
          Number.isFinite(retryAfter) && retryAfter > 0
            ? retryAfter * 1000
            : Math.min(30000, 1000 * 2 ** attempt);
        await sleep(wait);
        continue;
      }
      if (!res.ok) {
        const body = await res.text().catch(() => "");
        const err = new Error(`Teams ${res.status}`);
        err.status = res.status;
        err.body = body;
        throw err;
      }
      return res.json();
    }
  }

  /**
   * Hit one page of /v1/users/ME/conversations/{thread}/messages.
   * `endpointBase` is either a region name ("apac") or a full URL
   * ("https://apac.ng.msg.teams.microsoft.com").
   */
  async function teamsFetchPage(endpointBase, threadId, skypeToken, params) {
    const qs = new URLSearchParams(params).toString();
    const base = endpointBase.startsWith("http")
      ? endpointBase.replace(/\/$/, "")
      : `https://${endpointBase}.ng.msg.teams.microsoft.com`;
    const url = `${base}/v1/users/ME/conversations/${encodeURIComponent(
      threadId
    )}/messages?${qs}`;
    try {
      return await teamsFetchJson(url, skypeToken);
    } catch (e) {
      if (e.status) e.message = `Teams ${e.status} (${endpointBase})`;
      throw e;
    }
  }

  /** Figure out which messaging region serves this thread. */
  async function teamsResolveEndpoint(threadId, skypeToken) {
    if (cachedSkypeToken?.msgEndpoint) {
      try {
        await teamsFetchPage(cachedSkypeToken.msgEndpoint, threadId, skypeToken, {
          pageSize: 1,
        });
        return cachedSkypeToken.msgEndpoint;
      } catch (e) {
        if (e.status === 401 || e.status === 403) throw e;
        // 404 / 5xx → fall through and probe by region name
      }
    }
    const preferred = cachedSkypeToken?.region;
    const order = preferred
      ? [preferred, ...TEAMS_REGIONS.filter((r) => r !== preferred)]
      : TEAMS_REGIONS;
    let lastErr = null;
    for (const r of order) {
      try {
        await teamsFetchPage(r, threadId, skypeToken, { pageSize: 1 });
        return r;
      } catch (e) {
        lastErr = e;
        if (e.status === 401 || e.status === 403) throw e;
      }
    }
    throw new Error(
      `Teams internal API: no region reachable (${order.join(", ")}). ${
        lastErr ? "last=" + lastErr.message : ""
      }`
    );
  }

  function teamsMessageNormalize(m) {
    const created =
      m.composetime || m.originalarrivaltime || m.createdtime || m.composeTime;
    return {
      id: m.id || m.messageid,
      createdDateTime: created ? new Date(created).toISOString() : null,
      from: m.imdisplayname || m.from || "unknown",
      text: stripHtml(m.content || m.body?.content || ""),
    };
  }

  async function teamsCollectMessages(threadId, fromDate, toDate, onProgress) {
    onProgress({ page: 0, kept: 0, oldest: "exchanging AAD token for skypeToken..." });
    const skypeToken = await getSkypeToken();
    onProgress({
      page: 0,
      kept: 0,
      oldest: `skypeToken obtained${
        cachedSkypeToken?.region ? ` (region=${cachedSkypeToken.region})` : ""
      }`,
    });

    const endpoint = await teamsResolveEndpoint(threadId, skypeToken);
    onProgress({ page: 0, kept: 0, oldest: `endpoint=${endpoint}` });

    const kept = [];
    let page = 0;
    let startTime = toDate ? toDate.getTime() : Date.now();
    let done = false;

    while (!done) {
      page++;
      // First page: no startTime constraint so we get the latest messages.
      // Subsequent pages: walk backwards via startTime.
      const params =
        page === 1
          ? {
              pageSize: 200,
              view: "msnp24Equivalent|supportsMessageProperties",
            }
          : {
              pageSize: 200,
              startTime: String(startTime),
              view: "msnp24Equivalent|supportsMessageProperties",
            };
      let data;
      try {
        data = await teamsFetchPage(endpoint, threadId, skypeToken, params);
      } catch (e) {
        throw new Error(
          `Teams API page ${page} failed: ${e.message}${
            e.body ? "\n  " + e.body.slice(0, 200) : ""
          }`
        );
      }
      const batch = data.messages || data.value || data.posts || [];
      if (!batch.length) break;

      let pageOldest = null;
      for (const m of batch) {
        const norm = teamsMessageNormalize(m);
        if (!norm.createdDateTime) continue;
        const t = new Date(norm.createdDateTime);
        if (!pageOldest || t < pageOldest) pageOldest = t;
        if (toDate && t > toDate) continue;
        if (fromDate && t < fromDate) continue;
        kept.push(norm);
      }
      onProgress({
        page,
        kept: kept.length,
        oldest: pageOldest?.toISOString(),
      });

      if (fromDate && pageOldest && pageOldest < fromDate) {
        done = true;
        break;
      }
      // Advance: prefer the sync link from `_metadata`, else step `startTime`
      // backwards by the page's oldest message timestamp.
      const nextSync =
        data._metadata?.syncState ||
        data._links?.next ||
        data["@odata.nextLink"];
      if (nextSync && /^https?:\/\//.test(nextSync)) {
        let next;
        try {
          next = await teamsFetchJson(nextSync, skypeToken);
        } catch {
          break;
        }
        const nextBatch = next.messages || [];
        if (!nextBatch.length) break;
        startTime = new Date(
          teamsMessageNormalize(nextBatch[nextBatch.length - 1]).createdDateTime
        ).getTime();
        continue;
      }
      if (pageOldest) {
        const newStart = pageOldest.getTime() - 1;
        if (newStart >= startTime) break; // avoid infinite loop on no-progress
        startTime = newStart;
      } else {
        break;
      }
      await sleep(150);
    }
    return kept.sort((a, b) =>
      a.createdDateTime.localeCompare(b.createdDateTime)
    );
  }

  // ----- Top-level orchestration -----

  async function run(params, progress) {
    const fromDate = params.from ? parseDate(params.from, false) : null;
    const toDate = params.to ? parseDate(params.to, true) : null;
    const chatIdFromInput = parseChatIdFromUrl(params.chat);

    // Choose a Graph token: caller-supplied wins, else cached Graph token that
    // happens to bear Chat.Read. If neither exists we'll fall through to the
    // Teams internal API path.
    let graphToken = params.accessToken || null;
    if (graphToken) {
      const jwt = decodeJwt(graphToken);
      if (jwt && !GRAPH_AUDIENCES.has(jwt.aud)) {
        progress(
          `warning: supplied token has aud "${jwt.aud}" (not Graph); trying anyway`
        );
      }
    } else if (!params.useTeamsApi) {
      const t = findTokenByAudience([...GRAPH_AUDIENCES], true);
      if (t) graphToken = t.secret;
    }

    let chatInfo = null;
    let messages = null;
    let via = null;
    let graphError = null;

    if (graphToken && !params.useTeamsApi) {
      progress("[Graph] Token acquired; resolving chat...");
      try {
        if (chatIdFromInput) {
          chatInfo = await graphGetChat(graphToken, chatIdFromInput);
        } else {
          progress("[Graph] Searching /me/chats for a name match...");
          const all = await graphListChats(graphToken);
          const q = params.chat.toLowerCase();
          const matched = all
            .map((c) => ({ id: c.id, type: c.chatType, label: chatLabel(c) }))
            .filter((c) => c.label.toLowerCase().includes(q));
          if (matched.length === 0)
            throw new Error(`No chats matched "${params.chat}"`);
          if (matched.length > 1) {
            const list = matched
              .map((m, i) => `  [${i}] ${m.label} (${m.type}) id=${m.id}`)
              .join("\n");
            throw new Error(
              `Multiple chats matched. Paste a specific URL or chat id instead:\n${list}`
            );
          }
          chatInfo = matched[0];
        }
        progress(`[Graph] chat: ${chatInfo.label} (${chatInfo.id})`);
        progress("[Graph] Collecting messages...");
        messages = await graphCollectMessages(
          graphToken,
          chatInfo.id,
          fromDate,
          toDate,
          (p) =>
            progress(
              `  page ${p.page} | kept ${p.kept} | oldest ${p.oldest || "?"}`
            )
        );
        via = "graph";
      } catch (e) {
        graphError = e;
        progress(`[Graph] failed: ${e.message}`);
        if (e.body) progress(`  ${e.body.slice(0, 250)}`);
      }
    }

    if (!messages) {
      if (!chatIdFromInput) {
        const reason = graphError
          ? `Graph path failed (${graphError.message}); the Teams internal API requires a URL/ID`
          : "The Teams internal API needs a URL or chat ID (name search is not supported there)";
        throw new Error(
          `${reason}. In Teams, open the chat, click "..." → "Copy link to chat" and paste it back.`
        );
      }
      progress("[Teams-API] Collecting via the Teams internal API (experimental)...");
      chatInfo = chatInfo || {
        id: chatIdFromInput,
        type: "unknown",
        label: chatIdFromInput,
      };
      messages = await teamsCollectMessages(
        chatIdFromInput,
        fromDate,
        toDate,
        (p) =>
          progress(
            `  page ${p.page} | kept ${p.kept} | oldest ${p.oldest || "?"}`
          )
      );
      via = "teams-internal";
    }

    // Prefer the caller-supplied filename label; fall back to whatever we
    // resolved (which may be a raw chat id when the Teams internal API was
    // used and we could not look up a friendly name).
    const filenameLabel = (params.nameOverride || "").trim() || chatInfo.label;
    const payload = {
      chat: chatInfo,
      range: {
        from: fromDate?.toISOString() || null,
        to: toDate?.toISOString() || null,
      },
      via,
      collectedAt: new Date().toISOString(),
      count: messages.length,
      messages,
    };
    const filename = buildFilename(filenameLabel, fromDate, toDate);
    progress(`[done] ${messages.length} messages → ${filename} (${via})`);
    downloadJSON(payload, filename);

    return {
      count: messages.length,
      filename,
      chat: chatInfo,
      via,
    };
  }

  // ----- Port listener -----

  const listener = (port) => {
    if (port.name !== "tlc") return;
    const safePost = (m) => {
      try {
        port.postMessage(m);
      } catch {
        /* port closed; ignore */
      }
    };
    port.onMessage.addListener(async (msg) => {
      try {
        if (msg.type === "ping") {
          safePost({ type: "done", pong: true, version: TLC_VERSION });
        } else if (msg.type === "diagnose") {
          safePost({ type: "done", diagnose: diagnose() });
        } else if (msg.type === "collect") {
          const summary = await run(msg.params, (text) =>
            safePost({ type: "progress", text })
          );
          safePost({ type: "done", summary });
        } else {
          safePost({ type: "error", error: `Unknown message type: ${msg.type}` });
        }
      } catch (e) {
        safePost({ type: "error", error: String(e?.message || e) });
      }
    });
  };
  window.__tlcListener = listener;
  chrome.runtime.onConnect.addListener(listener);
})();
