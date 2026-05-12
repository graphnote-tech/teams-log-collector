/*
 * Teams Log Collector (browser console version)
 *
 * Collects messages from a Microsoft Teams group chat as a regular user,
 * using Microsoft Graph API via the access token cached by the logged-in
 * Teams web client. No admin rights required.
 *
 * HOW TO USE
 * ----------
 * 1. Open https://teams.microsoft.com in your browser and sign in.
 * 2. Open DevTools (F12) -> Console tab. Make sure the console's JavaScript
 *    context is the main teams.microsoft.com frame (not an iframe).
 * 3. Paste the ENTIRE contents of this file and press Enter.
 * 4. Run with ONE of the following chat selectors:
 *
 *      // (a) by URL copied from Teams ("Copy link to chat" / "Copy link to
 *      //     message"). Easiest — no search needed.
 *      await TeamsLogCollector.run({
 *        chatUrl: "https://teams.microsoft.com/l/message/19%3Axxx%40thread.v2/1700000000?...",
 *        from:    "2025-01-01",
 *        to:      "2025-12-31",
 *      })
 *
 *      // (b) by raw chat id (e.g. "19:xxxxx@thread.v2")
 *      await TeamsLogCollector.run({ chatId: "19:xxxxx@thread.v2", from: "...", to: "..." })
 *
 *      // (c) by chat name (partial match against topic or member names)
 *      await TeamsLogCollector.run({
 *        chatName: "Sales team",
 *        from:     "2025-01-01",
 *        to:       "2025-12-31",
 *      })
 *
 *    A JSON file will be downloaded automatically.
 *
 * OUTPUT SHAPE
 * ------------
 *   {
 *     "chat":        { "id": "...", "type": "group", "label": "Sales team" },
 *     "range":       { "from": "...", "to": "..." },
 *     "collectedAt": "2026-05-12T01:23:45.678Z",
 *     "count":       1234,
 *     "messages": [
 *       { "id": "...", "createdDateTime": "...", "from": "John Doe", "text": "..." },
 *       ...
 *     ]
 *   }
 *
 * NOTES
 * -----
 * - Access tokens expire ~1 hour after sign-in. If a long-running collection
 *   fails mid-way, reload the Teams tab to refresh the token, then re-run.
 * - Some tenants block direct Graph API calls via Conditional Access, or the
 *   Teams web app's Graph token may not include Chat.Read. In those cases the
 *   browser-extension version of this tool (see ./extension) is recommended —
 *   it can fall back to the Teams internal messaging API.
 * - This script only reads — it never writes/edits/deletes anything.
 */

(() => {
  "use strict";

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const GRAPH_AUDIENCES = new Set([
    "https://graph.microsoft.com",
    "00000003-0000-0000-c000-000000000000",
  ]);
  const CHAT_SCOPES = new Set([
    "chat.read",
    "chat.readbasic",
    "chat.readwrite",
    "chat.readwrite.all",
    "chatmessage.read",
    "chatmessage.read.all",
  ]);

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
    const tokens = [];
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
      tokens.push({
        key,
        secret: obj.secret,
        target: obj.target || "",
        expMs: (jwt.exp || 0) * 1000,
        aud: jwt.aud,
        scp: jwt.scp || "",
        tenant: jwt.tid,
        upn: jwt.upn || jwt.preferred_username,
      });
    }
    return tokens;
  }

  function findAccessToken() {
    const all = enumerateAccessTokens();
    const fresh = all.filter((t) => t.expMs > Date.now() + 60_000);
    const graph = fresh.filter((t) => GRAPH_AUDIENCES.has(t.aud));
    if (!graph.length) {
      const auds = [...new Set(fresh.map((t) => t.aud))].join(", ") || "(none)";
      throw new Error(
        "No Microsoft Graph access token found in localStorage.\n" +
          `Fresh access-token audiences in cache: ${auds}\n` +
          "Teams web caches tokens for its internal chat services rather than Graph,\n" +
          "so a Graph token with Chat.Read is not always present here. Options:\n" +
          "  1) Run TeamsLogCollector.diagnose() to inspect cached tokens.\n" +
          "  2) Open https://developer.microsoft.com/graph/graph-explorer, sign in, run\n" +
          '     any "Chat.Read" sample query, then paste the Access Token from the\n' +
          '     "Access token" tab into run({ accessToken: "..." }).\n' +
          "  3) Run this script in a tab that already holds a Graph+Chat.Read token."
      );
    }
    const withChat = graph.filter((t) =>
      t.scp
        .toLowerCase()
        .split(/\s+/)
        .some((s) => CHAT_SCOPES.has(s))
    );
    const pool = withChat.length ? withChat : graph;
    pool.sort((a, b) => b.expMs - a.expMs);
    if (!withChat.length) {
      console.warn(
        "[teams-log] Found a Graph token but no Chat.Read scope; trying anyway.\n" +
          `  scopes: ${pool[0].scp}`
      );
    }
    return pool[0].secret;
  }

  function diagnose() {
    const rows = enumerateAccessTokens().map((t) => ({
      aud: t.aud,
      scopes: t.scp,
      expiresInSec: Math.round((t.expMs - Date.now()) / 1000),
      target: t.target.slice(0, 80),
      upn: t.upn,
    }));
    console.table(rows);
    console.log(
      `[teams-log] ${rows.length} access-token entries in localStorage.`
    );
    return rows;
  }

  async function graphFetch(url, token) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.status === 429) {
      const retry = Number(res.headers.get("Retry-After") || 5);
      console.warn(`[teams-log] rate-limited; retry in ${retry}s`);
      await sleep(retry * 1000);
      return graphFetch(url, token);
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`Graph ${res.status} for ${url}\n${body}`);
    }
    return res.json();
  }

  async function listChats(token) {
    const chats = [];
    let url = `${GRAPH}/me/chats?$expand=members&$top=50`;
    while (url) {
      const data = await graphFetch(url, token);
      chats.push(...(data.value || []));
      url = data["@odata.nextLink"];
    }
    return chats;
  }

  function parseChatIdFromUrl(input) {
    if (!input) return null;
    let s = String(input).trim();
    // The id may appear url-encoded (%3A %40) or plain (: @).
    try {
      s = decodeURIComponent(s);
    } catch {
      /* keep raw */
    }
    // Match Teams thread id forms: 19:xxxxx@thread.v2 / @thread.tacv2 /
    // @unq.gbl.spaces (1:1) / @thread.skype (legacy).
    const m = s.match(
      /19:[A-Za-z0-9_\-+=/]+@(?:thread\.v2|thread\.tacv2|unq\.gbl\.spaces|thread\.skype)/
    );
    return m ? m[0] : null;
  }

  async function getChatById(token, chatId) {
    const url = `${GRAPH}/chats/${encodeURIComponent(chatId)}?$expand=members`;
    const chat = await graphFetch(url, token);
    return {
      id: chat.id,
      type: chat.chatType,
      label: chatLabel(chat),
    };
  }

  function chatLabel(chat) {
    if (chat.topic) return chat.topic;
    const names = (chat.members || [])
      .map((m) => m.displayName)
      .filter(Boolean)
      .slice(0, 5);
    return names.join(", ") || "(untitled)";
  }

  async function searchChatsByName(token, query) {
    const all = await listChats(token);
    const q = String(query).toLowerCase();
    return all
      .map((c) => ({
        id: c.id,
        type: c.chatType,
        label: chatLabel(c),
        lastUpdated: c.lastUpdatedDateTime,
      }))
      .filter((c) => c.label.toLowerCase().includes(q))
      .sort((a, b) => (b.lastUpdated || "").localeCompare(a.lastUpdated || ""));
  }

  function stripHtml(html) {
    if (!html) return "";
    const div = document.createElement("div");
    div.innerHTML = html;
    return (div.textContent || "").trim();
  }

  function senderName(m) {
    return (
      m.from?.user?.displayName ||
      m.from?.application?.displayName ||
      m.from?.device?.displayName ||
      (m.messageType === "systemEventMessage" ? "system" : "unknown")
    );
  }

  async function collectMessages(token, chatId, fromDate, toDate, onProgress) {
    const kept = [];
    let url = `${GRAPH}/chats/${encodeURIComponent(chatId)}/messages?$top=50`;
    let page = 0;
    while (url) {
      const data = await graphFetch(url, token);
      page++;
      const batch = data.value || [];
      let pageOldestLastModified = null;
      for (const m of batch) {
        const created = new Date(m.createdDateTime);
        const lastMod = new Date(m.lastModifiedDateTime || m.createdDateTime);
        if (!pageOldestLastModified || lastMod < pageOldestLastModified) {
          pageOldestLastModified = lastMod;
        }
        if (toDate && created > toDate) continue;
        if (fromDate && created < fromDate) continue;
        kept.push({
          id: m.id,
          createdDateTime: m.createdDateTime,
          from: senderName(m),
          text: stripHtml(m.body?.content),
        });
      }
      onProgress?.({
        page,
        collected: kept.length,
        oldestSeen: pageOldestLastModified?.toISOString(),
      });
      // Messages are returned ordered by lastModifiedDateTime desc.
      // Because lastModifiedDateTime >= createdDateTime always, once the page's
      // oldest lastModifiedDateTime is below `fromDate`, every subsequent
      // message also has createdDateTime below `fromDate` -> safe to stop.
      if (fromDate && pageOldestLastModified && pageOldestLastModified < fromDate) {
        break;
      }
      url = data["@odata.nextLink"];
      await sleep(150);
    }
    return kept.sort((a, b) =>
      a.createdDateTime.localeCompare(b.createdDateTime)
    );
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
    URL.revokeObjectURL(url);
  }

  function parseDate(d, endOfDay = false) {
    if (!d) return null;
    if (d instanceof Date) return d;
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      return new Date(d + (endOfDay ? "T23:59:59.999Z" : "T00:00:00.000Z"));
    }
    const parsed = new Date(d);
    if (Number.isNaN(parsed.getTime())) {
      throw new Error(`Invalid date: ${d}`);
    }
    return parsed;
  }

  function pickChat(matches) {
    if (matches.length === 0) {
      throw new Error("No chats matched the given name.");
    }
    if (matches.length === 1) return matches[0];
    console.log("[teams-log] Multiple chats matched:");
    matches.forEach((c, i) =>
      console.log(`  [${i}] ${c.label}  (type=${c.type}, id=${c.id})`)
    );
    const ans = window.prompt(
      `Multiple chats matched. Enter index 0-${matches.length - 1}:`,
      "0"
    );
    const idx = Number(ans);
    if (!Number.isInteger(idx) || idx < 0 || idx >= matches.length) {
      throw new Error("Selection cancelled or invalid.");
    }
    return matches[idx];
  }

  async function resolveChat(token, { chatUrl, chatId, chatName }) {
    const idFromUrl = chatUrl ? parseChatIdFromUrl(chatUrl) : null;
    if (chatUrl && !idFromUrl) {
      throw new Error(
        `Could not extract a Teams chat id from chatUrl: ${chatUrl}`
      );
    }
    const directId = idFromUrl || (chatId ? parseChatIdFromUrl(chatId) || chatId : null);
    if (directId) {
      console.log(`[teams-log] [2/4] resolving chat by id ${directId}`);
      return await getChatById(token, directId);
    }
    if (!chatName) {
      throw new Error(
        "Provide one of: chatUrl, chatId, or chatName."
      );
    }
    console.log(`[teams-log] [2/4] searching chats matching "${chatName}"...`);
    const matches = await searchChatsByName(token, chatName);
    console.log(`[teams-log]   found ${matches.length} match(es)`);
    return pickChat(matches);
  }

  async function run({ chatName, chatId, chatUrl, from, to, accessToken } = {}) {
    if (!chatName && !chatId && !chatUrl) {
      throw new Error("Provide one of: chatUrl, chatId, or chatName.");
    }
    const fromDate = parseDate(from, false);
    const toDate = parseDate(to, true);

    let token;
    if (accessToken) {
      console.log("[teams-log] [1/4] using caller-supplied access token");
      const jwt = decodeJwt(accessToken);
      if (jwt && !GRAPH_AUDIENCES.has(jwt.aud)) {
        console.warn(
          `[teams-log] supplied token audience is "${jwt.aud}", not Graph. Will try anyway.`
        );
      }
      token = accessToken;
    } else {
      console.log("[teams-log] [1/4] locating Graph access token...");
      token = findAccessToken();
    }

    const chat = await resolveChat(token, { chatUrl, chatId, chatName });
    console.log(`[teams-log] [3/4] collecting messages from "${chat.label}"`);
    console.log(`[teams-log]   from: ${fromDate?.toISOString() || "(beginning)"}`);
    console.log(`[teams-log]   to:   ${toDate?.toISOString() || "(now)"}`);

    const messages = await collectMessages(
      token,
      chat.id,
      fromDate,
      toDate,
      ({ page, collected, oldestSeen }) =>
        console.log(
          `[teams-log]   page ${page} | kept ${collected} | oldest in page ${oldestSeen}`
        )
    );

    const payload = {
      chat: { id: chat.id, type: chat.type, label: chat.label },
      range: {
        from: fromDate?.toISOString() || null,
        to: toDate?.toISOString() || null,
      },
      collectedAt: new Date().toISOString(),
      count: messages.length,
      messages,
    };

    const safeLabel = chat.label
      .replace(/[^\w぀-ヿ一-龯-]+/g, "_")
      .slice(0, 60);
    const filename = `teams-log_${safeLabel}_${new Date()
      .toISOString()
      .slice(0, 10)}.json`;
    console.log(
      `[teams-log] [4/4] downloading ${filename} (${messages.length} messages)`
    );
    downloadJSON(payload, filename);
    return payload;
  }

  window.TeamsLogCollector = {
    run,
    findAccessToken,
    diagnose,
    decodeJwt,
    enumerateAccessTokens,
    listChats,
    searchChatsByName,
    collectMessages,
    parseChatIdFromUrl,
    getChatById,
  };

  console.log(
    "%cTeamsLogCollector loaded.%c\n" +
      "Usage (pick one chat selector):\n" +
      '  await TeamsLogCollector.run({\n' +
      '    chatUrl:  "https://teams.microsoft.com/l/message/19%3A...%40thread.v2/...",\n' +
      '    // chatId:  "19:xxxxx@thread.v2",\n' +
      '    // chatName: "Sales team",\n' +
      '    from: "2025-01-01",   // optional\n' +
      '    to:   "2025-12-31",   // optional\n' +
      "  })",
    "color:#10b981;font-weight:bold",
    "color:inherit"
  );
})();
