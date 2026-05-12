// Popup UI for the Teams Log Collector extension.
//
// Responsibilities:
//   • read the form, find the active Teams tab and forward the request to the
//     content script via a long-lived port (so progress streams back),
//   • on version mismatch, reload the Teams tab so a fresh content.js is loaded
//     (a stale, previously-injected content script silently handling messages
//     was the single most painful bug during development).

const $ = (id) => document.getElementById(id);
const statusEl = $("status");

// Must match TLC_VERSION in content.js — used by the ping handshake to detect
// a stale content script that survived an extension reload.
const EXPECTED_VERSION = "0.0.1";

function clearStatus() {
  statusEl.textContent = "";
}

function log(line) {
  statusEl.textContent += (statusEl.textContent ? "\n" : "") + line;
  statusEl.scrollTop = statusEl.scrollHeight;
}

async function getTeamsTab() {
  const tabs = await chrome.tabs.query({
    url: ["https://teams.microsoft.com/*", "https://teams.cloud.microsoft/*"],
  });
  if (!tabs.length) {
    throw new Error(
      "No Microsoft Teams tab found. Open https://teams.microsoft.com, sign in, and try again."
    );
  }
  // Prefer the active one when multiple Teams tabs are open.
  return tabs.find((t) => t.active) || tabs[0];
}

/**
 * Ask the content script (if any) what version it is.
 * Resolves on pong; rejects on timeout / disconnect (i.e. no listener).
 */
function pingContent(tabId, timeoutMs = 400) {
  return new Promise((resolve, reject) => {
    let port;
    try {
      port = chrome.tabs.connect(tabId, { name: "tlc" });
    } catch (e) {
      reject(e);
      return;
    }
    const timer = setTimeout(() => {
      try { port.disconnect(); } catch {}
      reject(new Error("ping timeout"));
    }, timeoutMs);
    port.onMessage.addListener((msg) => {
      if (msg.type === "done" && msg.pong) {
        clearTimeout(timer);
        port.disconnect();
        resolve({ version: msg.version || "(unversioned)" });
      }
    });
    port.onDisconnect.addListener(() => {
      clearTimeout(timer);
      reject(new Error(chrome.runtime.lastError?.message || "disconnected"));
    });
    port.postMessage({ type: "ping" });
  });
}

/**
 * Make sure the Teams tab is running our latest content.js.
 *   • no listener        → inject content.js
 *   • listener, same ver → reuse
 *   • listener, diff ver → reload the tab, then inject
 *
 * We reload (instead of just re-injecting) because chrome.scripting cannot
 * evict already-registered onConnect listeners from old content scripts that
 * predate the listener-tracking logic.
 */
async function ensureInjected(tabId) {
  let pingResult = null;
  try {
    pingResult = await pingContent(tabId);
  } catch {
    // Not alive — fall through and inject.
  }
  if (pingResult && pingResult.version === EXPECTED_VERSION) {
    return;
  }
  if (pingResult && pingResult.version !== EXPECTED_VERSION) {
    log(
      `Stale content.js detected (got "${pingResult.version}", expected "${EXPECTED_VERSION}"). Reloading the Teams tab...`
    );
    await chrome.tabs.reload(tabId);
    await waitForTabComplete(tabId);
  }
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content.js"],
  });
  // Give the listener a moment to register.
  await new Promise((r) => setTimeout(r, 150));
}

function waitForTabComplete(tabId, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const deadline = Date.now() + timeoutMs;
    const tick = async () => {
      try {
        const tab = await chrome.tabs.get(tabId);
        if (tab.status === "complete") return resolve();
      } catch (e) {
        return reject(e);
      }
      if (Date.now() > deadline) return reject(new Error("tab reload timeout"));
      setTimeout(tick, 200);
    };
    tick();
  });
}

/**
 * Open a port, send `message`, stream `progress` events to the status pane,
 * and resolve with the final `done` payload.
 */
function openPort(tabId, message) {
  return new Promise((resolve, reject) => {
    let port;
    try {
      port = chrome.tabs.connect(tabId, { name: "tlc" });
    } catch (e) {
      reject(e);
      return;
    }
    let settled = false;
    port.onMessage.addListener((msg) => {
      if (msg.type === "progress") {
        log(msg.text);
      } else if (msg.type === "done") {
        settled = true;
        resolve(msg);
        port.disconnect();
      } else if (msg.type === "error") {
        settled = true;
        reject(new Error(msg.error));
        port.disconnect();
      }
    });
    port.onDisconnect.addListener(() => {
      if (!settled) {
        const err = chrome.runtime.lastError?.message;
        reject(
          new Error(
            err ||
              "Content script disconnected. Reload the Teams tab and the extension, then try again."
          )
        );
      }
    });
    port.postMessage(message);
  });
}

// ----- Button handlers -----

$("run").addEventListener("click", async () => {
  const chat = $("chat").value.trim();
  if (!chat) {
    clearStatus();
    log("Enter a chat URL, ID, or name.");
    return;
  }
  const params = {
    chat,
    from: $("from").value,
    to: $("to").value,
    nameOverride: $("nameOverride").value.trim(),
    accessToken: $("token").value.trim(),
    useTeamsApi: $("useTeamsApi").checked,
  };
  clearStatus();
  $("run").disabled = true;
  try {
    const tab = await getTeamsTab();
    log(`Teams tab: ${tab.url}`);
    await ensureInjected(tab.id);
    const result = await openPort(tab.id, { type: "collect", params });
    const s = result.summary;
    log(`✓ Done. ${s.count} messages → ${s.filename} (via ${s.via})`);
  } catch (e) {
    log(`✗ Error: ${e.message}`);
  } finally {
    $("run").disabled = false;
  }
});

$("diagnose").addEventListener("click", async () => {
  clearStatus();
  $("diagnose").disabled = true;
  try {
    const tab = await getTeamsTab();
    await ensureInjected(tab.id);
    const result = await openPort(tab.id, { type: "diagnose" });
    const tokens = result.diagnose || [];
    log(`Found ${tokens.length} cached access tokens in the Teams tab.`);
    log("");
    for (const t of tokens) {
      const tag = t.hasChatRead ? "[Chat.Read ✓]" : "             ";
      const expMin = Math.round(t.expiresInSec / 60);
      log(`${tag} aud=${t.aud}`);
      log(`              expires in ${expMin}min  scopes: ${t.scopes.slice(0, 120)}`);
    }
    const graph = tokens.find(
      (t) =>
        t.aud === "https://graph.microsoft.com" ||
        t.aud === "00000003-0000-0000-c000-000000000000"
    );
    log("");
    if (!graph) {
      log("⚠ No Microsoft Graph token is cached. The extension will use the Teams internal API.");
    } else if (!graph.hasChatRead) {
      log("⚠ A Graph token is cached but it lacks the Chat.Read scope.");
      log("  Options:");
      log("    • Paste a Graph Explorer token into Advanced → Manual access token");
      log("    • Or enable Advanced → Use Teams internal API");
    } else {
      log("✓ Graph + Chat.Read available. Collect should work as-is.");
    }
  } catch (e) {
    log(`✗ Error: ${e.message}`);
  } finally {
    $("diagnose").disabled = false;
  }
});

$("clear").addEventListener("click", () => clearStatus());

// ----- Persistence: remember form values between popup opens -----

const FIELDS = ["chat", "from", "to", "nameOverride"];
chrome.storage.local.get(FIELDS, (v) => {
  for (const f of FIELDS) if (v[f]) $(f).value = v[f];
});
for (const f of FIELDS) {
  $(f).addEventListener("input", (e) =>
    chrome.storage.local.set({ [f]: e.target.value })
  );
}
