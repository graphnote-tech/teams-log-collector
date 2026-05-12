# Teams Log Collector

A browser extension (Chrome / Edge / Manifest V3) and a companion console script
that **export Microsoft Teams group-chat or 1:1-chat messages to JSON** using
your existing signed-in session — **no admin rights, no app registration, no
Power Automate license required**.

> ⚠️ This is a personal-use tool. It only reads messages you can already see in
> Teams. It does not bypass any access controls.

## Why

A regular Microsoft 365 user can read their chats in the Teams web app, but
exporting them is surprisingly hard:

- The official Microsoft Graph `Chat.Read` permission requires an Azure AD app
  registration **with admin consent**. Many tenants restrict this.
- Graph Explorer is also blocked when the tenant disallows third-party app
  consent.
- The Teams web client uses its own internal services (`*.ng.msg.teams.microsoft.com`
  and `chatsvcagg.teams.microsoft.com`) with a different auth flow.

This project reuses the tokens that Teams web already holds for you and talks
to the same internal endpoints, so it works in restrictive tenants where the
Graph path is unavailable.

## Features

- Export an entire group-chat or 1:1-chat as JSON
- Optional date-range filter (`From` / `To`)
- Two collection paths, tried in order:
  1. **Microsoft Graph** — used when `Chat.Read` is available on the cached
     Graph token (rare in locked-down tenants, but the cleanest path when it
     works).
  2. **Teams internal API** — exchanges the cached
     `https://api.spaces.skype.com` AAD token for a `skypeToken` and talks to
     `*.ng.msg.teams.microsoft.com` directly, exactly the way the Teams web
     client does.
- Manual `accessToken` field for pasting a Graph Explorer token if you have it
- One-click **diagnose** that lists every cached access token, its audience,
  scopes and remaining lifetime
- Output filename includes the date range:
  `teams-log_<label>_2026-04-01_to_2026-04-30.json`

## Install (developer mode)

1. Clone or download this repository.
2. Open `chrome://extensions` (or `edge://extensions`).
3. Toggle **Developer mode** on.
4. Click **Load unpacked** and select the `extension/` directory.
5. Open <https://teams.microsoft.com> (or <https://teams.cloud.microsoft>) and
   sign in.
6. Click the extension's toolbar icon.

## Usage

1. In Teams, open the chat you want to export. Click `…` on the chat header and
   choose **Copy link to chat** (or copy a link to a specific message).
2. In the extension popup, paste that link into **Chat URL / ID / Name**.
3. (Optional) Set **From** and **To** dates.
4. (Optional) Set **Filename label** — useful because the Teams internal API
   path can't always resolve a human-readable chat name; pasting one here makes
   the output filename readable.
5. Click **Collect**. A `.json` file is downloaded when finished.

Tip: click **Diagnose** first to see which API paths your tenant is going to
allow.

## How it works

```text
            ┌───────── popup.html / popup.js ─────────┐
            │  UI (chat URL, date range, run/diagnose)│
            └──────────────────┬──────────────────────┘
                               │ chrome.tabs.connect (port "tlc")
                               ▼
            ┌──────────── content.js (in Teams tab) ──────────┐
            │ 1. read localStorage MSAL cache                 │
            │ 2. decode each AccessToken JWT, classify by aud │
            │ 3a. Graph path  → /v1.0/chats/{id}/messages     │
            │ 3b. Teams path  → POST authsvc/v1.0/authz       │
            │                   to exchange AAD → skypeToken  │
            │                 → GET {region}.ng.msg.teams     │
            │                       /v1/users/ME/conversations│
            │                       /{id}/messages            │
            │ 4. paginate, filter by date range               │
            │ 5. trigger JSON download via <a download>       │
            └─────────────────────────────────────────────────┘
```

A few non-obvious notes baked into the implementation:

- The Teams messaging service expects the skypeToken in an
  **`Authentication: skypetoken=...`** header — **not** `Authorization`.
- The `authz` response includes `regionGtms.msg`, which can be either the
  direct endpoint (`https://apac.ng.msg.teams.microsoft.com`) or a middleTier
  proxy (`https://teams.cloud.microsoft/api/mt/apac`). The proxy path uses a
  different URL shape, so the extension extracts the region letters and always
  builds the direct URL itself.
- Teams orders messages by `lastModifiedDateTime` (descending). When the
  page's oldest `lastModifiedDateTime` falls below your `From` date, the
  extension stops paginating — safe because `lastModified >= created` always
  holds.

## Limitations

- **Tenant policy:** if your tenant blocks both `Chat.Read` consent and direct
  access to `*.ng.msg.teams.microsoft.com`, neither path will work. There is no
  workaround short of asking your admin for delegated `Chat.Read`.
- **Channel messages are out of scope.** This tool collects *chats*
  (`19:xxx@thread.v2`, `19:xxx@unq.gbl.spaces`). Channel posts use a different
  endpoint (`/teams/{id}/channels/{id}/messages`) and are not supported.
- **Attachments are not downloaded** — only their references appear inside the
  message body. The exported file lists who/when/what; binary attachments live
  in SharePoint / OneDrive and require separate handling.
- **Tokens expire.** Cached tokens are typically valid for ~1 hour (Graph) and
  ~24 hours (skypeToken family). If a long collection fails mid-run, reload the
  Teams tab and try again — the extension will pick up the fresh tokens.

## Privacy & security

- The extension runs entirely on your machine and talks only to
  `teams.microsoft.com`, `teams.cloud.microsoft`, `*.ng.msg.teams.microsoft.com`,
  `api.spaces.skype.com` and `graph.microsoft.com`. No data is sent to any
  third-party server.
- Access tokens are read from your Teams tab's `localStorage` (the same place
  Teams itself stores them) and are only used in the headers of outgoing
  requests during a collection run. They are not persisted by the extension.
- `chrome.storage.local` stores only your form values (chat URL string, dates,
  filename label) so that they're remembered between popup opens.
- Review `extension/content.js` if you want to confirm exactly what is being
  requested.

## Standalone browser-console script

If you cannot install an extension on your machine, `teams-log-collector.js`
in the repo root is a self-contained version you can paste into the DevTools
console on `teams.microsoft.com`. It has the same dual-path logic but a
narrower set of features. See the comment header inside the file for usage.

## Icons

PNG icons (16/32/48/128 px) are checked in alongside the SVG source in
`extension/icons/`. If you want to change the artwork, edit `icon.svg` and
regenerate the PNGs:

```bash
# Uses uv (https://github.com/astral-sh/uv) — no sudo, no global installs.
cd extension/icons
uv run --with cairosvg python3 -c '
import cairosvg
for s in (16, 32, 48, 128):
    cairosvg.svg2png(url="icon.svg", write_to=f"icon{s}.png",
                     output_width=s, output_height=s)
'
```

If you don't have a Python toolchain handy, open
`extension/icons/build-pngs.html` in any browser and click **Generate** — it
rasterises the SVG via Canvas and offers each size as a download.

## Development

The extension is plain JavaScript — no build step, no dependencies.

- `extension/manifest.json` — Manifest V3 declaration
- `extension/popup.{html,css,js}` — popup UI and orchestration
- `extension/content.js` — token discovery, Graph + Teams-internal collection
- `teams-log-collector.js` — standalone console version

To make a change:

1. Edit the file.
2. Bump the `TLC_VERSION` constant in `content.js` **and** the `EXPECTED_VERSION`
   constant in `popup.js`. The popup compares the two via a ping handshake and
   force-reloads the Teams tab on mismatch, so this avoids stale content
   scripts running old logic.
3. Click the reload icon in `chrome://extensions`.

## License

[MIT](./LICENSE)
