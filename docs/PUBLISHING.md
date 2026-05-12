# Publishing to Chrome Web Store & Edge Add-ons

Step-by-step checklist for submitting **Teams Log Collector** to the two
extension stores. Treat the texts in this doc as ready-to-paste defaults; edit
to taste.

## Pre-flight

- [ ] Bump `version` in `extension/manifest.json` (it must be **higher than the
      version currently in the store** every time you submit an update).
- [ ] Bump `TLC_VERSION` in `extension/content.js` and `EXPECTED_VERSION` in
      `extension/popup.js` to match — same numeric prefix is fine, the suffix
      is for internal cache-busting.
- [ ] Smoke-test by loading the unpacked `extension/` directory in
      `chrome://extensions`.
- [ ] Make sure `PRIVACY.md` is reachable on a public URL. Easiest option:
      enable **GitHub Pages** for this repository (Settings → Pages → Source:
      `main` branch, `/` root). The privacy policy URL will then be
      `https://graphnote-tech.github.io/teams-log-collector/PRIVACY`.

## Build the upload bundle

```bash
cd extension
zip -r ../teams-log-collector-extension.zip . \
  -x "icons/build-pngs.html" \
  -x "icons/icon.svg"
```

Verify the zip contains only `manifest.json`, `popup.{html,css,js}`,
`content.js`, and the four PNG icons.

## Chrome Web Store

1. Go to <https://chrome.google.com/webstore/devconsole/> and sign in with the
   Google account that owns the developer profile. (One-time **\$5** signup
   fee.)
2. Click **Add new item** → upload `teams-log-collector-extension.zip`.
3. Fill the listing using the texts below.
4. Submit for review. Reviews typically take a few business days but can take
   longer for extensions that touch broad host permissions.

### Listing texts

**Name** (max 75 chars)

> Teams Log Collector

**Summary / Short description** (max 132 chars)

> Export Microsoft Teams group-chat and 1:1-chat messages to JSON. Works as a
> regular user; no admin rights needed.

**Detailed description**

> Teams Log Collector exports the chats you can already see in the Microsoft
> Teams web client to a JSON file you can keep, search, or archive.
>
> ✦ Works as a regular Microsoft 365 user — no admin rights, no Azure app
>   registration, no Power Automate license.
> ✦ Two collection paths, chosen automatically:
>     • Microsoft Graph when the cached Graph token has the Chat.Read scope.
>     • The Teams internal messaging API otherwise — the same one the Teams
>       web client itself uses.
> ✦ Date-range filter (From / To).
> ✦ Diagnose button shows exactly which tokens are cached, which audiences and
>   scopes they bear, and which collection path will be used.
> ✦ Output filename includes the date range, e.g.
>   teams-log_my-chat_2026-04-01_to_2026-04-30.json
>
> Everything runs locally in your browser. The extension talks only to
> Microsoft endpoints (graph.microsoft.com, authsvc.teams.microsoft.com,
> *.ng.msg.teams.microsoft.com, api.spaces.skype.com). No data is sent to the
> developer or any third party.
>
> Full source code: https://github.com/graphnote-tech/teams-log-collector
> Privacy policy:  https://graphnote-tech.github.io/teams-log-collector/PRIVACY
>
> Disclaimer: This extension is an independent project, not affiliated with or
> endorsed by Microsoft Corporation. "Microsoft Teams" is a trademark of
> Microsoft Corporation.

**Category**: Productivity

**Language**: English

### Privacy practices

**Single purpose** (free text)

> Export Microsoft Teams group-chat and 1:1-chat message history to a JSON
> file from the signed-in user's own session.

**Permission justifications**

- `storage` — Remember the form values (chat URL, date range, filename label)
  between popup opens.
- `scripting` — Inject the collection logic into the Microsoft Teams tab on
  demand, including tabs that pre-date an extension reload.
- Host permission `https://teams.microsoft.com/*` and
  `https://teams.cloud.microsoft/*` — The Microsoft Teams web client where the
  user is signed in and from whose page the cached access tokens are read.
- Host permission `https://*.teams.microsoft.com/*` — Includes
  `authsvc.teams.microsoft.com`, which exchanges the cached AAD token for a
  Skype token used by Teams' messaging service.
- Host permission `https://*.ng.msg.teams.microsoft.com/*` — Microsoft Teams'
  messaging service that returns the actual chat messages.
- Host permission `https://api.spaces.skype.com/*` — Auxiliary endpoint used
  by the Teams web client's auth flow.
- Host permission `https://graph.microsoft.com/*` — Alternate collection path
  when the cached Graph token bears the Chat.Read scope.

**Remote code** — None. All JavaScript executed by the extension is bundled in
the package.

**Data handling**

- The extension does **not** collect, transmit, sell, or share personal or
  sensitive user data.
- All requests go directly to Microsoft endpoints listed above.
- Form values are kept in `chrome.storage.local` and never leave the user's
  device.

**Privacy policy URL**

> https://graphnote-tech.github.io/teams-log-collector/PRIVACY

### Screenshots & promotional images

Required by Chrome Web Store:

- **Screenshot(s)** — at least 1, recommended 3–5. Size **1280×800** or
  **640×400**, PNG/JPEG. Suggested shots:
  1. Popup with a chat URL pasted and a date range set.
  2. Status pane showing progress lines during a collection.
  3. Diagnose output listing cached tokens.
- **Small promo tile** — **440×280**, required. The SVG source in
  `assets/promo-tile.svg` rasterises to this; regenerate with the same
  cairosvg one-liner used for the icons.

## Microsoft Edge Add-ons

1. Sign in at <https://partner.microsoft.com/dashboard/microsoftedge/>. (Edge
   add-on submission is **free**.)
2. Upload the same `teams-log-collector-extension.zip`.
3. Reuse the listing texts above. Edge's privacy answers are similar; the same
   privacy policy URL works.

## Risks to know before submitting

- **Microsoft trademark** — "Teams" is a Microsoft trademark. The name and
  description make the non-affiliation explicit; this typically passes review
  but is at the reviewer's discretion.
- **Internal API usage** — `*.ng.msg.teams.microsoft.com` is the same endpoint
  Microsoft Teams' own web client uses but is not part of the public Graph
  surface. Chrome Web Store does not generally flag this, but Microsoft could
  in principle restrict access at any time, in which case the Graph code path
  remains available for users whose tenants grant Chat.Read.
- **Broad host permissions** — Listed and individually justified above.
  Vague justifications are a common rejection cause; keep them specific.

## After approval

- Update README's installation section to mention the public Chrome Web Store
  / Edge Add-ons URLs in addition to "Load unpacked".
- Tag the released commit (`gh release create vX.Y.Z teams-log-collector-extension.zip`).
