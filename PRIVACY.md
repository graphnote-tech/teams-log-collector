# Privacy Policy — Teams Log Collector

*Last updated: 2026-05-12*

## TL;DR

**The Teams Log Collector browser extension runs entirely on your machine and
sends data only to Microsoft endpoints (the same ones the Microsoft Teams web
client itself uses). Nothing is sent to the developer or any third party.**

## What the extension does

When you click **Collect** in the extension popup, the extension:

1. Reads the access tokens that the Microsoft Teams web client has stored in
   the browser's `localStorage` for the open Teams tab.
2. Uses those tokens to call Microsoft endpoints (`graph.microsoft.com`,
   `authsvc.teams.microsoft.com`, `*.ng.msg.teams.microsoft.com`, and
   `api.spaces.skype.com`) to retrieve the chat messages you can already see
   in the Microsoft Teams web client.
3. Bundles the retrieved messages into a JSON file and triggers a download to
   your computer.

The extension never sends data to the extension developer or to any
third-party service.

## Data the extension reads

| Data | Where it lives | What we do with it |
| --- | --- | --- |
| Microsoft Teams access tokens | Your browser's `localStorage` for `teams.microsoft.com` / `teams.cloud.microsoft` | Used in the `Authorization` / `Authentication` header of outgoing requests to Microsoft endpoints during a collection run. Not persisted by the extension. |
| The chat messages you ask for | Microsoft Teams servers | Written to the JSON file you download. |
| Your form inputs (chat URL, dates, filename label) | `chrome.storage.local` | Remembered between popup opens so you don't have to retype them. Never transmitted anywhere. |

## Data we collect from you

**None.** The developer has no servers and no analytics, error-tracking, or
telemetry hooks of any kind in the extension. We do not know who has installed
the extension or how it is being used.

## Permissions and why they're needed

| Permission | Purpose |
| --- | --- |
| `storage` | Remembering form inputs between popup opens. |
| `scripting` | Injecting the content script into the Microsoft Teams tab when needed (covers the case of pre-existing tabs after an extension reload). |
| `host_permissions` for `https://teams.microsoft.com/*`, `https://teams.cloud.microsoft/*`, `https://*.teams.microsoft.com/*` | Reading the Microsoft Teams page where the user is signed in (where access tokens are cached). |
| `host_permissions` for `https://*.ng.msg.teams.microsoft.com/*` | Calling the Microsoft Teams messaging service to retrieve messages. |
| `host_permissions` for `https://api.spaces.skype.com/*` | Issuing the AAD-to-skypeToken exchange used by the Microsoft Teams web client. |
| `host_permissions` for `https://graph.microsoft.com/*` | Alternative path that uses Microsoft Graph when the access token has the `Chat.Read` scope. |

## Your data, your control

- The downloaded JSON file lives only on your computer. You can delete it at
  any time.
- Form values stored in `chrome.storage.local` are cleared if you uninstall
  the extension.
- The extension performs only read operations. It does not write, edit, or
  delete anything in Microsoft Teams.

## Open source

The full source code of this extension is published at
<https://github.com/graphnote-tech/teams-log-collector>. You can audit exactly
what the extension does at any time.

## Contact

If you have a privacy question or concern, please open an issue at
<https://github.com/graphnote-tech/teams-log-collector/issues>.

## Disclaimer

This extension is an independent project. It is not affiliated with, endorsed
by, or sponsored by Microsoft Corporation. "Microsoft", "Microsoft Teams", and
"Microsoft Graph" are trademarks of Microsoft Corporation.
