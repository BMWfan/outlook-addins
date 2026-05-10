# Unsubscribe Manager

Outlook Web Add-in that scans your inbox and displays an overview of all newsletter senders — grouped by company domain, unsubscribable with one click.

> For general Outlook add-in installation instructions, see the [main README](../README.md).

## Features

- **Scan** up to 500 emails from your inbox (configurable, default: 100)
- **Groups senders by domain** — instead of seeing "Amazon" 20 times, you get one entry with a `20×` badge
- **Detects `List-Unsubscribe` headers** automatically
- **One-click unsubscribe** — opens the unsubscribe link in the browser, or opens a pre-filled compose window for `mailto:` links
- **AI-powered display names** via Claude Sonnet — makes domain names readable (e.g. `rewe.de` → `REWE`)
- **Compact UI** — fits within Outlook's 250 px task pane height limit

## Usage

1. Open any email in Outlook Web (the add-in works from any open email)
2. Click **"Unsubscribe Manager"** in the toolbar
3. Set the number of emails to scan (10–500) in the input field
4. Click **"Scan"** — the add-in scans your inbox and groups all newsletter senders
5. Click **"Abmelden"** next to any sender to unsubscribe

The list is sorted by email count (most frequent senders first). Once you click unsubscribe, the button changes to **"✓ Abbestellt"** for that session.

## AI display names (optional)

The add-in calls the **Anthropic API** to turn raw domain names into readable company names. This requires adding your API key to `taskpane.js`:

```js
headers: {
  'Content-Type': 'application/json',
  'x-api-key': 'sk-ant-...',          // ← add this line
  'anthropic-version': '2023-06-01'    // ← add this line
}
```

Without the API key the request silently fails and the raw sender name or domain is shown instead — all other features still work normally.

## File structure

```
unsubscribe-manager/
├── manifest.xml      ← Add-in manifest for Microsoft Exchange
├── taskpane.html     ← UI and layout
├── taskpane.js       ← Logic: mailbox scan, grouping, AI enrichment
├── icon64.png
├── icon128.png
└── README.md
```

## Technical details

- Reads emails via **EWS** (`makeEwsRequestAsync`) with a fallback to the **Office REST API** for older Outlook hosts
- Extracts `List-Unsubscribe` headers; prefers HTTPS links over `mailto:` links
- `mailto:` unsubscribe links open an Outlook compose window with subject "Unsubscribe" pre-filled
- HTTPS unsubscribe links open in the default browser via `Office.context.ui.openBrowserWindow`
- System senders (mailer-daemon, postmaster, noreply@microsoft, bounce) are automatically excluded
- The add-in is hosted at `https://bmwfan.github.io/unsubscribe-manager/` — the `manifest.xml` already points to this URL

## Notes

- Requires the **ReadWriteMailbox** permission
- The 250 px height is an Outlook constraint for task panes in read mode; the UI is specifically designed for this limit
- Unsubscribed state is in-memory only and resets when the task pane is closed
