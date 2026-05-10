# Newsletter Manager

Outlook Web Add-in that automatically detects newsletters, moves them into a chosen folder, and keeps only the desired number of emails per sender.

> For general Outlook add-in installation instructions, see the [main README](../README.md).

## Features

- **Automatic newsletter detection** on every email you open, based on 7 signals
- **Confirm or reject** with one click — you stay in control
- **Choose a target folder** from your existing folders or create a new one directly in the panel
- **Set a limit** (1–50) — e.g. keep only the last 5 emails from REWE
- **Save rules** — stored locally and automatically suggested the next time you open an email from the same sender
- **Auto-cleanup** — older emails from the same sender in the target folder are moved to the trash

## Usage

1. Open an email in Outlook Web
2. Click **"Newsletter Manager"** in the toolbar
3. The task pane shows the analysis with all detected signals
4. Confirm as newsletter → choose a folder → set a limit → **"Apply rule & move email"**

## Newsletter Detection

An email is classified as a newsletter when **at least 2 of the following 7 signals** match:

| Signal | What is checked |
|--------|----------------|
| List-Unsubscribe Header | `list-unsubscribe` present in the email body |
| Unsubscribe link | Body contains "abmelden", "unsubscribe", or "abbestellen" |
| Bulk sender | Sender address contains newsletter, noreply, no-reply, info@, or marketing |
| "Newsletter" in subject | Subject contains newsletter, angebot, deal, %, or rabatt |
| Promo keywords | Body contains angebot, rabatt, jetzt kaufen, jetzt bestellen, or nur heute |
| No-reply sender | Sender address contains noreply, no-reply, or donotreply |
| HTML-only email | Body longer than 2,000 characters and contains a `<table>` element |

## File structure

```
newsletter-manager/
├── manifest.xml      ← Add-in manifest for Microsoft Exchange
├── taskpane.html     ← Complete app (UI + logic in a single file)
├── icon64.png
├── icon128.png
└── README.md
```

## Technical details

- Communicates with Exchange via **EWS** (`makeEwsRequestAsync`) for folder listing, moving, and deleting emails
- Folders are loaded on startup via EWS `FindFolder`; if EWS is unavailable, a demo fallback is shown
- Deleted emails are moved to **"Deleted Items"** (`DeleteType="MoveToDeletedItems"`), not permanently deleted
- Rules are stored in the browser's **LocalStorage** under the key `nlm_rules` (device-specific, no cloud sync)

## Notes

- Requires the **ReadWriteMailbox** permission to move and delete emails
- Primarily designed for Outlook Web; EWS functionality may be limited in Outlook Desktop

## Ideas for future improvements

- Claude AI API integration for smarter newsletter detection
- Export / import rules as JSON
- Automatically process the entire inbox on load
- Statistics: how many emails have been cleaned up?
