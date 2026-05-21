# Newsletter Manager

Outlook Add-in that automatically detects newsletters, moves them into a chosen folder, and creates server-side Exchange rules so future emails are processed even when Outlook is closed.

> For installation instructions, see the [main README](../README.md).

## Features

- **Automatic newsletter detection** on every email you open, based on email headers, content signals, and a trained Naive Bayes text classifier
- **Confirm or reject** with one click — corrections are remembered per sender
- **Choose a target folder** from your existing folders or create a new one directly in the panel
- **Exchange server-side rule** created automatically — future newsletters from the same address are moved without the add-in being open
- **Bulk move** — all existing emails from that sender are moved at once; personal emails from the same address are protected via header checks
- **Re-apply all rules** — processes newly arrived newsletters across all saved senders in one click and updates existing Exchange rules to the latest format
- **Folder Scanner** — scans the current folder for unknown newsletter senders not yet covered by a rule, grouped by sender with email count
- **Learned senders** — stores your confirm/deny decisions and applies them automatically on the next open

## How it works

### Newsletter detection

An email is classified as a newsletter when it matches any of these signals:

| Signal | What is checked |
|--------|-----------------|
| `List-Unsubscribe` header | Present in the email's internet message headers |
| `List-Id` / `List-Post` / other list headers | Any RFC 2369 mailing list header present |
| `Precedence: bulk` / `list` | Present in the email's internet message headers |
| Subject keywords | newsletter, angebot, deal, rabatt, %, sale, promo, … |
| Body keywords | unsubscribe, abmelden, opt-out, newsletter, … |
| Naive Bayes classifier | Text-based ML trained on your own confirm/deny feedback (active after 5 examples per class) |

### Exchange server-side rule

When you apply a rule, the add-in creates an inbox rule via Graph API with these conditions:

```
senderContains: [sender address]
headerContains: ['List-Unsubscribe', 'List-Id']
```

Using `headerContains` (rather than body keywords) ensures that personal emails from the same sender address are never moved — only emails sent via bulk-mailing infrastructure (Mailchimp, SendGrid, etc.) carry these headers.

### Folder Scanner

Click **↺ Scan** to scan the folder of your currently open email. The scanner fetches up to 500 messages, runs the same header-based detection, and shows each unknown newsletter sender once with the number of matching emails. From there you can apply a rule directly inline.

## Usage

1. Open an email in Outlook
2. Click **"Newsletter Manager"** in the toolbar
3. The task pane analyses the email and shows detected signals
4. Confirm as newsletter → choose a target folder → click **"Regel anwenden & E-Mail verschieben"**
5. Use **↺ Alle anwenden** periodically to process newly arrived newsletters across all rules

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

- Communicates with Microsoft 365 via the **Graph API** (`https://graph.microsoft.com/v1.0`)
- Authentication via **OAuth 2.0 / MSAL** using an Office dialog-based auth flow
- Exchange server-side rules are created/updated via `PATCH /me/mailFolders/inbox/messageRules/{id}`
- The current folder is resolved by converting the Office.js EWS item ID to a Graph REST ID via `convertToRestId`
- Rules and learned sender overrides are stored in **LocalStorage** (device-specific, no cloud sync)
- The Naive Bayes classifier state (vocabulary, class counts) is also stored in LocalStorage

## Required permissions

| Permission | Why |
|------------|-----|
| `Mail.ReadWrite` | Read and move emails |
| `MailboxSettings.ReadWrite` | Create and update inbox rules |
