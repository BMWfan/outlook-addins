# Paperless NGX — Outlook Add-in

Sends attachments from the currently open email directly to your [Paperless-NGX](https://docs.paperless-ngx.com/) document management server.

## Features

- Lists all (non-inline) file attachments of the open email
- Individual checkboxes per file plus a "Select all" toggle
- Optionally prefixes the document title with the email subject
- Per-file upload progress with live status indicators
- Settings (server URL + API token) stored locally in the browser

## Prerequisites

- A running Paperless-NGX instance reachable from your browser
- A Paperless-NGX API token (see below)
- Outlook on the Web or Outlook Desktop (Mailbox API 1.8+)

## Paperless-NGX Setup

### 1. Get your API token

In Paperless-NGX: **Profile icon → Edit Profile → API token**. Copy the token.

### 2. Allow CORS from Outlook Web

Paperless-NGX blocks cross-origin requests by default. Add the following to your Paperless configuration (`.env` or `docker-compose.yml` environment):

```
PAPERLESS_CORS_ALLOWED_HOSTS=https://bmwfan.github.io
```

If you self-host under a different domain, add that domain instead. Restart Paperless after the change.

## Installation

### 1. Host the files

Push to GitHub and enable GitHub Pages, or use Netlify/Vercel. Files must be served over HTTPS.

### 2. Update the manifest

Replace all occurrences of `bmwfan.github.io/outlook-addins` with your own hosting URL in `manifest.xml`.

### 3. Sideload in Outlook Web

1. Open [Outlook Web](https://outlook.office.com)
2. **Settings (gear) → Apps → Manage add-ins**
3. **Add a custom add-in → Add from file**
4. Upload `paperless-ngx/manifest.xml`

## Usage

1. Open any email with attachments
2. Click **Paperless NGX** in the ribbon
3. On first use: click the ⚙ icon and enter your server URL and API token, then **Save**
4. Check the attachments you want to send (or tick **Alle** for all)
5. Optionally enable **E-Mail-Betreff als Dokumenttitel verwenden**
6. Click **→ X Datei(en) an Paperless senden**
7. Each file shows a live status: spinning → ✓ success / ✗ error

Documents land in the Paperless-NGX inbox for tagging and classification.

## Troubleshooting

| Symptom | Likely cause |
|---------|-------------|
| Upload fails with network error | CORS not configured, or server unreachable from your browser |
| HTTP 401 | API token incorrect or expired |
| HTTP 400 | File rejected by Paperless (check server logs) |
| "Mailbox API not supported" message | Outlook version too old; update Outlook |
