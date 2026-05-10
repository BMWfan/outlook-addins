# Outlook Web Add-ins

A collection of self-hosted Outlook Web Add-ins for personal mailbox management.

## What are Outlook Web Add-ins?

Outlook Web Add-ins are small web applications that run embedded directly inside Outlook (Web, Desktop, Mobile). They consist of a `manifest.xml` (which describes the add-in to Microsoft) and a regular web page (`taskpane.html`) that you host yourself. Microsoft Exchange loads the add-in from your URL and renders it inside Outlook.

## Add-ins in this repository

| Add-in | Description |
|--------|-------------|
| [newsletter-manager](newsletter-manager/) | Automatically detects newsletters, moves them to a folder, and limits the number of emails kept per sender |
| [unsubscribe-manager](unsubscribe-manager/) | Scans your inbox, groups newsletter senders by domain, and lets you unsubscribe with one click |

## General Installation

### 1. Host your files

Each add-in must be reachable via a public **HTTPS URL**. Free options:

- **GitHub Pages** – Push the repository to GitHub, then enable Pages under `Settings → Pages → Branch: main / Root`. Files will be available at `https://<username>.github.io/<repo>/`.
- **Netlify Drop** – Drag the folder onto [app.netlify.com/drop](https://app.netlify.com/drop) and you're done.
- **Vercel** – Run `vercel deploy` inside the project folder.

### 2. Update manifest.xml

Replace the placeholder `https://YOUR-URL` in the add-in's `manifest.xml` with your actual hosting URL:

```xml
<SourceLocation DefaultValue="https://my-addin.netlify.app/taskpane.html"/>
```

### 3. Install the add-in in Outlook

**Via Outlook Web (easiest):**

1. Open [Outlook Web](https://outlook.office.com)
2. Click the **gear icon** (Settings) in the top right
3. Search for **"Manage add-ins"** or navigate to `Settings → Apps → Manage add-ins`
4. Click **"Add a custom add-in"** → **"Add from file"**
5. Upload the `manifest.xml` of the desired add-in and confirm

**Via PowerShell (Exchange Online):**

```powershell
New-App -Mailbox "your@email.com" -Url "https://your-url/manifest.xml"
```

**For organization-wide deployment:**

Add-ins can be rolled out to all users in an organization via the Exchange Admin Center (`admin.exchange.microsoft.com`).

### 4. Using the add-in

1. Open an email in Outlook Web or Desktop
2. Click the add-in name in the toolbar
3. The task pane (sidebar) opens

## Notes

- Add-ins require the **ReadWriteMailbox** permission depending on their features (e.g. moving or deleting emails)
- Locally stored data (rules, settings) lives in the **browser's LocalStorage** and is device-specific
- For regular use, GitHub Pages is recommended since code changes are published automatically
