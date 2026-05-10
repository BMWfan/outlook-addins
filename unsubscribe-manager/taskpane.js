/* ============================================================
   Unsubscribe Manager – taskpane.js
   Hosted at: https://bmwfan.github.io/unsubscribe-manager/
   ============================================================ */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById('btn-scan').addEventListener('click', startScan);
    document.getElementById('btn-clear').addEventListener('click', clearResults);
  }
});

// ── State ──────────────────────────────────────────────────
let senderMap = {}; // domain/company → { name, email, count, unsubLink, unsubscribed }
let scanning  = false;

// ── UI Helpers ─────────────────────────────────────────────
function setStatus(msg, mode = '') {
  const el  = document.getElementById('status');
  const txt = document.getElementById('status-text');
  el.className = mode;
  txt.textContent = msg;
}

function setProgress(pct) {
  const bar  = document.getElementById('progress-bar');
  const fill = document.getElementById('progress-fill');
  bar.style.display = pct >= 0 && pct < 100 ? 'block' : 'none';
  fill.style.width  = pct + '%';
}

function setSummary(count) {
  document.getElementById('summary-meta').textContent =
    count > 0 ? `${count} Sender` : '';
}

// ── Main scan flow ─────────────────────────────────────────
async function startScan() {
  if (scanning) return;
  scanning = true;

  const maxCount = parseInt(document.getElementById('scan-input').value, 10) || 100;
  document.getElementById('btn-scan').disabled = true;
  senderMap = {};
  renderList();
  setStatus('Postfach wird gelesen…', 'scanning');
  setProgress(0);

  try {
    const messages = await fetchMessages(maxCount);
    setStatus(`${messages.length} Mails geladen, analysiere…`, 'scanning');
    setProgress(30);
    processSenders(messages);
    setProgress(70);
    await enrichWithAI(messages);
    setProgress(100);
    renderList();
    setStatus(`Fertig. ${Object.keys(senderMap).length} Sender gefunden.`, 'done');
    setSummary(Object.keys(senderMap).length);
  } catch (err) {
    setStatus('Fehler: ' + err.message, 'error');
    console.error(err);
  } finally {
    scanning = false;
    document.getElementById('btn-scan').disabled = false;
    setProgress(100);
    document.getElementById('progress-bar').style.display = 'none';
  }
}

// ── Fetch messages via EWS REST ────────────────────────────
function fetchMessages(count) {
  return new Promise((resolve, reject) => {
    // Use Office.context.mailbox.makeEwsRequestAsync to get items
    const ewsRequest = buildFindItemRequest(count);
    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Fallback: try REST
        fetchMessagesRest(count).then(resolve).catch(reject);
        return;
      }
      try {
        const parsed = parseEwsResponse(result.value);
        resolve(parsed);
      } catch(e) {
        fetchMessagesRest(count).then(resolve).catch(reject);
      }
    });
  });
}

function buildFindItemRequest(count) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="message:From"/>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="message:InternetMessageHeaders"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${count}" Offset="0" BasePoint="Beginning"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;
}

function parseEwsResponse(xmlString) {
  const parser = new DOMParser();
  const xml    = parser.parseFromString(xmlString, 'text/xml');
  const items  = xml.querySelectorAll('Message');
  const result = [];

  items.forEach(item => {
    const fromName  = item.querySelector('Name')?.textContent || '';
    const fromEmail = item.querySelector('EmailAddress')?.textContent || '';
    const subject   = item.querySelector('Subject')?.textContent || '';
    const headers   = item.querySelectorAll('InternetMessageHeader');

    let unsubLink = '';
    headers.forEach(h => {
      if (h.getAttribute('HeaderName')?.toLowerCase() === 'list-unsubscribe') {
        const val = h.textContent;
        // prefer https link
        const https = val.match(/<(https?:\/\/[^>]+)>/)?.[1];
        const mailto = val.match(/<(mailto:[^>]+)>/)?.[1];
        unsubLink = https || mailto || '';
      }
    });

    if (fromEmail) {
      result.push({ fromName, fromEmail, subject, unsubLink });
    }
  });

  return result;
}

// Fallback: Office REST (older hosts)
async function fetchMessagesRest(count) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async tokenResult => {
      if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error('Token nicht verfügbar'));
        return;
      }
      const token    = tokenResult.value;
      const endpoint = Office.context.mailbox.restUrl;
      const url      = `${endpoint}/v2.0/me/mailFolders/inbox/messages` +
                       `?$top=${count}&$select=from,subject,internetMessageHeaders&$orderby=receivedDateTime desc`;
      try {
        const resp = await fetch(url, {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!resp.ok) throw new Error(`REST ${resp.status}`);
        const data = await resp.json();
        const msgs = (data.value || []).map(m => {
          const headers = m.internetMessageHeaders || [];
          const unsubH  = headers.find(h => h.name.toLowerCase() === 'list-unsubscribe');
          let unsubLink = '';
          if (unsubH) {
            const https  = unsubH.value.match(/<(https?:\/\/[^>]+)>/)?.[1];
            const mailto = unsubH.value.match(/<(mailto:[^>]+)>/)?.[1];
            unsubLink    = https || mailto || '';
          }
          return {
            fromName:  m.from?.emailAddress?.name  || '',
            fromEmail: m.from?.emailAddress?.address || '',
            subject:   m.subject || '',
            unsubLink
          };
        });
        resolve(msgs);
      } catch(e) { reject(e); }
    });
  });
}

// ── Process senders (group by domain) ─────────────────────
function processSenders(messages) {
  messages.forEach(msg => {
    const email  = msg.fromEmail.toLowerCase();
    const domain = extractDomain(email);
    if (!domain) return;

    // skip own domain / exchange system mails
    if (isSystemSender(email)) return;

    if (!senderMap[domain]) {
      senderMap[domain] = {
        key:         domain,
        displayName: msg.fromName || domain,
        email:       email,
        domain:      domain,
        count:       0,
        unsubLink:   '',
        unsubscribed: false,
        subjects:    []
      };
    }
    senderMap[domain].count++;
    if (msg.unsubLink && !senderMap[domain].unsubLink) {
      senderMap[domain].unsubLink = msg.unsubLink;
    }
    if (senderMap[domain].subjects.length < 3) {
      senderMap[domain].subjects.push(msg.subject);
    }
    // prefer longer/cleaner display names
    if (msg.fromName && msg.fromName.length > senderMap[domain].displayName.length) {
      senderMap[domain].displayName = msg.fromName;
    }
  });
}

function extractDomain(email) {
  const match = email.match(/@([\w.-]+)/);
  return match ? match[1] : null;
}

function isSystemSender(email) {
  const skip = ['noreply@microsoft', 'mailer-daemon', 'postmaster', 'bounce'];
  return skip.some(s => email.includes(s));
}

// ── AI enrichment via Anthropic API ───────────────────────
async function enrichWithAI(messages) {
  // Build a concise summary for Claude to label senders nicely
  const raw = Object.values(senderMap).map(s => ({
    domain:      s.domain,
    sample_name: s.displayName,
    subjects:    s.subjects.slice(0, 2)
  }));

  if (raw.length === 0) return;

  const prompt = `Du erhältst eine Liste von E-Mail-Absendern aus einem Postfach.
Deine Aufgabe: Gib für jeden Eintrag einen sauberen, kurzen deutschen Anzeigenamen zurück (max 30 Zeichen).
Nutze den sample_name oder leite ihn aus der Domain ab. Keine technischen Kürzel, echte Firmennamen bevorzugen.

Antworte NUR als JSON-Array (kein Markdown, keine Erklärung):
[{"domain":"example.com","label":"Beispiel GmbH"}, ...]

Eingabe:
${JSON.stringify(raw)}`;

  try {
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model:      'claude-sonnet-4-20250514',
        max_tokens: 1000,
        messages:   [{ role: 'user', content: prompt }]
      })
    });

    if (!resp.ok) return; // fail silently, display names already set
    const data = await resp.json();
    const text = data.content?.find(b => b.type === 'text')?.text || '';
    const labels = JSON.parse(text.replace(/```json|```/g, '').trim());

    labels.forEach(item => {
      if (senderMap[item.domain]) {
        senderMap[item.domain].displayName = item.label || senderMap[item.domain].displayName;
      }
    });
  } catch(_) {
    // AI enrichment optional – ignore errors
  }
}

// ── Render list ────────────────────────────────────────────
function renderList() {
  const container = document.getElementById('list-container');
  const entries   = Object.values(senderMap);

  if (entries.length === 0) {
    container.innerHTML = `<div id="empty">
      <div class="icon">📭</div>
      <div>Noch keine Analyse — klick auf <strong>Scannen</strong></div>
    </div>`;
    setSummary(0);
    return;
  }

  // Sort by count desc
  entries.sort((a, b) => b.count - a.count);

  container.innerHTML = entries.map(s => `
    <div class="sender-row" id="row-${CSS.escape(s.domain)}">
      <div class="sender-info">
        <div class="sender-name" title="${esc(s.displayName)}">${esc(s.displayName)}</div>
        <div class="sender-detail" title="${esc(s.domain)}">${esc(s.domain)}</div>
      </div>
      <span class="badge">${s.count}×</span>
      ${renderUnsubButton(s)}
    </div>
  `).join('');

  // Attach button listeners
  entries.forEach(s => {
    const btn = document.querySelector(`#row-${CSS.escape(s.domain)} .btn-unsub`);
    if (btn && !s.unsubscribed) {
      btn.addEventListener('click', () => handleUnsub(s));
    }
  });
}

function renderUnsubButton(s) {
  if (s.unsubscribed) {
    return `<button class="btn-unsub done-unsub" disabled>✓ Abbestellt</button>`;
  }
  if (s.unsubLink) {
    return `<button class="btn-unsub" data-domain="${esc(s.domain)}">Abmelden</button>`;
  }
  return `<button class="btn-unsub" style="opacity:0.35;cursor:default" disabled title="Kein Unsubscribe-Link gefunden">–</button>`;
}

function handleUnsub(sender) {
  if (sender.unsubLink.startsWith('mailto:')) {
    // Open compose window for mailto unsubscribe
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [{ emailAddress: sender.unsubLink.replace('mailto:', '').split('?')[0] }],
      subject:      'Unsubscribe',
      htmlBody:     'Please unsubscribe me from your mailing list.'
    });
  } else {
    // Open link in browser
    Office.context.ui.openBrowserWindow(sender.unsubLink);
  }
  // Mark as done
  senderMap[sender.domain].unsubscribed = true;
  renderList();
}

function clearResults() {
  senderMap = {};
  renderList();
  setStatus('Bereit. Anzahl wählen und scannen.', '');
  setSummary(0);
}

function esc(str) {
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}