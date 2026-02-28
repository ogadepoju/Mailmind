// MailMind — Outlook Taskpane v4
// Uses Office.js makeEwsRequestAsync to send — works on personal accounts

const BACKEND_URL = 'https://mailmind-6f1d.onrender.com';

let currentEmail = null;
let isDrafting = false;

// ─── OFFICE INIT ─────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailAndDraft();
    wireButtons();
  }
});

// ─── LOAD EMAIL + AUTO DRAFT ─────────────────────────────────────────────────

function loadEmailAndDraft() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showState('idle');
    showError('No email open', 'Open an email to get started.');
    return;
  }

  const from = item.from?.emailAddress || '';
  const fromName = item.from?.displayName || from;
  const subject = item.subject || 'No Subject';

  document.getElementById('meta-from').textContent = fromName || '—';
  document.getElementById('meta-subject').textContent = subject || '—';

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    const body = result.status === Office.AsyncResultStatus.Succeeded
      ? result.value.trim() : '';

    currentEmail = {
      id: item.itemId || '',
      conversationId: item.conversationId || '',
      subject,
      from,
      fromName,
      body: body.slice(0, 2000)
    };

    requestDraft();
  });
}

// ─── REQUEST DRAFT ────────────────────────────────────────────────────────────

async function requestDraft() {
  if (isDrafting || !currentEmail) return;
  isDrafting = true;

  setStatus('drafting', 'Drafting reply with Claude AI…');
  showState('idle');
  hideError();

  const manualBtn = document.getElementById('btn-manual-draft');
  if (manualBtn) manualBtn.disabled = true;

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);

    const res = await fetch(`${BACKEND_URL}/draft`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(currentEmail),
      signal: controller.signal
    });

    clearTimeout(timeout);

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Server error ${res.status}`);
    }

    const data = await res.json();
    if (!data.draft) throw new Error('Empty draft returned');

    document.getElementById('draft-textarea').value = data.draft;
    setStatus('ready', 'Draft ready — review and send');
    showState('draft');

  } catch (err) {
    const isTimeout = err.name === 'AbortError';
    const message = isTimeout ? 'Request timed out' : err.message;
    const hint = isTimeout ? 'Try regenerating.' : `Backend: ${BACKEND_URL}`;
    setStatus('error', message);
    showError(message, hint);
    showState('idle');
  } finally {
    isDrafting = false;
    if (manualBtn) manualBtn.disabled = false;
  }
}

// ─── SEND VIA EWS ────────────────────────────────────────────────────────────
// makeEwsRequestAsync works on personal outlook.com accounts
// No token or Graph API needed

async function sendReply() {
  const draftBody = document.getElementById('draft-textarea').value.trim();
  if (!draftBody) return;

  const btn = document.getElementById('btn-send');
  btn.textContent = 'Sending…';
  btn.disabled = true;

  try {
    await sendViaEWS(draftBody);
    showState('sent');
    setStatus('idle', '');
  } catch (err) {
    console.error('[MailMind] Send error:', err);
    btn.textContent = '✓ Send Reply';
    btn.disabled = false;
    setStatus('error', 'Send failed: ' + err.message);
    showError('Could not send', err.message);
  }
}

function sendViaEWS(bodyText) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    const itemId = item.itemId;
    const replySubject = currentEmail.subject?.toLowerCase().startsWith('re:')
      ? currentEmail.subject
      : `Re: ${currentEmail.subject}`;

    // Escape special XML characters in the body
    const escapedBody = bodyText
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
      .replace(/\n/g, '<br/>');

    // Build EWS CreateItem SOAP request
    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:SavedItemFolderId>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:Subject>${replySubject.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</t:Subject>
          <t:Body BodyType="HTML">${escapedBody}</t:Body>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>${currentEmail.from}</t:EmailAddress>
              <t:Name>${(currentEmail.fromName || currentEmail.from).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</t:Name>
            </t:Mailbox>
          </t:ToRecipients>
          <t:IsReplyAllowed>true</t:IsReplyAllowed>
          <t:InReplyTo>${itemId}</t:InReplyTo>
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Check EWS response for errors
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(result.value, 'text/xml');
        const responseClass = xmlDoc.querySelector('[ResponseClass]')?.getAttribute('ResponseClass');

        if (responseClass === 'Success') {
          resolve();
        } else {
          const messageText = xmlDoc.querySelector('MessageText')?.textContent || 'EWS send failed';
          reject(new Error(messageText));
        }
      } else {
        reject(new Error(result.error?.message || 'EWS request failed'));
      }
    });
  });
}

// ─── STATE MANAGEMENT ─────────────────────────────────────────────────────────

function showState(state) {
  ['view-draft', 'view-idle', 'view-sent'].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.classList.add('hidden'); el.style.display = 'none'; }
  });
  const target = document.getElementById(`view-${state}`);
  if (target) {
    target.classList.remove('hidden');
    target.style.display = state === 'draft' ? 'flex' : 'block';
  }
}

function setStatus(type, text) {
  const bar = document.getElementById('status-bar');
  const icon = document.getElementById('status-icon');
  const statusText = document.getElementById('status-text');
  bar.className = `status-bar ${type}`;
  bar.classList.remove('hidden');
  statusText.textContent = text;
  icon.innerHTML = '';
  if (type === 'drafting') {
    const spinner = document.createElement('div');
    spinner.className = 'status-spinner';
    icon.appendChild(spinner);
  } else if (type === 'ready') {
    icon.textContent = '●';
  } else if (type === 'error') {
    icon.textContent = '⚠';
  } else {
    bar.classList.add('hidden');
  }
}

function showError(title, detail) {
  const box = document.getElementById('error-box');
  document.getElementById('error-title').textContent = title;
  document.getElementById('error-detail').innerHTML = (detail || '').replace(/\n/g, '<br>');
  box.classList.remove('hidden');
}

function hideError() {
  document.getElementById('error-box')?.classList.add('hidden');
}

// ─── WIRE BUTTONS ─────────────────────────────────────────────────────────────

function wireButtons() {
  document.getElementById('btn-send')?.addEventListener('click', sendReply);
  document.getElementById('btn-regen')?.addEventListener('click', () => {
    document.getElementById('draft-textarea').value = '';
    requestDraft();
  });
  document.getElementById('btn-manual-draft')?.addEventListener('click', () => {
    hideError(); requestDraft();
  });
  document.getElementById('btn-retry')?.addEventListener('click', () => {
    hideError(); requestDraft();
  });
  document.getElementById('btn-after-sent')?.addEventListener('click', () => {
    showState('idle'); setStatus('idle', '');
  });
}
