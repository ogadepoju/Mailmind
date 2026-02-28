// MailMind — Outlook Taskpane v3
// Sends directly via Microsoft Graph API — no draft, no compose window

const BACKEND_URL = 'https://mailmind-6f1d.onrender.com';

// Microsoft Graph - these scopes allow reading + sending email
const GRAPH_SCOPES = ['https://graph.microsoft.com/Mail.Send'];

let currentEmail = null;
let isDrafting = false;
let graphToken = null;

// ─── OFFICE INIT ─────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailAndDraft();
    wireButtons();
  }
});

// ─── GET GRAPH TOKEN ─────────────────────────────────────────────────────────

async function getGraphToken() {
  if (graphToken) return graphToken;

  try {
    // Use Office SSO to get a token silently
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true
    });
    graphToken = token;
    return token;
  } catch (err) {
    console.error('[MailMind] SSO error:', err);
    throw new Error('Could not get auth token: ' + err.message);
  }
}

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
    const hint = isTimeout
      ? 'The AI took too long. Try regenerating.'
      : `Make sure the backend is running at:\n${BACKEND_URL}`;
    setStatus('error', message);
    showError(message, hint);
    showState('idle');
  } finally {
    isDrafting = false;
    if (manualBtn) manualBtn.disabled = false;
  }
}

// ─── SEND REPLY DIRECTLY ─────────────────────────────────────────────────────

async function sendReply() {
  const draftBody = document.getElementById('draft-textarea').value.trim();
  if (!draftBody) return;

  const btn = document.getElementById('btn-send');
  btn.textContent = 'Sending…';
  btn.disabled = true;

  try {
    const token = await getGraphToken();
    await sendViaGraph(token, draftBody);

    showState('sent');
    setStatus('idle', '');

  } catch (err) {
    console.error('[MailMind] Send error:', err);

    // If Graph API fails, fall back to EWS send
    try {
      await sendViaEWS(draftBody);
      showState('sent');
      setStatus('idle', '');
    } catch (ewsErr) {
      btn.textContent = '✓ Send Reply';
      btn.disabled = false;
      setStatus('error', 'Send failed: ' + err.message);
      showError('Could not send', err.message);
      showState('idle');
    }
  }
}

// ─── SEND VIA MICROSOFT GRAPH ────────────────────────────────────────────────

async function sendViaGraph(token, draftBody) {
  const subject = currentEmail.subject || '';
  const replySubject = subject.toLowerCase().startsWith('re:')
    ? subject
    : `Re: ${subject}`;

  const message = {
    subject: replySubject,
    body: {
      contentType: 'Text',
      content: draftBody
    },
    toRecipients: [
      {
        emailAddress: {
          address: currentEmail.from,
          name: currentEmail.fromName || currentEmail.from
        }
      }
    ]
  };

  // If we have a conversation ID, reply in thread
  if (currentEmail.id) {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${currentEmail.id}/reply`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ message, comment: draftBody })
      }
    );

    if (res.status === 202 || res.ok) return; // 202 Accepted = success for reply
    
    const errData = await res.json().catch(() => ({}));
    throw new Error(errData?.error?.message || `Graph error ${res.status}`);
  }

  // Fallback: send as new message
  const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ message, saveToSentItems: true })
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error(errData?.error?.message || `Graph error ${res.status}`);
  }
}

// ─── SEND VIA EWS (FALLBACK) ──────────────────────────────────────────────────

async function sendViaEWS(draftBody) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      async (result) => {
        try {
          // Create reply item and send
          Office.context.mailbox.displayReplyForm({
            htmlBody: draftBody.replace(/\n/g, '<br>'),
          });
          resolve();
        } catch (err) {
          reject(err);
        }
      }
    );
  });
}

// ─── STATE MANAGEMENT ─────────────────────────────────────────────────────────

function showState(state) {
  ['view-draft', 'view-idle', 'view-sent'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.classList.add('hidden');
      el.style.display = 'none';
    }
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
  const titleEl = document.getElementById('error-title');
  const detailEl = document.getElementById('error-detail');
  box.classList.remove('hidden');
  titleEl.textContent = title;
  detailEl.innerHTML = (detail || '').replace(/\n/g, '<br>');
}

function hideError() {
  document.getElementById('error-box')?.classList.add('hidden');
}

// ─── WIRE BUTTONS ─────────────────────────────────────────────────────────────

function wireButtons() {
  document.getElementById('btn-send')
    ?.addEventListener('click', sendReply);

  document.getElementById('btn-regen')
    ?.addEventListener('click', () => {
      document.getElementById('draft-textarea').value = '';
      requestDraft();
    });

  document.getElementById('btn-manual-draft')
    ?.addEventListener('click', () => {
      hideError();
      requestDraft();
    });

  document.getElementById('btn-retry')
    ?.addEventListener('click', () => {
      hideError();
      requestDraft();
    });

  document.getElementById('btn-after-sent')
    ?.addEventListener('click', () => {
      showState('idle');
      setStatus('idle', '');
    });
}
