// MailMind — Outlook Taskpane v2
// Auto-drafts on email open, sends via Outlook reply API

// ── Replace this with your Render URL once deployed ──
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
    showError('No email is currently open.', 'Open an email to get started.');
    return;
  }

  // Populate metadata immediately
  const from = item.from?.emailAddress || '';
  const fromName = item.from?.displayName || from;
  const subject = item.subject || 'No Subject';

  document.getElementById('meta-from').textContent = fromName || '—';
  document.getElementById('meta-subject').textContent = subject || '—';

  // Get body then auto-draft
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    const body = result.status === Office.AsyncResultStatus.Succeeded
      ? result.value.trim()
      : '';

    currentEmail = {
      id: item.itemId || '',
      conversationId: item.conversationId || '',
      subject,
      from,
      fromName,
      body: body.slice(0, 2000)
    };

    // Auto-draft immediately on open
    requestDraft();
  });
}

// ─── REQUEST DRAFT ────────────────────────────────────────────────────────────

async function requestDraft() {
  if (isDrafting || !currentEmail) return;
  isDrafting = true;

  setStatus('drafting', 'Drafting reply with Claude AI…');
  showState('idle'); // Show idle view while drafting (status bar shows progress)
  hideError();

  // Disable manual draft button while generating
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

    // Show draft
    document.getElementById('draft-textarea').value = data.draft;
    setStatus('ready', 'Draft ready — review and send');
    showState('draft');

  } catch (err) {
    const isTimeout = err.name === 'AbortError';
    const message = isTimeout ? 'Request timed out (30s)' : err.message;
    const hint = isTimeout
      ? 'The AI took too long. Try regenerating.'
      : `Make sure the MailMind backend is running at:\n${BACKEND_URL}`;

    setStatus('error', message);
    showError(message, hint);
    showState('idle');

  } finally {
    isDrafting = false;
    if (manualBtn) manualBtn.disabled = false;
  }
}

// ─── SEND REPLY ───────────────────────────────────────────────────────────────

async function sendReply() {
  const body = document.getElementById('draft-textarea').value.trim();
  if (!body) return;

  const btn = document.getElementById('btn-send');
  btn.textContent = 'Sending…';
  btn.disabled = true;

  try {
    // Use Outlook's built-in reply display (works on web + desktop + mobile)
    // This opens the reply compose window pre-filled with the draft
    Office.context.mailbox.item.displayReplyForm(body);

    // Show sent confirmation
    showState('sent');
    setStatus('idle', '');

  } catch (err) {
    btn.textContent = '✓ Send Reply';
    btn.disabled = false;
    setStatus('error', 'Failed to send: ' + err.message);
  }
}

// ─── STATE MANAGEMENT ─────────────────────────────────────────────────────────

function showState(state) {
  // Hide all views
  ['view-draft', 'view-idle', 'view-sent'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.classList.add('hidden');
      el.style.display = 'none';
    }
  });

  // Show target view
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

  // Clear old spinner
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
  detailEl.innerHTML = detail.includes('\n')
    ? detail.replace(/\n/g, '<br>')
    : detail;
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
