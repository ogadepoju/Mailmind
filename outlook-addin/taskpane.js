// MailMind — Outlook Taskpane v6
// Uses displayReplyAllForm with pre-filled body — works on all account types

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
    setStatus('ready', 'Draft ready — click Send to open pre-filled reply');
    showState('draft');

  } catch (err) {
    const isTimeout = err.name === 'AbortError';
    setStatus('error', isTimeout ? 'Request timed out' : err.message);
    showError(isTimeout ? 'Timed out' : err.message, `Backend: ${BACKEND_URL}`);
    showState('idle');
  } finally {
    isDrafting = false;
    if (manualBtn) manualBtn.disabled = false;
  }
}

// ─── SEND ────────────────────────────────────────────────────────────────────
// Opens a pre-filled reply compose window.
// User clicks Send once in the compose window — one click, already addressed.

function sendReply() {
  const draftBody = document.getElementById('draft-textarea').value.trim();
  if (!draftBody) return;

  const btn = document.getElementById('btn-send');
  btn.textContent = 'Opening…';
  btn.disabled = true;

  try {
    // Convert plain text to HTML for better formatting
    const htmlBody = draftBody
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\n\n/g, '</p><p>')
      .replace(/\n/g, '<br>');

    const fullHtml = `<p>${htmlBody}</p>`;

    // displayReplyForm opens a compose window pre-filled with:
    // - To: already set to the original sender
    // - Subject: Re: [original subject]  
    // - Body: your draft already typed in
    // User just clicks Send — one click
    Office.context.mailbox.item.displayReplyForm({
      htmlBody: fullHtml
    });

    // Show instructions
    showState('sending');

  } catch (err) {
    console.error('[MailMind] Open reply error:', err);
    btn.textContent = '✓ Open Reply';
    btn.disabled = false;
    setStatus('error', 'Could not open reply: ' + err.message);
  }
}

// ─── STATE MANAGEMENT ─────────────────────────────────────────────────────────

function showState(state) {
  ['view-draft', 'view-idle', 'view-sent', 'view-sending'].forEach(id => {
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
  document.getElementById('error-box')?.classList.remove('hidden');
  document.getElementById('error-title').textContent = title;
  document.getElementById('error-detail').innerHTML = (detail || '').replace(/\n/g, '<br>');
}

function hideError() {
  document.getElementById('error-box')?.classList.add('hidden');
}

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
  document.getElementById('btn-sending-back')?.addEventListener('click', () => {
    showState('draft');
    const btn = document.getElementById('btn-send');
    if (btn) { btn.textContent = '✓ Open Reply'; btn.disabled = false; }
  });
}
