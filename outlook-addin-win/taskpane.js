/* global Office */
'use strict';

(function () {

  // ── state ────────────────────────────────────────────────────────
  let enabled     = true;
  let tone        = 'professional';
  let draft       = '';
  let mailItem    = null;

  // ── helpers ──────────────────────────────────────────────────────
  const $  = id => document.getElementById(id);
  const show = id => $(id)?.classList.remove('hidden');
  const hide = id => $(id)?.classList.add('hidden');

  function store(key, val) { try { localStorage.setItem(key, String(val)); } catch (_) {} }
  function load(key, def)  { try { return localStorage.getItem(key) ?? def; } catch (_) { return def; } }

  // ── Office init ──────────────────────────────────────────────────
  Office.onReady(info => {
    if (info.host !== Office.HostType.Outlook) return;

    // Restore settings
    enabled = load('aiEnabled', 'true') === 'true';
    $('main-tog').checked = enabled;
    $('tog-lbl').textContent = enabled ? 'ON' : 'OFF';

    const savedKey = load('apiKey', '');
    $('api-inp').value = savedKey;

    // Wire UI
    wireToggle();
    wireTones();
    wireButtons();

    // Load email data
    mailItem = Office.context.mailbox.item;
    loadEmailPreview();
    renderLayout();
  });

  // ── email preview ────────────────────────────────────────────────
  function loadEmailPreview() {
    if (!mailItem) return;

    const from = mailItem.from;
    $('email-from').textContent = from
      ? (from.displayName && from.displayName !== from.emailAddress
          ? `${from.displayName} <${from.emailAddress}>`
          : from.emailAddress)
      : 'Unknown sender';

    $('email-subj').textContent = mailItem.subject || '(no subject)';

    mailItem.body.getAsync(Office.CoercionType.Text, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return;
      const text = res.value.replace(/[\r\n]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
      $('email-snip').textContent = text.slice(0, 220) + (text.length > 220 ? '…' : '');
    });
  }

  // ── layout renderer ──────────────────────────────────────────────
  function renderLayout() {
    const hasKey = load('apiKey', '').trim().startsWith('sk-ant-');

    // Always show API box if key is missing
    hasKey ? hide('api-box') : show('api-box');

    if (!enabled) {
      show('notice-off');
      hide('sec-email'); hide('sec-tone'); hide('sec-gen');
      hide('draft-card'); hide('loading'); hide('error'); hide('success');
      return;
    }

    hide('notice-off');
    show('sec-email');
    show('sec-tone');
    show('sec-gen');
  }

  // ── toggle ───────────────────────────────────────────────────────
  function wireToggle() {
    $('main-tog').addEventListener('change', e => {
      enabled = e.target.checked;
      $('tog-lbl').textContent = enabled ? 'ON' : 'OFF';
      store('aiEnabled', enabled);
      renderLayout();
    });
  }

  // ── tone buttons ─────────────────────────────────────────────────
  function wireTones() {
    document.querySelectorAll('.tone-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tone-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        tone = btn.dataset.tone;
      });
    });
  }

  // ── action buttons ───────────────────────────────────────────────
  function wireButtons() {
    $('gen-btn')   .addEventListener('click', generate);
    $('btn-insert').addEventListener('click', insertDraft);
    $('btn-regen') .addEventListener('click', generate);
    $('btn-disc')  .addEventListener('click', () => { hide('draft-card'); hide('success'); });
    $('api-save')  .addEventListener('click', saveKey);
    $('api-inp')   .addEventListener('keydown', e => { if (e.key === 'Enter') saveKey(); });
  }

  // ── save API key ─────────────────────────────────────────────────
  function saveKey() {
    const key = $('api-inp').value.trim();
    if (!key) return;
    store('apiKey', key);
    const pill = $('api-pill');
    pill.textContent = 'SAVED ✓';
    pill.className = 'api-pill ok';
    setTimeout(() => { hide('api-box'); renderLayout(); }, 900);
  }

  // ── generate draft ───────────────────────────────────────────────
  async function generate() {
    const apiKey = load('apiKey', '');
    if (!apiKey) { showApiRequired(); return; }

    hide('draft-card'); hide('error'); hide('success');
    show('loading');
    $('gen-btn').disabled = true;

    try {
      const body   = await getBody();
      const sender = mailItem?.from?.displayName || 'the sender';
      const subj   = mailItem?.subject || '';

      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          system: buildSystem(tone),
          messages: [{
            role: 'user',
            content: `Write a reply to this email.\n\nFrom: ${sender}\nSubject: ${subj}\n\n---\n${body.slice(0, 4000)}\n---`,
          }],
        }),
      });

      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error?.message || `HTTP ${resp.status}`);

      draft = data.content?.[0]?.text || '';
      if (!draft) throw new Error('Empty response — please try again.');

      $('draft-body').textContent = draft;
      hide('loading');
      show('draft-card');

    } catch (err) {
      hide('loading');
      $('error').textContent = '⚠ ' + err.message;
      show('error');
    } finally {
      $('gen-btn').disabled = false;
    }
  }

  // ── insert draft into Outlook reply window ───────────────────────
  function insertDraft() {
    if (!draft) return;

    // displayReplyForm opens a Reply window pre-populated with our HTML
    mailItem.displayReplyForm({ htmlBody: toHtml(draft) });

    hide('draft-card');
    show('success');
    setTimeout(() => hide('success'), 6000);
  }

  // ── helpers ──────────────────────────────────────────────────────
  function getBody() {
    return new Promise((res, rej) => {
      mailItem.body.getAsync(Office.CoercionType.Text, result => {
        result.status === Office.AsyncResultStatus.Succeeded
          ? res(result.value || '')
          : rej(new Error(result.error?.message || 'Could not read email body'));
      });
    });
  }

  function showApiRequired() {
    show('api-box');
    $('error').textContent = '⚠ Please enter and save your Anthropic API key first.';
    show('error');
  }

  function buildSystem(t) {
    const map = {
      professional: 'professional, clear, and business-appropriate',
      friendly:     'warm, friendly, and approachable while remaining professional',
      concise:      'brief and direct — no filler, just the key points',
      formal:       'highly formal and structured, suitable for official correspondence',
      empathetic:   'empathetic, understanding, and supportive',
      assertive:    'confident, direct, and assertive while still respectful',
    };
    return `You are an expert email assistant. Write a reply that is ${map[t] || map.professional}.
Rules:
- Write ONLY the reply body — no subject, no headers, no meta-commentary
- No placeholder text like [Your Name]
- End with an appropriate sign-off (e.g. "Best regards,") but no name
- Sound human, not robotic`;
  }

  function toHtml(text) {
    const esc = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000">'
      + esc.split('\n').map(l => `<p style="margin:0 0 6px">${l||'&nbsp;'}</p>`).join('')
      + '</div>';
  }

})();
