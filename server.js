const express    = require('express');
const nodemailer = require('nodemailer');
const multer     = require('multer');
const XLSX       = require('xlsx');
const path       = require('path');

require('dotenv').config();
const crypto = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3000;
const BASE_URL = (process.env.BASE_URL || `http://localhost:${PORT}`).replace(/\/$/, '');
const APP_PASSWORD = process.env.APP_PASSWORD || '7817808959';
const DAILY_SEND_LIMIT = 300;
const sessions = new Set();

const SMTP = {
  host: process.env.SMTP_HOST || 'smtp-relay.brevo.com',
  port: +(process.env.SMTP_PORT || 587),
  user: process.env.SMTP_USER || '',
  pass: process.env.SMTP_PASS || ''
};

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function cookieToken(req) {
  const raw = req.headers.cookie || '';
  const part = raw.split(';').map(s => s.trim()).find(s => s.startsWith('gate='));
  return part ? part.slice(5) : '';
}

function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace(/^Bearer\s+/i, '')
    || cookieToken(req)
    || String(req.query.token || '');
  if (!sessions.has(token)) return res.status(401).json({ error: 'Pehle password daalo' });
  next();
}

app.post('/api/login', (req, res) => {
  if (String(req.body?.password || '') !== APP_PASSWORD)
    return res.status(401).json({ error: 'Galat password' });
  const token = crypto.randomBytes(24).toString('hex');
  sessions.add(token);
  res.setHeader('Set-Cookie', `gate=${token}; HttpOnly; SameSite=Lax; Path=/`);
  res.json({ ok: true, token });
});

app.get('/api/session', requireAuth, (req, res) => res.json({ ok: true }));
app.get('/api/ping', requireAuth, (req, res) => res.json({ ok: true, running: state.running }));

app.post('/api/logout', (req, res) => {
  sessions.delete(cookieToken(req));
  res.setHeader('Set-Cookie', 'gate=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0');
  res.json({ ok: true });
});

const upload = multer({ storage: multer.memoryStorage() });

// ── Campaign state ─────────────────────────────────────────────────────────────
let state = {
  running: false,
  stop:    false,
  total:   0,
  sent:    0,
  failed:  0,
  log:     []           // [{ email, status, error?, time }]
};

let dailyUsage = { date: new Date().toISOString().slice(0, 10), triggered: 0 };

let clients = [];       // SSE connections

// ── Helpers ───────────────────────────────────────────────────────────────────
function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function getDailyUsage() {
  const today = new Date().toISOString().slice(0, 10);
  if (dailyUsage.date !== today) dailyUsage = { date: today, triggered: 0 };
  return dailyUsage;
}

async function sendViaBrevoApi({ to, fromEmail, displayName, subject, html, text }) {
  const apiKey = process.env.BREVO_API_KEY;
  const r = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': apiKey,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      sender: { name: displayName, email: fromEmail },
      to: [{ email: to }],
      subject,
      htmlContent: html,
      textContent: text,
      replyTo: { email: fromEmail, name: displayName }
    })
  });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data.message || data.error || 'Brevo API fail');
  return data;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function bodyToHtml(raw) {
  const text = String(raw || '').trim();
  if (/<[a-z][\s\S]*>/i.test(text)) return text;
  return text
    .split(/\n{2,}/)
    .map(p => `<p style="margin:0 0 16px;line-height:1.7;color:#333;font-size:16px;">${
      escapeHtml(p).replace(/\n/g, '<br>')
    }</p>`)
    .join('');
}

function wrapEmail({ displayName, inner, unsubUrl }) {
  const brand = escapeHtml(displayName);
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Georgia,'Times New Roman',serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:32px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #e5e7eb;">
          <tr>
            <td style="background:#111827;padding:22px 32px;">
              <p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;letter-spacing:1.5px;text-transform:uppercase;color:#9ca3af;">${brand}</p>
            </td>
          </tr>
          <tr>
            <td style="height:4px;background:#6366f1;font-size:0;line-height:0;">&nbsp;</td>
          </tr>
          <tr>
            <td style="padding:36px 32px 28px;font-family:Arial,Helvetica,sans-serif;">
              ${inner}
            </td>
          </tr>
          <tr>
            <td style="padding:0 32px 28px;font-family:Arial,Helvetica,sans-serif;">
              <p style="margin:0;color:#111827;font-size:15px;line-height:1.6;">Warm regards,</p>
              <p style="margin:8px 0 0;color:#111827;font-size:15px;font-weight:700;">${brand}</p>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 32px;background:#f9fafb;border-top:1px solid #e5e7eb;font-family:Arial,Helvetica,sans-serif;">
              <p style="margin:0;font-size:12px;color:#9ca3af;line-height:1.5;">
                You received this email from ${brand}.
                <a href="${unsubUrl}" style="color:#6b7280;">Unsubscribe</a>
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function push(event, data) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  clients.forEach(res => { try { res.write(msg); } catch(_){} });
}

// ── API: Upload Excel/CSV → return email list ──────────────────────────────────
app.post('/api/upload', requireAuth, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Koi file nahi mili' });

  let wb;
  try { wb = XLSX.read(req.file.buffer, { type: 'buffer' }); }
  catch(e) { return res.status(400).json({ error: 'File parse nahi hui: ' + e.message }); }

  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  let emails = [];

  // Column jisme "email" ho
  const col = rows.length ? Object.keys(rows[0]).find(k => /email/i.test(k)) : null;
  if (col) {
    emails = rows.map(r => String(r[col]).trim()).filter(e => emailRe.test(e));
  } else {
    // Koi bhi cell jo email jaisi ho
    emails = Object.values(ws)
      .map(c => String(c.v ?? '').trim())
      .filter(e => emailRe.test(e));
  }

  emails = [...new Set(emails)];
  if (!emails.length) return res.status(400).json({ error: '"email" column nahi mili ya koi valid email nahi' });

  res.json({ emails, count: emails.length });
});

// ── API: Start campaign ────────────────────────────────────────────────────────
app.post('/api/start', requireAuth, async (req, res) => {
  if (state.running) return res.status(409).json({ error: 'Campaign pehle se chal rahi hai' });

  const { subject, body, emails, delay, from, fromName } = req.body;
  const fromEmail = String(from || '').trim();
  const displayName = String(fromName || 'Mygrow Software').trim() || 'Mygrow Software';

  if (!subject || !body)
    return res.status(400).json({ error: 'Subject ya email body khaali hai' });
  if (!emails?.length)
    return res.status(400).json({ error: 'Koi email nahi mili list mein' });
  const usage = getDailyUsage();
  const remaining = DAILY_SEND_LIMIT - usage.triggered;
  if (emails.length > remaining)
    return res.status(400).json({ error: remaining > 0
      ? `Aaj sirf ${remaining} emails ki limit baaki hai. Maximum ${DAILY_SEND_LIMIT} emails per day bhej sakte ho.`
      : `Your today limit is complete. Aaj maximum ${DAILY_SEND_LIMIT} emails already trigger ho chuki hain.` });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(fromEmail))
    return res.status(400).json({ error: 'From email galat hai — Brevo mein verified sender daalo (Gmail/domain), smtp-brevo.com nahi' });
  if (/smtp-brevo\.com$/i.test(fromEmail))
    return res.status(400).json({ error: 'From mein SMTP login mat daalo. Apna verified email use karo (jaise Gmail).' });

  const useApi = !!process.env.BREVO_API_KEY;
  let transporter;

  if (!useApi) {
    if (!SMTP.user || !SMTP.pass)
      return res.status(400).json({ error: 'SMTP/API key missing. Render pe BREVO_API_KEY lagao (SMTP free Render pe block hai).' });
    try {
      transporter = nodemailer.createTransport({
        host: SMTP.host,
        port: SMTP.port,
        secure: SMTP.port === 465,
        connectionTimeout: 12000,
        greetingTimeout: 12000,
        auth: { user: SMTP.user, pass: SMTP.pass }
      });
      await transporter.verify();
    } catch (e) {
      return res.status(400).json({
        error: 'SMTP fail (Render Free 587 block karta hai). Brevo API key banao aur BREVO_API_KEY env mein daalo. ' + e.message
      });
    }
  }

  const delaySec = Math.max(1, +delay || 10);
  usage.triggered += emails.length;

  // State reset
  state = { running: true, stop: false, total: emails.length, sent: 0, failed: 0, log: [] };

  res.json({ ok: true, total: emails.length });

  // ── Background sending loop ──────────────────────────────────────────────────
  (async () => {
    push('start', { total: state.total });

    for (let i = 0; i < emails.length; i++) {
      if (state.stop) break;

      const to = emails[i].trim();
      const unsubUrl = `${BASE_URL}/unsubscribe?e=${encodeURIComponent(to)}`;
      const html = wrapEmail({
        displayName,
        inner: bodyToHtml(body),
        unsubUrl
      });
      const text = String(body).trim() + `\n\n— ${displayName}\nUnsubscribe: ${unsubUrl}`;

      try {
        if (useApi) {
          const info = await sendViaBrevoApi({ to, fromEmail, displayName, subject, html, text });
          console.log('SENT', to, info.messageId || 'api');
        } else {
          const info = await transporter.sendMail({
            from: `"${displayName}" <${fromEmail}>`,
            replyTo: fromEmail,
            to,
            subject,
            html,
            text
          });
          console.log('SENT', to, info.response || info.messageId);
        }
        state.sent++;
        state.log.unshift({ email: to, status: 'sent',   time: new Date().toLocaleTimeString() });
      } catch(e) {
        console.error('FAIL', to, e.message);
        state.failed++;
        state.log.unshift({ email: to, status: 'failed', time: new Date().toLocaleTimeString(), error: e.message });
      }

      push('update', { sent: state.sent, failed: state.failed, total: state.total, last: state.log[0] });

      if (i < emails.length - 1 && !state.stop) await sleep(delaySec * 1000);
    }

    state.running = false;
    push('done', { sent: state.sent, failed: state.failed });
  })();
});

// ── API: Stop campaign ─────────────────────────────────────────────────────────
app.post('/api/stop', requireAuth, (req, res) => {
  state.stop = true;
  res.json({ ok: true });
});

// ── API: Live progress (SSE) ───────────────────────────────────────────────────
app.get('/api/progress', requireAuth, (req, res) => {
  res.setHeader('Content-Type',  'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection',    'keep-alive');
  res.flushHeaders();
  clients.push(res);
  if (state.total) {
    push('update', { sent: state.sent, failed: state.failed, total: state.total, last: state.log[0] || null });
  }
  req.on('close', () => { clients = clients.filter(c => c !== res); });
});

// ── Unsubscribe page ───────────────────────────────────────────────────────────
app.get('/unsubscribe', (req, res) => {
  res.send(`<h2 style="font-family:sans-serif;padding:40px">
    ✅ ${req.query.e || 'Email'} unsubscribe ho gaya.</h2>`);
});

app.listen(PORT, '0.0.0.0', () => console.log(`\n✅ ${BASE_URL}\n`));
