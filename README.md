# BulkMailer — Bulk Email Marketing Tool

Send email campaigns to thousands of recipients with a configurable delay between
each send, a live progress dashboard, and Excel/CSV list import.

## Quick Start

```bash
npm install
npm start
```

Then open **http://localhost:3000** in your browser.

## How it works

```
public/index.html  →  Single-page UI (SMTP config, compose, recipients, dashboard)
server.js          →  Express backend (upload parsing, SMTP verify, delay loop, SSE)
```

1. Upload an Excel/CSV sheet **or** paste emails directly into the text area.
2. Fill in your SMTP credentials and compose your message.
3. Set the delay (default 10 s) and click **Start Campaign**.
4. Live progress updates stream to the browser via SSE — no polling needed.
5. Hit **Stop Campaign** any time to halt after the current email finishes.

---

## SMTP Providers (recommended)

| Provider | Host | Port | Notes |
|----------|------|------|-------|
| **Brevo** (Sendinblue) | `smtp-relay.brevo.com` | 587 | Free 300/day |
| **SendGrid** | `smtp.sendgrid.net` | 587 | Free 100/day |
| **Mailgun** | `smtp.mailgun.org` | 587 | Free 5 000/month |
| **AWS SES** | `email-smtp.<region>.amazonaws.com` | 587 | Sandbox until verified |

> ⚠ **Do not use personal Gmail SMTP** for bulk sending — daily limit is ~500 and
> the account can be flagged or suspended.

---

## Excel / CSV Format

The uploader auto-detects the email column.  
**Easiest format** — a column header that contains the word `email`:

```
name,email,city
Alice,alice@example.com,Delhi
Bob,bob@example.com,Mumbai
```

If no `email` header is found, every cell in the sheet is scanned for
email-shaped values — so any layout works.

---

## Deliverability checklist

Delay alone does **not** guarantee inbox delivery. These matter more:

1. **SPF / DKIM / DMARC** — set DNS records on your sending domain. Without them
   Gmail and Outlook route you straight to spam.
2. **Reputable SMTP provider** — use Brevo, SendGrid, Mailgun, or AWS SES, not
   personal Gmail.
3. **Unsubscribe link** — appended automatically. For 5 000+/day campaigns
   (Gmail/Yahoo 2024 rules) replace it with a real one-click handler.
4. **Permission-based list** — only email people who opted in. Purchased or
   scraped lists produce high spam complaint rates.
5. **Domain warm-up** — start with 20–50/day on a fresh domain, grow gradually.
   Blasting 1 000 emails from a cold domain triggers spam filters regardless of delay.

---

## Estimated campaign time

| Recipients | Delay | Approx. time |
|-----------|-------|--------------|
| 100 | 10 s | ~17 min |
| 500 | 10 s | ~1.4 h |
| 1 000 | 10 s | ~2.8 h |
| 5 000 | 10 s | ~13.9 h |

---

## API Reference

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/upload-list` | Parse Excel/CSV, return email array |
| `POST` | `/api/send-campaign` | Start campaign (responds immediately, sends in background) |
| `POST` | `/api/stop-campaign` | Set stop flag |
| `GET`  | `/api/progress` | SSE stream of live progress |
| `GET`  | `/unsubscribe?e=email` | Demo unsubscribe handler |
