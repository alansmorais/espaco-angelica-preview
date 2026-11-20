# Espaço Angélica — Isolated Read-Only Preview

Dedicated, standalone GitHub Pages preview repository for the Espaço Angélica do Carmo interface.

- **Source repository:** `alansmorais/espaco-angelica-clean`
- **Source branch:** `feature/google-calendar-sync-preview`
- **Source commit:** `50d96a07f6b0d6f6aed95f1503d58e3a0210a6f4`
- **Preview URL:** https://alansmorais.github.io/espaco-angelica-preview/

## Isolation Guarantees

- Completely separate from production (`espacoangelica.alansmsolutions.com`) and its GitHub Pages environment.
- Strictly read-only: no writes to Google Sheets, no Google Calendar mutations, no email/WhatsApp dispatches, and no production Apps Script endpoint configured.
- Excludes production workflows, Google Apps Script code, credentials, `.env` files, and real customer/booking records.
