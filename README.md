# Family Finances 2026

A phone app for the family budget. It reads and writes the **Family Finances 2026** Google Sheet, so every phone that uses it sees the same data.

- **Current situation**: what came in, what was spent, what is in the joint account, and what is expected to be left over on the 9th.
- **Add transaction**: saved as a new row in the Transakcijos tab.
- **Budget**: edit this month's budget or plan next month (Seimos Biudzetas tab).
- **Pending**: payments caught from Revolut notifications on Android (MacroDroid → bridge → Laukiancios tab), edited and approved into Transakcijos.
- **Analytics, Transactions, Fixed costs**: drill-downs, trends and the costs posted on the 10th.

Financial months run from the 10th to the 9th.

## How it connects

The app talks to a small Google Apps Script web app attached to the sheet (`apps-script/Code.gs`). The script runs as the sheet owner and accepts requests only with the access key, so no Google login is needed on the phones and the data never leaves the sheet.

### One-time setup
1. In the sheet: **Extensions → Apps Script**, paste `apps-script/Code.gs`, set `KEY` to a long random value, save.
2. **Deploy → New deployment → Web app**. Execute as **Me**, Who has access **Anyone**. Authorize, copy the URL ending in `/exec`.
3. Open the app, paste the URL and the key.
4. On the second phone: in the app open Settings (⚙) → **Copy link**, send it, and open it there.

### Install on a phone
- Android (Chrome): menu ⋮ → **Install app**.
- iPhone: Share → **Add to Home Screen**.

No financial data is stored in this repository.
