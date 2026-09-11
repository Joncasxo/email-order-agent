# Email Order Agent

A Windows desktop agent (Python/Tkinter, Lithuanian UI) that watches a Microsoft 365 mailbox via Microsoft Graph and uses the Anthropic Claude API to detect and extract purchase orders from incoming email — including forwarded chains, Excel/PDF/ODS attachments and photos — into a local SQLite order list.

## Features

- **AI order extraction** — each candidate email is classified and parsed by Claude with a strict structured-output schema (customer, contact details, line items with quantities and notes, delivery dates, contract numbers, confidence score).
- **Conversation-chain awareness** — replies and forwards (Re:/FW:) are matched to already registered orders and recorded as updates instead of duplicates; unseen updates are flagged in the list.
- **Attachment parsing** — order content is read from `.xlsx`/`.xls`/`.ods` spreadsheets, PDFs and images, and merged with the email body into a single order.
- **Color-aware HTML parsing** — text the client highlighted in red or yellow (e.g. the chosen row of a price table) is preserved as inline markers, so the AI treats highlighting as the client's selection.
- **Ambiguity warnings** — a red warning banner is raised when a request is risky: "correct the registration, do NOT reorder", unclear or missing quantities, possible duplicates of an earlier order.
- **Supplier inquiry tracking** — inquiries sent to suppliers carry the order code in the subject; replies are matched back to the order and flagged until reviewed.
- **Contact export** — client contacts collected from emails can be exported to vCard for import into a phone.
- **Read-only viewer mode** — a synced database copy and viewer app let a manager follow the order list without write access.
- **Low API cost** — a keyword pre-filter skips obviously irrelevant mail before any API call; typical usage is a few EUR per month with the Sonnet model.

## Requirements

- Windows 10/11
- Python 3.12+ with `pywin32`, `msal`, `requests`, `anthropic`, `openpyxl` (see `requirements.txt`)
- A Microsoft 365 mailbox
- An Anthropic API key

## Setup

1. Install dependencies:

   ```
   pip install -r requirements.txt
   ```

2. Register an app in Entra ID: portal.azure.com → **App registrations** → **New registration** — public client, redirect URI `http://localhost`, delegated permission **Mail.Read** (grant admin consent if your tenant requires it). Note the **Application (client) ID** and **Directory (tenant) ID**.

3. Run the app (`Paleisti.bat`), open **Settings** and paste the Anthropic API key, the Entra Client ID / Tenant ID and your mailbox address.

4. The first mail check opens a Microsoft sign-in window once; afterwards the token is refreshed silently.

## Configuration

- **Sender filters** — empty means all incoming mail is checked. The typical setup is direct client emails; optionally add a colleague who forwards orders to you — the real client is recognized inside the forwarded message (From headers and signature).
- **Check interval** — how often the mailbox is polled.
- **Keyword filter** — the pre-filter word list applied before any AI call (emails with unreadable attachments such as PDFs or photos always pass through).
- **Model choice** — pick a more accurate or a cheaper Claude model depending on cost sensitivity.

## How it works

Mail is pulled from Microsoft Graph every N minutes → a keyword pre-filter drops irrelevant messages → Claude performs structured extraction over the body, forwarded chain and attachments → results land in a local SQLite database and are shown in the Tkinter UI, where orders get statuses, notes and Excel export.

## Notes

- The UI and the AI prompts are in Lithuanian.
- The extraction prompt in `extractor.py` is tailored to a specific purchasing workflow (contract handling, order-code format, typical clients) — adapt its context lines to your own company before use.
- All data stays local: orders live in a SQLite database on your machine, and the Anthropic API key and the mail token cache are encrypted with Windows DPAPI, bound to your user account.

## License

MIT — see [LICENSE](LICENSE).
