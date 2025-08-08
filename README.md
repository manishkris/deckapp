# Gmail Inbox Cleaner Agent

Automates unsubscribing from unwanted emails and tidying your Gmail inbox with safe defaults.

- OAuth with Gmail API (no password)
- Only touches UNREAD mail in INBOX by default
- Targets typical newsletters/promotions using List-Unsubscribe header
- Dry-run first (labels for review), then apply
- Weekly scheduling guide
- Email summary report and a CSV ledger of unsubscribed domains

## 1) Prerequisites

- Python 3.9+
- A Google Cloud project with Gmail API enabled
- OAuth client (Desktop) JSON downloaded as `credentials.json` into the project root

Enable Gmail API and create OAuth client:
- Go to Google Cloud Console → APIs & Services → Enable APIs → enable "Gmail API"
- Credentials → Create Credentials → OAuth client ID → Application type: Desktop
- Download the JSON to this folder as `credentials.json`

## 2) Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 3) Configure

Edit `config.yaml` as needed. Defaults match your preferences:
- Inbox-only, unread-only
- Keep senders: Bhavna, Apsarabollywood, Krishnan, Vignesh, Visalakshi
- Unsubscribe on all supported methods (List-Unsubscribe: One-Click URL, HTTPS, mailto)
- Archive after unsubscribe
- Delete disabled until you opt-in

## 4) First run (dry-run)

```bash
python inbox_agent.py --dry-run
```
The browser will open for OAuth. A token is stored in `token.json` for future runs.

Dry-run will:
- Scan UNREAD messages in INBOX (optionally only Promotions/Updates)
- Stage candidates with label `AI/To-Unsubscribe`
- Email you a summary report
- Make no destructive changes

## 5) Apply actions

After reviewing the staged label, run:
```bash
python inbox_agent.py --apply
```
This will:
- Attempt to unsubscribe via One-Click/URL/mailto
- Archive the message, add label `AI/Unsubscribed`
- Record domain/sender in `data/unsubscribed_ledger.csv`

Deletion remains OFF by default. If you enable it later, messages go to `AI/Trash-Staging` for N days before purge.

## 6) Schedule weekly

Using cron (runs every Sunday at 3:30am local time):
```bash
crontab -e
# Add:
30 3 * * 0 cd /path/to/project && /path/to/project/.venv/bin/python inbox_agent.py --apply >> agent.log 2>&1
```

Or run dry-run weekly:
```bash
30 3 * * 0 cd /path/to/project && /path/to/project/.venv/bin/python inbox_agent.py --dry-run >> agent.log 2>&1
```

## 7) Troubleshooting

- If labels are missing, the agent creates them automatically
- To reset auth, delete `token.json` and re-run
- If your org blocks browser auth, create an OAuth client with appropriate access and use the device flow opened by the script

## Security notes

- OAuth tokens are stored locally in `token.json`
- The agent only acts on UNREAD emails unless you change the config
- Unsubscribe URLs are requested over HTTPS where possible