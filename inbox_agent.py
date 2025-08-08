import argparse
import base64
import csv
import os
import re
import sys
import time
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, List, Optional, Tuple

import requests
import tldextract
import yaml
from dateutil import tz
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from rich.console import Console
from rich.table import Table

console = Console()

# If modifying scopes, delete token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
]

DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.yaml")


def load_config(config_path: str) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def ensure_dir(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)


def get_gmail_service(credentials_path: str, token_path: str):
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    return service


def get_label_map(service) -> Dict[str, str]:
    resp = service.users().labels().list(userId="me").execute()
    labels = resp.get("labels", [])
    return {label["name"]: label["id"] for label in labels}


def ensure_label(service, name: str) -> str:
    label_map = get_label_map(service)
    if name in label_map:
        return label_map[name]
    body = {
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show",
        "name": name,
    }
    created = service.users().labels().create(userId="me", body=body).execute()
    return created["id"]


def find_header(headers: List[Dict[str, str]], name: str) -> Optional[str]:
    for h in headers:
        if h.get("name", "").lower() == name.lower():
            return h.get("value")
    return None


def parse_list_unsubscribe(value: str) -> Tuple[List[str], List[str]]:
    if not value:
        return [], []
    # Split by commas not inside quotes or angle brackets
    parts = re.split(r",\s*(?=(?:[^<]*<[^>]*>)*[^>]*$)", value)
    urls: List[str] = []
    mailtos: List[str] = []
    for part in parts:
        token = part.strip().strip("<>")
        if token.lower().startswith("mailto:"):
            mailtos.append(token)
        elif token.lower().startswith("http"):
            urls.append(token)
    return urls, mailtos


def get_sender_domain(from_value: str) -> str:
    # Extract email from From header and then domain
    m = re.search(r"<([^>]+)>", from_value)
    email_addr = m.group(1) if m else from_value
    email_addr = email_addr.strip().strip("\"')")
    at_idx = email_addr.rfind("@")
    domain = email_addr[at_idx + 1 :] if at_idx != -1 else email_addr
    ext = tldextract.extract(domain)
    if ext.suffix:
        return f"{ext.domain}.{ext.suffix}".lower()
    return domain.lower()


def sender_matches_keep(from_value: str, keep_substrings: List[str]) -> bool:
    hay = from_value.lower()
    return any(sub.lower() in hay for sub in keep_substrings)


def list_messages(service, label_ids: List[str], q: str, max_messages: int) -> List[Dict[str, str]]:
    messages: List[Dict[str, str]] = []
    page_token = None
    while True:
        params = {
            "userId": "me",
            "labelIds": label_ids,
            "q": q,
            "maxResults": min(500, max_messages - len(messages)),
        }
        if page_token:
            params["pageToken"] = page_token
        resp = service.users().messages().list(**params).execute()
        msgs = resp.get("messages", [])
        messages.extend(msgs)
        page_token = resp.get("nextPageToken")
        if not page_token or len(messages) >= max_messages:
            break
    return messages


def get_message_metadata(service, msg_id: str) -> Dict:
    return (
        service
        .users()
        .messages()
        .get(
            userId="me",
            id=msg_id,
            format="metadata",
            metadataHeaders=[
                "From",
                "To",
                "Subject",
                "Date",
                "List-Unsubscribe",
                "List-Unsubscribe-Post",
                "Precedence",
                "Auto-Submitted",
            ],
        )
        .execute()
    )


def add_labels(service, msg_id: str, add_label_ids: List[str], remove_label_ids: Optional[List[str]] = None) -> None:
    body = {
        "addLabelIds": add_label_ids or [],
        "removeLabelIds": remove_label_ids or [],
    }
    service.users().messages().modify(userId="me", id=msg_id, body=body).execute()


def try_unsubscribe(urls: List[str], mailtos: List[str], post_header: Optional[str]) -> Tuple[str, str]:
    # Returns (method, detail) where method in {"one_click", "url_get", "mailto", "none"}
    headers = {"User-Agent": "InboxCleaner/1.0 (+https://github.com/)"}
    # Prefer One-Click POST if declared
    if post_header and any(p.strip().lower() == "list-unsubscribe=one-click" for p in post_header.split(";")):
        for u in urls:
            try:
                resp = requests.post(u, headers=headers, timeout=12)
                if 200 <= resp.status_code < 300:
                    return "one_click", u
            except Exception:
                continue
    # Fallback to GET on https URL
    for u in urls:
        try:
            resp = requests.get(u, headers=headers, timeout=12, allow_redirects=True)
            if 200 <= resp.status_code < 400:
                return "url_get", u
        except Exception:
            continue
    # Fallback to mailto: we "send" an email using Gmail API later, just indicate intent here
    if mailtos:
        return "mailto", mailtos[0]
    return "none", ""


def send_mail(service, to_addr: str, subject: str, html_body: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["to"] = to_addr
    msg["subject"] = subject
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()


def send_unsubscribe_mail(service, mailto_url: str, from_subject: str) -> bool:
    # mailto may include parameters; parse rudimentarily
    # Examples: mailto:unsub@example.com, mailto:list@example.com?subject=unsubscribe
    m = re.match(r"mailto:([^?]+)(?:\?(.*))?", mailto_url, re.IGNORECASE)
    if not m:
        return False
    to_addr = m.group(1)
    query = m.group(2) or ""
    params = dict(re.findall(r"([^=&]+)=([^&]*)", query))
    subject = params.get("subject", "unsubscribe")
    body_text = params.get("body", "unsubscribe")

    msg = MIMEMultipart("alternative")
    msg["to"] = to_addr
    msg["subject"] = subject
    msg.attach(MIMEText(body_text or "unsubscribe", "plain", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        return True
    except Exception:
        return False


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def write_ledger(ledger_path: str, rows: List[List[str]]) -> None:
    if not rows:
        return
    ensure_dir(ledger_path)
    file_exists = os.path.exists(ledger_path)
    with open(ledger_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp_utc", "sender", "domain", "method", "detail", "subject"]) 
        for r in rows:
            writer.writerow(r)


def summarize_to_html(summary: Dict) -> str:
    def section(title: str, items: List[str]) -> str:
        if not items:
            return ""
        lis = "".join(f"<li>{x}</li>" for x in items)
        return f"<h3>{title}</h3><ul>{lis}</ul>"

    parts = [f"<h2>Inbox Cleaner Report {datetime.now().strftime('%Y-%m-%d %H:%M')}</h2>"]
    parts.append(f"<p>Dry-run: <b>{str(summary['dry_run']).lower()}</b></p>")
    parts.append(f"<p>Scanned: {summary['scanned']} | Candidates: {summary['candidates']} | Unsubscribed: {summary['unsubscribed']} | Archived: {summary['archived']}</p>")

    if summary.get("samples"):
        if summary["samples"].get("candidates"):
            parts.append(section("Candidates", summary["samples"]["candidates"]))
        if summary["samples"].get("unsubscribed"):
            parts.append(section("Unsubscribed", summary["samples"]["unsubscribed"]))
    return "".join(parts)


def run(config_path: str, apply_actions: bool) -> None:
    config = load_config(config_path)

    # Respect CLI flags overriding config
    dry_run = not apply_actions if apply_actions is not None else config["mode"].get("dry_run", True)

    service = get_gmail_service(
        credentials_path=config["oauth"]["credentials_file"],
        token_path=config["oauth"]["token_file"],
    )

    # Prepare labels
    label_names = config["labels"]
    id_to_unsub = ensure_label(service, label_names["to_unsubscribe"])  # noqa: F841
    id_unsubscribed = ensure_label(service, label_names["unsubscribed"])  # noqa: F841
    id_trash_staging = ensure_label(service, label_names["trash_staging"])  # noqa: F841
    id_keep = ensure_label(service, label_names["keep"])  # noqa: F841

    # Build search
    label_filter = ["INBOX"] if config["mode"].get("inbox_only", True) else []
    q_parts = []
    if config["mode"].get("unread_only", True):
        q_parts.append("is:unread")
    q = " ".join(q_parts)

    max_scan = int(config["limits"].get("max_messages_per_run", 500))
    label_ids = label_filter.copy()

    if config["filters"].get("gmail_category_label_ids"):
        label_ids.extend(config["filters"]["gmail_category_label_ids"])  # CATEGORY_PROMOTIONS, CATEGORY_UPDATES

    console.print(f"Scanning up to {max_scan} messages with labels {label_ids} and query '{q or '(none)'}'...")
    messages = list_messages(service, label_ids=label_ids, q=q, max_messages=max_scan)

    keep_substrings = config["keep"].get("senders_contains", [])

    counters = {
        "scanned": 0,
        "candidates": 0,
        "unsubscribed": 0,
        "archived": 0,
    }
    samples = {"candidates": [], "unsubscribed": []}
    ledger_rows: List[List[str]] = []

    max_unsubs = int(config["limits"].get("max_unsubscribes_per_run", 150))

    for idx, m in enumerate(messages):
        if counters["candidates"] >= max_unsubs and dry_run:
            # Do not stage too many in dry-run; allow review
            break

        try:
            meta = get_message_metadata(service, m["id"])
        except HttpError as e:
            console.print(f"Failed to get message {m['id']}: {e}")
            continue

        headers = meta.get("payload", {}).get("headers", [])
        from_value = find_header(headers, "From") or ""
        subject = find_header(headers, "Subject") or ""
        list_unsub = find_header(headers, "List-Unsubscribe")
        list_unsub_post = find_header(headers, "List-Unsubscribe-Post")

        counters["scanned"] += 1

        # Skip keep senders
        if sender_matches_keep(from_value, keep_substrings):
            continue

        # Require List-Unsubscribe when configured, typical for promos/newsletters
        if config["filters"].get("require_list_unsubscribe", True) and not list_unsub:
            continue

        urls, mailtos = parse_list_unsubscribe(list_unsub or "")
        if not urls and not mailtos:
            continue

        counters["candidates"] += 1
        if len(samples["candidates"]) < int(config["report"].get("max_samples_per_bucket", 20)):
            samples["candidates"].append(f"{from_value} — {subject}")

        if dry_run:
            # Stage with label only
            try:
                add_labels(service, meta["id"], add_label_ids=[ensure_label(service, label_names["to_unsubscribe"])])
            except Exception:
                pass
            continue

        # Apply mode: try to unsubscribe
        method, detail = try_unsubscribe(urls, mailtos, list_unsub_post)
        if method == "mailto":
            ok = send_unsubscribe_mail(service, detail, subject)
            if not ok:
                method = "none"
        if method != "none":
            counters["unsubscribed"] += 1
            if len(samples["unsubscribed"]) < int(config["report"].get("max_samples_per_bucket", 20)):
                samples["unsubscribed"].append(f"{from_value} — {subject} ({method})")
            # Record ledger
            if config["ledger"].get("enabled", True):
                ledger_rows.append([
                    now_iso(),
                    from_value,
                    get_sender_domain(from_value),
                    method,
                    detail,
                    subject,
                ])
            # Label and optionally archive
            try:
                add_labels(
                    service,
                    meta["id"],
                    add_label_ids=[ensure_label(service, label_names["unsubscribed"])],
                    remove_label_ids=["INBOX"] if config["actions"].get("archive_after_unsubscribe", True) else None,
                )
                if config["actions"].get("archive_after_unsubscribe", True):
                    counters["archived"] += 1
            except Exception:
                pass

        # Cap unsubscribes per run
        if counters["unsubscribed"] >= max_unsubs:
            break

    # Write ledger
    if ledger_rows and config["ledger"].get("enabled", True):
        write_ledger(config["ledger"]["path"], ledger_rows)

    # Prepare and send summary
    summary = {
        "dry_run": dry_run,
        "scanned": counters["scanned"],
        "candidates": counters["candidates"],
        "unsubscribed": counters["unsubscribed"],
        "archived": counters["archived"],
        "samples": samples,
    }

    table = Table(title="Inbox Cleaner Summary")
    table.add_column("Metric")
    table.add_column("Count")
    for k in ["scanned", "candidates", "unsubscribed", "archived"]:
        table.add_row(k, str(summary[k]))
    console.print(table)

    if config["report"].get("send_email_summary", True):
        recipient = str(config["report"].get("recipient", "me"))
        try:
            send_mail(
                service,
                to_addr=recipient,
                subject="Inbox Cleaner Report",
                html_body=summarize_to_html(summary),
            )
            console.print(f"Sent summary email to {recipient}")
        except Exception as e:
            console.print(f"Failed to send summary email: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gmail Inbox Cleaner Agent")
    parser.add_argument("--config", default=DEFAULT_CONFIG_PATH, help="Path to config.yaml")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", dest="apply", action="store_false", help="Do not execute actions; label only")
    mode.add_argument("--apply", dest="apply", action="store_true", help="Execute actions (unsubscribe, archive)")
    args = parser.parse_args()

    try:
        run(config_path=args.config, apply_actions=args.apply)
    except KeyboardInterrupt:
        console.print("Interrupted by user")
        sys.exit(1)