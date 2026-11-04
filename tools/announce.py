#!/usr/bin/env python3
"""Open an Outlook draft (default mail app if Outlook is not installed) announcing the next journal club
session to everyone who signed up, with all addresses in BCC.

Setup (once):
  1. Share the Google Sheet behind the sign-up form as "Anyone with the link
     can view", and note its ID from the URL
     (https://docs.google.com/spreadsheets/d/<ID>/edit).
  2. Save a config file at ~/.journal-club.json:
       {"sheet_id": "<ID>", "from_note": "Kimia"}
     Keep this file on your laptop only; it is not in the repo.

Run:  python3 tools/announce.py            (opens a draft for the next session)
      python3 tools/announce.py --dry-run  (prints what it would send)
"""
import csv, io, json, os, re, sys, urllib.parse, urllib.request, datetime, subprocess

CONFIG = os.path.expanduser("~/.journal-club.json")
DATA_URL = "https://raw.githubusercontent.com/TaatiTeam/journal-club/main/_data/journal_club.yml"
PAGE_URL = "https://taatiteam.github.io/journal-club/"


def load_config():
    try:
        with open(CONFIG) as f:
            return json.load(f)
    except FileNotFoundError:
        sys.exit(f"Missing {CONFIG}. See the top of this script for setup.")


def fetch(url):
    with urllib.request.urlopen(url, timeout=30) as r:
        return r.read().decode("utf-8")


def emails_from_sheet(sheet_id):
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    rows = list(csv.reader(io.StringIO(fetch(csv_url))))
    if not rows:
        return []
    header = [h.strip().lower() for h in rows[0]]
    col = next((i for i, h in enumerate(header) if "email" in h), None)
    if col is None:
        sys.exit(f"No column with 'email' in the sheet header: {rows[0]}")
    seen, out = set(), []
    for row in rows[1:]:
        if col < len(row):
            e = row[col].strip().lower()
            if "@" in e and e not in seen:
                seen.add(e); out.append(e)
    return out


def next_session():
    """Minimal parse of the `upcoming:` list without needing PyYAML."""
    text = fetch(DATA_URL)
    m = re.search(r"^upcoming:\s*\n((?:  .*\n|\s*\n)*?)(?=^\S|\Z)", text, re.M)
    if not m or not m.group(1).strip():
        return None
    block = m.group(1)
    first = re.split(r"^  - ", block, flags=re.M)[1:]
    if not first:
        return None
    entry = {}
    for line in ("  - " + first[0]).splitlines():
        kv = re.match(r"\s*-?\s*(\w+):\s*(.*?)\s*(?:#.*)?$", line)
        if kv:
            entry[kv.group(1)] = kv.group(2).strip().strip('"')
    return entry


def main():
    dry = "--dry-run" in sys.argv
    cfg = load_config()
    emails = emails_from_sheet(cfg["sheet_id"])
    s = next_session()
    if not s or not s.get("date"):
        sys.exit("No upcoming session in _data/journal_club.yml. Add one, push, and run again.")
    date = datetime.date.fromisoformat(s["date"]).strftime("%A, %B %-d")
    title = s.get("title") or "title to be announced"
    subject = f"Journal club {date}: {s.get('speaker', 'TBD')}"
    body = (
        f"Hi all,\n\nThis week's Computer Vision Journal Club:\n\n"
        f"  {date}, 1:00–2:00 PM, Vector Institute\n"
        f"  Speaker: {s.get('speaker', 'TBD')}\n"
        f"  Title: {title}\n"
        + (f"  Paper: {s['paper']}\n" if s.get("paper") else "")
        + (f"\n{s['abstract']}\n" if s.get("abstract") else "")
        + f"\nZoom link and schedule: {PAGE_URL}\n\nSee you there,\n{cfg.get('from_note', '')}\n"
    )
    if dry:
        print(f"BCC ({len(emails)}): {', '.join(emails)}\n\nSubject: {subject}\n\n{body}")
        return
    if not emails:
        sys.exit("The sheet has no email addresses yet.")
    q = urllib.parse.quote
    link = f"mailto:?bcc={','.join(emails)}&subject={q(subject)}&body={q(body)}"
    # Open in Outlook if it is installed, otherwise the default mail app.
    app = ["-a", "Microsoft Outlook"] if os.path.isdir("/Applications/Microsoft Outlook.app") else []
    subprocess.run(["open", *app, link], check=True)
    print(f"Opened an Outlook draft to {len(emails)} people for {date}. Check it and press Send.")


if __name__ == "__main__":
    main()
