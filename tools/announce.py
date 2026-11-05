#!/usr/bin/env python3
"""Open an Outlook draft (default mail app if Outlook is not installed) announcing the next journal club
session to everyone who signed up, with all addresses in BCC.

Setup (once):
  1. Share the Google Sheet behind the sign-up form as "Anyone with the link
     can view", and note its ID from the URL
     (https://docs.google.com/spreadsheets/d/<ID>/edit).
  2. Save a config file at ~/.journal-club.json:
       {"sheet_id": "<ID>", "from_note": "Kimia",
        "to": ["kimia@example.org"], "cc": ["babak@example.org"],
        "recipients": ["someone@example.org"]}
     to / cc: who goes in the To and Cc fields. Everyone else (recipients
     listed by hand plus everyone who signed up through the form) goes in
     Bcc, so members never see each other's addresses.
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


ZOOM_URL = ""


def next_session():
    """First entry of `upcoming:` in the schedule YAML."""
    global ZOOM_URL
    text = fetch(DATA_URL)
    try:
        import yaml  # PyYAML, present on most Python installs
        data = yaml.safe_load(text)
        ZOOM_URL = (data.get("meeting") or {}).get("zoom_url") or ""
        up = data.get("upcoming") or []
        if not up:
            return None
        e = {k: (v.isoformat() if hasattr(v, "isoformat") else ("" if v is None else str(v).strip())) for k, v in up[0].items()}
        return e
    except ImportError:
        pass
    # Fallback: minimal parse of simple one-line fields.
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
    norm = lambda xs: [x.strip().lower() for x in xs if "@" in x]
    to, cc = norm(cfg.get("to", [])), norm(cfg.get("cc", []))
    emails = []
    for e in norm(cfg.get("recipients", [])) + emails_from_sheet(cfg["sheet_id"]):
        if e not in emails and e not in to and e not in cc:
            emails.append(e)
    s = next_session()
    if not s or not s.get("date"):
        sys.exit("No upcoming session in _data/journal_club.yml. Add one, push, and run again.")
    date = datetime.date.fromisoformat(s["date"]).strftime("%A, %B %-d")
    speaker = s.get("speaker") or "a speaker to be announced"
    title = s.get("title") or ""
    where = s.get("location") or "the Vector Institute"
    subject = f"Journal club next week: {speaker}" + (f' on "{title}"' if title else "")
    zoom = ZOOM_URL
    passcode = cfg.get("zoom_passcode", "")
    zoom_line = (f"Zoom: {zoom}" + (f" (passcode {passcode})" if passcode else "")) if zoom else ""
    body = f"Hi everyone,\n\nPlease join us next week for the Computer Vision Journal Club. {speaker} will be presenting"
    body += f' "{title}".\n\n' if title else ".\n\n"
    body += f"When: {date}, 1:00–2:00 PM\nWhere: {where}"
    body += ", or on Zoom\n" if zoom else "\n"
    if zoom_line:
        body += zoom_line + "\n"
    if s.get("paper"):
        body += f"Paper: {s['paper']}\n"
    if s.get("abstract"):
        body += f"\nAbstract:\n{s['abstract'].strip()}\n"
    body += (
        f"\nThe full schedule and past sessions are at {PAGE_URL}. "
        "If you'd like to present in a future week, just reply to this email.\n\n"
        f"See you there,\n{cfg.get('from_note', '')}\n"
    )
    if dry:
        print(f"TO: {', '.join(to)}\nCC: {', '.join(cc)}\nBCC ({len(emails)}): {', '.join(emails)}\n\nSubject: {subject}\n\n{body}")
        return
    if not (to or cc or emails):
        sys.exit("Nobody to send to: add `to`/`cc`/`recipients` in the config or wait for sign-ups.")
    q = urllib.parse.quote
    link = (f"mailto:{','.join(to)}?cc={','.join(cc)}&bcc={','.join(emails)}"
            f"&subject={q(subject)}&body={q(body)}")
    # Open in Outlook if it is installed, otherwise the default mail app.
    app = ["-a", "Microsoft Outlook"] if os.path.isdir("/Applications/Microsoft Outlook.app") else []
    subprocess.run(["open", *app, link], check=True)
    print(f"Opened an Outlook draft for {date}: To {len(to)}, Cc {len(cc)}, Bcc {len(emails)}. Check it and press Send.")


if __name__ == "__main__":
    main()
