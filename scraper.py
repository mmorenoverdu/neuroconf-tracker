#!/usr/bin/env python3
"""
NeuroConf Tracker — weekly scraper
Runs via GitHub Actions every Monday. Visits each conference website,
looks for date/deadline changes compared to what's currently in index.html,
updates the DATA block, and prepends a CHANGELOG entry.

Usage:
    python scraper.py          # dry run (prints changes, no write)
    python scraper.py --write  # write changes to index.html
"""

import re
import sys
import json
import datetime
import time
import urllib.request
import urllib.error
from html.parser import HTMLParser

WRITE = "--write" in sys.argv
INDEX_PATH = "index.html"

# ── Conferences to scrape ─────────────────────────────────────────────────────
# Each entry: id, short name, URL, and a list of regex patterns to search for
# dates/deadlines in the page text. The scraper does a best-effort text search;
# conference websites vary wildly so manual review is always recommended.
SCRAPE_TARGETS = [
    {
        "id": 1, "short": "NCM 2026",
        "url": "https://ncm-society.org/2026-meeting/",
        "patterns": [
            r"(?:abstract|poster)\s+deadline[:\s]+([A-Za-z]+ \d{1,2},?\s*\d{4})",
            r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})[^\n]*(?:abstract|poster|deadline)",
            r"(?:early.bird|early registration)[:\s]+([A-Za-z]+ \d{1,2},?\s*\d{4})",
        ],
    },
    {
        "id": 2, "short": "ESOC 2026",
        "url": "https://eso-stroke.org/esoc2026/",
        "patterns": [
            r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})[^\n]*(?:abstract|submission|deadline)",
            r"(?:abstract|submission)[^\n]{0,40}(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        ],
    },
    {
        "id": 3, "short": "WCNR 2026",
        "url": "https://wfnr-congress.org/",
        "patterns": [
            r"(?:abstract|deadline)[^\n]{0,60}(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
            r"([A-Za-z]+ \d{1,2},?\s*\d{4})[^\n]*(?:abstract|deadline|submission)",
        ],
    },
    {
        "id": 4, "short": "SfN 2026",
        "url": "https://www.sfn.org/meetings/neuroscience-2026",
        "patterns": [
            r"(?:abstract|submission)[^\n]{0,60}([A-Za-z]+ \d{1,2}[,\s]*\d{4})",
            r"(?:registration)[^\n]{0,60}([A-Za-z]+ \d{1,2}[,\s]*\d{4})",
        ],
    },
    {
        "id": 5, "short": "FENS 2026",
        "url": "https://fensforum.org/",
        "patterns": [
            r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})[^\n]*(?:abstract|registration|deadline)",
            r"(?:deadline|abstract)[^\n]{0,60}(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        ],
    },
    {
        "id": 8, "short": "BSC 2027",
        "url": "https://www.brainstimjrnl.com/",
        "patterns": [
            r"(?:abstract|poster|symposi)[^\n]{0,60}(\w+ \d{1,2},?\s*\d{4})",
            r"(?:deadline)[^\n]{0,60}(\w+ \d{1,2},?\s*\d{4})",
        ],
    },
    {
        "id": 15, "short": "EAN 2026",
        "url": "https://www.ean.org/congress2026",
        "patterns": [
            r"(?:abstract)[^\n]{0,60}(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
            r"(?:registration)[^\n]{0,60}(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        ],
    },
    {
        "id": 16, "short": "MDS 2026",
        "url": "https://www.mdscongress.org/",
        "patterns": [
            r"(?:abstract|submission)[^\n]{0,60}([A-Za-z]+ \d{1,2},?\s*\d{4})",
            r"October\s+\d{1,2}[–-]\d{1,2},?\s*\d{4}",
        ],
    },
    # For TBD conferences we still visit to detect when info appears
    {
        "id": 6,  "short": "RIO",        "url": "https://riogroup.weebly.com/",
        "patterns": [r"(\d{4})", r"(?:conference|meeting|congress)[^\n]{0,60}(\d{4})"],
    },
    {
        "id": 7,  "short": "WSC",        "url": "https://worldstrokecongress.org/",
        "patterns": [r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})", r"(?:deadline|abstract)[^\n]{0,60}(\d{4})"],
    },
    {
        "id": 12, "short": "RASEN 2026", "url": "https://reunion.sen.es/",
        "patterns": [r"(\d{4})", r"(?:abstract|inscripci|fecha)[^\n]{0,60}(\d{4})"],
    },
    {
        "id": 14, "short": "WPC 2026",   "url": "https://wpc2026.org/",
        "patterns": [r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})", r"(?:deadline|abstract)[^\n]{0,60}(\d{4})"],
    },
    {
        "id": 19, "short": "EFNR 2026",  "url": "https://efnr.org/",
        "patterns": [r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})", r"(?:congress|conference)[^\n]{0,80}(\d{4})"],
    },
    {
        "id": 20, "short": "INS 2026",   "url": "https://ins-congress.com/",
        "patterns": [r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})", r"(?:deadline|abstract)[^\n]{0,60}(\d{4})"],
    },
]

# ── HTML text extractor ──────────────────────────────────────────────────────
class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self._skip = False
        self.text_parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head"):
            self._skip = False

    def handle_data(self, data):
        if not self._skip:
            stripped = data.strip()
            if stripped:
                self.text_parts.append(stripped)

    def get_text(self):
        return " ".join(self.text_parts)


def fetch_text(url: str, timeout: int = 20) -> str | None:
    """Fetch URL and return visible text content."""
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; NeuroConfBot/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw_html = resp.read().decode("utf-8", errors="replace")
        parser = TextExtractor()
        parser.feed(raw_html)
        return parser.get_text()
    except Exception as e:
        print(f"  ⚠ Could not fetch {url}: {e}")
        return None


def search_patterns(text: str, patterns: list[str]) -> list[str]:
    """Return unique non-overlapping matches across all patterns."""
    found = []
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            hit = (m.group(1) if m.lastindex else m.group(0)).strip()
            if hit and hit not in found:
                found.append(hit)
    return found[:6]  # cap at 6 to keep notes short


# ── Read current index.html ───────────────────────────────────────────────────
def read_index() -> str:
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return f.read()


def write_index(content: str):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        f.write(content)


# ── Extract current notes from DATA block ─────────────────────────────────────
def extract_current_notes(html: str) -> dict[int, dict]:
    """Return {id: {notes, abstractDeadline, earlyReg, status, location}} from current DATA block."""
    result = {}
    # Match each conference object roughly
    conf_blocks = re.findall(
        r'\{\s*id:\s*(\d+).*?changed:\s*(?:true|false)\s*\}',
        html, re.DOTALL
    )
    for block in conf_blocks:
        # id
        id_m = re.search(r'id:\s*(\d+)', block)
        if not id_m:
            continue
        cid = int(id_m.group(1))
        def extract_field(b, field):
            m = re.search(rf'{field}:\s*"([^"]*)"', b)
            return m.group(1) if m else ""
        result[cid] = {
            "notes":            extract_field(block, "notes"),
            "abstractDeadline": extract_field(block, "abstractDeadline"),
            "earlyReg":         extract_field(block, "earlyReg"),
            "status":           extract_field(block, "status"),
            "location":         extract_field(block, "location"),
        }
    return result


# ── Build changelog entry ─────────────────────────────────────────────────────
def build_changelog_entry(changes: list[dict]) -> str:
    today = datetime.date.today().isoformat()
    items_js = json.dumps(changes, ensure_ascii=False, indent=4)
    return f"""  {{
    date: "{today}",
    changes: {items_js}
  }}"""


# ── Main scrape + patch ───────────────────────────────────────────────────────
def main():
    print("── NeuroConf Tracker weekly scraper ──\n")
    html = read_index()
    current_data = extract_current_notes(html)
    changes = []

    for target in SCRAPE_TARGETS:
        cid   = target["id"]
        short = target["short"]
        url   = target["url"]
        print(f"Checking {short} ({url})")
        time.sleep(1.5)  # polite delay

        text = fetch_text(url)
        if not text:
            continue

        hits = search_patterns(text, target["patterns"])
        if not hits:
            print(f"  → No pattern matches found")
            continue

        print(f"  → Found: {hits[:3]}")

        # Compare to existing notes
        old = current_data.get(cid, {})
        old_notes = old.get("notes", "")
        # Check if any hit is NOT already in the existing notes
        new_info = [h for h in hits if h not in old_notes and len(h) > 4]

        if new_info:
            change_text = f"Possible new info detected: {'; '.join(new_info[:3])}"
            changes.append({
                "type":  "updated",
                "conf":  short,
                "text":  change_text,
            })
            print(f"  ✦ Change flagged: {change_text}")

            # Mark conference as changed in html
            # Locate the conference block by id and set changed:false → changed:true
            html = re.sub(
                rf'(id:\s*{cid},.*?changed:\s*)false',
                r'\1true',
                html, count=1, flags=re.DOTALL
            )
        else:
            print(f"  ✓ No new info vs current notes")

    # ── Update CHANGELOG block ─────────────────────────────────────────────
    if changes:
        print(f"\n{len(changes)} change(s) detected.")
        new_entry = build_changelog_entry(changes)

        # Replace CHANGELOG block
        html = re.sub(
            r'(// CHANGELOG_START\n)const CHANGELOG = \[([^\]]*)\];(\n// CHANGELOG_END)',
            lambda m: (
                f"{m.group(1)}const CHANGELOG = [\n"
                f"{new_entry}"
                + (f",\n{m.group(2).strip()}" if m.group(2).strip() else "")
                + f"\n];\n{m.group(3)}"
            ),
            html, flags=re.DOTALL
        )

        # Reset all changed:true back to false first, then set new ones
        # (already done above per conference)

        # Update last-verified date
        today_str = datetime.date.today().strftime("%-d %B %Y")
        html = re.sub(
            r'(Last verified: )[^·]+',
            f'\\1{today_str} ',
            html
        )

        if WRITE:
            write_index(html)
            print(f"✅ index.html updated and saved.")
        else:
            print("ℹ  Dry run — pass --write to save changes.")
    else:
        print("\n✓ No changes detected this week.")

        # Still reset all changed flags (so old highlights don't persist)
        if WRITE:
            html = re.sub(r'(changed:\s*)true', r'\1false', html)
            write_index(html)
            print("✅ index.html updated (changed flags reset).")

    print("\nDone.")


if __name__ == "__main__":
    main()
