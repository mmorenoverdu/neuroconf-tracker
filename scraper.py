#!/usr/bin/env python3
"""
NeuroConf Tracker — weekly scraper
Runs via GitHub Actions every Monday. Does two things:

  1. SYNC FROM EXCEL
     Reads Conference_tracker_list.xlsx from the repo.
     Any row whose short-name is not already in index.html is treated as a
     new conference. The scraper visits its website, extracts whatever dates /
     deadlines it can find, builds a new CONFERENCES entry, and inserts it.

  2. UPDATE EXISTING CONFERENCES
     For conferences already in the tracker that have a real website, the
     scraper visits the page and flags any text that looks new compared to
     the stored notes.

Both kinds of changes are written into the CHANGELOG block so visitors see
a "What changed this week" panel at the top of the page.

Usage:
    python scraper.py          # dry run — prints what would change, no write
    python scraper.py --write  # writes changes to index.html
"""

import re, sys, json, time, datetime
import urllib.request, urllib.error
from html.parser import HTMLParser

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("⚠  openpyxl not installed — Excel sync disabled. Run: pip install openpyxl")

WRITE        = "--write" in sys.argv
INDEX_PATH   = "index.html"
EXCEL_PATH   = "Conference_tracker_list.xlsx"

# ── HTML text extractor ───────────────────────────────────────────────────────
class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self._skip = False
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head", "nav", "footer"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head", "nav", "footer"):
            self._skip = False

    def handle_data(self, data):
        if not self._skip:
            s = data.strip()
            if s:
                self.parts.append(s)

    def get_text(self):
        return " ".join(self.parts)


def fetch_text(url: str, timeout: int = 25) -> str | None:
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; NeuroConfBot/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            html = resp.read().decode("utf-8", errors="replace")
        p = TextExtractor()
        p.feed(html)
        return p.get_text()
    except Exception as e:
        print(f"    ⚠ fetch failed: {e}")
        return None


# ── Date pattern helpers ──────────────────────────────────────────────────────
# Patterns that capture a date-like string near deadline/abstract/registration keywords
DATE_PATTERNS = [
    r"(?:abstract|poster|submission)\s*(?:deadline|due)[^\n]{0,60}?(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    r"(?:abstract|poster|submission)\s*(?:deadline|due)[^\n]{0,60}?([A-Za-z]+\s+\d{1,2},?\s*\d{4})",
    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})[^\n]{0,60}(?:abstract|deadline|submission)",
    r"([A-Za-z]+\s+\d{1,2},?\s*\d{4})[^\n]{0,60}(?:abstract|deadline|submission)",
    r"(?:early.bird|early registration)[^\n]{0,60}?(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    r"(?:early.bird|early registration)[^\n]{0,60}?([A-Za-z]+\s+\d{1,2},?\s*\d{4})",
    r"(?:registration)[^\n]{0,60}?(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
]

CONF_DATE_PATTERNS = [
    r"(?:conference|congress|meeting|forum|symposium)[^\n]{0,80}?(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    r"(\d{1,2}[–\-]\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    r"([A-Za-z]+\s+\d{1,2}[–\-]\d{1,2},?\s*\d{4})",
]

LOCATION_PATTERNS = [
    r"(?:will be held|taking place|venue|location)[^\n]{0,80}?in ([A-Z][A-Za-z\s,]+(?:USA|UK|Germany|France|Spain|Italy|Netherlands|Switzerland|Portugal|Korea|Japan|Australia|Canada|Brazil|Austria|Belgium|Sweden|Norway|Denmark|Finland|Poland|Czech|Hungary|Greece|Turkey|China|Singapore|India))",
    r"([A-Z][A-Za-z\s]+,\s*(?:USA|UK|Germany|France|Spain|Italy|Netherlands|Switzerland|Portugal|South Korea|Japan|Australia|Canada|Brazil|Austria|Belgium|Sweden|Norway|Denmark|Finland|Poland|Czech Republic|Hungary|Greece|Turkey|China|Singapore|India))",
]


def search_text(text: str, patterns: list) -> list:
    found = []
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            hit = (m.group(1) if m.lastindex else m.group(0)).strip()
            hit = re.sub(r"\s+", " ", hit)
            if hit and hit not in found and len(hit) > 3:
                found.append(hit)
    return found[:4]


def build_notes_from_hits(abstract_hits, location_hits, conf_date_hits, url):
    parts = []
    if abstract_hits:
        parts.append(f"Possible abstract deadline: {abstract_hits[0]}")
    if conf_date_hits:
        parts.append(f"Possible dates: {conf_date_hits[0]}")
    if location_hits:
        parts.append(f"Location found: {location_hits[0]}")
    if not parts:
        parts.append("No structured dates found — check website directly.")
    return ". ".join(parts) + f" Source: {url}"


# ── Read / write index.html ───────────────────────────────────────────────────
def read_index() -> str:
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return f.read()

def write_index(content: str):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        f.write(content)


# ── Extract existing conference IDs and short names from DATA block ───────────
def extract_existing(html: str) -> dict:
    """
    Returns a dict of everything already in the tracker, for duplicate detection.
    Keys: set of normalised strings (short name, full name words, website domain).
    Value: id.
    We match an Excel row as "existing" if ANY of these keys overlap.
    """
    result = {}   # normalised_key -> id
    # Extract each conference block
    for m in re.finditer(
        r'\{\s*id:\s*(\d+).*?short:"([^"]*)".*?name:"([^"]*)".*?website:"([^"]*)".*?changed:\s*(?:true|false)\s*\}',
        html, re.DOTALL
    ):
        cid     = int(m.group(1))
        short   = m.group(2).strip().lower()
        name    = m.group(3).strip().lower()
        website = m.group(4).strip().lower()

        # Store multiple keys all pointing to same id
        def add(k):
            k = re.sub(r"[^a-z0-9]", "", k)  # strip everything non-alphanumeric
            if k:
                result[k] = cid

        add(short)
        # Also add significant words from the full name (≥4 chars)
        for word in re.split(r'\W+', name):
            if len(word) >= 4:
                add(word)
        # Add website domain (most reliable unique key)
        domain = re.sub(r'^https?://(www\.)?', '', website).split('/')[0]
        add(domain)

    return result


def is_already_in(excel_name: str, excel_website: str, existing: dict) -> bool:
    """
    Returns True if this Excel row already exists in the tracker.
    Matches on: acronym/short name words, full name words, website domain.
    """
    def norm(s): return re.sub(r"[^a-z0-9]", "", s.lower())

    candidates = set()
    # All words from the Excel name
    for word in re.split(r'\W+', excel_name.lower()):
        w = norm(word)
        if len(w) >= 3:
            candidates.add(w)
    # Website domain
    domain = re.sub(r'^https?://(www\.)?', '', excel_website.lower()).split('/')[0]
    candidates.add(norm(domain))

    return any(c in existing for c in candidates)


def get_max_id(html: str) -> int:
    ids = re.findall(r'\{\s*id:\s*(\d+)', html)
    return max(int(i) for i in ids) if ids else 0


# ── Extract existing notes for change-detection ───────────────────────────────
def extract_notes(html: str) -> dict:
    """Returns {id: notes_string}"""
    result = {}
    for m in re.finditer(r'id:\s*(\d+).*?notes:"([^"]*)"', html, re.DOTALL):
        result[int(m.group(1))] = m.group(2)
    return result


# ── Read Excel sheet ──────────────────────────────────────────────────────────
def read_excel() -> list[dict]:
    """
    Returns list of {name, org, website} dicts from the spreadsheet.
    Expects columns: Conference, Organization, Website  (row 1 = header).
    """
    if not HAS_OPENPYXL:
        return []
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        # Detect header row
        header = [str(c).strip().lower() if c else "" for c in rows[0]]
        try:
            ci_name = header.index("conference")
            ci_org  = header.index("organization")
            ci_url  = header.index("website")
        except ValueError:
            # Fallback: assume columns 0, 1, 2
            ci_name, ci_org, ci_url = 0, 1, 2
        result = []
        for row in rows[1:]:
            if not row or not row[ci_name]:
                continue
            result.append({
                "name":    str(row[ci_name]).strip(),
                "org":     str(row[ci_org]).strip()  if row[ci_org]  else "",
                "website": str(row[ci_url]).strip()  if row[ci_url]  else "",
            })
        return result
    except Exception as e:
        print(f"  ⚠ Could not read Excel: {e}")
        return []


# ── Build a new JS conference object ─────────────────────────────────────────
def build_conf_js(cid: int, name: str, org: str, website: str,
                  notes: str, city: str = "TBD", country: str = "TBD") -> str:
    # Escape for JS string
    def esc(s): return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")
    short = name[:20].strip()
    return (
        f"  {{ id:{cid}, "
        f'name:"{esc(name)}", '
        f'short:"{esc(short)}", '
        f'org:"{esc(org)}", '
        f'city:"{esc(city)}", country:"{esc(country)}", '
        f'dates:{{start:"TBD",end:"TBD"}}, '
        f'abstractDeadline:"TBD", earlyReg:"TBD", finalReg:"TBD", '
        f'status:"tbd", website:"{esc(website)}", '
        f'notes:"{esc(notes)}", changed:true }}'
    )


# ── Insert a new conference into the DATA block ───────────────────────────────
def insert_conference(html: str, conf_js: str) -> str:
    """Appends the new conference before the closing ]; of the DATA block."""
    return re.sub(
        r'(// DATA_END)',
        f"{conf_js},\n];\n\\1",
        html,
        count=1,
        flags=re.DOTALL
    )


# ── Update CHANGELOG block ────────────────────────────────────────────────────
def update_changelog(html: str, changes: list[dict]) -> str:
    if not changes:
        return html
    today_str = datetime.date.today().isoformat()
    entry_js  = json.dumps(changes, ensure_ascii=False, indent=4)
    new_entry = f"  {{\n    date: \"{today_str}\",\n    changes: {entry_js}\n  }}"

    def replacer(m):
        existing = m.group(2).strip()
        inner = new_entry
        if existing:
            inner += f",\n  {existing}"
        return f"{m.group(1)}const CHANGELOG = [\n{inner}\n];\n{m.group(3)}"

    return re.sub(
        r'(// CHANGELOG_START\n)const CHANGELOG = \[([^\]]*)\];(\n// CHANGELOG_END)',
        replacer,
        html,
        flags=re.DOTALL
    )


# ── Update last-verified date ─────────────────────────────────────────────────
def update_last_verified(html: str) -> str:
    today_str = datetime.date.today().strftime("%-d %B %Y")
    return re.sub(r'(Last verified: )[^\s·]+\s+[^\s·]+\s+[^\s·]+', f'\\g<1>{today_str}', html)


# ── SCRAPE TARGETS for existing conferences ───────────────────────────────────
EXISTING_TARGETS = [
    { "id": 3,  "short": "WCNR 2026",   "url": "https://wfnr-congress.org/"                                                                     },
    { "id": 4,  "short": "SfN 2026",    "url": "https://www.sfn.org/meetings/neuroscience-2026"                                                  },
    { "id": 5,  "short": "FENS 2026",   "url": "https://fensforum.org/"                                                                          },
    { "id": 8,  "short": "BSC 2027",    "url": "https://www.brainstimjrnl.com/"                                                                  },
    { "id": 14, "short": "WPC 2026",    "url": "https://wpc2026.org/"                                                                            },
    { "id": 15, "short": "EAN 2026",    "url": "https://www.ean.org/congress2026"                                                                 },
    { "id": 16, "short": "MDS 2026",    "url": "https://www.mdscongress.org/"                                                                     },
    { "id": 19, "short": "ECNR 2026",   "url": "https://efnr.org/"                                                                               },
    { "id": 22, "short": "CSM 2027",    "url": "https://csm.apta.org/"                                                                           },
    { "id": 6,  "short": "RIO",         "url": "https://riogroup.weebly.com/"                                                                    },
    { "id": 7,  "short": "WSC",         "url": "https://worldstrokecongress.org/"                                                                 },
    { "id": 11, "short": "INPA 2026",   "url": "https://inpa.world/global-conference/"                                                           },
    { "id": 12, "short": "RASEN 2026",  "url": "https://reunion.sen.es/"                                                                         },
    { "id": 13, "short": "IAPRD 2026",  "url": "https://www.iaprd-world-congress.com/"                                                           },
    { "id": 17, "short": "EurPhysio",   "url": "https://www.physiocongress.eu/"                                                                   },
    { "id": 18, "short": "WorldPhysio", "url": "https://world.physio/congress"                                                                    },
    { "id": 23, "short": "AEF 2026",    "url": "https://congresoaef.com/"                                                                        },
]


# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("── NeuroConf Tracker weekly scraper ──\n")
    html    = read_index()
    changes = []

    # ── STEP 1: Excel sync ────────────────────────────────────────────────────
    print("STEP 1: Checking Excel for new conferences…")
    excel_rows = read_excel()
    if not excel_rows:
        print("  (No rows read from Excel — skipping sync)\n")
    else:
        existing = extract_existing(html)
        max_id   = get_max_id(html)
        added    = 0

        for row in excel_rows:
            name    = row["name"]
            org     = row["org"]
            website = row["website"]

            if is_already_in(name, website, existing):
                print(f"  ✓ Already tracked: {name}")
                continue

            print(f"  ✦ New conference found: {name}")
            time.sleep(1.5)
            text   = fetch_text(website) if website else None
            notes  = "No website provided." if not website else ""

            if text:
                abs_hits  = search_text(text, DATE_PATTERNS[:4])
                loc_hits  = search_text(text, LOCATION_PATTERNS)
                date_hits = search_text(text, CONF_DATE_PATTERNS)
                city = country = "TBD"
                if loc_hits:
                    # Try to split "City, Country"
                    parts = loc_hits[0].split(",")
                    if len(parts) >= 2:
                        city    = parts[0].strip()
                        country = parts[-1].strip()
                notes = build_notes_from_hits(abs_hits, loc_hits, date_hits, website)
            elif website:
                notes = f"Website unreachable during scrape — check manually: {website}"

            max_id   += 1
            conf_js   = build_conf_js(max_id, name, org, website, notes, city if text else "TBD", country if text else "TBD")
            html      = insert_conference(html, conf_js)
            # Register the new entry so subsequent rows don't double-insert
            new_key = re.sub(r"[^a-z0-9]", "", name.lower())
            existing[new_key] = max_id
            domain = re.sub(r'^https?://(www\.)?', '', website.lower()).split('/')[0]
            existing[re.sub(r"[^a-z0-9]", "", domain)] = max_id
            added += 1
            changes.append({
                "type": "new",
                "conf": name,
                "text": f"Added automatically from Excel. {notes[:120]}{'…' if len(notes)>120 else ''}"
            })

        print(f"  → {added} new conference(s) added.\n")

    # ── STEP 2: Scrape existing conferences for changes ───────────────────────
    print("STEP 2: Checking existing conferences for updates…")
    existing_notes = extract_notes(html)

    for target in EXISTING_TARGETS:
        cid   = target["id"]
        short = target["short"]
        url   = target["url"]
        print(f"  Checking {short}…")
        time.sleep(1.5)

        text = fetch_text(url)
        if not text:
            continue

        abs_hits  = search_text(text, DATE_PATTERNS[:4])
        date_hits = search_text(text, CONF_DATE_PATTERNS)
        all_hits  = abs_hits + date_hits

        old_notes = existing_notes.get(cid, "")
        new_info  = [h for h in all_hits if h not in old_notes and len(h) > 5]

        if new_info:
            change_text = f"Possible new info: {'; '.join(new_info[:3])}"
            print(f"    ✦ Flagged: {change_text}")
            changes.append({"type": "updated", "conf": short, "text": change_text})
            # Mark as changed
            html = re.sub(
                rf'(id:\s*{cid},.*?changed:\s*)false',
                r'\1true', html, count=1, flags=re.DOTALL
            )
        else:
            print(f"    ✓ No new info")

    # ── STEP 3: Write changes ─────────────────────────────────────────────────
    print()
    if changes:
        print(f"{len(changes)} change(s) detected.")
        html = update_changelog(html, changes)
        html = update_last_verified(html)
    else:
        print("✓ No changes detected — resetting any stale 'changed' flags.")
        html = re.sub(r'(changed:\s*)true', r'\1false', html)

    if WRITE:
        write_index(html)
        print("✅ index.html saved.")
    else:
        print("ℹ  Dry run — pass --write to save.")

    print("\nDone.")


if __name__ == "__main__":
    main()
