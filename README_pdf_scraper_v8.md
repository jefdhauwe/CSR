# PDF Scraper v8

Bulk PDF downloader for academic bibliography Excel files.  
Handles 15+ publishers, WHO IRIS (Cloudflare-protected), PubMed/PMC, and falls back to saving abstracts as `.txt` when no PDF is available.

---

## Features

| Feature | Detail |
|---|---|
| **15+ publisher handlers** | BioMedCentral, MDPI, Lancet, ScienceDirect, BMJ, Tandfonline, Oxford, PLOS, Wiley, Springer, WHO IRIS + generic |
| **WHO IRIS (Cloudflare)** | 4-strategy cascade: DSpace 7 API → OAI-PMH → session cookie → Playwright browser |
| **Parallel processing** | Configurable worker pool processes N Excel rows simultaneously |
| **Proxy rotation** | Free proxies, manual proxy list, or Tor exit-node rotation on 403/429 blocks |
| **Fallback chain** | Bibcite URL → PubMed URL → Abstract `.txt` |
| **Resume support** | Skip already-downloaded rows; configurable start position |
| **PDF validation** | Checks `%PDF` magic bytes, page count, and file size |
| **Duplicate detection** | MD5 hash comparison across all downloads |
| **Auto-save** | Excel output saved every N rows so progress is never lost |

---

## File Structure

```
pdf_scraper_v8.py       ← main script (edit CONFIG inside main())
journal_handler.py      ← journal-specific URL patterns (keep in same folder)
infontd_biblio.xlsx     ← your input Excel file
infontd_pdfs/           ← downloaded PDFs saved here (auto-created)
abstracts/              ← abstract .txt fallbacks saved here (auto-created)
infontd_processed_*.xlsx← output Excel with download status columns
```

---

## Installation

### Core dependencies
```bash
pip install requests beautifulsoup4 pandas openpyxl PyPDF2
```

### Playwright (required for WHO IRIS Cloudflare bypass)
```bash
pip install playwright
playwright install chromium
```

### Proxy rotation (optional)
```bash
pip install free-proxy       # free public proxy rotation
pip install stem             # Tor control (only if you run Tor locally)
```

---

## Quick Start

1. Put your Excel file in the same folder as the script.
2. Open `pdf_scraper_v8.py` and edit the `CONFIG` block inside `main()`.
3. Run:

```bash
python pdf_scraper_v8.py
```

### Test a single URL first
```bash
python pdf_scraper_v8.py --test-url "http://apps.who.int/iris/bitstream/10665/258983/1/9789241550116-eng.pdf"
```

Or set `"test_url"` in CONFIG and run normally.

---

## Configuration

All settings live in the `CONFIG` dict inside `main()` — no command-line flags needed. Just edit the values and save.

```python
CONFIG = {

    # ── Files ──────────────────────────────────────────────────────────────
    "input_file":           "infontd_biblio.xlsx",   # your input Excel
    "output_file":          None,                    # None = auto timestamped
    "download_dir":         "infontd_pdfs",          # where PDFs go
    "abstract_dir":         "abstracts",             # where .txt fallbacks go

    # ── Filtering ──────────────────────────────────────────────────────────
    "domain_filter":        "infontd",   # filter rows by 'Domains' column
    "type_filter":          "biblio",    # filter rows by 'Type' column (None = all)
    "start_from_row":       None,        # resume from row N (None = start from 1)

    # ── Performance ────────────────────────────────────────────────────────
    "workers":              5,           # parallel row workers
    "playwright_sessions":  2,           # max concurrent browser sessions
    "request_delay":        1.0,         # max random delay (seconds) per request
    "autosave_interval":    50,          # save Excel every N rows

    # ── API Keys ───────────────────────────────────────────────────────────
    "ncbi_api_key":         "YOUR_KEY",  # NCBI key for faster PubMed lookups

    # ── Proxy Settings ─────────────────────────────────────────────────────
    "use_free_proxies":     False,       # rotate free public proxies on blocks
    "manual_proxies":       [],          # your own proxies ["http://host:port"]
    "use_tor":              False,       # rotate Tor exit node on blocks
    "tor_socks_port":       9050,        # Tor SOCKS port
    "tor_control_port":     9051,        # Tor control port
    "tor_password":         None,        # Tor control password (if set)
    "proxy_max_failures":   3,           # failures before a proxy is blacklisted

    # ── Debug ──────────────────────────────────────────────────────────────
    "verbose":              False,       # detailed per-URL logs
    "test_url":             None,        # set a URL here to test it and exit
}
```

### Key settings explained

**`workers`** — How many Excel rows are processed at the same time. Each worker is a thread making network requests independently. Higher = faster, but also more memory and risk of rate-limiting. Good values:
- `1` — Sequential, safe for debugging or slow connections
- `5` — Default, good balance
- `10` — Fast machine with stable internet

**`request_delay`** — Each worker waits a random `0 – request_delay` seconds before each HTTP request. This keeps traffic polite and avoids triggering rate limits. Set to `0` for maximum speed (aggressive).

**`playwright_sessions`** — Playwright launches a real Chromium browser to solve Cloudflare challenges (used for WHO IRIS). Each session uses ~200MB RAM. Keep this at 2 unless you have a lot of RAM.

**`start_from_row`** — Useful when resuming an interrupted run. Set to the row number where you want to continue. Already-downloaded rows are always skipped regardless of this setting.

---

## Proxy Options

Proxies are used automatically when a request returns HTTP 403 or 429. The scraper rotates to a new proxy and retries up to 3 times.

### Option 1: Free public proxies
```python
"use_free_proxies": True
```
Requires `pip install free-proxy`. Fetches a random working proxy from a public list. Free proxies are unreliable but cost nothing and are good enough for occasional retries.

### Option 2: Your own proxy list
```python
"manual_proxies": [
    "http://proxy1.example.com:8080",
    "http://user:password@proxy2.example.com:3128",
]
```
These are tried first (round-robin) before falling back to free proxies. Use a VPN's SOCKS/HTTP proxy address here, or a paid proxy service.

### Option 3: Tor (advanced)
```python
"use_tor": True,
"tor_socks_port": 9050,
"tor_control_port": 9051,
"tor_password": None,
```
Requires Tor to be running locally and `pip install stem`. On each blocked request, sends a `NEWNYM` signal to get a new Tor exit IP.

> **Note:** Tor exit nodes are often pre-blocked by academic publishers. Most effective for WHO IRIS and government sites.

**Install Tor on Windows:** download from [torproject.org](https://www.torproject.org/download/)  
**Install Tor on Linux:** `sudo apt install tor && sudo systemctl start tor`

---

## Excel Input Format

The script expects these columns (names are case-sensitive):

| Column | Used for |
|---|---|
| `Node ID` | Used in output filenames (`nid_123_title.pdf`) |
| `Title` | Used in output filenames |
| `Domains` | Filtered by `domain_filter` setting |
| `Type` | Filtered by `type_filter` setting |
| `Bibcite URL` | Primary URL tried for PDF download |
| `PubMed URL` | Fallback URL (resolves PMC links) |
| `Abstract (English)` | Last-resort fallback; saved as `.txt` |

---

## Output Excel Columns

The script adds/updates these columns in the output file:

| Column | Values / Description |
|---|---|
| `download_status` | `success_bibcite` / `success_pubmed` / `success_abstract_txt` / `failed` / `skipped` / `duplicate` |
| `source_used` | Which source column worked |
| `url_used` | The actual URL that succeeded |
| `download_filename` | Filename only (`nid_123_My_Title.pdf`) |
| `download_filepath` | Full local path to the downloaded file |
| `download_error` | Short error for failed rows |
| `detailed_errors` | All errors from each source attempted |
| `handler_used` | Which publisher handler was used |
| `download_timestamp` | When the row was processed |

---

## WHO IRIS — How the 4 Strategies Work

WHO IRIS migrated to DSpace 7 and is behind Cloudflare, blocking plain `urllib`/`requests`. The scraper tries these in order:

1. **DSpace 7 REST API** — Calls `/server/api/core/handles/{id}` to get the item UUID, then walks bundles → bitstreams → content URL. Fastest, no browser needed.

2. **OAI-PMH XML harvest** — Calls the OAI endpoint (`/oai/request?verb=GetRecord`) which typically bypasses Cloudflare because it serves XML to data harvesters.

3. **Session cookie** — Visits the handle page first to get a Cloudflare clearance cookie, then downloads the canonical bitstream URL. Works when the challenge is mild.

4. **Playwright browser** — Launches a real Chromium browser that executes the Cloudflare JS challenge automatically. Falls back to visible (headed) mode if headless is detected. Last resort but most reliable.

---

## Troubleshooting

**`ModuleNotFoundError: No module named 'playwright'`**  
```bash
pip install playwright && playwright install chromium
```

**`❌ Input file not found`**  
Edit `CONFIG["input_file"]` to match your filename exactly, including `.xlsx`.

**WHO IRIS always fails even with Playwright**  
Visit `https://iris.who.int` in your browser once, then immediately run the script — the Cloudflare cookie may help. Alternatively enable `"use_tor": True` to route through a different IP.

**Very slow on large files**  
Increase `"workers"` to `8–10`. Lower `"request_delay"` to `0.5`. Don't set `"autosave_interval"` too low (each save rewrites the full Excel file).

**High memory usage**  
Lower `"playwright_sessions"` to `1` and `"workers"` to `3`.

**`py` command uses wrong Python / Playwright not found**  
Always use `python` (not `py`) inside your activated virtual environment:
```powershell
.\venv\Scripts\Activate.ps1
python pdf_scraper_v8.py
```
