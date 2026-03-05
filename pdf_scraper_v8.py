"""
PDF Scraper v8 - Parallel + Playwright + Proxy Rotation
=========================================================

What's new vs v7:
  - Free proxy rotation via the `free-proxy` library (optional)
  - Tor support for exit-node rotation (optional, requires Tor + stem)
  - ProxyManager rotates proxies on 403/429/Cloudflare responses
  - All settings are configured in the CONFIG block inside main()
  - No mandatory CLI arguments — just edit the CONFIG dict and run

Usage:
    python pdf_scraper_v8.py                  # full run with CONFIG settings
    python pdf_scraper_v8.py --test-url URL   # single URL test

Dependencies (core):
    pip install requests beautifulsoup4 pandas openpyxl PyPDF2 playwright
    playwright install chromium

Dependencies (optional proxy):
    pip install free-proxy          # free proxy rotation
    pip install stem                # Tor control (needs Tor running locally)
"""

import os, re, time, random, json, hashlib, sys, argparse, tempfile, threading
import urllib.request, urllib.error, urllib.parse, http.cookiejar
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, unquote, quote

import requests
from bs4 import BeautifulSoup
import pandas as pd
import PyPDF2

try:
    from journal_handler import JournalSpecificHandler
    JOURNAL_HANDLER_AVAILABLE = True
except Exception:
    JOURNAL_HANDLER_AVAILABLE = False
    print("⚠️  journal_handler.py not found — some journal-specific patterns won't work.")

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

try:
    from fp.fp import FreeProxy
    FREE_PROXY_AVAILABLE = True
except ImportError:
    FREE_PROXY_AVAILABLE = False

try:
    from stem import Signal
    from stem.control import Controller as TorController
    TOR_AVAILABLE = True
except ImportError:
    TOR_AVAILABLE = False

# Limit concurrent Playwright browser sessions to avoid OOM
_PLAYWRIGHT_SEMAPHORE = threading.Semaphore(2)
# Single lock for console output so lines don't interleave
_PRINT_LOCK = threading.Lock()

# User-agent pool for rotation on retries
_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0",
]


def tprint(*args, **kwargs):
    """Thread-safe print."""
    with _PRINT_LOCK:
        print(*args, **kwargs)


# ============================================================================
# Proxy Manager
# ============================================================================

class ProxyManager:
    """
    Rotates proxies on HTTP 403 / 429 / Cloudflare responses.

    Priority order:
      1. Manual proxy list (always tried first if provided)
      2. Free proxy rotation via `free-proxy` library
      3. Tor exit-node rotation via stem (if Tor is running locally)
      4. Direct connection (no proxy)

    Usage:
        pm = ProxyManager(use_free_proxies=True, use_tor=False)
        proxy = pm.get_proxy()          # returns {"http": ..., "https": ...} or None
        pm.report_failure(proxy_url)    # blacklists a bad proxy
        pm.rotate()                     # force get a fresh proxy (Tor: sends NEWNYM)
    """

    def __init__(self,
                 manual_proxies=None,
                 use_free_proxies=True,
                 use_tor=False,
                 tor_host="127.0.0.1",
                 tor_socks_port=9050,
                 tor_control_port=9051,
                 tor_password=None,
                 max_failures=3,
                 verbose=False):

        self.manual_proxies    = list(manual_proxies or [])
        self.use_free_proxies  = use_free_proxies and FREE_PROXY_AVAILABLE
        self.use_tor           = use_tor and TOR_AVAILABLE
        self.tor_host          = tor_host
        self.tor_socks_port    = tor_socks_port
        self.tor_control_port  = tor_control_port
        self.tor_password      = tor_password
        self.max_failures      = max_failures
        self.verbose           = verbose

        self._lock             = threading.Lock()
        self._blacklist        = {}   # proxy_url -> failure_count
        self._current_free     = None
        self._manual_idx       = 0

        if self.use_tor:
            self._tor_proxy = f"socks5://{tor_host}:{tor_socks_port}"
        else:
            self._tor_proxy = None

        if self.use_free_proxies and not FREE_PROXY_AVAILABLE:
            tprint("⚠️  ProxyManager: free-proxy not installed. Run: pip install free-proxy")
        if self.use_tor and not TOR_AVAILABLE:
            tprint("⚠️  ProxyManager: stem not installed. Run: pip install stem")

    def _proxy_dict(self, url):
        return {"http": url, "https": url} if url else None

    def _is_blacklisted(self, url):
        return self._blacklist.get(url, 0) >= self.max_failures

    def get_proxy(self):
        """
        Return a proxy dict or None (direct connection).
        Tries manual → free → Tor in order.
        """
        with self._lock:
            # 1. Manual proxies (round-robin, skip blacklisted)
            for _ in range(len(self.manual_proxies)):
                url = self.manual_proxies[self._manual_idx % len(self.manual_proxies)]
                self._manual_idx += 1
                if not self._is_blacklisted(url):
                    return self._proxy_dict(url)

            # 2. Free proxy
            if self.use_free_proxies:
                proxy_url = self._fetch_free_proxy()
                if proxy_url and not self._is_blacklisted(proxy_url):
                    self._current_free = proxy_url
                    return self._proxy_dict(proxy_url)

            # 3. Tor
            if self.use_tor and self._tor_proxy and not self._is_blacklisted(self._tor_proxy):
                return self._proxy_dict(self._tor_proxy)

        return None  # direct connection

    def _fetch_free_proxy(self):
        try:
            return FreeProxy(rand=True, timeout=1, https=True).get()
        except Exception:
            try:
                return FreeProxy(rand=True, timeout=2, https=False).get()
            except Exception:
                return None

    def report_failure(self, proxy_dict):
        """Mark a proxy as failed. After max_failures it is blacklisted."""
        if not proxy_dict:
            return
        url = proxy_dict.get("https") or proxy_dict.get("http")
        if not url:
            return
        with self._lock:
            self._blacklist[url] = self._blacklist.get(url, 0) + 1
            if self._blacklist[url] >= self.max_failures and self.verbose:
                tprint(f"  🚫 Proxy blacklisted after {self.max_failures} failures: {url}")

    def rotate(self):
        """
        Force a new proxy.  For Tor this sends a NEWNYM signal to get a new exit node.
        """
        if self.use_tor:
            try:
                with TorController.from_port(port=self.tor_control_port) as ctrl:
                    if self.tor_password:
                        ctrl.authenticate(password=self.tor_password)
                    else:
                        ctrl.authenticate()
                    ctrl.signal(Signal.NEWNYM)
                    time.sleep(1)   # give Tor time to establish new circuit
                    if self.verbose:
                        tprint("  🔄 Tor: new exit node requested")
            except Exception as e:
                if self.verbose:
                    tprint(f"  ⚠️ Tor rotate failed: {e}")
        elif self.use_free_proxies:
            with self._lock:
                self._current_free = self._fetch_free_proxy()

    def apply_to_session(self, session, proxy_dict):
        """Apply a proxy dict to a requests Session."""
        if proxy_dict:
            session.proxies.update(proxy_dict)
        else:
            session.proxies = {}

    @property
    def enabled(self):
        return bool(self.manual_proxies or self.use_free_proxies or self.use_tor)

    def status(self):
        parts = []
        if self.manual_proxies:
            parts.append(f"{len(self.manual_proxies)} manual proxies")
        if self.use_free_proxies:
            parts.append("free-proxy rotation")
        if self.use_tor:
            parts.append("Tor exit-node rotation")
        return ", ".join(parts) if parts else "disabled (direct connection)"


# ============================================================================
# WHO IRIS Downloader  (urllib strategies + Playwright fallback)
# ============================================================================

class WhoIrisDownloader:
    """
    Self-contained WHO IRIS PDF downloader.

    Strategy order (fastest/lightest first):
      1. DSpace 7 REST API  (/server/api/core/handles/…)
      2. OAI-PMH XML harvest
      3. Canonical bitstream URL with session cookie
      4. Playwright (solves Cloudflare, used as last resort)
    Returns bytes on success, None on failure.
    """

    IRIS_BASE   = "https://iris.who.int"
    DSPACE7_API = "https://iris.who.int/server/api"

    UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
          "AppleWebKit/537.36 (KHTML, like Gecko) "
          "Chrome/124.0.0.0 Safari/537.36")

    BASE_HEADERS = {
        "User-Agent": UA,
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
    }

    def __init__(self, verbose=False, proxy_manager=None):
        self.verbose = verbose
        self.proxy_manager = proxy_manager
        jar    = http.cookiejar.CookieJar()
        self._opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(jar),
            urllib.request.HTTPRedirectHandler(),
        )

    def log(self, msg):
        if self.verbose:
            tprint(msg)

    # ── helpers ─────────────────────────────────────────────────────────────

    @staticmethod
    def parse_url(url):
        url = url.lstrip("-").strip()
        m = re.search(r"(?:handle/|bitstream(?:/handle)?/)(\d+/\d+)", url)
        handle = m.group(1) if m else None
        m2 = re.search(r"/([^/?#]+\.pdf)", url, re.I)
        filename = m2.group(1) if m2 else None
        return handle, filename

    def _get(self, url, extra=None, timeout=25):
        """GET via urllib opener (keeps session cookies). Proxy injected if available."""
        h = {**self.BASE_HEADERS, **(extra or {}),
             "Accept": (extra or {}).get("Accept", "text/html,*/*")}

        # Inject proxy into opener if proxy_manager provides one
        if self.proxy_manager and self.proxy_manager.enabled:
            proxy = self.proxy_manager.get_proxy()
            if proxy:
                proxy_url = proxy.get("https") or proxy.get("http", "")
                proxy_handler = urllib.request.ProxyHandler({
                    "http": proxy_url, "https": proxy_url
                })
                jar = http.cookiejar.CookieJar()
                self._opener = urllib.request.build_opener(
                    proxy_handler,
                    urllib.request.HTTPCookieProcessor(jar),
                    urllib.request.HTTPRedirectHandler(),
                )

        req = urllib.request.Request(url, headers=h)
        with self._opener.open(req, timeout=timeout) as r:
            return r.geturl(), dict(r.headers), r.read()

    def _is_pdf(self, headers, data):
        ct = headers.get("Content-Type", headers.get("content-type", "")).lower()
        return "pdf" in ct or data[:4] == b"%PDF"

    # ── Strategy 1: DSpace 7 REST API ────────────────────────────────────────

    def _strategy_dspace7(self, handle):
        self.log("    [IRIS S1] DSpace 7 REST API")
        prefix, suffix = handle.split("/", 1)
        try:
            _, _, body = self._get(
                f"{self.DSPACE7_API}/core/handles/{prefix}/{suffix}",
                {"Accept": "application/json"}
            )
            data = json.loads(body)
            item_href = (data.get("_links", {}).get("resource", {}).get("href")
                         or data.get("_links", {}).get("self", {}).get("href"))
            if not item_href:
                return None

            _, _, body = self._get(item_href.rstrip("/") + "/bundles",
                                   {"Accept": "application/json"})
            bundles = json.loads(body).get("_embedded", {}).get("bundles", [])

            for bundle in bundles:
                if bundle.get("name") != "ORIGINAL":
                    continue
                bs_href = bundle.get("_links", {}).get("bitstreams", {}).get("href")
                if not bs_href:
                    continue
                _, _, body = self._get(bs_href, {"Accept": "application/json"})
                for bs in json.loads(body).get("_embedded", {}).get("bitstreams", []):
                    content_url = bs.get("_links", {}).get("content", {}).get("href")
                    if not content_url:
                        continue
                    self.log(f"      → {content_url[:70]}")
                    try:
                        _, hdrs, pdf = self._get(
                            content_url,
                            {"Accept": "application/pdf,*/*",
                             "Referer": f"{self.IRIS_BASE}/handle/{handle}"}
                        )
                        if self._is_pdf(hdrs, pdf[:8]):
                            self.log(f"      ✓ DSpace7 API ({len(pdf)//1024} KB)")
                            return pdf
                    except Exception as e:
                        self.log(f"      ✗ {e}")
        except Exception as e:
            self.log(f"      ✗ DSpace7 API failed: {e}")
        return None

    # ── Strategy 2: OAI-PMH ──────────────────────────────────────────────────

    def _strategy_oai(self, handle):
        self.log("    [IRIS S2] OAI-PMH")
        oai_id = f"oai:iris.who.int:{handle}"
        for prefix in ["ore", "oai_dc"]:
            url = (f"{self.IRIS_BASE}/oai/request?verb=GetRecord"
                   f"&identifier={quote(oai_id)}&metadataPrefix={prefix}")
            try:
                _, _, body = self._get(url, {"Accept": "text/xml,*/*"})
                xml = body.decode("utf-8", errors="replace")
                links = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', xml, re.I)
                links += re.findall(r'<dc:identifier>([^<]+\.pdf[^<]*)</dc:identifier>', xml, re.I)
                for link in links:
                    full = link if link.startswith("http") else self.IRIS_BASE + link
                    try:
                        _, hdrs, pdf = self._get(full, {"Accept": "application/pdf,*/*"})
                        if self._is_pdf(hdrs, pdf[:8]):
                            self.log(f"      ✓ OAI-PMH ({len(pdf)//1024} KB)")
                            return pdf
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"      ✗ OAI {prefix}: {e}")
        return None

    # ── Strategy 3: Session cookie + canonical bitstream URL ─────────────────

    def _strategy_session(self, handle, filename):
        self.log("    [IRIS S3] Session cookie + bitstream URL")
        handle_url = f"{self.IRIS_BASE}/handle/{handle}"
        try:
            self._get(handle_url)   # establishes session cookie
        except Exception as e:
            self.log(f"      ✗ Could not load handle page: {e}")
            return None

        time.sleep(1.2)

        candidates = []
        if filename:
            candidates += [
                f"{self.IRIS_BASE}/bitstream/handle/{handle}/{filename}?sequence=1&isAllowed=y",
                f"{self.IRIS_BASE}/bitstream/handle/{handle}/{filename}?sequence=1",
            ]
        candidates.append(
            f"{self.IRIS_BASE}/bitstream/handle/{handle}?sequence=1&isAllowed=y"
        )

        for url in candidates:
            try:
                _, hdrs, pdf = self._get(
                    url,
                    {"Accept": "application/pdf,*/*", "Referer": handle_url}
                )
                if self._is_pdf(hdrs, pdf[:8]):
                    self.log(f"      ✓ Session bitstream ({len(pdf)//1024} KB)")
                    return pdf
            except Exception as e:
                self.log(f"      ✗ {url[:60]}: {e}")
        return None

    # ── Strategy 4: Playwright ────────────────────────────────────────────────

    def _strategy_playwright(self, handle, filename):
        if not PLAYWRIGHT_AVAILABLE:
            self.log("    [IRIS S4] Playwright not installed — skipping")
            return None
        self.log("    [IRIS S4] Playwright (acquires semaphore)")

        _PLAYWRIGHT_SEMAPHORE.acquire()
        try:
            return self._run_playwright(handle, filename)
        finally:
            _PLAYWRIGHT_SEMAPHORE.release()

    def _run_playwright(self, handle, filename, headless=True):
        handle_url = f"{self.IRIS_BASE}/handle/{handle}"
        intercepted = {}

        try:
            with sync_playwright() as pw:
                browser = pw.chromium.launch(
                    headless=headless,
                    args=["--disable-blink-features=AutomationControlled",
                          "--no-sandbox", "--disable-dev-shm-usage"],
                )
                context = browser.new_context(
                    user_agent=self.UA,
                    accept_downloads=True,
                    viewport={"width": 1280, "height": 800},
                    locale="en-US",
                )

                def on_response(response):
                    ct = response.headers.get("content-type", "")
                    if "pdf" in ct.lower() and response.status == 200:
                        try:
                            intercepted["body"] = response.body()
                            intercepted["url"]  = response.url
                        except Exception:
                            pass

                page = context.new_page()
                page.on("response", on_response)

                # Load handle page (solves Cloudflare)
                try:
                    page.goto(handle_url, wait_until="networkidle", timeout=60_000)
                except Exception:
                    try:
                        page.goto(handle_url, wait_until="domcontentloaded", timeout=30_000)
                        page.wait_for_timeout(5_000)
                    except Exception as e:
                        self.log(f"      ✗ Playwright handle page failed: {e}")
                        browser.close()
                        return None

                # Cloudflare headless detection → retry headed
                if headless and ("Just a moment" in page.title()
                                 or "cf-browser-verification" in page.content()):
                    self.log("      ⚠ Cloudflare challenge detected, retrying headed")
                    browser.close()
                    return self._run_playwright(handle, filename, headless=False)

                if "body" in intercepted:
                    self.log(f"      ✓ Playwright intercepted on load ({len(intercepted['body'])//1024} KB)")
                    browser.close()
                    return intercepted["body"]

                # Try clicking PDF link
                for sel in ["a[href*='.pdf']", "a[href*='bitstream']",
                            "a:has-text('PDF')", "a:has-text('Download')"]:
                    try:
                        el = page.query_selector(sel)
                        if el:
                            href = el.get_attribute("href") or ""
                            pdf_link = href if href.startswith("http") else self.IRIS_BASE + href
                            intercepted.clear()
                            try:
                                with page.expect_download(timeout=30_000) as dl:
                                    page.goto(pdf_link, wait_until="commit", timeout=30_000)
                                tmp = tempfile.mktemp(suffix=".pdf")
                                dl.value.save_as(tmp)
                                data = Path(tmp).read_bytes()
                                os.unlink(tmp)
                                if data[:4] == b"%PDF":
                                    self.log(f"      ✓ Playwright download ({len(data)//1024} KB)")
                                    browser.close()
                                    return data
                            except Exception:
                                page.wait_for_timeout(2_000)
                                if "body" in intercepted:
                                    data = intercepted["body"]
                                    self.log(f"      ✓ Playwright intercepted ({len(data)//1024} KB)")
                                    browser.close()
                                    return data
                            break
                    except Exception:
                        pass

                # Direct bitstream navigation
                if filename:
                    for seq in ["1", "2"]:
                        bs_url = (f"{self.IRIS_BASE}/bitstream/handle/{handle}/"
                                  f"{filename}?sequence={seq}&isAllowed=y")
                        intercepted.clear()
                        try:
                            with page.expect_download(timeout=25_000) as dl:
                                page.goto(bs_url, wait_until="commit", timeout=25_000)
                            tmp = tempfile.mktemp(suffix=".pdf")
                            dl.value.save_as(tmp)
                            data = Path(tmp).read_bytes()
                            os.unlink(tmp)
                            if data[:4] == b"%PDF":
                                self.log(f"      ✓ Playwright bitstream ({len(data)//1024} KB)")
                                browser.close()
                                return data
                        except Exception:
                            page.wait_for_timeout(2_000)
                            if "body" in intercepted:
                                data = intercepted["body"]
                                self.log(f"      ✓ Playwright intercepted bitstream ({len(data)//1024} KB)")
                                browser.close()
                                return data

                browser.close()
        except Exception as e:
            self.log(f"      ✗ Playwright error: {e}")
        return None

    # ── Public entry point ────────────────────────────────────────────────────

    def download(self, url):
        """Run all strategies in order. Returns PDF bytes or None."""
        handle, filename = self.parse_url(url)
        if not handle:
            self.log(f"    ✗ Cannot extract handle from: {url}")
            return None

        for strategy in [
            lambda: self._strategy_dspace7(handle),
            lambda: self._strategy_oai(handle),
            lambda: self._strategy_session(handle, filename),
            lambda: self._strategy_playwright(handle, filename),
        ]:
            try:
                result = strategy()
            except Exception as e:
                self.log(f"    ✗ Strategy exception: {e}")
                result = None
            if result:
                return result
        return None


# ============================================================================
# Generic Playwright Downloader
# ============================================================================

class PlaywrightDownloader:
    """
    Generic Playwright-based PDF downloader.

    Used as a last-resort fallback for ANY URL after requests-based strategies
    fail (e.g. JS-rendered pages, cookie walls, Cloudflare, login redirects).

    Three capture methods, tried in order:
      A. Response interception  — grabs PDF bytes the moment the browser
                                  receives a content-type:application/pdf response
      B. expect_download()      — catches the file when the server sends
                                  Content-Disposition: attachment
      C. PDF link click         — parses the loaded page for PDF/download links
                                  and navigates to them

    Returns PDF bytes on success, None on failure.
    """

    UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
          "AppleWebKit/537.36 (KHTML, like Gecko) "
          "Chrome/124.0.0.0 Safari/537.36")

    # Domains that are known to be paywalled and won't yield a free PDF even
    # with a real browser — skip Playwright for these to save time.
    SKIP_DOMAINS = {
        'nejm.org', 'nature.com', 'science.org', 'cell.com',
        'jamanetwork.com', 'annals.org',
    }

    def __init__(self, verbose=False, proxy=None):
        self.verbose = verbose
        self.proxy   = proxy   # dict {"http": ..., "https": ...} or None

    def log(self, msg):
        if self.verbose:
            tprint(msg)

    @staticmethod
    def _domain(url):
        return urlparse(url).netloc.lower().lstrip('www.')

    def should_skip(self, url):
        d = self._domain(url)
        return any(skip in d for skip in PlaywrightDownloader.SKIP_DOMAINS)

    def download(self, url, headless=True):
        """
        Navigate to `url` in a real browser and return PDF bytes, or None.
        """
        if not PLAYWRIGHT_AVAILABLE:
            return None
        if self.should_skip(url):
            self.log(f"    [PW] Skipping known paywall: {self._domain(url)}")
            return None

        _PLAYWRIGHT_SEMAPHORE.acquire()
        try:
            return self._run(url, headless)
        finally:
            _PLAYWRIGHT_SEMAPHORE.release()

    def _run(self, url, headless=True):
        intercepted = {}

        def on_response(response):
            ct = response.headers.get("content-type", "")
            if "pdf" in ct.lower() and response.status == 200:
                try:
                    body = response.body()
                    if body[:4] == b"%PDF":
                        intercepted["body"] = body
                        intercepted["url"]  = response.url
                        self.log(f"    [PW] Intercepted PDF ({len(body)//1024} KB) ← {response.url[:70]}")
                except Exception:
                    pass

        try:
            with sync_playwright() as pw:
                launch_kwargs = dict(
                    headless=headless,
                    args=["--disable-blink-features=AutomationControlled",
                          "--no-sandbox", "--disable-dev-shm-usage"],
                )

                browser = pw.chromium.launch(**launch_kwargs)

                ctx_kwargs = dict(
                    user_agent=self.UA,
                    accept_downloads=True,
                    viewport={"width": 1280, "height": 800},
                    locale="en-US",
                    extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
                )
                # Wire proxy if provided
                if self.proxy:
                    proxy_url = self.proxy.get("https") or self.proxy.get("http", "")
                    if proxy_url:
                        ctx_kwargs["proxy"] = {"server": proxy_url}

                context = browser.new_context(**ctx_kwargs)
                page    = context.new_page()
                page.on("response", on_response)

                # ── Method A: navigate and intercept ─────────────────────
                self.log(f"    [PW] Navigating: {url[:80]}")
                try:
                    with page.expect_download(timeout=25_000) as dl_info:
                        page.goto(url, wait_until="commit", timeout=30_000)

                    # Browser triggered a file download
                    tmp = tempfile.mktemp(suffix=".pdf")
                    dl_info.value.save_as(tmp)
                    data = Path(tmp).read_bytes()
                    os.unlink(tmp)
                    if data[:4] == b"%PDF":
                        self.log(f"    [PW] ✓ Download captured ({len(data)//1024} KB)")
                        browser.close()
                        return data

                except Exception:
                    # No download triggered — that's fine, fall through
                    pass

                # Give JS a moment to settle
                try:
                    page.wait_for_load_state("networkidle", timeout=8_000)
                except Exception:
                    pass

                # ── Method B: check interception buffer ──────────────────
                if "body" in intercepted:
                    browser.close()
                    return intercepted["body"]

                # Cloudflare headless detection? Retry headed.
                title = page.title()
                if headless and ("Just a moment" in title or
                                 "cf-browser-verification" in page.content()):
                    self.log("    [PW] ⚠ Cloudflare detected — retrying in headed mode")
                    browser.close()
                    return self._run(url, headless=False)

                # ── Method C: find and click PDF/download link on page ───
                self.log("    [PW] Scanning page for PDF links…")
                pdf_link = None

                # Priority 1: links that end in .pdf
                for sel in [
                    "a[href$='.pdf']",
                    "a[href*='.pdf?']",
                    "a[href*='/pdf/']",
                    "a[href*='download']",
                    "a[href*='bitstream']",
                ]:
                    try:
                        el = page.query_selector(sel)
                        if el:
                            href = el.get_attribute("href") or ""
                            if href and not href.startswith('#'):
                                pdf_link = (href if href.startswith("http")
                                            else f"{urlparse(url).scheme}://{urlparse(url).netloc}{href}")
                                self.log(f"    [PW] Found link [{sel}]: {pdf_link[:70]}")
                                break
                    except Exception:
                        pass

                # Priority 2: visible buttons/links with PDF-related text
                if not pdf_link:
                    for text in ["PDF", "Download PDF", "Full Text PDF",
                                 "Download", "Get PDF", "Open PDF"]:
                        try:
                            el = page.get_by_text(text, exact=False).first
                            href = el.get_attribute("href") if el else None
                            if href and href.startswith("http") and not href.startswith('#'):
                                pdf_link = href
                                self.log(f"    [PW] Found text link '{text}': {pdf_link[:70]}")
                                break
                        except Exception:
                            pass

                if pdf_link:
                    intercepted.clear()
                    try:
                        with page.expect_download(timeout=25_000) as dl_info:
                            page.goto(pdf_link, wait_until="commit", timeout=25_000)
                        tmp = tempfile.mktemp(suffix=".pdf")
                        dl_info.value.save_as(tmp)
                        data = Path(tmp).read_bytes()
                        os.unlink(tmp)
                        if data[:4] == b"%PDF":
                            self.log(f"    [PW] ✓ Link download ({len(data)//1024} KB)")
                            browser.close()
                            return data
                    except Exception:
                        page.wait_for_timeout(3_000)

                    if "body" in intercepted:
                        self.log(f"    [PW] ✓ Link intercept ({len(intercepted['body'])//1024} KB)")
                        browser.close()
                        return intercepted["body"]

                self.log("    [PW] ✗ No PDF found")
                browser.close()

        except Exception as e:
            self.log(f"    [PW] ✗ Error: {e}")

        return None


# ============================================================================
# PDF Validator  (unchanged from v6)
# ============================================================================

class PDFValidator:
    @staticmethod
    def validate_pdf(filepath):
        try:
            if not os.path.exists(filepath):
                return False, "File does not exist"
            size = os.path.getsize(filepath)
            if size < 100:
                return False, f"File too small ({size} bytes)"
            with open(filepath, 'rb') as f:
                if f.read(5) != b'%PDF-':
                    return False, "Invalid PDF header"
            with open(filepath, 'rb') as f:
                pdf = PyPDF2.PdfReader(f)
                pages = len(pdf.pages)
                if pages == 0:
                    return False, "Zero pages"
                return True, f"{pages} pages, {size/1024:.1f} KB"
        except Exception as e:
            return False, f"Validation error: {str(e)[:50]}"

    @staticmethod
    def get_file_hash(filepath):
        h = hashlib.md5()
        with open(filepath, 'rb') as f:
            buf = f.read(65536)
            while buf:
                h.update(buf)
                buf = f.read(65536)
        return h.hexdigest()

    @staticmethod
    def sanitize_filename(url, nid=None, title=None):
        if title and str(title).strip():
            clean = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean = re.sub(r'[\s]+', '_', clean).strip('_')[:150]
            fn = f"nid_{nid}_{clean}.pdf" if nid else f"{clean}.pdf"
        else:
            parsed = urlparse(url)
            fn = os.path.basename(unquote(parsed.path))
            if not fn or len(fn) < 5:
                fn = f"{parsed.netloc.replace('www.','')}_{hashlib.md5(url.encode()).hexdigest()[:8]}.pdf"
            if not fn.lower().endswith('.pdf'):
                fn += '.pdf'
            if nid:
                fn = f"nid_{nid}_{fn}"
        return re.sub(r'[<>:"/\\|?*]', '_', fn)[:255]


# ============================================================================
# PMC Downloader  (unchanged from v6)
# ============================================================================

class PMCDownloader:
    def __init__(self, api_key=None):
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0'})

    def is_pmc_url(self, url):
        return 'ncbi.nlm.nih.gov' in url.lower() and ('pmc' in url.lower() or 'pubmed' in url.lower())

    def get_pmc_id(self, url):
        m = re.search(r'PMC(\d+)', url, re.I)
        return m.group(1) if m else None

    def get_pdf_link_from_oa_service(self, pmc_id):
        try:
            from xml.etree import ElementTree as ET
            params = {'id': f'PMC{pmc_id}'}
            if self.api_key:
                params['api_key'] = self.api_key
            r = self.session.get("https://www.ncbi.nlm.nih.gov/pmc/utils/oa/oa.fcgi",
                                 params=params, timeout=10)
            r.raise_for_status()
            for link in ET.fromstring(r.content).findall('.//link'):
                if 'pdf' in link.get('format', '').lower() and link.get('href'):
                    return link.get('href')
        except Exception:
            pass
        return None

    def download(self, url, output_path):
        pmc_id = self.get_pmc_id(url)
        if not pmc_id:
            return False, "Could not extract PMC ID"
        pdf_url = self.get_pdf_link_from_oa_service(pmc_id)
        if not pdf_url:
            for pattern in [
                f"https://www.ncbi.nlm.nih.gov/pmc/articles/PMC{pmc_id}/pdf/",
                f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{pmc_id}/pdf/",
            ]:
                try:
                    params = {'api_key': self.api_key} if self.api_key else {}
                    if self.session.head(pattern, params=params, timeout=5).status_code == 200:
                        pdf_url = pattern
                        break
                except Exception:
                    continue
        if not pdf_url:
            return False, "Could not find PDF link"
        if pdf_url.startswith('ftp://'):
            return self._download_ftp(pdf_url, output_path)
        try:
            params = {'api_key': self.api_key} if self.api_key else {}
            r = self.session.get(pdf_url, params=params, timeout=30, stream=True)
            r.raise_for_status()
            with open(output_path, 'wb') as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)
            return True, "Downloaded from PMC"
        except Exception as e:
            return False, f"Download error: {str(e)[:100]}"

    def _download_ftp(self, ftp_url, output_path):
        import tarfile, zipfile
        from urllib.request import urlopen as _urlopen
        try:
            arc = output_path + '.archive'
            with open(arc, 'wb') as f:
                f.write(_urlopen(ftp_url, timeout=30).read())
            ok = self._extract(arc, output_path)
            if os.path.exists(arc):
                os.remove(arc)
            return ok, "Extracted from FTP" if ok else "FTP extraction failed"
        except Exception as e:
            return False, f"FTP error: {str(e)[:100]}"

    def _extract(self, arc, out):
        import tarfile, zipfile
        try:
            if tarfile.is_tarfile(arc):
                with tarfile.open(arc, 'r:*') as t:
                    pdfs = sorted([m for m in t.getmembers() if m.name.lower().endswith('.pdf')],
                                  key=lambda x: x.name.count('/'))
                    if pdfs:
                        data = t.extractfile(pdfs[0]).read()
                        if data[:4] == b'%PDF':
                            Path(out).write_bytes(data)
                            return True
            elif zipfile.is_zipfile(arc):
                with zipfile.ZipFile(arc) as z:
                    pdfs = sorted([n for n in z.namelist() if n.lower().endswith('.pdf')],
                                  key=lambda x: x.count('/'))
                    if pdfs:
                        data = z.read(pdfs[0])
                        if data[:4] == b'%PDF':
                            Path(out).write_bytes(data)
                            return True
        except Exception:
            pass
        return False


# ============================================================================
# PubMed Handler  (unchanged from v6)
# ============================================================================

class PubMedHandler:
    def __init__(self, pmc_downloader):
        self.pmc = pmc_downloader
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0'})

    def extract_pmid(self, url):
        m = re.search(r'/pubmed/(\d+)', url)
        return m.group(1) if m else None

    def get_pmc_from_pmid(self, pmid):
        try:
            params = {'dbfrom': 'pubmed', 'db': 'pmc',
                      'id': pmid, 'retmode': 'json'}
            if self.pmc.api_key:
                params['api_key'] = self.pmc.api_key
            data = self.session.get(
                "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/elink.fcgi",
                params=params, timeout=10).json()
            for ls in data.get('linksets', []):
                for lsdb in ls.get('linksetdbs', []):
                    if lsdb.get('linkname') == 'pubmed_pmc':
                        links = lsdb.get('links', [])
                        if links:
                            return links[0]
        except Exception:
            pass
        return None

    def download(self, pubmed_url, output_path):
        pmid = self.extract_pmid(pubmed_url)
        if not pmid:
            return False, "Could not extract PMID", None
        pmc_id = self.get_pmc_from_pmid(pmid)
        if pmc_id:
            ok, msg = self.pmc.download(
                f"https://www.ncbi.nlm.nih.gov/pmc/articles/PMC{pmc_id}/",
                output_path)
            return ok, msg, ('pubmed_pmc' if ok else None)
        return False, "No free PMC version available", None


# ============================================================================
# Abstract Saver  (unchanged from v6)
# ============================================================================

class AbstractSaver:
    def __init__(self, output_dir="abstracts"):
        self.output_dir = output_dir
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)

    def save_abstract(self, node_id, title, abstract):
        if not abstract or pd.isna(abstract):
            return None
        if title and str(title).strip():
            clean = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean = re.sub(r'[\s]+', '_', clean).strip('_')[:150]
            fn = f"nid_{node_id}_{clean}.txt"
        else:
            fn = f"nid_{node_id}_abstract.txt"
        fp = os.path.join(self.output_dir, fn)
        try:
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(f"{'='*80}\nInfoNTD Abstract\n{'='*80}\n\n"
                        f"Node ID: {node_id}\nTitle: {title}\n"
                        f"Date Saved: {datetime.now():%Y-%m-%d %H:%M:%S}\n\n"
                        f"{'='*80}\nABSTRACT\n{'='*80}\n\n{abstract}\n\n{'='*80}\n")
            return fp
        except Exception:
            return None


# ============================================================================
# Enhanced PDF Finder  (updated WHO IRIS handler)
# ============================================================================

class EnhancedPDFFinder:
    def __init__(self, verbose=False, proxy_manager=None, playwright_fallback=True):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': random.choice(_USER_AGENTS)
        })
        self.verbose            = verbose
        self.proxy_manager      = proxy_manager
        self.playwright_fallback = playwright_fallback and PLAYWRIGHT_AVAILABLE
        self.journal_handler    = JournalSpecificHandler() if JOURNAL_HANDLER_AVAILABLE else None
        self.handler_stats      = {}
        self._stats_lock        = threading.Lock()

    def _get_with_proxy_retry(self, url, stream=False, timeout=30):
        """
        GET a URL with automatic proxy rotation on 403/429.
        Returns a requests.Response.
        Tries up to 3 times: direct → proxy1 → proxy2.
        """
        last_exc = None
        proxy = None

        for attempt in range(3):
            # Rotate user-agent on each retry
            self.session.headers.update({'User-Agent': random.choice(_USER_AGENTS)})

            if attempt > 0 and self.proxy_manager and self.proxy_manager.enabled:
                if proxy:
                    self.proxy_manager.report_failure(proxy)
                    self.proxy_manager.rotate()
                proxy = self.proxy_manager.get_proxy()
                self.proxy_manager.apply_to_session(self.session, proxy)
                if self.verbose and proxy:
                    tprint(f"    🔀 Retry {attempt} via proxy: {list(proxy.values())[0]}")

            try:
                r = self.session.get(url, timeout=timeout, stream=stream, allow_redirects=True)
                if r.status_code in (403, 429) and attempt < 2:
                    if self.verbose:
                        tprint(f"    ⚠️ HTTP {r.status_code} — will retry with proxy")
                    last_exc = requests.HTTPError(response=r)
                    continue
                return r
            except Exception as e:
                last_exc = e

        raise last_exc or requests.ConnectionError(f"All attempts failed for {url}")

    def log(self, msg):
        if self.verbose:
            tprint(msg)

    def _track_handler(self, name):
        with self._stats_lock:
            self.handler_stats[name] = self.handler_stats.get(name, 0) + 1

    def _pw_downloader(self):
        """Build a PlaywrightDownloader with the current proxy (if any)."""
        proxy = self.proxy_manager.get_proxy() if (
            self.proxy_manager and self.proxy_manager.enabled) else None
        return PlaywrightDownloader(verbose=self.verbose, proxy=proxy)

    def playwright_download_to_tmp(self, url):
        """
        Run the generic Playwright downloader on url.
        Returns a __BYTES_FILE__:/path sentinel string, or None.
        """
        if not self.playwright_fallback:
            return None
        self.log(f"  → Generic Playwright fallback for: {url[:70]}")
        data = self._pw_downloader().download(url)
        if data:
            fname = os.path.basename(urlparse(url).path) or "download.pdf"
            if not fname.lower().endswith('.pdf'):
                fname += '.pdf'
            tmp = tempfile.mktemp(suffix=f"_{fname}")
            Path(tmp).write_bytes(data)
            self._track_handler('Playwright_Generic')
            return f"__BYTES_FILE__:{tmp}"
        return None

    # ── URL normalisation (unchanged) ────────────────────────────────────────

    def normalize_url(self, url):
        url = url.strip()
        if re.match(r'^10\.\d+/', url):
            return f'https://doi.org/{url}'
        if re.match(r'^\d{7,8}$', url):
            return f'https://pubmed.ncbi.nlm.nih.gov/{url}/'
        if url.startswith('dx.doi.org'):
            return 'https://' + url
        if 'apps.who.int/iris' in url:
            m = re.search(r'10665/(\d+)', url)
            if m:
                return f"https://iris.who.int/handle/10665/{m.group(1)}"
        return url

    def expand_short_url(self, url):
        if any(d in url.lower() for d in ['tinyurl.com', 'bit.ly', 'goo.gl', 't.co', 'ow.ly']):
            try:
                r = self.session.head(url, timeout=5, allow_redirects=True)
                if r.url != url:
                    return r.url
            except Exception:
                pass
        return url

    # ── Main entry ───────────────────────────────────────────────────────────

    def find_pdf(self, url):
        url = self.normalize_url(url)
        url = self.expand_short_url(url)
        if not url.startswith('http'):
            return None

        domain = urlparse(url).netloc.lower()

        if 'iris.who.int' in url or 'apps.who.int/iris' in url:
            self._track_handler('WHO_IRIS')
            return self._handle_who_iris(url)

        handler_map = [
            ('biomedcentral.com',      'BioMedCentral',  self._handle_biomedcentral),
            ('mdpi.com',               'MDPI',            self._handle_mdpi),
            ('thelancet.com',          'Lancet',          self._handle_lancet),
            ('tandfonline.com',        'Tandfonline',     self._handle_tandfonline),
            ('bmj.com',                'BMJ',             self._handle_bmj),
            ('oup.com',                'Oxford',          self._handle_oxford),
            ('oxfordjournals.org',     'Oxford',          self._handle_oxford),
            ('sciencedirect.com',      'ScienceDirect',   self._handle_sciencedirect),
            ('elsevier.com',           'ScienceDirect',   self._handle_sciencedirect),
            ('plos.org',               'PLOS',            self._handle_plos),
            ('plosntds.org',           'PLOS',            self._handle_plos),
            ('wiley.com',              'Wiley',           self._handle_wiley),
            ('onlinelibrary.wiley.com','Wiley',           self._handle_wiley),
            ('springeropen.com',       'Springer',        self._handle_springer),
            ('springer.com',           'Springer',        self._handle_springer),
        ]
        for pat, name, fn in handler_map:
            if pat in domain:
                self.log(f"  → Using {name} handler")
                self._track_handler(name)
                return fn(url)

        if self.journal_handler and self.journal_handler.can_handle(url):
            try:
                self._track_handler('Journal_Specific')
                return self.journal_handler.find_pdf(url)
            except Exception:
                pass

        self._track_handler('Generic')
        return self._handle_generic(url)

    # ── WHO IRIS — upgraded handler ──────────────────────────────────────────

    def _handle_who_iris(self, url):
        """
        Upgraded WHO IRIS handler.
        Fast urllib strategies first; falls back to Playwright for Cloudflare.
        Returns a URL string for simple cases, or saves bytes and returns filepath.
        """
        # Already a direct content API URL?
        if '/server/api/core/bitstreams/' in url and '/content' in url:
            self.log("    ✓ Direct bitstream API URL")
            return url
        if '/bitstream/handle/' in url and url.lower().endswith('.pdf'):
            self.log("    ✓ Direct bitstream URL")
            return url

        # Use WhoIrisDownloader — returns bytes
        self.log("    → WhoIrisDownloader (multi-strategy)")
        downloader = WhoIrisDownloader(verbose=self.verbose, proxy_manager=self.proxy_manager)
        pdf_bytes = downloader.download(url)
        if pdf_bytes:
            # Write to a temp file so the caller can validate it normally
            _, filename = WhoIrisDownloader.parse_url(url)
            tmp = tempfile.mktemp(suffix=f"_{filename or 'iris.pdf'}")
            Path(tmp).write_bytes(pdf_bytes)
            self.log(f"    ✓ WHO IRIS bytes written to temp: {tmp}")
            return f"__BYTES_FILE__:{tmp}"   # sentinel for download_pdf()
        return None

    # ── All other handlers (identical to v6) ─────────────────────────────────

    def _handle_biomedcentral(self, url):
        m = re.search(r'10\.1186/[^\s/?]+', url)
        if m:
            doi = m.group(0).rstrip('/')
            for pdf_url in [
                f"https://link.springer.com/content/pdf/{doi}.pdf",
                f"https://bmcinfectdis.biomedcentral.com/track/pdf/{doi}",
            ]:
                try:
                    r = self.session.head(pdf_url, timeout=5, allow_redirects=True)
                    if r.status_code == 200 and 'pdf' in r.headers.get('Content-Type','').lower():
                        return pdf_url
                except Exception:
                    continue
        return None

    def _handle_mdpi(self, url):
        if '/htm' in url:
            return url.replace('/htm', '/pdf')
        return url.rstrip('/') + '/pdf' if not url.endswith('.pdf') else url

    def _handle_lancet(self, url):
        if '/action/showPdf' in url:
            return url
        for pat in [r'PII[S:]?(S?\d+-\d+X?\([^)]+\))',
                    r'/article/(S?\d+-\d+X?\([^/)]+\))',
                    r'pii=(S[^&]+)']:
            m = re.search(pat, url)
            if m:
                pii = m.group(1).replace('(', '%28').replace(')', '%29')
                return f"https://www.thelancet.com/action/showPdf?pii={pii}"
        return None

    def _handle_sciencedirect(self, url):
        m = re.search(r'(S\d{4}\d{3,4}X?\d{2}\d{5}[\dX]?)', url)
        if m:
            pii = m.group(1)
            for pdf_url in [
                f"https://www.sciencedirect.com/science/article/pii/{pii}/pdfft?download=true",
                f"https://reader.elsevier.com/reader/sd/pii/{pii}",
            ]:
                try:
                    if self.session.head(pdf_url, timeout=5, allow_redirects=True).status_code == 200:
                        return pdf_url
                except Exception:
                    continue
        return None

    def _handle_bmj(self, url):
        url = url.replace('.full.pdf+html', '.full.pdf')
        if not url.endswith('.pdf') and '/content/' in url:
            url = url.split('.full.pdf')[0].rstrip('/') + '.full.pdf'
        return url

    def _handle_tandfonline(self, url):
        url = url.replace('/doi/abs/', '/doi/pdf/').replace('/doi/full/', '/doi/pdf/')
        if '?' not in url:
            url += '?needAccess=true'
        return url

    def _handle_oxford(self, url):
        url = url.replace('.full.pdf+html', '.full.pdf')
        if 'academic.oup.com' in url:
            m = re.search(r'/doi/([0-9.]+/[^/\s?]+)', url)
            if m:
                journal = urlparse(url).path.split('/')[1]
                return f"https://academic.oup.com/{journal}/article-pdf/doi/{m.group(1)}/pdf"
        elif 'oxfordjournals.org' in url:
            if url.endswith('.pdf'):
                return url
            if '/content/' in url:
                return url.split('.full.pdf')[0].rstrip('/') + '.full.pdf'
        return None

    def _handle_plos(self, url):
        m = re.search(r'10\.1371/journal\.[^?\s]+', url)
        if m:
            return f"https://journals.plos.org/plosone/article/file?id={m.group(0)}&type=printable"
        return None

    def _handle_wiley(self, url):
        return url.replace('/doi/', '/doi/pdf/').replace('/epdf/', '/pdf/')

    def _handle_springer(self, url):
        for pat in [r'10\.1186[^?&]+', r'10\.1186/[^\s?]+']:
            m = re.search(pat, url)
            if m:
                doi = m.group(0).replace('%2F', '/')
                return f"https://link.springer.com/content/pdf/{doi}.pdf"
        return None

    def _handle_generic(self, url):
        if url.lower().endswith('.pdf'):
            return url
        if self.journal_handler:
            try:
                r = self.journal_handler.find_pdf(url)
                if r:
                    return r
            except Exception:
                pass
        return url


# ============================================================================
# Thread-safe state containers
# ============================================================================

class ThreadSafeStats:
    """Wraps a stats dict with a lock."""
    def __init__(self, keys):
        self._d = {k: 0 for k in keys}
        self._lock = threading.Lock()

    def inc(self, key, n=1):
        with self._lock:
            self._d[key] = self._d.get(key, 0) + n

    def get(self, key):
        with self._lock:
            return self._d.get(key, 0)

    def snapshot(self):
        with self._lock:
            return dict(self._d)


class ThreadSafeSet:
    def __init__(self):
        self._s = set()
        self._lock = threading.Lock()

    def add_if_absent(self, val):
        """Add val and return True if it was new, False if already present."""
        with self._lock:
            if val in self._s:
                return False
            self._s.add(val)
            return True

    def __contains__(self, val):
        with self._lock:
            return val in self._s


# ============================================================================
# Main Excel PDF Scraper  (v7 — parallel)
# ============================================================================

class ExcelPDFScraper:
    def __init__(self, download_dir="infontd_pdfs", abstract_dir="abstracts",
                 delay=1, ncbi_api_key=None, autosave_interval=50,
                 start_from_row=None, domain_filter=None, type_filter="biblio",
                 verbose=False, workers=5, proxy_manager=None, playwright_fallback=True):

        self.download_dir      = download_dir
        self.abstract_dir      = abstract_dir
        self.delay             = delay
        self.autosave_interval = autosave_interval
        self.start_from_row    = start_from_row
        self.domain_filter     = domain_filter or []
        self.type_filter       = type_filter
        self.verbose           = verbose
        self.workers           = workers
        self.proxy_manager     = proxy_manager

        Path(self.download_dir).mkdir(parents=True, exist_ok=True)
        Path(self.abstract_dir).mkdir(parents=True, exist_ok=True)

        self.validator      = PDFValidator()
        self.pmc_downloader = PMCDownloader(api_key=ncbi_api_key)
        self.pubmed_handler = PubMedHandler(self.pmc_downloader)
        self.abstract_saver = AbstractSaver(self.abstract_dir)
        self.pdf_finder     = EnhancedPDFFinder(
            verbose=verbose,
            proxy_manager=proxy_manager,
            playwright_fallback=playwright_fallback,
        )

        # Thread-safe shared state
        self.downloaded_hashes = ThreadSafeSet()
        self.processed_urls    = ThreadSafeSet()
        self._df_lock          = threading.Lock()   # for DataFrame writes
        self._save_lock        = threading.Lock()   # for Excel auto-save

        self.stats = ThreadSafeStats([
            'total_rows', 'infontd_rows', 'valid_urls',
            'already_downloaded', 'duplicate_urls', 'manually_skipped',
            'success_bibcite', 'success_pubmed', 'success_abstract_txt',
            'abstracts_deleted', 'failed', 'skipped', 'duplicates', 'corrupt',
            'pmc_downloads',
        ])

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _valid_url(self, val):
        if pd.isna(val) or not val or str(val).strip() == '':
            return None
        url = str(val).strip()
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url if url.startswith('www.') else None
        return url

    def is_already_downloaded(self, row):
        return (pd.notna(row.get('download_status')) and
                row.get('download_status') in ('success_bibcite', 'success_pubmed'))

    def delete_abstract_file(self, nid, title=None):
        for fn in self._abstract_filenames(nid, title):
            fp = os.path.join(self.abstract_dir, fn)
            if os.path.exists(fp):
                try:
                    os.remove(fp)
                    self.stats.inc('abstracts_deleted')
                except Exception:
                    pass

    def _abstract_filenames(self, nid, title):
        fns = [f"nid_{nid}_abstract.txt"]
        if title and str(title).strip():
            clean = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean = re.sub(r'[\s]+', '_', clean).strip('_')[:150]
            fns.insert(0, f"nid_{nid}_{clean}.txt")
        return fns

    # ── Download a single PDF ─────────────────────────────────────────────────

    def download_pdf(self, url, nid, title=None):
        """
        Download PDF from URL.
        Returns (status, filepath, message).

        Fallback chain for non-PMC, non-WHO-IRIS URLs:
          1. find_pdf()  → publisher-specific URL construction
          2. _get_with_proxy_retry()  → download the constructed URL
          3. Generic Playwright browser  → last resort if requests fails
        """
        filename = self.validator.sanitize_filename(url, nid, title=title)
        filepath = os.path.join(self.download_dir, filename)

        if self.pmc_downloader.is_pmc_url(url):
            self.stats.inc('pmc_downloads')
            ok, msg = self.pmc_downloader.download(url, filepath)
            if not ok:
                return 'failed', None, f'PMC: {msg}'
        else:
            pdf_url = self.pdf_finder.find_pdf(url)
            if not pdf_url:
                return 'failed', None, f'Could not find PDF URL (src: {url[:100]})'

            # WHO IRIS / any handler that returned pre-downloaded bytes
            if pdf_url.startswith("__BYTES_FILE__:"):
                tmp_path = pdf_url[len("__BYTES_FILE__:"):]
                try:
                    import shutil
                    shutil.move(tmp_path, filepath)
                except Exception as e:
                    return 'failed', None, f'Could not move temp file: {e}'
            else:
                # Standard requests-based download
                requests_ok = False
                fail_reason = ''
                try:
                    r = self.pdf_finder._get_with_proxy_retry(pdf_url, stream=True, timeout=30)
                    r.raise_for_status()
                    ct = r.headers.get('Content-Type', '')
                    if 'html' in ct.lower() and 'pdf' not in ct.lower():
                        fail_reason = f'Got HTML not PDF (CT: {ct})'
                    else:
                        with open(filepath, 'wb') as f:
                            for chunk in r.iter_content(8192):
                                f.write(chunk)
                        requests_ok = True
                except requests.HTTPError as e:
                    sc = e.response.status_code if e.response else '?'
                    fail_reason = f'HTTP {sc}'
                except requests.Timeout:
                    fail_reason = 'Timeout'
                except Exception as e:
                    fail_reason = f'{type(e).__name__}: {str(e)[:80]}'

                # ── Generic Playwright fallback ───────────────────────────
                if not requests_ok:
                    self.pdf_finder.log(
                        f"    ⚠ requests failed ({fail_reason}) — trying Playwright"
                    )
                    pw_result = self.pdf_finder.playwright_download_to_tmp(
                        # Use the original page URL (not the constructed pdf_url)
                        # so Playwright loads the article page and finds the link itself
                        url if pdf_url == url else pdf_url
                    )
                    if pw_result and pw_result.startswith("__BYTES_FILE__:"):
                        tmp_path = pw_result[len("__BYTES_FILE__:"):]
                        try:
                            import shutil
                            shutil.move(tmp_path, filepath)
                        except Exception as e:
                            return 'failed', None, f'Playwright move failed: {e}'
                    else:
                        # Also try with the original landing page URL
                        if pdf_url != url:
                            pw_result2 = self.pdf_finder.playwright_download_to_tmp(url)
                            if pw_result2 and pw_result2.startswith("__BYTES_FILE__:"):
                                tmp_path = pw_result2[len("__BYTES_FILE__:"):]
                                try:
                                    import shutil
                                    shutil.move(tmp_path, filepath)
                                except Exception as e:
                                    return 'failed', None, f'Playwright move failed: {e}'
                            else:
                                return 'failed', None, (
                                    f'{fail_reason} | Playwright also failed '
                                    f'(url: {pdf_url[:80]})'
                                )
                        else:
                            return 'failed', None, (
                                f'{fail_reason} | Playwright also failed '
                                f'(url: {pdf_url[:80]})'
                            )

        ok, msg = self.validator.validate_pdf(filepath)
        if not ok:
            sz = os.path.getsize(filepath) if os.path.exists(filepath) else 0
            if os.path.exists(filepath):
                os.remove(filepath)
            return 'corrupt', None, f'Invalid PDF: {msg} (size: {sz}B)'

        h = self.validator.get_file_hash(filepath)
        if not self.downloaded_hashes.add_if_absent(h):
            os.remove(filepath)
            return 'duplicate', None, f'Duplicate (hash: {h[:16]})'

        return 'success', filepath, msg

    # ── Process one row ───────────────────────────────────────────────────────

    def process_row(self, row):
        """
        Returns (status, filepath, url_used, source_used, message,
                 detailed_errors, handler_used).
        Thread-safe: all shared state accessed via locks.
        """
        if self.is_already_downloaded(row):
            return (row.get('download_status'), row.get('download_filepath'),
                    row.get('url_used',''), row.get('source_used',''),
                    'Previously downloaded', row.get('detailed_errors',''),
                    row.get('handler_used',''))

        nid          = row.get('Node ID')
        title        = row.get('Title')
        bibcite_url  = self._valid_url(row.get('Bibcite URL'))
        pubmed_url   = self._valid_url(row.get('PubMed URL'))
        abstract     = row.get('Abstract (English)')
        has_abstract = pd.notna(abstract) and str(abstract).strip() != ''
        errors       = []
        handler_used = ''

        # Step 1: Bibcite URL
        if bibcite_url:
            if not self.processed_urls.add_if_absent(bibcite_url):
                self.stats.inc('duplicate_urls')
                errors.append("Bibcite: Duplicate URL")
            else:
                status, fp, msg = self.download_pdf(bibcite_url, nid, title=title)
                if status == 'success':
                    self.stats.inc('success_bibcite')
                    self.delete_abstract_file(nid, title)
                    return 'success_bibcite', fp, bibcite_url, 'Bibcite URL', msg, '', handler_used
                elif status == 'duplicate':
                    self.stats.inc('duplicates')
                    return 'duplicate', fp, bibcite_url, 'Bibcite URL', msg, '', handler_used
                else:
                    errors.append(f"Bibcite: {msg}")
        else:
            errors.append("Bibcite: No URL")

        # Step 2: PubMed URL
        if pubmed_url:
            if not self.processed_urls.add_if_absent(pubmed_url):
                errors.append("PubMed: Duplicate URL")
            else:
                status, fp, msg = self.download_pdf(pubmed_url, nid, title=title)
                if status == 'success':
                    self.stats.inc('success_pubmed')
                    self.delete_abstract_file(nid, title)
                    return ('success_pubmed', fp, pubmed_url, 'PubMed URL',
                            msg, ' | '.join(errors), handler_used)
                elif status == 'duplicate':
                    self.stats.inc('duplicates')
                    return ('duplicate', fp, pubmed_url, 'PubMed URL',
                            msg, ' | '.join(errors), handler_used)
                else:
                    errors.append(f"PubMed: {msg}")
        else:
            errors.append("PubMed: No URL")

        # Step 3: Abstract
        if has_abstract:
            fp = self.abstract_saver.save_abstract(nid, title, abstract)
            if fp:
                self.stats.inc('success_abstract_txt')
                return ('success_abstract_txt', fp, '', 'Abstract (English)',
                        f'Abstract saved ({os.path.getsize(fp)}B)',
                        ' | '.join(errors), '')
            else:
                self.stats.inc('failed')
                errors.append("Abstract: save failed")
                return ('failed', None, '', 'Abstract (English)',
                        'Could not save abstract', ' | '.join(errors), '')
        else:
            errors.append("Abstract: none")

        self.stats.inc('skipped')
        return 'skipped', None, '', 'none', 'No data available', ' | '.join(errors), ''

    # ── Parallel Excel processing ─────────────────────────────────────────────

    def process_excel_file(self, input_file, output_file=None):
        tprint(f"\n{'='*80}")
        tprint(f"PDF Scraper v7  —  {self.workers} parallel workers")
        tprint(f"Fallback chain: Bibcite URL → PubMed URL → Abstract TXT")
        tprint(f"WHO IRIS: DSpace7 API → OAI-PMH → Session → Playwright")
        tprint(f"{'='*80}\n")

        tprint(f"Reading: {input_file}")
        df = pd.read_excel(input_file, engine='openpyxl')

        if output_file is None:
            output_file = f"infontd_processed_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

        self.stats.inc('total_rows', len(df))
        tprint(f"Total rows: {len(df)}")

        # Filter by domain — OR across all entries in domain_filter
        domain_filters = self.domain_filter
        if isinstance(domain_filters, str):
            domain_filters = [domain_filters]
        if domain_filters:
            mask = df['Domains'].str.contains(domain_filters[0], case=False, na=False)
            for d in domain_filters[1:]:
                mask = mask | df['Domains'].str.contains(d, case=False, na=False)
        else:
            mask = pd.Series([True] * len(df), index=df.index)
        work_df = df[mask].copy()
        if self.type_filter:
            work_df = work_df[work_df['Type'].str.contains(
                self.type_filter, case=False, na=False)]
        self.stats.inc('infontd_rows', len(work_df))
        tprint(f"Work rows after filter: {len(work_df)}")

        # Ensure output columns
        for col in ['download_status','download_filepath','download_filename',
                    'url_used','source_used','download_error','detailed_errors',
                    'handler_used','download_timestamp']:
            if col not in df.columns:
                df[col] = ''

        already = df['download_status'].isin(['success_bibcite','success_pubmed']).sum()
        if already:
            tprint(f"Already processed: {already} — will skip\n")

        tprint(f"{'='*80}")
        tprint(f"Starting parallel run (workers={self.workers})")
        tprint(f"{'='*80}\n")

        start = time.time()
        processed_count = 0
        rows_done = 0
        total = len(work_df)
        counter_lock = threading.Lock()

        def _task(idx, row_dict):
            """Worker: process one row, return (idx, result_tuple)."""
            if self.delay > 0:
                time.sleep(random.uniform(0, self.delay))
            result = self.process_row(row_dict)
            return idx, result

        with ThreadPoolExecutor(max_workers=self.workers) as ex:
            futures = {}
            for idx in work_df.index:
                row = df.loc[idx]
                with counter_lock:
                    processed_count += 1
                    pos = processed_count

                if self.start_from_row and pos < self.start_from_row:
                    self.stats.inc('manually_skipped')
                    continue

                future = ex.submit(_task, idx, row.to_dict())
                futures[future] = (idx, pos)

            for future in as_completed(futures):
                idx, pos = futures[future]
                try:
                    idx, (status, filepath, url_used, source_used,
                          message, detailed_errors, handler_used) = future.result()
                except Exception as e:
                    tprint(f"  ✗ Worker exception for row {idx}: {e}")
                    continue

                nid = df.at[idx, 'Node ID'] if 'Node ID' in df.columns else idx

                # Update DataFrame under lock
                with self._df_lock:
                    df.at[idx, 'download_status']    = status
                    df.at[idx, 'download_filepath']   = filepath or ''
                    df.at[idx, 'download_filename']   = os.path.basename(filepath) if filepath else ''
                    df.at[idx, 'url_used']            = url_used
                    df.at[idx, 'source_used']         = source_used
                    df.at[idx, 'download_error']      = message
                    df.at[idx, 'detailed_errors']     = detailed_errors
                    df.at[idx, 'handler_used']        = handler_used
                    df.at[idx, 'download_timestamp']  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                icon = '✓' if status.startswith('success') else ('⊙' if status=='duplicate' else '✗')
                tprint(f"  [{pos}/{total}] NID {nid} {icon} {status}: {message[:80]}")

                with counter_lock:
                    rows_done += 1
                    rd = rows_done

                # Auto-save
                if rd % self.autosave_interval == 0:
                    with self._save_lock:
                        self._auto_save(df, output_file)
                    self._print_progress()

        # Final save
        tprint(f"\n{'='*80}")
        tprint("Saving final results…")
        df.to_excel(output_file, index=False, engine='openpyxl')
        tprint(f"Saved: {output_file}")

        self._print_summary(time.time() - start, output_file)
        self._print_handler_stats()

    # ── Output helpers ────────────────────────────────────────────────────────

    def _auto_save(self, df, path):
        try:
            df.to_excel(path, index=False, engine='openpyxl')
            tprint(f"\n  💾 Auto-saved → {path}")
        except Exception as e:
            tprint(f"  ⚠ Auto-save failed: {e}")

    def _print_progress(self):
        s = self.stats.snapshot()
        tprint(f"\n  Progress — Bibcite: {s['success_bibcite']} | "
               f"PubMed: {s['success_pubmed']} | "
               f"Abstract TXT: {s['success_abstract_txt']} | "
               f"Failed: {s['failed']} | Skipped: {s['skipped']}\n")

    def _print_summary(self, elapsed, output_file):
        s = self.stats.snapshot()
        tprint(f"\n{'='*80}\nCOMPLETE\n{'='*80}")
        tprint(f"Total rows:     {s['total_rows']}")
        tprint(f"Work rows:      {s['infontd_rows']}")
        tprint(f"\nResults:")
        tprint(f"  ✓ Bibcite PDF:   {s['success_bibcite']}")
        tprint(f"  ✓ PubMed PDF:    {s['success_pubmed']}")
        tprint(f"  ✓ Abstract TXT:  {s['success_abstract_txt']}")
        tprint(f"  ⊙ Duplicates:    {s['duplicates']}")
        tprint(f"  ✗ Failed:        {s['failed']}")
        tprint(f"  ○ Skipped:       {s['skipped']}")
        total = s['success_bibcite'] + s['success_pubmed'] + s['success_abstract_txt']
        tprint(f"\nTotal files:    {total}")
        tprint(f"Time:           {elapsed/60:.1f} min")
        tprint(f"Output:         {output_file}\n{'='*80}\n")

    def _print_handler_stats(self):
        hs = self.pdf_finder.handler_stats
        if not hs:
            return
        tprint(f"\n{'='*80}\nHANDLER STATISTICS\n{'='*80}")
        for name, count in sorted(hs.items(), key=lambda x: -x[1]):
            tprint(f"  {name:22} {count:4} uses")
        tprint(f"{'='*80}\n")


# ============================================================================
# Single-URL Test Mode
# ============================================================================

def test_single_url(url, verbose=True):
    tprint(f"\n{'='*80}\nSINGLE-LINK TEST  —  v7\n{'='*80}\n")
    tprint(f"URL: {url}\n")
    finder = EnhancedPDFFinder(verbose=verbose)
    norm = finder.normalize_url(url)
    if norm != url:
        tprint(f"Normalized: {norm}")
    pdf_url = finder.find_pdf(norm)
    if pdf_url and pdf_url.startswith("__BYTES_FILE__:"):
        tmp = pdf_url[len("__BYTES_FILE__:"):]
        sz = os.path.getsize(tmp)
        tprint(f"\n✅ SUCCESS (WHO IRIS Playwright bytes, {sz//1024} KB) → {tmp}")
        return True, tmp, "WHO IRIS download"
    if pdf_url:
        tprint(f"\nFound PDF URL: {pdf_url}")
        try:
            r = requests.head(pdf_url, timeout=10, allow_redirects=True)
            ct = r.headers.get('Content-Type','?')
            tprint(f"HTTP {r.status_code}  Content-Type: {ct}")
            if r.status_code == 200:
                tprint(f"\n✅ SUCCESS\n{'='*80}")
                return True, pdf_url, "Success"
        except Exception as e:
            tprint(f"Verification error: {e}")
        return True, pdf_url, "URL found (verify manually)"
    tprint(f"\n❌ FAILED\n{'='*80}")
    return False, None, "Could not find PDF"


# ============================================================================
# Main  —  edit CONFIG below, then run:  python pdf_scraper_v8.py
# ============================================================================

def main():
    # ┌─────────────────────────────────────────────────────────────────────┐
    # │                        C O N F I G                                  │
    # │  Edit these values to control the scraper. No CLI flags needed.     │
    # └─────────────────────────────────────────────────────────────────────┘
    CONFIG = {

        # ── Files ──────────────────────────────────────────────────────────
        # Path to the input Excel file
        "input_file":           "site_pages_export_processed.xlsx",

        # Output Excel file (None = auto-generate timestamped filename)
        "output_file":          "site_pages_export_processed_v2.xlsx",

        # Folder where downloaded PDFs are saved
        "download_dir":         "infontd_pdfs",

        # Folder where abstract .txt fallback files are saved
        "abstract_dir":         "abstracts",

        # ── Filtering ──────────────────────────────────────────────────────
        # Only process rows where the 'Domains' column matches ANY of these values.
        # Can be a single string or a list of strings — all are OR-combined.
        # Each entry is matched as a substring (case-insensitive).
        # Example:  ["infontd", "infontd_org", "leprosy_information_org"]
        "domain_filter":        ["infontd", "infontd_org | leprosy_information_org"],

        # Only process rows where the 'Type' column contains this string
        # Set to None or "" to process all types
        "type_filter":          "biblio",

        # Skip rows before this position (useful to resume; None = start from 1)
        "start_from_row":       None,

        # ── Performance ────────────────────────────────────────────────────
        # Number of rows processed in parallel
        # Recommended: 5 for normal use, 1 for debugging, 10 for fast machines
        "workers":              5,

        # ── Playwright ─────────────────────────────────────────────────────
        # Use Playwright as a generic fallback for ALL URLs when requests fails.
        # WHO IRIS always uses Playwright internally regardless of this setting.
        # Disable if you want faster runs and don't need JS-challenge bypass.
        # Requires:  pip install playwright && playwright install chromium
        "playwright_fallback":  True,

        # Max concurrent Playwright browser sessions (each uses ~200MB RAM)
        "playwright_sessions":  4,

        # Max random delay (seconds) between requests per worker
        # Keeps traffic polite. 0 = no delay (aggressive)
        "request_delay":        4.0,

        # Auto-save Excel output every N rows processed
        "autosave_interval":    50,

        # ── API Keys ───────────────────────────────────────────────────────
        # NCBI API key for PubMed / PMC (higher rate limits)
        "ncbi_api_key":         "72f60e78d388b5c3a3c46dc5854038099c08",

        # ── Proxy Settings ─────────────────────────────────────────────────
        # Enable free proxy rotation (requires:  pip install free-proxy)
        # Proxies are rotated automatically on HTTP 403 / 429 responses
        "use_free_proxies":     True,

        # List of your own proxy URLs to use (tried first, round-robin)
        # Format: ["http://host:port", "http://user:pass@host:port"]
        # Leave empty [] to skip manual proxies
        "manual_proxies":       [],

        # Enable Tor exit-node rotation (requires Tor running + pip install stem)
        # Sends NEWNYM signal on blocked requests to get a new exit IP
        "use_tor":              False,

        # Tor SOCKS proxy port (default Tor port)
        "tor_socks_port":       9050,

        # Tor control port (used to send NEWNYM)
        "tor_control_port":     9051,

        # Tor control password (None if no password set in torrc)
        "tor_password":         None,

        # How many times a proxy can fail before it's blacklisted
        "proxy_max_failures":   3,

        # ── Debug ──────────────────────────────────────────────────────────
        # Print detailed per-URL handler logs
        "verbose":              True,

        # Set to a URL string to test a single URL and exit
        # Example:  "http://apps.who.int/iris/bitstream/10665/258983/1/9789241550116-eng.pdf"
        "test_url":             None,
    }
    # └─────────────────────────────────────────────────────────────────────┘

    # Allow --test-url as the only CLI override (convenient for quick tests)
    if len(sys.argv) > 1:
        p = argparse.ArgumentParser(add_help=True)
        p.add_argument('--test-url', type=str, default=None)
        args, _ = p.parse_known_args()
        if args.test_url:
            CONFIG["test_url"] = args.test_url.lstrip('-')

    # ── Apply playwright semaphore from config ───────────────────────────
    global _PLAYWRIGHT_SEMAPHORE
    _PLAYWRIGHT_SEMAPHORE = threading.Semaphore(CONFIG["playwright_sessions"])

    # ── Single-URL test mode ─────────────────────────────────────────────
    if CONFIG["test_url"]:
        test_single_url(CONFIG["test_url"], verbose=True)
        return

    # ── Build proxy manager ──────────────────────────────────────────────
    proxy_manager = ProxyManager(
        manual_proxies   = CONFIG["manual_proxies"],
        use_free_proxies = CONFIG["use_free_proxies"],
        use_tor          = CONFIG["use_tor"],
        tor_socks_port   = CONFIG["tor_socks_port"],
        tor_control_port = CONFIG["tor_control_port"],
        tor_password     = CONFIG["tor_password"],
        max_failures     = CONFIG["proxy_max_failures"],
        verbose          = CONFIG["verbose"],
    )

    # ── Startup banner ───────────────────────────────────────────────────
    tprint(f"\n{'='*80}")
    tprint(f"PDF Scraper v8  —  Parallel + Playwright + Proxy Rotation")
    tprint(f"{'='*80}")
    tprint(f"  Input file     : {CONFIG['input_file']}")
    tprint(f"  Download dir   : {CONFIG['download_dir']}")
    tprint(f"  Workers        : {CONFIG['workers']}")
    tprint(f"  Request delay  : 0 – {CONFIG['request_delay']}s")
    tprint(f"  Proxy          : {proxy_manager.status()}")
    tprint(f"  Playwright     : {'available' if PLAYWRIGHT_AVAILABLE else 'NOT installed'}")
    tprint(f"  PW fallback    : {'enabled' if CONFIG['playwright_fallback'] and PLAYWRIGHT_AVAILABLE else 'disabled'}")
    tprint(f"  Verbose        : {CONFIG['verbose']}")
    tprint(f"{'='*80}\n")

    if not os.path.exists(CONFIG["input_file"]):
        tprint(f"❌  Input file not found: {CONFIG['input_file']}")
        tprint(f"    Edit CONFIG['input_file'] in main() and re-run.")
        sys.exit(1)

    # ── Run ──────────────────────────────────────────────────────────────
    scraper = ExcelPDFScraper(
        download_dir     = CONFIG["download_dir"],
        abstract_dir     = CONFIG["abstract_dir"],
        delay            = CONFIG["request_delay"],
        ncbi_api_key     = CONFIG["ncbi_api_key"],
        autosave_interval= CONFIG["autosave_interval"],
        start_from_row   = CONFIG["start_from_row"],
        domain_filter    = CONFIG["domain_filter"],
        type_filter      = CONFIG["type_filter"],
        verbose          = CONFIG["verbose"],
        workers          = CONFIG["workers"],
        proxy_manager    = proxy_manager,
        playwright_fallback = CONFIG["playwright_fallback"],
    )
    scraper.process_excel_file(CONFIG["input_file"], CONFIG["output_file"])


if __name__ == "__main__":
    main()