"""
Excel-Based PDF Scraper - v4
Updated for new input format:
  1. Try "Bibcite URL" → download PDF
  2. Try "PubMed URL"  → download PDF
  3. Fallback: save "Abstract (English)" as .txt file
Outputs enriched Excel with full status info per row.
"""

import os
import re
import time
import requests
from urllib.parse import urljoin, urlparse, unquote
from urllib.request import urlopen
from bs4 import BeautifulSoup
from pathlib import Path
import hashlib
import PyPDF2
import pandas as pd
from datetime import datetime
import tarfile
import zipfile

# Import journal handler if available
try:
    from journal_handler import JournalSpecificHandler
    JOURNAL_HANDLER_AVAILABLE = True
except Exception:
    JOURNAL_HANDLER_AVAILABLE = False


# ============================================================================
# PDF Validator
# ============================================================================

class PDFValidator:
    @staticmethod
    def validate_pdf(filepath):
        try:
            with open(filepath, 'rb') as f:
                f.seek(0, 2)
                size = f.tell()
                f.seek(0)
                if size < 100:
                    return False, "File too small"
                header = f.read(5)
                if header != b'%PDF-':
                    return False, "Invalid PDF header"
                f.seek(0)
                try:
                    pdf_reader = PyPDF2.PdfReader(f)
                    num_pages = len(pdf_reader.pages)
                    if num_pages == 0:
                        return False, "Zero pages"
                    return True, f"{num_pages} pages"
                except Exception as e:
                    return False, f"PyPDF2 error: {str(e)[:50]}"
        except Exception as e:
            return False, f"Read error: {str(e)[:50]}"

    @staticmethod
    def get_file_hash(filepath):
        hasher = hashlib.md5()
        with open(filepath, 'rb') as f:
            buf = f.read(65536)
            while buf:
                hasher.update(buf)
                buf = f.read(65536)
        return hasher.hexdigest()

    @staticmethod
    def sanitize_filename(url, nid=None, title=None):
        if title and str(title).strip():
            # Use title: replace spaces/special chars with underscores, strip punctuation
            clean_title = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean_title = re.sub(r'[\s]+', '_', clean_title)
            clean_title = clean_title.strip('_')
            clean_title = clean_title[:150]  # cap title length
            filename = f"nid_{nid}_{clean_title}.pdf" if nid else f"{clean_title}.pdf"
        else:
            # Fallback to URL-derived name
            parsed = urlparse(url)
            filename = os.path.basename(unquote(parsed.path))
            if not filename or len(filename) < 5:
                domain = parsed.netloc.replace('www.', '')
                path_hash = hashlib.md5(url.encode()).hexdigest()[:8]
                filename = f"{domain}_{path_hash}.pdf"
            if not filename.lower().endswith('.pdf'):
                filename += '.pdf'
            if nid:
                filename = f"nid_{nid}_{filename}"
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        return filename[:255]


# ============================================================================
# PMC Downloader
# ============================================================================

class PMCDownloader:
    def __init__(self, api_key=None):
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0'
        })

    def is_pmc_url(self, url):
        return 'pmc.ncbi.nlm.nih.gov' in url or 'ncbi.nlm.nih.gov/pmc' in url

    def get_pmc_id(self, url):
        match = re.search(r'PMC(\d+)', url)
        return match.group(1) if match else None

    def get_pdf_link_from_oa_service(self, pmc_id):
        url = "https://www.ncbi.nlm.nih.gov/pmc/utils/oa/oa.fcgi"
        params = {'id': f'PMC{pmc_id}'}
        if self.api_key:
            params['api_key'] = self.api_key
        try:
            response = self.session.get(url, params=params, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'xml')
            for link in soup.find_all('link'):
                href = link.get('href', '')
                if href:
                    return href
        except Exception:
            pass
        return None

    def download_from_ftp(self, ftp_url, output_path):
        try:
            archive_path = output_path + '.tar.gz'
            with urlopen(ftp_url, timeout=60) as response:
                with open(archive_path, 'wb') as f:
                    while True:
                        chunk = response.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
            extracted = self._extract_pdf_from_archive(archive_path, output_path)
            if os.path.exists(archive_path):
                os.remove(archive_path)
            return extracted
        except Exception:
            if os.path.exists(output_path + '.tar.gz'):
                os.remove(output_path + '.tar.gz')
            return False

    def _extract_pdf_from_archive(self, archive_path, output_path):
        try:
            if tarfile.is_tarfile(archive_path):
                with tarfile.open(archive_path, 'r:*') as tar:
                    pdf_members = [m for m in tar.getmembers() if m.name.endswith('.pdf')]
                    if not pdf_members:
                        return False
                    pdf_member = pdf_members[0]
                    for member in pdf_members:
                        if member.name.count('/') <= 2:
                            pdf_member = member
                            break
                    pdf_file = tar.extractfile(pdf_member)
                    if pdf_file:
                        with open(output_path, 'wb') as f:
                            f.write(pdf_file.read())
                        with open(output_path, 'rb') as f:
                            if f.read(5) == b'%PDF-':
                                return True
                        os.remove(output_path)
            elif zipfile.is_zipfile(archive_path):
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    pdf_files = [f for f in zip_ref.namelist() if f.endswith('.pdf')]
                    if not pdf_files:
                        return False
                    with zip_ref.open(pdf_files[0]) as pdf:
                        with open(output_path, 'wb') as f:
                            f.write(pdf.read())
                    with open(output_path, 'rb') as f:
                        if f.read(5) == b'%PDF-':
                            return True
                    os.remove(output_path)
        except Exception:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception:
                    pass
        return False

    def download(self, url, output_path):
        pmc_id = self.get_pmc_id(url)
        if not pmc_id:
            return False, "Could not extract PMC ID"
        pdf_url = self.get_pdf_link_from_oa_service(pmc_id)
        if not pdf_url:
            patterns = [
                f"https://www.ncbi.nlm.nih.gov/pmc/articles/PMC{pmc_id}/pdf/",
                f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{pmc_id}/pdf/",
            ]
            for pattern in patterns:
                try:
                    params = {'api_key': self.api_key} if self.api_key else {}
                    response = self.session.head(pattern, params=params, timeout=5)
                    if response.status_code == 200:
                        pdf_url = pattern
                        break
                except Exception:
                    continue
        if not pdf_url:
            return False, "Could not find PDF link"
        if pdf_url.startswith('ftp://'):
            if self.download_from_ftp(pdf_url, output_path):
                return True, "Downloaded from FTP archive"
            else:
                return False, "FTP download failed"
        else:
            try:
                params = {'api_key': self.api_key} if self.api_key else {}
                response = self.session.get(pdf_url, params=params, timeout=30, stream=True)
                response.raise_for_status()
                content_type = response.headers.get('Content-Type', '').lower()
                if 'pdf' in content_type:
                    with open(output_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    return True, "Downloaded from HTTPS"
                elif 'gzip' in content_type or 'tar' in content_type or 'octet-stream' in content_type:
                    archive_path = output_path + '.archive'
                    with open(archive_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    extracted = self._extract_pdf_from_archive(archive_path, output_path)
                    os.remove(archive_path)
                    if extracted:
                        return True, "Extracted from archive"
                    else:
                        return False, "Could not extract PDF"
                else:
                    return False, f"Unexpected content type: {content_type}"
            except Exception as e:
                return False, f"Download error: {str(e)[:100]}"


# ============================================================================
# Universal PDF Finder
# ============================================================================

class UniversalPDFFinder:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        if JOURNAL_HANDLER_AVAILABLE:
            self.journal_handler = JournalSpecificHandler()
        else:
            self.journal_handler = None

    def verify_pdf_url(self, url):
        try:
            response = self.session.head(url, timeout=5, allow_redirects=True)
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '').lower()
                if 'pdf' in content_type and 'html' not in content_type:
                    return True
        except Exception:
            pass
        return False

    def find_pdf(self, url):
        if self.journal_handler and self.journal_handler.can_handle(url):
            try:
                pdf_url = self.journal_handler.find_pdf(url)
                if pdf_url:
                    return pdf_url
            except Exception:
                pass
        if url.lower().endswith('.pdf'):
            if self.verify_pdf_url(url):
                return url
        transformations = [
            lambda u: u.rstrip('/') + '/pdf',
            lambda u: u.rstrip('/') + '.pdf',
            lambda u: u.replace('/article/', '/pdf/'),
            lambda u: u.replace('/doi/full/', '/doi/pdf/'),
        ]
        for transform in transformations:
            try:
                pdf_url = transform(url)
                if pdf_url != url and self.verify_pdf_url(pdf_url):
                    return pdf_url
            except Exception:
                continue
        try:
            response = self.session.get(url, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            for link in soup.find_all('a', href=True):
                href = link['href']
                text = link.get_text().lower()
                if '.pdf' in href.lower() or 'pdf' in text or 'download' in text:
                    absolute_url = urljoin(url, href)
                    if self.verify_pdf_url(absolute_url):
                        return absolute_url
        except Exception:
            pass
        return None


# ============================================================================
# Main Scraper
# ============================================================================

class ExcelPDFScraper:
    """
    Scraper with URL fallback chain:
      1. Bibcite URL  → try to download PDF
      2. PubMed URL   → try to download PDF
      3. Abstract (English) → save as .txt file
    """

    def __init__(self, download_dir="infontd_pdfs", abstract_dir="infontd_abstracts", delay=3, ncbi_api_key=None,
                 autosave_interval=50, start_from_row=None, type_filter=None):
        self.download_dir = download_dir
        self.abstract_dir = abstract_dir
        self.delay = delay
        self.autosave_interval = autosave_interval
        self.start_from_row = start_from_row
        self.type_filter = type_filter  # e.g. "biblio" to only process rows where Type contains this string

        Path(self.download_dir).mkdir(parents=True, exist_ok=True)
        Path(self.abstract_dir).mkdir(parents=True, exist_ok=True)

        self.validator = PDFValidator()
        self.pmc_downloader = PMCDownloader(api_key=ncbi_api_key)
        self.universal_finder = UniversalPDFFinder()

        self.downloaded_hashes = set()
        self.processed_urls = set()

        self.stats = {
            'total_rows': 0,
            'infontd_rows': 0,
            'valid_urls': 0,
            'already_downloaded': 0,
            'duplicate_urls': 0,
            'manually_skipped': 0,
            'success_bibcite': 0,
            'success_pubmed': 0,
            'success_abstract_txt': 0,
            'abstracts_deleted': 0,
            'failed': 0,
            'skipped': 0,
            'duplicates': 0,
            'corrupt': 0,
            'pmc_downloads': 0,
        }

    def is_already_downloaded(self, row):
        """Check if a PDF was successfully downloaded (excludes abstracts)."""
        return (pd.notna(row.get('download_status')) and
                row.get('download_status') in ('success_bibcite', 'success_pubmed'))

    def delete_abstract_if_exists(self, nid, title=None):
        """Delete abstract .txt file if it exists. Returns True if deleted."""
        # Try to construct the abstract filename the same way save_abstract_txt does
        if title and str(title).strip():
            clean_title = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean_title = re.sub(r'[\s]+', '_', clean_title).strip('_')[:150]
            filename = f"nid_{nid}_{clean_title}.txt"
        else:
            filename = f"nid_{nid}_abstract.txt"
        
        filepath = os.path.join(self.abstract_dir, filename)
        
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                self.stats['abstracts_deleted'] += 1
                return True
            except Exception as e:
                print(f"  ⚠️ Could not delete abstract: {str(e)[:50]}")
                return False
        return False

    def _valid_url(self, val):
        if pd.isna(val) or not val or str(val).strip() == '':
            return None
        url = str(val).strip()
        if not url.startswith(('http://', 'https://')):
            if url.startswith('www.'):
                url = 'http://' + url
            else:
                return None
        return url

    def download_pdf(self, url, nid, title=None):
        """Try to download a PDF from url. Returns (status, filepath, message)."""
        filename = self.validator.sanitize_filename(url, nid, title=title)
        filepath = os.path.join(self.download_dir, filename)

        if self.pmc_downloader.is_pmc_url(url):
            self.stats['pmc_downloads'] += 1
            success, message = self.pmc_downloader.download(url, filepath)
            if not success:
                return 'failed', None, f'PMC: {message} (URL: {url[:100]})'
        else:
            pdf_url = self.universal_finder.find_pdf(url)
            if not pdf_url:
                return 'failed', None, f'Could not find PDF URL on page (Source URL: {url[:100]})'
            try:
                response = self.universal_finder.session.get(pdf_url, timeout=30, stream=True)
                response.raise_for_status()
                
                # Check content type
                content_type = response.headers.get('Content-Type', 'unknown')
                if 'html' in content_type.lower() and 'pdf' not in content_type.lower():
                    return 'failed', None, f'Received HTML instead of PDF (Content-Type: {content_type}, URL: {pdf_url[:100]})'
                
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                        
            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code if e.response else 'unknown'
                return 'failed', None, f'HTTP {status_code} error: {str(e)[:100]} (URL: {pdf_url[:100]})'
            except requests.exceptions.Timeout:
                return 'failed', None, f'Request timeout after 30s (URL: {pdf_url[:100]})'
            except requests.exceptions.ConnectionError as e:
                return 'failed', None, f'Connection error: {str(e)[:100]} (URL: {pdf_url[:100]})'
            except Exception as e:
                error_type = type(e).__name__
                return 'failed', None, f'{error_type}: {str(e)[:100]} (URL: {pdf_url[:100]})'

        is_valid, validation_msg = self.validator.validate_pdf(filepath)
        if not is_valid:
            file_size = os.path.getsize(filepath) if os.path.exists(filepath) else 0
            if os.path.exists(filepath):
                os.remove(filepath)
            return 'corrupt', None, f'Invalid PDF: {validation_msg} (Size: {file_size} bytes, URL: {url[:100]})'

        file_hash = self.validator.get_file_hash(filepath)
        if file_hash in self.downloaded_hashes:
            os.remove(filepath)
            return 'duplicate', None, f'Duplicate file (MD5: {file_hash[:8]}...)'

        self.downloaded_hashes.add(file_hash)
        file_size = os.path.getsize(filepath) / 1024
        return 'success', filepath, f'{validation_msg}, {file_size:.1f} KB'

    def save_abstract_txt(self, nid, abstract_text, title=None):
        """Save abstract as a .txt file. Returns (status, filepath, message)."""
        if title and str(title).strip():
            clean_title = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean_title = re.sub(r'[\s]+', '_', clean_title).strip('_')[:150]
            filename = f"nid_{nid}_{clean_title}.txt"
        else:
            filename = f"nid_{nid}_abstract.txt"
        filepath = os.path.join(self.abstract_dir, filename)
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(str(abstract_text))
            file_size = os.path.getsize(filepath)
            return 'success_abstract_txt', filepath, f'Abstract saved as TXT ({file_size} bytes)'
        except Exception as e:
            return 'failed', None, f'Could not save abstract: {str(e)[:100]}'
    
    def delete_abstract_file(self, nid, title=None):
        """Delete abstract.txt file if it exists (from previous run)."""
        # Try both filename patterns
        if title and str(title).strip():
            clean_title = re.sub(r'[^\w\s-]', '', str(title).strip())
            clean_title = re.sub(r'[\s]+', '_', clean_title).strip('_')[:150]
            filename1 = f"nid_{nid}_{clean_title}.txt"
        else:
            filename1 = f"nid_{nid}_abstract.txt"
        
        filename2 = f"nid_{nid}_abstract.txt"  # Fallback pattern
        
        deleted = False
        for filename in [filename1, filename2]:
            filepath = os.path.join(self.abstract_dir, filename)
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                    deleted = True
                    self.stats['abstracts_deleted'] += 1
                    print(f"  🗑️  Deleted previous abstract: {filename}")
                except Exception as e:
                    print(f"  ⚠️  Could not delete {filename}: {str(e)[:50]}")
        
        return deleted

    def process_row(self, row):
        """
        Process one row with fallback chain.
        Returns: (status, filepath, url_used, source_used, message, detailed_errors)
        """
        # Already done?
        if self.is_already_downloaded(row):
            return (row.get('download_status'), row.get('download_filepath'),
                    row.get('url_used', ''), row.get('source_used', ''), 
                    'Previously downloaded', row.get('detailed_errors', ''))

        nid = row.get('Node ID')
        title = row.get('Title')
        bibcite_url = self._valid_url(row.get('Bibcite URL'))
        pubmed_url = self._valid_url(row.get('PubMed URL'))
        abstract = row.get('Abstract (English)')
        has_abstract = pd.notna(abstract) and str(abstract).strip() != ''

        # Track all errors from each source
        error_log = []

        # ── Step 1: Bibcite URL ──────────────────────────────────────────────
        if bibcite_url:
            if bibcite_url in self.processed_urls:
                self.stats['duplicate_urls'] += 1
                error_log.append(f"Bibcite: Duplicate URL (already processed)")
                # Still try pubmed
            else:
                self.processed_urls.add(bibcite_url)
                status, filepath, msg = self.download_pdf(bibcite_url, nid, title=title)
                if status == 'success':
                    self.stats['success_bibcite'] += 1
                    # Delete abstract file if it exists from previous run
                    self.delete_abstract_file(nid, title)
                    return 'success_bibcite', filepath, bibcite_url, 'Bibcite URL', msg, ''
                elif status == 'duplicate':
                    self.stats['duplicates'] += 1
                    return 'duplicate', filepath, bibcite_url, 'Bibcite URL', msg, ''
                else:
                    # Failed or corrupt - log the error
                    error_log.append(f"Bibcite: {msg}")
        else:
            error_log.append("Bibcite: No URL provided")

        # ── Step 2: PubMed URL ───────────────────────────────────────────────
        if pubmed_url:
            if pubmed_url not in self.processed_urls:
                self.processed_urls.add(pubmed_url)
                status, filepath, msg = self.download_pdf(pubmed_url, nid, title=title)
                if status == 'success':
                    self.stats['success_pubmed'] += 1
                    detailed_errors = ' | '.join(error_log) if error_log else ''
                    # Delete abstract file if it exists from previous run
                    self.delete_abstract_file(nid, title)
                    return 'success_pubmed', filepath, pubmed_url, 'PubMed URL', msg, detailed_errors
                elif status == 'duplicate':
                    self.stats['duplicates'] += 1
                    detailed_errors = ' | '.join(error_log) if error_log else ''
                    return 'duplicate', filepath, pubmed_url, 'PubMed URL', msg, detailed_errors
                else:
                    # Failed or corrupt
                    error_log.append(f"PubMed: {msg}")
            else:
                error_log.append("PubMed: Duplicate URL (already processed)")
        else:
            error_log.append("PubMed: No URL provided")

        # ── Step 3: Abstract as .txt ─────────────────────────────────────────
        if has_abstract:
            status, filepath, msg = self.save_abstract_txt(nid, abstract, title=title)
            detailed_errors = ' | '.join(error_log) if error_log else ''
            if status == 'success_abstract_txt':
                self.stats['success_abstract_txt'] += 1
                return 'success_abstract_txt', filepath, '', 'Abstract (English)', msg, detailed_errors
            else:
                self.stats['failed'] += 1
                error_log.append(f"Abstract: {msg}")
                detailed_errors = ' | '.join(error_log)
                return status, None, '', 'Abstract (English)', msg, detailed_errors
        else:
            error_log.append("Abstract: No abstract text available")

        # ── Nothing worked ───────────────────────────────────────────────────
        self.stats['skipped'] += 1
        detailed_errors = ' | '.join(error_log)
        return 'skipped', None, '', 'none', 'No data available', detailed_errors

    def process_excel_file(self, input_file, output_file=None):
        print(f"\n{'='*80}")
        print(f"Excel PDF Scraper v4 - InfoNTD Edition")
        print(f"Fallback chain: Bibcite URL → PubMed URL → Abstract TXT")
        print(f"{'='*80}\n")

        df = pd.read_excel(input_file)
        self.stats['total_rows'] = len(df)
        print(f"Total rows: {len(df)}")

        # Filter infontd rows
        infontd_mask = df['Domains'].str.contains('infontd', case=False, na=False)
        infontd_df = df[infontd_mask]
        self.stats['infontd_rows'] = len(infontd_df)
        print(f"InfoNTD rows: {len(infontd_df)}")

        # Optional: filter by Type column
        if self.type_filter:
            type_mask = infontd_df['Type'].str.contains(self.type_filter, case=False, na=False)
            infontd_df = infontd_df[type_mask]
            print(f"After type filter '{self.type_filter}': {len(infontd_df)} rows")

        # Count rows with at least one usable source
        has_data = (
            infontd_df['Bibcite URL'].notna() |
            infontd_df['PubMed URL'].notna() |
            infontd_df['Abstract (English)'].notna()
        ).sum()
        self.stats['valid_urls'] = has_data
        print(f"Rows with at least one data source: {has_data}")

        # Initialize / ensure output columns
        for col in ['download_status', 'download_filepath', 'download_filename', 'url_used',
                    'source_used', 'download_error', 'detailed_errors', 'download_timestamp']:
            if col not in df.columns:
                df[col] = ''

        already_success = df['download_status'].isin(
            ['success_bibcite', 'success_pubmed', 'success_abstract_txt']
        ).sum()
        if already_success > 0:
            print(f"Found {already_success} already processed — will skip\n")

        print(f"{'='*80}")
        print(f"Starting... Auto-save every {self.autosave_interval} rows")
        if self.start_from_row:
            print(f"Starting from row {self.start_from_row}")
        print(f"{'='*80}\n")

        start_time = time.time()
        processed = 0
        rows_updated = 0

        for idx in infontd_df.index:
            row = df.loc[idx]
            processed += 1

            if self.start_from_row and processed < self.start_from_row:
                self.stats['manually_skipped'] += 1
                if processed == 1 or processed % 100 == 0:
                    print(f"[{processed}/{len(infontd_df)}] Skipping (before start position)...")
                continue

            nid = row.get('Node ID', idx)
            bibcite = str(row.get('Bibcite URL', ''))[:50]
            print(f"[{processed}/{len(infontd_df)}] NID {nid}: {bibcite}...")

            status, filepath, url_used, source_used, message, detailed_errors = self.process_row(row)

            if status in ('already_downloaded',):
                self.stats['already_downloaded'] += 1
                print(f"  ○ SKIP: {message}")
                continue

            if status == 'skipped_duplicate_url':
                print(f"  ⊙ DUPLICATE URL: {message}")
                continue

            # Update dataframe
            df.at[idx, 'download_status'] = status
            df.at[idx, 'download_filepath'] = filepath if filepath else ''
            df.at[idx, 'download_filename'] = os.path.basename(filepath) if filepath else ''
            df.at[idx, 'url_used'] = url_used
            df.at[idx, 'source_used'] = source_used
            df.at[idx, 'download_error'] = message
            df.at[idx, 'detailed_errors'] = detailed_errors
            df.at[idx, 'download_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            rows_updated += 1

            # Print result
            icons = {
                'success_bibcite': '✓ BIBCITE',
                'success_pubmed': '✓ PUBMED',
                'success_abstract_txt': '✓ ABSTRACT TXT',
                'duplicate': '⊙ DUPLICATE',
                'corrupt': '✗ CORRUPT',
                'skipped': '○ SKIP',
                'failed': '✗ FAILED',
            }
            icon = icons.get(status, f'? {status}')
            print(f"  {icon}: {message}")
            if url_used:
                print(f"    URL: {url_used[:80]}")

            # Auto-save
            if rows_updated > 0 and rows_updated % self.autosave_interval == 0:
                self._auto_save(df, output_file or self._default_output(input_file), rows_updated)

            # Progress summary
            if processed % 50 == 0:
                self._print_progress()

            time.sleep(self.delay)

        # Final save
        if output_file is None:
            output_file = self._default_output(input_file)

        print(f"\n{'='*80}")
        print(f"Saving final results to: {output_file}")
        df.to_excel(output_file, index=False)
        print(f"✓ Saved!")

        elapsed = time.time() - start_time
        self._print_summary(elapsed, output_file)
        return output_file

    def _default_output(self, input_file):
        base, ext = os.path.splitext(input_file)
        return f"{base}_processed{ext}"

    def _auto_save(self, df, output_file, rows_updated):
        try:
            print(f"\n  💾 AUTO-SAVING... ({rows_updated} rows updated)")
            df.to_excel(output_file, index=False)
            print(f"  ✓ Auto-saved to: {output_file}\n")
        except Exception as e:
            print(f"  ⚠ Auto-save failed: {str(e)[:50]}\n")

    def _print_progress(self):
        print(f"\n  Progress — Bibcite: {self.stats['success_bibcite']} | "
              f"PubMed: {self.stats['success_pubmed']} | "
              f"Abstract TXT: {self.stats['success_abstract_txt']} | "
              f"Failed: {self.stats['failed']} | "
              f"Skipped: {self.stats['skipped']}\n")

    def _print_summary(self, elapsed, output_file):
        print(f"\n{'='*80}")
        print(f"COMPLETE")
        print(f"{'='*80}\n")
        print(f"Input rows:     {self.stats['total_rows']}")
        print(f"InfoNTD rows:   {self.stats['infontd_rows']}")
        print(f"With data:      {self.stats['valid_urls']}")
        if self.stats['manually_skipped']:
            print(f"Skipped (start pos): {self.stats['manually_skipped']}")
        if self.stats['already_downloaded']:
            print(f"Already done:   {self.stats['already_downloaded']}")
        print(f"\nResults:")
        print(f"  ✓ Bibcite PDF:   {self.stats['success_bibcite']}")
        print(f"  ✓ PubMed PDF:    {self.stats['success_pubmed']}")
        print(f"  ✓ Abstract TXT:  {self.stats['success_abstract_txt']}")
        if self.stats['abstracts_deleted']:
            print(f"  🗑️  Abstracts deleted (replaced with PDFs): {self.stats['abstracts_deleted']}")
        print(f"  ⊙ Duplicates:    {self.stats['duplicates']}")
        print(f"  ✗ Failed:        {self.stats['failed']}")
        print(f"  ○ Skipped:       {self.stats['skipped']} (no data at all)")
        print(f"  PMC downloads:   {self.stats['pmc_downloads']}")
        total = (self.stats['success_bibcite'] + self.stats['success_pubmed'] +
                 self.stats['success_abstract_txt'] + self.stats['already_downloaded'])
        print(f"\nTotal files:    {total}")
        print(f"Time:           {elapsed/60:.1f} min")
        print(f"Output Excel:   {output_file}")
        print(f"\nOutput columns added/updated:")
        print(f"  download_status   — success_bibcite / success_pubmed / success_abstract_txt / failed / skipped / duplicate / corrupt")
        print(f"  source_used       — which source was used (Bibcite URL / PubMed URL / Abstract (English) / none)")
        print(f"  url_used          — the actual URL that worked (empty for abstract TXT)")
        print(f"  download_filename — filename only, e.g. nid_123_article.pdf")
        print(f"  download_filepath — full local path to the saved file")
        print(f"  download_error    — details / error message")
        print(f"  detailed_errors   — all errors from each source attempted (Bibcite | PubMed | Abstract)")
        print(f"  download_timestamp")
        print(f"\n{'='*80}\n")


# ============================================================================
# Main
# ============================================================================

def main():
    NCBI_API_KEY = "72f60e78d388b5c3a3c46dc5854038099c08"

    scraper = ExcelPDFScraper(
        download_dir="infontd_pdfs",
        abstract_dir="abstracts",
        delay=3,
        ncbi_api_key=NCBI_API_KEY,
        autosave_interval=50,
        start_from_row=None,  # Set to a row number to resume from a specific position
        type_filter="biblio"  # Set to None to process all types, or e.g. "biblio" to filter by Type column
    )

    print(f"NCBI API Key: {'*' * 28}{NCBI_API_KEY[-8:]}")
    print(f"Auto-save: Every 50 rows")

    scraper.process_excel_file('site_pages_export.xlsx')


if __name__ == "__main__":
    main()
