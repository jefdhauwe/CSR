"""
Journal-Specific Handler
Handles PDF finding for journals with predictable URL patterns
"""

import re
import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup


class JournalSpecificHandler:
    """
    Handles PDF downloads for journals with specific URL patterns.
    Each journal has custom logic for finding PDFs.
    """
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        })
        
        # Map domains to handler methods
        self.handlers = {
            'leprosyreview.org': self._handle_leprosyreview,
            'journals.plos.org': self._handle_plos,
            'plosntds.org': self._handle_plos,  # Same handler as PLOS
            'bmj.com': self._handle_bmj,
            'thelancet.com': self._handle_lancet,
            'springer.com': self._handle_springer,
            'link.springer.com': self._handle_springer,
            'wiley.com': self._handle_wiley,
            'onlinelibrary.wiley.com': self._handle_wiley,
            'tandfonline.com': self._handle_tandfonline,
            'sciencedirect.com': self._handle_sciencedirect,
            'tinyurl.com': self._handle_tinyurl,  # NEW
            'apps.who.int': self._handle_who_iris,  # NEW
            'who.int': self._handle_who_int,  # NEW
            'washntds.org': self._handle_washntds,  # NEW
            'mdpi.com': self._handle_mdpi,  # NEW
            'academic.oup.com': self._handle_oup,  # NEW
            'ajtmh.org': self._handle_ajtmh,  # NEW
            'dcidj.org': self._handle_dcidj,  # NEW
        }
    
    def can_handle(self, url):
        """Check if this handler can process the URL."""
        domain = urlparse(url).netloc.lower()
        for supported_domain in self.handlers.keys():
            if supported_domain in domain:
                return True
        return False
    
    def find_pdf(self, url):
        """
        Find PDF URL for a journal article.
        Returns: PDF URL or None
        """
        domain = urlparse(url).netloc.lower()
        
        # Find the right handler
        for supported_domain, handler_func in self.handlers.items():
            if supported_domain in domain:
                try:
                    return handler_func(url)
                except Exception as e:
                    # If specific handler fails, try generic scraping
                    return self._generic_scrape(url)
        
        return None
    
    def verify_pdf_url(self, url, use_get=False):
        """Verify that a URL returns a PDF."""
        try:
            if use_get:
                response = self.session.get(url, timeout=5, stream=True, allow_redirects=True)
            else:
                response = self.session.head(url, timeout=5, allow_redirects=True)
            
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '').lower()
                if 'pdf' in content_type and 'html' not in content_type:
                    return True
        except:
            pass
        return False
    
    # ========================================================================
    # Journal-Specific Handlers
    # ========================================================================
    
    def _handle_leprosyreview(self, url):
        """
        Handle leprosyreview.org URLs.
        Pattern: /article/80/2/19-7204
        """
        # Extract article components
        match = re.search(r'/article/(\d+)/(\d+)/(\d+-\d+)', url)
        if not match:
            return self._generic_scrape(url)
        
        volume, issue, page_id = match.groups()
        
        # Try multiple PDF URL patterns
        patterns = [
            # Pattern 1: /article/pdf/...
            f"https://leprosyreview.org/article/pdf/{volume}/{issue}/{page_id}",
            
            # Pattern 2: /pdf/...
            f"https://leprosyreview.org/pdf/{volume}/{issue}/{page_id}",
            
            # Pattern 3: /admin/public/pdf/...
            f"https://leprosyreview.org/admin/public/pdf/{volume}/{issue}/{page_id}",
            
            # Pattern 4: /admin/public/assets/pdf/...
            f"https://leprosyreview.org/admin/public/assets/pdf/{volume}/{issue}/{page_id}.pdf",
            
            # Pattern 5: /assets/pdf/...
            f"https://leprosyreview.org/assets/pdf/{volume}/{issue}/{page_id}.pdf",
            
            # Pattern 6: /assets/articles/...
            f"https://leprosyreview.org/assets/articles/{volume}/{issue}/{page_id}.pdf",
            
            # Pattern 7: Direct file with underscores
            f"https://leprosyreview.org/articles/{volume}_{issue}_{page_id}.pdf",
            
            # Pattern 8: Admin storage
            f"https://leprosyreview.org/admin/public/storage/articles/{volume}_{issue}_{page_id}.pdf",
        ]
        
        for pdf_url in patterns:
            if self.verify_pdf_url(pdf_url):
                return pdf_url
            # Try GET if HEAD fails (some servers block HEAD)
            if self.verify_pdf_url(pdf_url, use_get=True):
                return pdf_url
        
        # Fallback: scrape the page
        return self._generic_scrape(url)
    
    def _handle_plos(self, url):
        """
        Handle PLOS journals (including plosntds.org).
        Works for: journals.plos.org and plosntds.org
        """
        # Extract DOI
        match = re.search(r'10\.\d+/[^\s&]+', url)
        if match:
            doi = match.group(0)
            
            # Try multiple PLOS PDF patterns
            patterns = [
                f"https://journals.plos.org/plosone/article/file?id={doi}&type=printable",
                f"https://journals.plos.org/plosntds/article/file?id={doi}&type=printable",
                f"https://journals.plos.org/plospathogens/article/file?id={doi}&type=printable",
            ]
            
            for pdf_url in patterns:
                if self.verify_pdf_url(pdf_url):
                    return pdf_url
        
        # Try extracting from URL structure
        if '/article?id=' in url or '/article/' in url:
            # Try adding /file?type=printable
            if '/article?id=' in url:
                pdf_url = url.replace('/article?id=', '/article/file?id=') + '&type=printable'
            else:
                pdf_url = url.rstrip('/') + '/file?type=printable'
            
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return None
    
    def _handle_bmj(self, url):
        """Handle BMJ journals."""
        if '/content/' in url:
            pdf_url = url.rstrip('/') + '.full.pdf'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        return None
    
    def _handle_lancet(self, url):
        """
        Handle The Lancet with multiple URL patterns.
        """
        # Pattern 1: Extract PII from URL
        match = re.search(r'(PIIS\d+-\d+\([^)]+\)[^/]+)', url)
        if match:
            article_id = match.group(1)
            pdf_url = f"https://www.thelancet.com/action/showPdf?pii={article_id}"
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Pattern 2: Try /pdfs/ directory
        if '/journals/' in url or '/article/' in url:
            # Replace /journals/ or /article/ with /pdfs/
            for old_path in ['/journals/', '/article/']:
                if old_path in url:
                    pdf_url = url.replace(old_path, '/pdfs/')
                    if not pdf_url.endswith('.pdf'):
                        pdf_url += '.pdf'
                    if self.verify_pdf_url(pdf_url):
                        return pdf_url
        
        # Pattern 3: Try /retrieve/pii/ pattern
        if 'article/' in url:
            parts = url.split('article/')
            if len(parts) > 1:
                article_part = parts[1].split('/')[0]
                pdf_url = f"{parts[0]}retrieve/pii/{article_part}"
                if self.verify_pdf_url(pdf_url):
                    return pdf_url
        
        return self._generic_scrape(url)
    
    def _handle_springer(self, url):
        """Handle Springer journals."""
        match = re.search(r'10\.\d+/[^\s]+', url)
        if match:
            doi = match.group(0).rstrip('/')
            pdf_url = f"https://link.springer.com/content/pdf/{doi}.pdf"
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        return None
    
    def _handle_wiley(self, url):
        """Handle Wiley journals."""
        if '/doi/' in url and '/pdfdirect/' not in url:
            pdf_url = url.replace('/doi/', '/doi/pdfdirect/')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        return None
    
    def _handle_tandfonline(self, url):
        """
        Handle Taylor & Francis / Tandfonline with multiple patterns.
        """
        # Pattern 1: /doi/full/ → /doi/pdf/
        if '/doi/full/' in url:
            pdf_url = url.replace('/doi/full/', '/doi/pdf/')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Pattern 2: /doi/abs/ → /doi/pdf/
        if '/doi/abs/' in url:
            pdf_url = url.replace('/doi/abs/', '/doi/pdf/')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Pattern 3: Just /doi/ → /doi/pdf/
        if '/doi/' in url and '/doi/pdf/' not in url:
            # Insert 'pdf/' after 'doi/'
            pdf_url = url.replace('/doi/', '/doi/pdf/')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Pattern 4: Add ?download=true parameter
        if '/doi/' in url:
            pdf_url = url.rstrip('/') + '?download=true'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return self._generic_scrape(url)
    
    def _handle_sciencedirect(self, url):
        """
        Handle ScienceDirect with multiple patterns.
        """
        # Pattern 1: Extract PII and use /pdfft
        if '/pii/' in url:
            # Already has PII, try adding /pdfft
            if not url.endswith('/pdfft'):
                pdf_url = url.rstrip('/') + '/pdfft'
                if self.verify_pdf_url(pdf_url):
                    return pdf_url
        
        # Pattern 2: /article/pii/... format
        if '/article/pii/' in url:
            if not url.endswith('/pdfft'):
                pdf_url = url.rstrip('/') + '/pdfft'
                if self.verify_pdf_url(pdf_url):
                    return pdf_url
        
        # Pattern 3: Extract PII from various URL formats
        pii_match = re.search(r'(S\d{4}\w{10,})', url)
        if pii_match:
            pii = pii_match.group(1)
            pdf_url = f"https://www.sciencedirect.com/science/article/pii/{pii}/pdfft"
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Pattern 4: Generic /article/ → /article/pii/
        if '/article/' in url or '/science/article/' in url:
            pdf_url = url.replace('/article/', '/article/pii/').replace('/pii/pii/', '/pii/') + '/pdfft'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return self._generic_scrape(url)
    
    # ========================================================================
    # NEW HANDLERS FOR PROBLEMATIC DOMAINS
    # ========================================================================
    
    def _handle_tinyurl(self, url):
        """
        Handle TinyURL redirects.
        Follow redirect to get real URL, then process normally.
        """
        try:
            # Follow redirects to get real URL
            response = self.session.get(url, timeout=10, allow_redirects=True)
            real_url = response.url
            
            # If we got redirected to a different domain, try to find PDF there
            if real_url != url and 'tinyurl.com' not in real_url:
                # Check if real URL is a direct PDF
                if real_url.lower().endswith('.pdf'):
                    if self.verify_pdf_url(real_url):
                        return real_url
                
                # Try to find PDF on the real URL's page
                return self._generic_scrape(real_url)
        except:
            pass
        return None
    
    def _handle_who_iris(self, url):
        """
        Handle WHO IRIS repository URLs.
        Pattern: apps.who.int/iris/bitstream/handle/...
        """
        try:
            # Try direct access with proper headers
            response = self.session.get(url, timeout=15, allow_redirects=True)
            
            # Check if we got redirected to a PDF
            if response.url.lower().endswith('.pdf'):
                if self.verify_pdf_url(response.url):
                    return response.url
            
            # If URL contains /retrieve, it might be a direct download
            if '/retrieve' in url or '/bitstream/' in url:
                # Try accessing directly
                if self.verify_pdf_url(url, use_get=True):
                    return url
            
            # Try to find PDF link in the page
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Look for download links
            for link in soup.find_all('a', href=True):
                href = link['href']
                if '.pdf' in href.lower() or 'download' in href.lower():
                    absolute_url = urljoin(url, href)
                    if self.verify_pdf_url(absolute_url):
                        return absolute_url
        except:
            pass
        return None
    
    def _handle_who_int(self, url):
        """
        Handle old who.int PDF URLs.
        Many have moved to apps.who.int/iris
        """
        try:
            # Try following redirects
            response = self.session.get(url, timeout=15, allow_redirects=True)
            final_url = response.url
            
            # If redirected to IRIS, use IRIS handler
            if 'apps.who.int/iris' in final_url:
                return self._handle_who_iris(final_url)
            
            # Check if final URL is a PDF
            if final_url.lower().endswith('.pdf'):
                if self.verify_pdf_url(final_url):
                    return final_url
            
            # Try the original URL as-is (might work)
            if self.verify_pdf_url(url, use_get=True):
                return url
        except:
            pass
        return None
    
    def _handle_washntds(self, url):
        """Handle WASHNTDs.org URLs."""
        # Try common patterns
        patterns = [
            url.replace('/article/', '/pdf/'),
            url.replace('/post/', '/pdf/'),
            url + '.pdf' if not url.endswith('.pdf') else url,
        ]
        
        for pdf_url in patterns:
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        # Fallback to scraping
        return self._generic_scrape(url)
    
    def _handle_mdpi(self, url):
        """
        Handle MDPI journal URLs.
        Pattern: mdpi.com/journal/article/12345 → mdpi.com/journal/article/12345/pdf
        """
        if '/article/' in url or '/articles/' in url:
            # Try adding /pdf
            pdf_url = url.rstrip('/') + '/pdf'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
            
            # Try /pdf?version=...
            pdf_url = url.rstrip('/') + '/pdf?download=true'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return self._generic_scrape(url)
    
    def _handle_oup(self, url):
        """
        Handle Oxford University Press (academic.oup.com).
        Pattern: /article/123/456/789 → /article-pdf/123/456/789
        """
        if '/article/' in url:
            # Try replacing /article/ with /article-pdf/
            pdf_url = url.replace('/article/', '/article-pdf/')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
            
            # Try adding /article-pdf at end
            pdf_url = url.rstrip('/') + '/article-pdf'
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return self._generic_scrape(url)
    
    def _handle_ajtmh(self, url):
        """
        Handle American Journal of Tropical Medicine and Hygiene.
        Pattern: ajtmh.org/view/journals/tpmd/123/4/article-p789.xml
        """
        if '/view/journals/' in url:
            # Replace .xml with .pdf
            pdf_url = url.replace('.xml', '.pdf')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
            
            # Try adding /pdf
            pdf_url = url.replace('/article-', '/pdf/article-')
            if self.verify_pdf_url(pdf_url):
                return pdf_url
        
        return self._generic_scrape(url)
    
    def _handle_dcidj(self, url):
        """Handle DCIDJ (Disability, CBR & Inclusive Development) journal."""
        # Try common patterns
        if '/article/view/' in url:
            # Pattern: /article/view/123 → /article/download/123/456
            parts = url.split('/article/view/')
            if len(parts) == 2:
                article_id = parts[1].split('/')[0]
                pdf_url = f"{parts[0]}/article/download/{article_id}/{article_id}"
                if self.verify_pdf_url(pdf_url):
                    return pdf_url
        
        return self._generic_scrape(url)
    
    def _generic_scrape(self, url):
        """
        Generic page scraping for finding PDF links.
        Fallback when specific patterns don't work.
        """
        try:
            response = self.session.get(url, timeout=15)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Strategy 1: Direct PDF links
            for link in soup.find_all('a', href=True):
                href = link['href']
                if '.pdf' in href.lower():
                    absolute_url = urljoin(url, href)
                    if self.verify_pdf_url(absolute_url):
                        return absolute_url
            
            # Strategy 2: Download/PDF buttons
            keywords = ['pdf', 'download', 'full text', 'full-text', 'download pdf']
            for link in soup.find_all('a', href=True):
                text = link.get_text(strip=True).lower()
                classes = ' '.join(link.get('class', [])).lower()
                
                if any(kw in text or kw in classes for kw in keywords):
                    href = link['href']
                    if not href.startswith('#'):
                        absolute_url = urljoin(url, href)
                        if self.verify_pdf_url(absolute_url):
                            return absolute_url
            
            # Strategy 3: Meta tags
            for meta in soup.find_all('meta'):
                content = meta.get('content', '')
                if '.pdf' in content.lower() and content.startswith('http'):
                    if self.verify_pdf_url(content):
                        return content
            
            # Strategy 4: iframes/embeds
            for tag in soup.find_all(['iframe', 'embed'], src=True):
                src = tag['src']
                if '.pdf' in src.lower():
                    absolute_url = urljoin(url, src)
                    if self.verify_pdf_url(absolute_url):
                        return absolute_url
        
        except Exception as e:
            pass
        
        return None


# Quick test function
if __name__ == "__main__":
    handler = JournalSpecificHandler()
    
    test_urls = [
        "https://leprosyreview.org/article/80/2/19-7204",
        "https://leprosyreview.org/article/75/4/36-7375",
    ]
    
    print("Testing Journal-Specific Handler\n")
    
    for url in test_urls:
        print(f"URL: {url}")
        print(f"Can handle: {handler.can_handle(url)}")
        
        if handler.can_handle(url):
            pdf_url = handler.find_pdf(url)
            if pdf_url:
                print(f"✓ Found PDF: {pdf_url}")
            else:
                print(f"✗ Could not find PDF")
        print()
