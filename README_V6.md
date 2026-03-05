# PDF Scraper Enhanced v6 - Complete Usage Guide

## 🎉 **Implementation Complete!**

All 11 handlers implemented with single-link test mode.

---

## 📁 **Files Provided**

✅ **pdf_scraper_v6_enhanced.py** - Complete enhanced scraper (1,442 lines)
✅ **requirements_v6.txt** - Python dependencies
✅ **journal_handler.py** - Keep your existing file (unchanged)
✅ **README_V6.md** - This guide

---

## 🚀 **Quick Start**

### Step 1: Install Dependencies
```bash
pip install -r requirements_v6.txt
```

### Step 2: Test Single URL
```bash
python pdf_scraper_v6_enhanced.py --test-url "http://bmcinfectdis.biomedcentral.com/articles/10.1186/s12879-016-1593-0"
```

### Step 3: Run Full Scraper
```bash
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx
```

---

## 🧪 **Testing Mode** (NEW!)

### Test Any URL:
```bash
python pdf_scraper_v6_enhanced.py --test-url "YOUR_URL_HERE"
```

### Example Tests:

```bash
# BioMedCentral (should work)
python pdf_scraper_v6_enhanced.py --test-url "http://bmcinfectdis.biomedcentral.com/articles/10.1186/s12879-016-1593-0"

# MDPI (should work)
python pdf_scraper_v6_enhanced.py --test-url "https://www.mdpi.com/2414-6366/8/3/101/htm"

# DOI-only (should normalize and work)
python pdf_scraper_v6_enhanced.py --test-url "10.1007/978-3-319-28498-9_12"

# WHO IRIS (should work with new handler)
python pdf_scraper_v6_enhanced.py --test-url "http://apps.who.int/iris/handle/10665/259357"

# The Lancet (should work)
python pdf_scraper_v6_enhanced.py --test-url "http://thelancet.com/journals/lancet/article/PIIS0140-6736(16)31253-3/fulltext"
```

### Expected Output:
```
================================================================================
SINGLE-LINK TEST MODE
================================================================================

Testing URL: http://bmcinfectdis.biomedcentral.com/articles/10.1186/s12879-016-1593-0

STEP 1: URL Normalization
  → No normalization needed

STEP 2: Domain Detection
  → Domain: bmcinfectdis.biomedcentral.com
  → Using BioMedCentral handler

STEP 3: Finding PDF
    ✓ Found PDF: https://link.springer.com/content/pdf/10.1186/s12879-016-1593-0.pdf

STEP 4: Verification (quick check)
  → HTTP Status: 200
  → Content-Type: application/pdf

================================================================================
✅ SUCCESS!
================================================================================
PDF URL: https://link.springer.com/content/pdf/10.1186/s12879-016-1593-0.pdf
================================================================================
```

---

## 📊 **Full Scraper Mode**

### Basic Usage:
```bash
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx
```

### With Options:
```bash
# Start from specific row
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --start-from-row 1000

# With verbose logging
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --verbose

# Faster (no delay)
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --delay 0

# Custom output file
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --output my_results.xlsx

# Different type filter
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --type-filter "article"
```

---

## 🎯 **What's New in v6**

### 1. **11 Publisher-Specific Handlers**
- ✅ BioMedCentral (+205 PDFs expected)
- ✅ MDPI (+79 PDFs)
- ✅ The Lancet (+55 PDFs)
- ✅ ScienceDirect (+116 PDFs)
- ✅ BMJ (+31 PDFs)
- ✅ Tandfonline (+26 PDFs)
- ✅ **WHO IRIS (+62 PDFs)** - NEW!
- ✅ Oxford (+105 PDFs)
- ✅ PLOS (+20 PDFs)
- ✅ Wiley (+40 PDFs)
- ✅ Springer (+25 PDFs)

### 2. **URL Normalization**
- DOI-only → https://doi.org/
- PMID-only → https://pubmed.ncbi.nlm.nih.gov/
- dx.doi.org → https://doi.org/
- WHO IRIS old → new system
- TinyURL expansion

### 3. **Single-Link Test Mode**
- Test any URL instantly
- See which handler is used
- Detailed step-by-step output
- Quick verification

### 4. **Enhanced Error Tracking**
- Detailed errors from each source
- HTTP status codes
- Content-Type information
- Handler used
- Full diagnostic trail

### 5. **Handler Statistics** (NEW!)
Shows which handlers were used:
```
HANDLER STATISTICS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BioMedCentral:   205 uses
ScienceDirect:   116 uses
Oxford:          105 uses
MDPI:            79 uses
... etc.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## 📋 **Excel Output**

### New Columns Added:
- `download_status` - success_bibcite / success_pubmed / success_abstract_txt / failed / skipped
- `download_filepath` - Full path to downloaded file
- `download_filename` - Just the filename
- `url_used` - The URL that successfully worked
- `source_used` - Which source succeeded (Bibcite URL / PubMed URL / Abstract)
- `download_error` - Main error or success message
- `detailed_errors` - ALL errors from each attempt (Bibcite | PubMed | Abstract)
- **`handler_used`** - Which handler processed the URL (NEW!)
- `download_timestamp` - When processed

### Example Row:
```
download_status: success_bibcite
download_filename: nid_123_Article_Title.pdf
url_used: http://bmcinfectdis.biomedcentral.com/articles/10.1186/s12879-016-1593-0
source_used: Bibcite URL
download_error: 15 pages, 234.5 KB
detailed_errors: (empty - succeeded on first try)
handler_used: BioMedCentral
```

---

## 🔍 **WHO IRIS Details**

### Three Access Methods:

**1. Direct Bitstream API** (Best!)
```
https://iris.who.int/server/api/core/bitstreams/{UUID}/content
```
- Direct PDF download
- No scraping needed
- Your example: `2a12a9b9-a61e-460d-b1a4-59bbccd9a612`

**2. Handle Page Scraping**
```
https://iris.who.int/handle/10665/{ID}
```
- Scrapes HTML for PDF links
- Finds bitstream API links
- Fallback to traditional bitstream

**3. Direct Bitstream**
```
https://iris.who.int/bitstream/handle/10665/{ID}/{filename}.pdf
```
- If filename known from old URL
- Direct access attempt

### Old System Migration:
```
http://apps.who.int/iris/handle/10665/259357
↓ (automatically converted)
https://iris.who.int/handle/10665/259357
↓ (scraped for PDF links)
https://iris.who.int/server/api/core/bitstreams/{UUID}/content
```

---

## 📈 **Expected Results**

### Current Performance:
```
Total rows:      6,622
Success:         4,176 (63%)
  - PDFs:        3,700
  - Abstracts:   476
Failed:          2,446 (37%)
```

### After v6 Improvements:
```
Total rows:      6,622
Success:         5,091 (77%) ⬆️ +14 points
  - PDFs:        4,615 (+915!)
  - Abstracts:   476
Failed:          1,531 (23%) ⬇️ -14 points

Improvement: +915 PDFs! 🎉
```

### Handler Impact Breakdown:
```
BioMedCentral:    +205 PDFs (22%)
ScienceDirect:    +116 PDFs (13%)
Oxford:           +105 PDFs (11%)
URL Normalization: +86 PDFs (9%)
MDPI:             +79 PDFs (9%)
TinyURL:          +74 PDFs (8%)
WHO IRIS:         +62 PDFs (7%)
The Lancet:       +55 PDFs (6%)
BMJ:              +31 PDFs (3%)
Tandfonline:      +26 PDFs (3%)
Others:           +76 PDFs (8%)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TOTAL:            +915 PDFs
```

---

## 🎯 **Command Line Options**

```
usage: pdf_scraper_v6_enhanced.py [options]

Required (choose one):
  --test-url URL        Test a single URL (test mode)
  --input FILE          Input Excel file (full scraper mode)

Optional:
  --output FILE         Output Excel file (default: auto-generated)
  --delay SECONDS       Delay between requests (default: 3)
  --start-from-row N    Start from specific row
  --type-filter TEXT    Filter by Type column (default: biblio)
  --verbose             Enable detailed logging
  --no-verbose          Disable verbose output (default)

Examples:
  # Test single URL
  python pdf_scraper_v6_enhanced.py --test-url "http://example.com/article"
  
  # Run full scraper
  python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx
  
  # Run with verbose logging
  python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --verbose
  
  # Resume from row 1000
  python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --start-from-row 1000
```

---

## 🔧 **Troubleshooting**

### "ModuleNotFoundError: No module named 'bs4'"
```bash
pip install beautifulsoup4
```

### "ModuleNotFoundError: No module named 'openpyxl'"
```bash
pip install openpyxl
```

### "File not found: infontd_biblio.xlsx"
```bash
# Specify your file
python pdf_scraper_v6_enhanced.py --input YOUR_FILE.xlsx
```

### Test Mode Not Working
```bash
# Make sure URL is in quotes
python pdf_scraper_v6_enhanced.py --test-url "http://example.com/article"
```

### WHO IRIS Not Working
```bash
# Test WHO IRIS specifically
python pdf_scraper_v6_enhanced.py --test-url "http://apps.who.int/iris/handle/10665/259357" --verbose
```

---

## 💡 **Pro Tips**

### Tip 1: Test Your Failed URLs First
```bash
# Create a test script
cat failed_links.txt | head -10 | while read url; do
    python pdf_scraper_v6_enhanced.py --test-url "$url"
done
```

### Tip 2: Run with Verbose for Debugging
```bash
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --verbose
```

### Tip 3: Start Small
```bash
# Test first 100 rows
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --start-from-row 1
# Then continue from row 101
python pdf_scraper_v6_enhanced.py --input infontd_biblio.xlsx --start-from-row 101
```

### Tip 4: Analyze Results
```python
import pandas as pd

df = pd.read_excel('infontd_processed_*.xlsx')

# Check handler statistics
print(df['handler_used'].value_counts())

# Check success rate per handler
for handler in df['handler_used'].unique():
    if handler:
        subset = df[df['handler_used'] == handler]
        success = len(subset[subset['download_status'] == 'success_bibcite'])
        total = len(subset)
        print(f"{handler}: {success}/{total} ({success/total*100:.0f}%)")
```

---

## ✅ **Success Checklist**

- [ ] Installed dependencies (`pip install -r requirements_v6.txt`)
- [ ] Tested single URL (`--test-url`)
- [ ] Tested with your failed URLs
- [ ] Reviewed handler statistics
- [ ] Ran full scraper
- [ ] Checked output Excel
- [ ] Compared before/after results

---

## 📞 **Support**

### If Something Doesn't Work:

1. **Test in verbose mode:**
   ```bash
   python pdf_scraper_v6_enhanced.py --test-url "YOUR_URL" --verbose
   ```

2. **Check the detailed_errors column** in output Excel

3. **Check handler_used column** to see which handler processed the URL

4. **Review the handler statistics** at the end of the run

---

## 🎉 **Summary**

**What You Get:**
- ✅ Complete standalone scraper
- ✅ 11 publisher-specific handlers
- ✅ WHO IRIS support (3 methods)
- ✅ Single-link test mode
- ✅ Enhanced error tracking
- ✅ Handler statistics
- ✅ URL normalization
- ✅ Resume capability

**Expected Impact:**
- ✅ +915 PDFs
- ✅ 77% success rate (up from 63%)
- ✅ +14 percentage points improvement

**Ready to use!** 🚀
