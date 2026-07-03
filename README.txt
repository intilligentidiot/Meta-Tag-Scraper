Meta Tag Scraper - Streamlit App
=================================

WHAT IT DOES:
  - Paste one or more website URLs
  - Scrapes Meta Title and Meta Description from each page
  - Shows SEO status: OK / Too Short / Too Long / Missing
  - Export results as Excel (.xlsx) file with color-coded cells

FILES:
  meta_scraper.py   - Main app (Python/Streamlit)
  requirements.txt  - Python dependencies

HOW TO RUN:
  Step 1 - Install dependencies (first time only):
    python -m pip install -r requirements.txt

  Step 2 - Start the app:
    python -m streamlit run meta_scraper.py

  Step 3 - Open browser at:
    http://localhost:8501

SEO RULES:
  Meta Title       : 30-60 characters ideal
  Meta Description : 70-160 characters ideal
