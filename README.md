# Slide QA Copilot

A Streamlit application for source-based review of slide claims and their cited evidence.

## What it does

- Accepts PowerPoint `.pptx` files directly.
- Extracts referenced paragraphs, text-run hyperlinks, grouped shapes, and table-cell content.
- Supports several cited sources for one slide claim.
- Provides slide-range, keyword, multi-reference, mapping-issue, and select-all filters for larger decks.
- Uses Gemini URL Context to compare selected claims with their cited public sources.
- Separates unsupported claims from claims that remain unverified because a source was inaccessible.
- Generates one Excel QA workbook with review results, source retrieval details, and optional API-usage information.

## Google Slides option

The optional `extract_slide_references.gs` script can be added to a Google Slides presentation through **Extensions → Apps Script**. It creates a Google Sheet containing the extracted slide text and URLs. Download that extraction as CSV and upload it to the app when the Google Slides → PPTX conversion does not preserve the links correctly.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

- Branch: `main`
- Main file path: `app.py`

## Data handling

This is an internal beta. Use only content approved for external AI processing. Where practical, remove client names, logos, comments, and unnecessary confidential identifiers before upload. Final QA decisions remain with the analyst.
