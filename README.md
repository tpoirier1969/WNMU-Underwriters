# WNMU Underwriter Intake v0.1.0

This is a local-first browser app for turning underwriter contract files into a reviewable dataset for quarterly reporting.

## What it does now

- Accepts PDF and DOCX files by drag-and-drop
- Extracts text from DOCX files directly in the browser
- Extracts text from text-based PDFs directly in the browser
- Falls back to browser OCR for scanned PDFs
- Parses likely contract fields into a working record:
  - Underwriter / sponsor
  - Contract type
  - Program / campaign
  - Contact info
  - Start and end dates
  - Amount
  - Credits / spots
  - Program count
  - Audio credit copy
- Stores records locally in the browser
- Flags likely duplicates
- Flags overlapping date ranges only when the same underwriter appears in more than one active contract window
- Exports a quarterly CSV and JSON backup
- Provides manual correction and manual entry for messy documents

## Important limitations in this first pass

- OCR quality depends on the scan quality of the PDF. The sample contract layout is the kind of thing this build is aimed at, but you should expect to review imported fields.
- PDF OCR is capped at 4 pages per file in this version to keep the browser from bogging down.
- The quarterly narrative is a worksheet summary, not a finished regulatory filing template.
- No Supabase wiring is included yet. This build is intentionally local-first so you can start using it without cloud setup.

## Fastest workflow

1. Open `index.html` in a browser.
2. Drop in a batch of PDF and DOCX files.
3. Review flagged records in the table.
4. Fix any questionable fields in the right-hand editor.
5. Export the quarterly CSV.
6. Copy the narrative text if you want a quick summary draft.

## Suggested next step

Once you like the field set and duplicate rules, the next pass should wire the same schema to Supabase so the same intake database can be shared with your other WNMU project.
