# WNMU Underwriter Intake v0.2.0

This is a local-first browser app for ingesting underwriting contracts, reviewing them, and preparing quarterly reporting data.

## What changed in this pass

- Split into **3 top sections**:
  - **Ingest**
  - **Contracts**
  - **Quarterly**
- Contract grid columns reordered to:
  - **Underwriter**
  - **Start**
  - **End**
  - **Flags**
  - **Type**
  - **Program / Day / Day-part**
  - **Amount**
  - **Source**
- Added **sortable column headers** with ascending / descending toggle
- Added **popup editor modal** triggered by the **Open** button
- Added blank editable field for **Program / Day / Day-part**
- Added blank editable field for **Exact credit run dates / times**
- Overlap logic only flags contracts that overlap for the **same underwriter**, not different ones
- Quarterly CSV export now includes placement and exact run schedule fields

## Files

- `index.html`
- `app.css`
- `app.js`

## How to run

Open `index.html` in a modern browser.

## Supported import types

- PDF
- DOCX
- JSON backup from this app

## Notes

- Scanned PDFs use browser OCR and can be slower.
- DOCX import uses Mammoth in-browser text extraction.
- This pass is still **local-first**. It is structured so a later Supabase-backed version can reuse the record shape instead of starting over.
