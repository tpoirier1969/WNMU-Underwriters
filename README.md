# WNMU Underwriter Intake v0.5.2

This pass cleans up the Supabase side so the app behaves better inside a shared Supabase project instead of living in a generic public-table junk drawer.

## What changed

- Keeps the three real sections: **Ingest / Contracts / Quarterly**
- Moves the cloud setup to a **namespaced custom schema**: `wnmu_underwriters`
- Uses a dedicated contracts table: `wnmu_underwriter_contracts`
- Pre-creates a future detail table for exact run timestamps: `wnmu_underwriter_credit_runs`
- Uses namespaced trigger/function/policy names instead of generic shared-instance names
- Keeps the existing browser-first workflow and local fallback
- Auto-upgrades legacy config defaults if your old config still says `underwriter_contracts`

## Files in this zip

- `index.html` - app shell
- `app.css` - styling
- `app.js` - app logic, parsing, duplicate checks, quarterly report, cloud sync
- `README.md` - setup notes
- `supabase-schema.sql` - run this in Supabase SQL editor

## Supabase setup for this pass

1. In **Supabase > API Settings > Exposed schemas**, add: `wnmu_underwriters`
2. Run `supabase-schema.sql` in the SQL editor.
3. Keep your existing `config.js` file and change only these lines if needed:
   - `enableSupabase: true`
   - `workspaceKey: 'wnmu-underwriters'`
   - `cloudSchema: 'wnmu_underwriters'`
   - `cloudTable: 'wnmu_underwriter_contracts'`
   - `cloudRunsTable: 'wnmu_underwriter_credit_runs'`
4. Open the app and use **Push all to cloud**.
5. On another machine, use the same config and **Pull from cloud**.

## Important truth, not brochure copy

This is **better isolated**, not fully private.

It now sits in its own schema and uses explicitly named objects, which is the right move for a shared Supabase project. But because the browser build still uses anon-key client access, this is not the final hardening step. Real privacy later means tighter RLS and/or auth.

## Import behavior

- PDF: tries text layer first, then OCR fallback for scanned contracts
- DOCX: extracts text automatically
- JSON backup: imports previously exported app data

## Known limits in this pass

- PDF OCR can be slow on big scans
- Exact credit run dates / times are still mainly manual entry for now
- The future `wnmu_underwriter_credit_runs` table is created now so we can wire detailed run records cleanly in a later pass


## v0.4.1 change

- Underwriter names in **Quarterly** are now clickable and open the full contract modal, including raw extracted text.


## v0.5.2 changes

- Keeps the **Open** button on one line instead of breaking it into two
- Shrinks and recenters the contract modal so it stays on screen better
- Adds **previous / next** arrow buttons in the modal to move through contracts
- Tunes OCR import handling for scanned PDFs with higher-resolution rendering and preprocessing
- Improves underwriter extraction so scanned contracts are less likely to misread the sponsor name
