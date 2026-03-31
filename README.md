# WNMU Underwriter Intake v0.3.1

This pass fixes the fake-tab behavior, tightens the quarterly layout, and adds quick Supabase sync.

## What changed

- Real top-level sections: **Ingest / Contracts / Quarterly**
- Inactive sections are actually hidden instead of living on one endless page
- Contracts list sorts from the header arrows
- Contract columns are in this order:
  - Underwriter
  - Start
  - End
  - Flags
  - Type
  - Program / Day / Day-part
  - Amount
  - Source
- **Open** launches a modal popup editor
- Added editable field for **Program / Day / Day-part**
- Added editable field for **Exact credit run dates / times**
- Quarterly grid wraps better and avoids the ridiculous side-scroll problem
- Supabase **Pull from cloud** and **Push all to cloud**
- Local browser storage still works if cloud is not configured yet

## Files

- `index.html` - app shell
- `app.css` - styling
- `app.js` - app logic, parsing, duplicate checks, quarterly report, cloud sync
- `config.js` - fill this in with your Supabase values
- `supabase-schema.sql` - run this in Supabase SQL editor

## Quick Supabase setup

1. Run `supabase-schema.sql` in your Supabase project.
2. Open `config.js`.
3. Change these values:
   - `enableSupabase: true`
   - `supabaseUrl`
   - `supabaseAnonKey`
   - `workspaceKey` (leave `wnmu-underwriters` unless you want a different shared set)
4. Open `index.html`.
5. Use **Push all to cloud** once to move your local records up.
6. On another machine, use the same `config.js`, open the app, and hit **Pull from cloud**.

## Important note

This is intentionally a quick shared-browser setup. The included SQL policy allows browser access with the anon key, which is fast but not hardened. If you want, the next pass can move this to real auth and tighter RLS.

## Import behavior

- PDF: tries text layer first, then OCR fallback for scanned contracts
- DOCX: extracts text automatically
- JSON backup: imports previously exported app data

## Known limits in this pass

- PDF OCR can be slow on big scans
- Exact credit run dates / times are still manual entry for now
- Program / day / day-part parsing is best-effort, not psychic
