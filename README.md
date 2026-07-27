# IDOT Job IDR Generator

This version keeps a persistent contract lookup record in `idot_contract_index.sqlite`.

## What changed

- A normal lookup checks the SQLite index first, so previously recorded jobs are instant.
- Current/recent letting pages are refreshed only when stale.
- An exact contract-number web search is used as a fallback.
- On a miss, the app indexes a mixed batch from both newer and oldest unchecked archives.
- Each letting page is requested once. The old code requested multiple `?page=` variants that usually repeated the same contract list.
- Misses are cached for only five minutes and are invalidated when archive coverage grows.
- The sidebar shows index coverage and includes a **Build / Refresh Full Job Index** button.
- The sidebar can download the SQLite database as a backup.

## Files

- `app.py` — Streamlit application.
- `build_index.py` — optional command-line full-index builder.
- `idot_contract_index.sqlite` — persistent job/letting record; generated automatically.
- `requirements.txt` — Python dependencies.
- `packages.txt` — installs LibreOffice for exact Excel-to-PDF conversion on supported Linux hosts.

Your existing `IDR_Template.xlsx` or `IDR_template.xlsx` must remain in the same folder as `app.py`.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Build the complete visible archive index

Use either the sidebar button or run:

```bash
python build_index.py
```

Force a refresh:

```bash
python build_index.py --force
```

Only refresh the six newest lettings:

```bash
python build_index.py --recent 6 --force
```

## Persistence

Keep `idot_contract_index.sqlite` beside `app.py`. Back it up with the sidebar download button. On hosting systems with temporary local storage, seed deployments with a prebuilt copy of the database or attach persistent storage.

## Recommended lookup behavior

1. Search by the five-character contract suffix, such as `62K33`.
2. The full form, such as `001-62K33`, also works.
3. A direct IDOT contract-detail URL is still accepted.
4. For maximum older-job coverage, build the full index once.
