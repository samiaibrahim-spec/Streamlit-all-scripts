# CB Reporting Hub

A single web app that combines all three SA360 reporting tools. Users select a report type, upload their file, and download the output. No Python, no installs, no setup.

## Reports Available

**Weekly WoW Report** — Generates a formatted Excel report comparing the most recent two weeks of campaign performance with percent change rows. Upload a weekly SA360 CSV or Excel export.

**Monthly Campaign Summary** — Generates Excel summary tables (one per segment) with MoM and YoY comparisons, plus a text insights file with automated performance narratives. Upload a monthly SA360 CSV or Excel export.

**Keyword Analysis** — Generates top 10 keyword tables segmented by Customer Type and Brand/NonBrand. Upload an SA360 keyword CSV or Excel export.

## How to use

1. Open the app URL in your browser
2. Select a report from the dropdown
3. Upload your data file
4. Review the data preview and classifications
5. Click the Generate button
6. Download the output file(s)

Bookmark the link for weekly use. Share it with any new team members.

## Deploying

The app is a single file (app.py) with one requirements.txt.

**Streamlit Community Cloud:**
1. Push app.py and requirements.txt to a GitHub repo
2. Connect the repo at https://share.streamlit.io
3. Deploy and share the URL

**Local:**
```
pip install -r requirements.txt
streamlit run app.py
```

## Adding a new report

1. Write the report logic as a function (load, process, generate output)
2. Write a `run_<report_name>()` function that handles the Streamlit UI (file uploader, buttons, download)
3. Add the report name to the selectbox options in `main()`
4. Add an `elif` branch in `main()` to call your new function
5. Commit to GitHub — the app updates automatically

## Updating configuration

All configuration for each report lives in app.py at the top of its respective section. Common things to update:

- **WEEKLY_TABLES** — add or rename weekly report table sections
- **WEEKLY_METRICS** — raw metric columns pulled from weekly data
- **STANDARD_COLS / VBB_COLS** — columns shown in the weekly Excel output
- **WEEKLY_COLUMN_ALIASES** — maps variant column names to standard ones
- **TOTAL_ACTIONS_COMPONENTS** — which metrics sum into Total Actions
- **MONTHLY_SUMMARY_METRICS** — metrics shown in monthly summary tables
- **parse_campaign_name()** — how Customer Type, Engine, and Brand are extracted
