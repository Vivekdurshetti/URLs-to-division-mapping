URL TO DIVISION MAPPING TOOL
============================

HOW TO RUN
==========

1. Open index.html in a browser (it loads style.css and script.js from the same folder,
   so keep the three files together).

2. Upload your report. Either:
   - one Excel workbook that has an "XP pages" tab and an "XM pages" tab, or
   - separate files (.xlsx / .xls / .csv), one for XP and one for XM.
   You can select or drop several files at once.

3. Use the three tabs:
   - Dashboard          charts and KPIs for every field
   - URL Lookup         paste URLs, find their division and metadata
   - Browse All Fields  every row, every column, with filters and export


HOW XP / XM ARE DETECTED
========================

- A tab or file name containing "XP" or "XM" is labelled automatically.
- If a name says neither, the "treat it as" dropdown decides (XP or XM).
- Tabs in a multi-tab workbook that are neither XP nor XM are skipped.
- Uploading a new file for a source replaces that source's rows.
- Use the x on a source chip to remove a source.


SUPPORTED HEADERS (not case sensitive)
======================================

ID                              Language
Path                            Page URLs (or Page URL / URL)
Division Names (or Divisions)   Sub-Division (or Sub Divisions)
Page Types                      Template Name
Page cretaed date               Page name
Branch Template                 Page Status

Any other column in the sheet is also loaded and shown in the dashboard's
"Field completeness" table and in the Browse tab.
Multiple divisions in one cell can be separated with | or ;


FEATURES
========

Dashboard
  - KPIs: total pages, XP and XM counts, divisions, languages, page types,
    division coverage, URLs present in both XP and XM
  - Stacked bars (XP vs XM) by division, page type, sub-division, language,
    status, created year and template. Click a bar to open those pages
    in the Browse tab.
  - Data quality: pages without a division, without a URL, duplicate URLs
    within a sheet, URLs found in both XP and XM
  - Field completeness: % of rows filled for every column, per source

URL Lookup
  - Paste URLs and search XP + XM, or only one of them
  - A URL found in both sheets shows one row per match with a Source column
  - Duplicate and missing detection, search, filter, pagination
  - Click a cell to copy it; copy a whole column from the header icon
  - Export Excel / CSV (includes Source)

Browse All Fields
  - Filter by source, division, sub-division, language, page type and data
    quality issue; free-text search across every field
  - Choose which columns to show; export the filtered rows to Excel / CSV

SUPPORTED FILES
===============

.xlsx
.xls
.csv
