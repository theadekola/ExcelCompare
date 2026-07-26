# Version 2.0.4

- Removed the separate Compare and merge group.
- Added Compare only and Compare and merge operation buttons inside Column mapping and actions.
- Moved Start Comparison directly below the operation controls.
- Shows merge settings, preview and save controls only when Compare and merge is selected.
- Added a responsive scrollable layout for laptop and desktop screen sizes.
- Improved splitter, table and preview resizing.

# Version 2.0.4

- Automatically selects Workbook 1 columns for comparison after loading and previewing.
- Automatically selects the first common mapped column as the initial record key.
- Total Records now represents Workbook 1 row count in compare-only mode.
- When compare-and-merge mode is enabled, Total Records represents the combined comparison outcomes.

## 2.0.4

- Fixed Compare selection so key columns no longer depend on Merge selection.
- Comparison uses the Compare checkboxes; merge output uses the Merge checkboxes.
- Mapping rows now populate immediately after loading or changing worksheets.
- Select Key Columns is available immediately after loading.
- Output Column now follows the Compare and Merge checkboxes automatically.
- Added clearer validation when no comparison column or key column is selected.


## 2.0.4

- Auto mapping is no longer performed automatically when worksheets load.
- Added **Tick Compare** and **Tick Merge** checkboxes beside Auto-map columns.
- Auto-map now selects comparison and merge columns only according to those checkboxes.
- Default mode maps columns for comparison only; merge selection remains off until explicitly chosen.

# Version 2.0.4

- Added Compare only / Compare and merge mode checkbox.
- Merge controls and preview are enabled only when merge mode is selected.
- Compare-only mode clears and disables merge output controls.

# Version 1.9.2

- Fixed the Results status filter so **New Records** correctly displays records whose internal status is **New Record**.
- Added explicit mappings for All Results, New Records, Unmatched, Modified, and Matched.
- Filtered Excel, PDF, and CSV exports now follow the corrected visible-row selection.

# Version 1.9.1

- Renamed Added to New Record.
- Renamed Deleted to Unmatched.
- Renamed Unchanged to Matched.
- Added status icons throughout the dashboard, filters, results, Excel, PDF and CSV exports.
- Renamed All Changes to All Results and Total to Total Records.
- Kept the existing status colours: green, red, orange and grey.

# Version 1.9.0

- Results now always retain New Records, Unmatched, Modified and Matched records.
- Status filter correctly displays each selected category.
- Excel Comparison Results sheet now mirrors the on-screen table with one output field per column.
- Modified values remain in one field cell using Workbook 1 -> Workbook 2.
- Full-row status colour coding is preserved in Excel output.

# Version 1.7.0

- Preserves rows with missing keys instead of skipping them.
- Preserves duplicate-key rows instead of overwriting earlier records.
- Keeps missing Workbook 1 and Workbook 2 rows in comparison exports.
- Applies New Records, Unmatched, Modified and Matched colours across complete exported rows.
- Adds Excel, PDF and CSV export buttons directly to the Results toolbar.
- Reduces the Results search-box width.
- Includes New Record and Unmatched fields in the Detailed Changes worksheet.

# Changelog

## 1.6.0

- Commercial dark-sidebar navigation and separate Compare, Results, Reports, Audit Log, Settings and About pages.
- Side-by-side workbook previews and worksheet selection.
- Column mapping table and dedicated key-column selection dialog.
- Searchable, filterable, colour-coded comparison results.
- Report preview and in-application audit history.
- Improved error reporting for legacy spreadsheet conversion failures.
- Compare and merge workflows retained from version 1.2.

# Version 1.6.0

- Redesigned the loading workflow with side-by-side workbook previews.
- New Record independent worksheet selection.
- New Record visual column mapping and automatic header matching.
- New Record mapped-column comparison when header names differ.
- New Record full outer, Workbook 1 and matched-only merge modes.
- New Record merge conflict policies and merged `.xlsx` export.

# Changelog

## 1.6.0

- New Record native CSV reader.
- New Record native TSV and TAB reader.
- New Record `.xls`, `.xlsb`, `.xlsm`, `.xlt`, `.xltx`, `.xltm`, `.ods` and `.fods` support through temporary LibreOffice conversion.
- New Record automatic LibreOffice detection.
- New Record a legacy-format installation helper using Windows Package Manager.
- Expanded the Open File dialog to show all supported spreadsheet types.
- New Record `.xls`, `.ods`, `.csv` and `.tsv` sample files.
- Updated help documentation and the PDF user manual.
- Updated AAT-Tech Ltd branding and version metadata.
- Improved background-thread safety when loading and comparing files.

## 1.0.0

- Initial commercial package.

## 1.6.0

- New Record native Microsoft Excel 97-2003 `.xls` reading with libxls 1.6.3.
- `.xls` no longer requires LibreOffice or Microsoft Excel.
- Kept LibreOffice fallback for `.xlsb`, `.ods`, `.fods`, and `.xlt`.
- New Record clear parser errors for malformed, encrypted, or unsupported legacy workbooks.

## 1.6.0
- Redesigned Compare Files workspace to match the two-panel workflow.
- Moved each workbook selector beneath its preview panel.
- New Record selectable maximum number of output columns for merging.
- New Record in-app merged workbook preview before saving.
- Save is disabled until a valid preview has been generated.

## 1.9.0
- Excel, PDF, and CSV exports now use the currently visible Results-page records.
- Status selection supports All Results, New Record, Unmatched, Modified, and Matched.
- Search-filtered rows are also respected during export.
- Excel opens directly on the row-based Comparison Results sheet.
- CSV now matches the on-screen one-record-per-row layout.
- Detailed field export is limited to modified cells and no longer expands deleted records into one row per field.

## 2.0.4
- Aligned Start Comparison to the right of Compare only and Compare and merge.
- Placed Worksheet selector, workbook path and Browse button on one row under each preview.
- Improved compact use of horizontal space on laptop and desktop displays.
