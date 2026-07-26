# Missing records and exports

Version 1.7.0 never silently removes a source row during comparison.

- Rows missing a selected key are retained with a generated record label.
- Duplicate-key rows are retained separately.
- Workbook 1-only rows appear as Unmatched and use the red/pink status colour.
- Workbook 2-only rows appear as New Record and use the green status colour.
- Columns available on only one side remain present with blank values on the other side.
- Excel, PDF and CSV exports are available directly from the Results toolbar.
