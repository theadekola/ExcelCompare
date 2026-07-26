# Compare and Merge Workspace

Excel Compare Professional 1.2 introduces a two-workbook mapping workspace.

## Workflow

1. Select Workbook 1 and Workbook 2.
2. Select each workbook with **Browse**. It loads and previews automatically.
3. Select a worksheet independently for each workbook. Worksheet names do not need to match.
4. Review the first 50 data rows in the side-by-side preview.
5. Open **Column mapping**.
6. Map each Workbook 1 column to its corresponding Workbook 2 column.
7. Tick **Key** on one or more mappings that uniquely identify a record, such as Employee ID, Invoice Number, Student ID, or Account Code.
8. Tick **Compare** for fields whose values should be checked.
9. Tick **Use** for fields that should be present in the merged workbook and edit the output column name when required.
10. Choose a merge row policy and conflict policy.
11. Click **Compare selected columns** for an audit report or **Merge and save Excel workbook** to create a new `.xlsx` file.

## Merge modes

- **Full merge** includes records from both workbooks.
- **Workbook 1** keeps every Workbook 1 record and adds mapped Workbook 2 values where a key matches.
- **Matched only** includes records whose key exists in both workbooks.

## Conflict policies

- Prefer Workbook 2
- Prefer Workbook 1
- Mark both conflicting values in the merged cell

The original files are never modified.
