# Spreadsheet format support

Excel Compare Professional 1.1 supports the following input formats.

## Native support

- `.xlsx` through QXlsx
- `.csv` through the built-in UTF-8 delimited-text reader
- `.tsv` and `.tab` through the built-in tab-delimited reader

## Formats converted through LibreOffice

- `.xls`
- `.xlsb`
- `.xlsm`
- `.xlt`
- `.xltx`
- `.xltm`
- `.ods`
- `.fods`

For these formats, Excel Compare Professional starts LibreOffice in headless mode, converts the selected file into a temporary `.xlsx` workbook, reads the converted workbook, and automatically removes the temporary file.

The original spreadsheet is never modified.

## Install legacy-format support

Run:

```bat
tools\install-libreoffice.bat
```

Alternatively, install LibreOffice manually. The application looks for `soffice.exe` in the Windows PATH and in the standard Program Files locations.

## Limitations

- Macros and VBA code are not executed or compared.
- Password-protected or encrypted files must be unlocked before comparison.
- CSV and TSV files are treated as a single worksheet.
- Formula results are compared as loaded values. Formula source expressions are not currently compared.
