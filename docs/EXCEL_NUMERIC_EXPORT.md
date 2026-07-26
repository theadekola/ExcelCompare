# Excel numeric export

Version 2.0.4 writes exported values using native Excel cell types.

- Dates are written as Excel date cells and displayed as `dd/mm/yyyy`.
- Financial and ordinary numeric values are written as numeric cells.
- Numeric identifiers with leading zeros use a custom zero format, preserving values such as `0000057` without Excel's "Number Stored as Text" warning.
- Alphanumeric identifiers such as `LNS0004237` remain exact text.
- Status row colours are retained.
