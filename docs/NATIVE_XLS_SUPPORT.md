# Native XLS support

Excel Compare Professional 1.4 reads Microsoft Excel 97-2003 `.xls` files directly with libxls 1.6.3.

No Microsoft Excel or LibreOffice installation is required for `.xls` files.

## Reader selection

- `.xlsx`, `.xlsm`, `.xltx`, `.xltm`: QXlsx
- `.xls`: libxls
- `.csv`, `.tsv`, `.tab`: Qt native text reader
- `.xlsb`, `.ods`, `.fods`, `.xlt`: LibreOffice conversion fallback

## Important limitations

The native XLS reader compares calculated cell values. It does not execute VBA macros. Password-encrypted XLS files may be rejected. Very unusual formatting, charts and embedded objects are not part of the row-data comparison.

libxls is distributed under the BSD 2-Clause licence. Its source is downloaded and compiled as a static library during the CMake configuration step.
