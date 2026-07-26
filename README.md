# Excel Compare Professional

**Developed and published by AAT-Tech Ltd**

Excel Compare Professional is a Windows desktop application built with C++20 and Qt 6 for comparing, reconciling, reporting, and merging spreadsheet data.

## Features

- Compare large spreadsheets using selected key columns
- Detect new, unmatched, modified, and matched records
- Native support for legacy `.xls` files through libxls
- Support for `.xlsx`, `.xlsm`, `.xltx`, `.csv`, `.tsv`, and `.tab`
- Optional conversion support for `.xlsb`, `.ods`, `.fods`, and `.xlt`
- Side-by-side workbook previews
- Manual and automatic column mapping
- Compare-only and compare-and-merge workflows
- Colour-coded results
- Export filtered results to Excel, PDF, or CSV
- Preserve identifiers, prefixes, and leading zeros
- Audit log, reports, settings, and help documentation

## Technology

- C++20
- Qt 6
- CMake
- QXlsx 1.5.0
- libxls 1.6.3
- Inno Setup

## Build requirements

- Windows 10 or Windows 11 x64
- Visual Studio 2022 or compatible Build Tools with the v143 toolset
- Qt 6.5 or newer using the MSVC 2022 64-bit kit
- CMake 3.24 or newer
- Git
- Inno Setup 6 or 7 for creating the installer

## Build the application

Open a Visual Studio Developer Command Prompt and run:

```bat
cd /d C:\path\to\ExcelCompareProfessional
set QT_ROOT=C:\Qt\6.8.3\msvc2022_64
scripts\build-cmake.bat
```

The application executable is generated under:

```text
build\Release\ExcelCompareProfessional.exe
```

## Build the Windows installer

```bat
scripts\build-installer.bat
```

The installer is generated under:

```text
dist\ExcelCompareProfessional-Setup-v2.0.4.exe
```

## Repository policy

Generated build folders, deployment DLLs, installers, executables, and machine-specific Visual Studio files are intentionally excluded from source control. Windows installers should be published through **GitHub Releases**, not committed to the repository.

## Sample files

The `samples` folder contains demonstration spreadsheet files for testing comparison and export features.

## Documentation

Product and technical documentation is available in the `docs` folder. Third-party licence notices are provided in `docs/THIRD_PARTY_NOTICES.md`.

## Licence

This project is proprietary software owned by AAT-Tech Ltd. See `LICENSE` for the repository licence terms. Third-party components remain subject to their respective licences.

Copyright © 2026 AAT-Tech Ltd. All rights reserved.
