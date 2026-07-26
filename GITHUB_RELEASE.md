# Publishing a GitHub Release

1. Build and test the application on Windows.
2. Run `scripts\build-installer.bat`.
3. Create a Git tag matching the application version, for example `v2.0.4`.
4. In GitHub, open **Releases**, select **Draft a new release**, and choose the tag.
5. Upload `ExcelCompareProfessional-Setup-v2.0.4.exe` as the release asset.
6. Add the SHA-256 checksum file to the release assets.
7. Do not commit the installer, deployment DLLs, `build`, `package`, or `dist` folders to the main repository.
