# Release Checklist

1. Update `VERSION` in `pdf2ppt.py` and Windows version fields in `version_info.txt`.
2. Update `CHANGELOG.md`.
3. Run `powershell -ExecutionPolicy Bypass -File .\release.ps1`.
4. Confirm all automated tests pass.
5. Extract the portable ZIP into a clean directory and convert a test PDF.
6. Scan the final EXE with Microsoft Defender.
7. Verify the hashes in `release/SHA256SUMS.txt`.
8. Confirm the portable ZIP contains README, privacy notice, MIT license, and third-party licenses.
9. If a code-signing certificate is available, sign the EXE before creating the ZIP and regenerate hashes.
10. Commit the release changes, create a tag such as `v1.5.0`, and push the tag.
