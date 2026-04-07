# Code Signing Guide

The Microsoft Store MSI/EXE requirements currently state that the installer binary and all Portable Executable files inside it must be digitally signed by a trusted Certificate Authority chain.

Official reference:
https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msi/app-package-requirements

## Recommended signing order

1. Build the PyInstaller app EXE.
2. Sign the app EXE.
3. Build the Inno Setup installer so it packages the signed app EXE.
4. Sign the installer EXE.
5. Recalculate SHA256 hashes.
6. Update `store/submission_metadata.json` and `store/PARTNER_CENTER_SUBMISSION_v1.0.2.md` if the hash changes.

## Supported build script options

`build_store_installer.ps1` now supports optional signing parameters:

- `-SignToolPath`
- `-CertificateThumbprint`
- `-PfxPath`
- `-PfxPassword`
- `-TimestampUrl`

## Example with a certificate in the Windows certificate store

```powershell
powershell -ExecutionPolicy Bypass -File .\build_store_installer.ps1 `
  -SkipDependencyInstall `
  -CertificateThumbprint "YOUR_CERT_THUMBPRINT"
```

## Example with a PFX file

```powershell
powershell -ExecutionPolicy Bypass -File .\build_store_installer.ps1 `
  -SkipDependencyInstall `
  -PfxPath "C:\path\to\certificate.pfx" `
  -PfxPassword "YOUR_PFX_PASSWORD"
```

## Current blocker

This repository is release-ready, but the published `v1.0.2` installer is still unsigned. Rebuild and re-upload the assets after signing before using the package URL in Partner Center.
