# MSIX Migration Guide

Checked on: 2026-04-08

## Why switch to MSIX

The current EXE/MSI Store path requires a trusted CA code-signing certificate for the installer and the PE files inside it.

MSIX is the better route if you cannot obtain that certificate yourself because Microsoft Store supports MSIX submissions and handles Store distribution signing and hosting more directly than the EXE/MSI route.

Official references:
- https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msi/app-package-requirements
- https://learn.microsoft.com/mt-mt/windows/apps/publish/faq/submit-your-app
- https://learn.microsoft.com/en-us/windows/msix/package/packaging-uwp-apps

## What was added to this repository

- `build_msix_package.ps1`
  Builds the app EXE, stages a packaged desktop layout, and creates an `.msix` package.
- `msix/msix_config.json`
  Holds the package identity, publisher, version, and executable settings.
- `msix/AppxManifest.template.xml`
  Template manifest for a full-trust packaged desktop app.
- `msix/generate_msix_assets.ps1`
  Generates placeholder visual assets required by the manifest.

## Important Store identity note

Before uploading the final package to Partner Center, replace these fields in `msix/msix_config.json` with the values from your reserved Store identity:

- `identityName`
- `publisher`
- optionally `publisherDisplayName`

If these values do not match Partner Center, the submitted package identity will not line up with the Store listing.

## Desktop app compatibility notes

This app is a classic desktop program and should be packaged as a full-trust desktop app.

Important runtime assumptions:
- File picking should continue to work because the app already uses standard file picker dialogs.
- Writing output into user-selected folders should continue to work for a full-trust desktop app.
- Excel automation is still a risk area and must be tested after packaging because COM behavior can differ slightly once the app is running from inside a package.

## Recommended next steps

1. Install the Windows SDK if `makeappx.exe` is missing.
2. Reserve the Store app name in Partner Center.
3. Update `msix/msix_config.json` with the reserved identity and publisher.
4. Run `build_msix_package.ps1`.
5. If you want local sideload testing, rerun with `-CreateTestCertificate`.
6. Test file selection, workbook generation, and Excel-installed/non-Excel-installed scenarios.

## Current verified state

The repository has now been updated with the Partner Center identity values:

- Identity name: `Vurctne.MasterBudgetAutomationTool`
- Publisher: `CN=E75204F6-F77B-4E0C-89C6-AC00A663F6A0`
- Publisher display name: `Vurctne`

Verified local MSIX build artifact:
- `dist_msix/Vurctne.MasterBudgetAutomationTool_1.0.2.0_x64.msix`
- SHA256: `AF30F6B28B05A55941A660600493D1DFE333FA087634F2A7483E675E3B460700`

Verified local installation result:
- Package full name: `Vurctne.MasterBudgetAutomationTool_1.0.2.0_x64__m3ks3abq6d0tt`
- Start menu entry: `Master Budget Automation Tool`
- App launch from packaged install: successful
