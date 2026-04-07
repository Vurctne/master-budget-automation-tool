# Microsoft Store Publish Checklist

Checked on: 2026-04-08

## Official references

- Get started with Microsoft Store:
  https://learn.microsoft.com/ms-my/windows/apps/publish/get-started
- Create an app submission for your MSI/EXE app:
  https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msi/create-app-submission
- App package requirements for MSI/EXE apps:
  https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msi/app-package-requirements

## Current project status

This project is now prepared for the Microsoft Store EXE submission path.

Ready now:
- Public GitHub repository
- Public GitHub Release page
- Versioned public installer URL
- Silent-installable Inno Setup installer
- Release notes draft
- Store listing draft
- Certification notes draft
- Privacy policy draft
- SHA256 hashes for release assets

Still required before submission:
- Code-sign the inner app EXE and the installer EXE
- Confirm the final publisher name exactly matches the Partner Center publisher
- Upload screenshots and logos in Partner Center

## Package requirement checkpoints

Microsoft Learn currently states that MSI/EXE submissions must meet all of the following:
- The installer binary must be `.exe` or `.msi`
- The installer and all PE files inside it must be code-signed with a trusted CA certificate
- The package URL must be an HTTPS direct link
- The package URL must be versioned
- The binary behind that URL must not change after submission
- Silent install is required
- The installer must be a standalone offline installer, not a downloader stub

## Current package values for this project

- App type: `EXE`
- Architecture: `x64`
- Language: `en-au`
- Install scope: `per-user`
- Silent install: `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-`
- Silent uninstall: `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-`
- Installer URL:
  `https://github.com/Vurctne/master-budget-automation-tool/releases/download/v1.0.2/MasterBudgetAutomationTool_Setup_1.0.2.exe`
- Installer SHA256:
  `E3D8E55489494C47C6CFA4E2E463FBA27983E80E7263763998F8627646A12100`

## Supporting URLs

- Repository:
  `https://github.com/Vurctne/master-budget-automation-tool`
- Release page:
  `https://github.com/Vurctne/master-budget-automation-tool/releases/tag/v1.0.2`
- Privacy policy:
  `https://github.com/Vurctne/master-budget-automation-tool/blob/main/store/PRIVACY_POLICY.md`

## Validation already completed

- `python -m py_compile app.py budget_automation.py app_metadata.py`
- `powershell -ExecutionPolicy Bypass -File .\build_store_installer.ps1 -SkipDependencyInstall`
- Public release asset URL responds successfully

## Recommended final sequence

1. Obtain a code-signing certificate.
2. Rebuild with signing enabled.
3. Verify the signed installer hash and update the Store metadata if the hash changes.
4. Reserve the product name in Partner Center.
5. Fill the submission fields using `store/PARTNER_CENTER_SUBMISSION_v1.0.2.md`.
6. Upload screenshots and logos.
7. Submit for certification.
