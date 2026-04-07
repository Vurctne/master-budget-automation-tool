# Partner Center MSIX Submission Pack for v1.0.2

Prepared on: 2026-04-08

Use this file when you are ready to submit the MSIX package to Microsoft Store instead of the EXE installer route.

## Package to upload

- File:
  `dist_msix/Vurctne.MasterBudgetAutomationTool_1.0.2.0_x64.msix`
- SHA256:
  `AF30F6B28B05A55941A660600493D1DFE333FA087634F2A7483E675E3B460700`

## Identity values already matched to Partner Center

- Package/Identity/Name:
  `Vurctne.MasterBudgetAutomationTool`
- Package/Identity/Publisher:
  `CN=E75204F6-F77B-4E0C-89C6-AC00A663F6A0`
- Package/Properties/PublisherDisplayName:
  `Vurctne`

## Important note about the current package

The current MSIX is signed with a local test certificate for sideload validation.

For Microsoft Store submission:
- keep the package identity values exactly as above
- upload the MSIX package through Partner Center
- do not use the local test certificate as your final distribution trust model

The package identity is the important part for Store submission. The Store will validate the package against the reserved product identity.

## Availability

- Markets: All desired markets
- Discoverability: Available in Microsoft Store
- Pricing: Free
- Free trial: None

## Properties

- Category: Business
- Privacy policy URL:
  `https://github.com/Vurctne/master-budget-automation-tool/blob/main/store/PRIVACY_POLICY.md`
- Website:
  `https://github.com/Vurctne/master-budget-automation-tool`
- Contact email:
  `ivan.wang@education.vic.gov.au`
- Publisher display name:
  `Vurctne`

## Store listing

- Product name:
  `Master Budget Automation Tool`
- Short description:
  `Import Compass Expense Sub-Program data into a Master Budget workbook without OFFSET/MATCH formulas.`
- Main description:
  Use `store/listing_content.en-AU.md`
- Certification notes:
  Use `store/notes_for_certification.txt`

## Local validation already completed

- MSIX package generation succeeded
- Local machine certificate trust setup succeeded
- Local MSIX installation succeeded
- Packaged app launch succeeded from Start menu

## Remaining checks recommended before Store submission

1. Open the packaged app and run a real import end to end.
2. Test with Excel installed.
3. Test with Excel not installed.
4. Confirm output workbook save flow works under packaged execution.
5. Upload screenshots and required logo assets in Partner Center.
