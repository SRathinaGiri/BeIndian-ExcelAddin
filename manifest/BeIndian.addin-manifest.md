# BeIndian Excel Add-in Manifest

- Add-in name: BeIndian Excel Audit Tool
- Target file: `dist/BeIndian2_VBA.xlsm` during development, save as `BeIndian2_VBA.xlam` for installation.
- Ribbon tab: `BeIndian`
- Minimum Excel version: Excel 2019 for Windows.
- VBA project modules:
  - `modBI_Core`
  - `modBI_Validators`
  - `modBI_TextDate`
  - `modBI_Forensics`
  - `modBI_Benford`
  - `modBI_Financial`
  - `modBI_Ribbon`
- Initial Ribbon callbacks:
  - `BI_Tool_ValidatePAN`
  - `BI_Tool_BenfordFirstDigit`
  - `BI_Tool_AltmanZScore`
  - `BI_Tool_EMISchedule`
  - `BI_Tool_NotYetPorted`
  - `BI_ShowAbout`

## Installation

1. Open `dist/BeIndian2_VBA.xlsm`.
2. Confirm macros are enabled.
3. Save As `Excel Add-in (*.xlam)`.
4. In Excel, go to File > Options > Add-ins > Manage Excel Add-ins > Go.
5. Browse to the saved `.xlam` and enable it.
