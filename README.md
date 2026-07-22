# BeIndian Excel Add-in

BeIndian is a VBA-based Microsoft Excel add-in for Indian audit analytics, forensic review, sampling, text/date utilities, and financial calculations.

This repository contains the Excel Add-in version of BeIndian, implemented as a VBA `.xlam` add-in that can be used in Excel 2019, Excel 2021, and Microsoft 365 on Windows. BeIndian is also available as a Microsoft 365 LAMBDA template for Excel versions that support LAMBDA functions.

## Current Release

- Version: `1.0.0 (Excel Add-in)`
- Main add-in file: [`dist/BeIndian2_VBA.xlam`](dist/BeIndian2_VBA.xlam)
- Macro workbook/source package: [`dist/BeIndian2_VBA.xlsm`](dist/BeIndian2_VBA.xlsm)
- Help workbook: [`dist/BeIndian_Addin_Help.xlsx`](dist/BeIndian_Addin_Help.xlsx)

## Feature Areas

The add-in adds a **BeIndian** ribbon tab with tools grouped for audit workflows:

- Forensic Audit: score templates, Altman Z-Score, Beneish M-Score, Ohlson O-Score, Piotroski F-Score, Benford tests, related-party style tests, round-number tests, duplication tests.
- Audit Sampling: MUS sample size, extraction, and evaluation helpers.
- Audit: duplicate, unique, missing-number, formula-list, fuzzy matching, and Luhn checks.
- Statistics: descriptive statistics, ANOVA, correlation, covariance, regression, histogram, moving average, smoothing, ranking, top/bottom tests, Z-score.
- Financial: EMI, fixed deposit, depreciation, payback, and discounted payback schedules.
- Array: array matching, filtering, sorting, row occurrence, and related utilities.
- Data: consolidate, pivot, unpivot, summary, and running total.
- Date: financial year, quarter, Indian workday/network-day helpers, and calendars.
- Text: contains/begins/ends checks, extraction, split, word count, number-to-words, fill down, and reverse text.
- Indian: PAN validator, Indian number-to-words, rupee symbol, Indian workday helpers.
- Help: all function forms, About, and the help workbook.

Most tools can be used through ribbon forms. Many functions are also available as worksheet UDFs with the `BI_` prefix.

## XLTX Lambda Template vs XLAM Add-in

Do not load the BeIndian `.xltx` LAMBDA template and this `.xlam` add-in in the same Excel session.

Both versions use `BI_` function names. If both are loaded together, Excel may resolve a formula to the workbook/template Lambda name instead of the add-in VBA function, which can make testing and results confusing.

Use one version at a time:

- Use `BeIndian2_VBA.xlam` for Excel 2019, Excel 2021, and Microsoft 365 on Windows.
- Use the `.xltx` LAMBDA template in Excel versions that support LAMBDA functions.

## Installation

1. Download `dist/BeIndian2_VBA.xlam`.
2. In Excel, open **File > Options > Add-ins**.
3. Choose **Excel Add-ins** in the Manage box and click **Go**.
4. Click **Browse**, select `BeIndian2_VBA.xlam`, and enable it.
5. If Windows blocks macros, right-click the downloaded `.xlam`, open **Properties**, choose **Unblock**, then reopen Excel.

For frequent use, place the `.xlam` in a trusted folder or Excel's AddIns folder before enabling it.

## Using the Add-in

- Select the input cell or range before opening a tool form.
- Forms normally prefill the input from the current Excel selection.
- Output defaults to the cell or range one column to the right.
- Use the **New sheet** option when you want results written to a separate sheet without changing the source sheet.
- For score models such as Altman, Beneish, Ohlson, and Piotroski, first create the score template, enter the required values, then run the calculation.

## Repository Layout

- `dist/` - release-ready Excel files and help workbook.
- `src/vba/` - exported VBA modules.
- `src/forms/` - exported VBA UserForms.
- `src/customUI/` - ribbon XML and icon assets.
- `manifest/` - add-in manifest/supporting metadata.
- `analysis/` - verification and development analysis artifacts.

## Verification Status

The current add-in build has been checked for:

- `.xlsm` and `.xlam` package integrity.
- Ribbon XML/image relationship validity.
- About form version and branding.
- Help workbook layout readability.
- Basic Excel COM open/load smoke test.
- Representative UDF execution smoke test.

Before using for live professional work, run a manual acceptance pass in your Excel environment, especially for the functions you rely on in an audit file.

## Notes

BeIndian is intended to support audit analytics and spreadsheet review. It does not replace professional judgment, audit documentation, or independent verification of results.

## Author

CA S. Rathinagiri
