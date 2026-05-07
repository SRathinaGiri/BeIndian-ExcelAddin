# Project Specification: BeIndian Excel Add-in (VBA)

## 1. Objective
Convert the "BeIndian" audit tool library from an O365-only LAMBDA template (`.xltx`) into a universal Excel Add-in (`.xlam`). This ensures compatibility with Excel 2019, 2021, and Microsoft 365, serving the Chartered Accountancy community for forensic and statutory audits.

## 2. Source Material
- **Logic Source:** `BeIndian.xltx` (Analyze the Name Manager for 100+ LAMBDA definitions).
- **Documentation:** Refer to the provided PDF for functional logic and expected outputs.
- **Reference URL:** https://github.com/SRathinaGiri/BeIndian

## 3. Technical Requirements
- **Language:** VBA (Visual Basic for Applications).
- **Target Format:** Excel Add-in (`.xlam`).
- **Compatibility:** Windows Excel 2019 and later.
- **UI:** Custom Ribbon Tab (XML-based).

## 4. Development Tasks

### A. VBA Module Generation (UDFs)
- Transform all LAMBDA functions into `Public Function` User Defined Functions (UDFs).
- **Requirement:** Every function must include an `IRibbonControl` parameter if called by a button, or be a standard Function for formula bar usage.
- **Key Functions to Port:** 
    - PAN/GST/Aadhaar Validators.
    - Benford's Law Analysis.
    - Altman Z-Score calculation.
    - Tax computation logic for the Indian context.

### B. Ribbon UI Construction
- Create a `customUI14.xml` file for the Ribbon.
- **Tab Label:** "BeIndian"
- **Groups:**
    1. **Validators:** Buttons for ID and Tax identity verification.
    2. **Forensics:** Tools for Benford's Law and outlier detection.
    3. **Analytics:** Financial ratios and audit sampling tools.
- **Handshake:** Ensure `onAction` attributes in XML match the Subroutine names in VBA exactly.

### C. Error Handling & Standards
- Implement `On Error GoTo` in all subroutines.
- Use `Option Explicit` in all modules.
- Include Javadoc-style comments for each function: `Description`, `Parameters`, `Returns`.

## 5. Deployment Instructions
1. Package all modules and the XML into a single `.xlsm`.
2. Provide instructions to "Save As" `.xlam`.
3. Generate a manifest for the Excel Add-ins dialog.