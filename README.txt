ROUND BAR STOCK PURCHASING + 1D CUT NESTING TEMPLATE
====================================================

AUTHORING HONESTY
-----------------
This package was authored in a non-Windows environment, so the authoring environment did NOT run:
- Windows Excel COM
- actual XLSM generation
- actual VBA import into Excel
- actual self-tests against Excel

The files in this package are written so that on a Windows PC with:
- Excel 365 desktop
- Python 3.x
they will build the workbook FROM SCRATCH using Excel COM, import VBA, save XLSM, and run the required self-tests.

NO FAKE CLAIMS:
- No claim is made here that the XLSM was already built.
- No claim is made here that COM validation already happened.
- The validation JSON included in this package is a pre-run scaffold only.
- The real validation JSON is generated/overwritten by build_xlsm.py on Windows.

PACKAGE CONTENTS
----------------
1) ROUND_BAR_Nesting_Template.xlsm
   - Generated on Windows by build_xlsm.py

2) modRoundBarNesting.bas
   - Single VBA module containing:
     - EnsureInputValidations
     - EnsurePrintButton
     - ResetDemoData
     - GenerateDemoData
     - RefreshAll
     - RunBasicNesting
     - RunFinalNesting
     - BuildCutlistFromFinal
     - PrintCutlist
     - extensive debug logging and defensive helpers

3) build_xlsm.py
   - Creates workbook FROM SCRATCH using Excel COM DispatchEx
   - Creates exact required sheets in exact required order
   - Creates required tables and named ranges
   - Writes only safe formulas
   - Attempts DataValidation.Add but does not hard-fail on validation issues
   - Imports VBA module
   - Injects ThisWorkbook Workbook_Open code
   - Saves as XLSM
   - Runs self-tests for seeds 11, 22, 33, 44, 55
   - Writes validation JSON

4) build_xlsm.ps1
   - Closes Excel
   - Installs pywin32 if missing
   - Runs build_xlsm.py
   - Tees output to build.log
   - Prints failure tail if build fails

5) RUN_ME.ps1
   - Convenience wrapper around build_xlsm.ps1

6) RUN_ME.cmd
   - Double-click entry point for Windows users

7) ROUND_BAR_Nesting_Template.validation.json
   - Pre-run scaffold only
   - Real file is overwritten by build_xlsm.py after Windows COM build/test

PLATFORM / API ASSUMPTIONS
--------------------------
- Windows 10/11
- Microsoft Excel 365 desktop installed locally
- Python 3.x installed locally
- Excel macro settings allow VBA macros to run
- Excel Trust Center option "Trust access to the VBA project object model" is enabled
  (required so Python can import the VBA module and inject ThisWorkbook code)

HOW TO RUN
----------
Method A (recommended):
1) Put all package files in one folder.
2) Right-click RUN_ME.cmd and run.
3) Wait for build to finish.
4) Review:
   - ROUND_BAR_Nesting_Template.xlsm
   - ROUND_BAR_Nesting_Template.validation.json
   - build.log

Method B (PowerShell):
1) Open PowerShell in the package folder.
2) Run:
   powershell -ExecutionPolicy Bypass -File .\RUN_ME.ps1

Method C (direct):
1) Open PowerShell in the package folder.
2) Run:
   powershell -ExecutionPolicy Bypass -File .\build_xlsm.ps1

WORKBOOK STRUCTURE
------------------
Sheets in exact order:
1) INPUT_PARTS
2) MATERIAL_SUMMARY
3) PURCHASE_PLAN
4) NEST_BASIC
5) NEST_FINAL
6) CUTLIST

Named ranges on INPUT_PARTS:
- DefaultKerf_in
- DefaultFaceAllow_in
- MinRemnantKeep_in
- MaxPartRows
- MaxPieces
- LastFinalRun
- LastCutlistGen
- Materials

CORE WORKFLOW
-------------
1) Enter or generate rows in INPUT_PARTS.
2) Run RefreshAll.
3) Run RunBasicNesting for deterministic Next-Fit baseline.
4) Run RunFinalNesting for improved deterministic bar packing.
5) Run BuildCutlistFromFinal.
6) Use PrintCutlist for print-ready grouped output with page breaks per StockKey.

LOGGING
-------
All VBA logging goes to Immediate Window using:
- LogInfo(proc, msg)
- LogError(proc, errNum, errDesc)

All Python build logging goes to console and build.log.

DETERMINISM
-----------
- BASIC nesting = deterministic Next-Fit
- FINAL nesting = deterministic Best-Fit Decreasing on already-open bars,
  with deterministic new-bar choice logic
- Demo generation is deterministic for a given seed

KNOWN LIMITATIONS
-----------------
1) Real XLSM generation cannot happen unless run on Windows with Excel desktop.
2) VBA import requires "Trust access to the VBA project object model".
3) Material dropdown is configured to allow free text by disabling validation error alert.
4) NEST_FINAL uses improved deterministic packing, but it is still 1D bar nesting only.
5) FINAL nesting does not optimize cost by price-per-length because pricing data was not requested.
6) Page-break behavior may vary slightly with printer driver, but macro resets breaks and reapplies them each time.

FUTURE IMPROVEMENTS
-------------------
- Add optional stock pricing and best-buy-length cost optimization
- Add saw setup grouping / machine grouping
- Add remnant inventory carry-forward sheet
- Add export to CSV/PDF package
- Add lockable released cutlist snapshots
- Add part-family grouping and tooling notes