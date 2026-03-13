from __future__ import annotations

import json
import traceback
from datetime import datetime
from pathlib import Path

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None


XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52
XL_SRC_RANGE = 1
XL_YES = 1
XL_VALIDATE_WHOLE_NUMBER = 1
XL_VALIDATE_DECIMAL = 2
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1

INPUT_PARTS_TABLE_START_ROW = 110

SHEETS = [
    "INPUT_PARTS",
    "MATERIAL_SUMMARY",
    "PURCHASE_PLAN",
    "NEST_BASIC",
    "NEST_FINAL",
    "CUTLIST",
]

TABLES = {
    "INPUT_PARTS": (
        "tblParts",
        [
            "Job",
            "PartNo",
            "Material",
            "Diameter_in",
            "FinishLen_in",
            "ExtraStock_in",
            "Qty",
            "Kerf_in",
            "FaceAllow_in",
            "RequiredCutLen_in",
            "TotalLenReq_in",
            "StockKey",
        ],
    ),
    "MATERIAL_SUMMARY": (
        "tblMat",
        [
            "StockKey",
            "Material",
            "Diameter_in",
            "TotalReq_in",
            "TotalReq_ft",
            "BuyLen1_in",
            "BuyLen2_in",
            "PreferBuyLen1",
            "Notes",
        ],
    ),
    "PURCHASE_PLAN": (
        "tblPurchase",
        [
            "StockKey",
            "BuyLen1_in",
            "BuyLen2_in",
            "BarsNeeded_BASIC",
            "BarsNeeded_FINAL",
            "PurchasedLen_FINAL",
            "TotalReq_in",
            "TotalDrop_FINAL",
            "Waste%_FINAL",
            "ScrapLen_FINAL",
            "RemnantLen_FINAL",
            "Scrap%_FINAL",
            "Remnant%_FINAL",
        ],
    ),
    "NEST_FINAL": (
        "tblNestFinal",
        [
            "StockKey",
            "BuyLen_in",
            "PieceID",
            "CutLen_in",
            "Bar#",
            "RemainingBefore_in",
            "RemainingAfter_in",
            "IsBarEnd",
            "LeftoverAtEnd_in",
            "RemnantStatus",
        ],
    ),
    "CUTLIST": (
        "tblCutlist",
        [
            "StockKey",
            "Bar#",
            "BuyLen_in",
            "Pieces",
            "UsedLen_in",
            "Leftover_in",
            "RemnantStatus",
            "Util%",
        ],
    ),
}

BASIC_HEADERS = [
    "StockKey",
    "BuyLen_in",
    "PieceID",
    "CutLen_in",
    "Bar#",
    "RemainingBefore_in",
    "RemainingAfter_in",
    "IsBarEnd",
    "LeftoverAtEnd_in",
    "RemnantStatus",
]

SEEDS = [11, 22, 33, 44, 55]
TOL = 1e-6
HERE = Path(__file__).resolve().parent


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(level: str, msg: str) -> None:
    print(f"{_ts()} | {level:<5} | {msg}", flush=True)


def log_info(msg: str) -> None:
    log("INFO", msg)


def log_warn(msg: str) -> None:
    log("WARN", msg)


def log_error(msg: str) -> None:
    log("ERROR", msg)


def write_validation_json(validation_path: Path, payload: dict) -> None:
    validation_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def make_headers_row(ws, headers, row=1, start_col=1):
    for idx, header in enumerate(headers, start=start_col):
        ws.Cells(row, idx).Value = header
        ws.Cells(row, idx).Font.Bold = True


def create_table(ws, start_row: int, start_col: int, headers, table_name: str):
    end_col = start_col + len(headers) - 1
    make_headers_row(ws, headers, start_row, start_col)
    rng = ws.Range(ws.Cells(start_row, start_col), ws.Cells(start_row + 1, end_col))
    log_info(f"Creating table {table_name} on {ws.Name} at {rng.Address}")
    lo = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    return lo


def set_named_range(wb, name: str, ws_name: str, cell_address: str):
    refers_to = f"='{ws_name}'!{cell_address}"
    try:
        wb.Names(name).Delete()
    except Exception:
        pass
    wb.Names.Add(Name=name, RefersTo=refers_to)


def try_validation_decimal(rng, warnings, name: str):
    try:
        rng.Validation.Delete()
    except Exception:
        pass
    try:
        rng.Validation.Add(
            XL_VALIDATE_DECIMAL,
            XL_VALID_ALERT_STOP,
            XL_BETWEEN,
            "0.000001",
            "1000000",
        )
        rng.Validation.IgnoreBlank = True
        rng.Validation.InputTitle = name
        rng.Validation.ErrorTitle = f"Invalid {name}"
        rng.Validation.InputMessage = f"Enter numeric {name} > 0"
        rng.Validation.ErrorMessage = f"{name} must be numeric and > 0"
        rng.Validation.ShowError = True
    except Exception as exc:
        warnings.append(f"Validation warning ({name}): {exc}")
        log_warn(f"Validation warning ({name}): {exc}")


def try_validation_whole(rng, warnings, name: str):
    try:
        rng.Validation.Delete()
    except Exception:
        pass
    try:
        rng.Validation.Add(
            XL_VALIDATE_WHOLE_NUMBER,
            XL_VALID_ALERT_STOP,
            XL_BETWEEN,
            "1",
            "1000000",
        )
        rng.Validation.IgnoreBlank = True
        rng.Validation.InputTitle = name
        rng.Validation.ErrorTitle = f"Invalid {name}"
        rng.Validation.InputMessage = f"Enter whole number {name} >= 1"
        rng.Validation.ErrorMessage = f"{name} must be a whole number >= 1"
        rng.Validation.ShowError = True
    except Exception as exc:
        warnings.append(f"Validation warning ({name}): {exc}")
        log_warn(f"Validation warning ({name}): {exc}")


def try_validation_list_allow_free_text(rng, warnings, formula1: str, name: str):
    try:
        rng.Validation.Delete()
    except Exception:
        pass
    try:
        rng.Validation.Add(
            XL_VALIDATE_LIST,
            XL_VALID_ALERT_STOP,
            XL_BETWEEN,
            formula1,
        )
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        rng.Validation.InputTitle = name
        rng.Validation.ErrorTitle = name
        rng.Validation.InputMessage = "Choose from list or type a new value"
        rng.Validation.ErrorMessage = "Choose from list or type a new value"
        rng.Validation.ShowError = False
    except Exception as exc:
        warnings.append(f"Validation warning ({name}): {exc}")
        log_warn(f"Validation warning ({name}): {exc}")


def normalize_2d(value):
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [[value]]
    if len(value) == 0:
        return []
    if not isinstance(value[0], tuple):
        return [list(value)]
    return [list(r) for r in value]


def get_table_rows(lo):
    try:
        body = lo.DataBodyRange
        if body is None:
            return []
        return normalize_2d(body.Value)
    except Exception as exc:
        log_warn(f"Could not read table rows for {lo.Name}: {exc}")
        return []


def scan_sheet_for_error_strings(ws):
    errs = []
    try:
        used = ws.UsedRange
        values = normalize_2d(used.Value)
        for r_idx, row in enumerate(values, start=1):
            for c_idx, val in enumerate(row, start=1):
                if isinstance(val, str) and val.startswith("#"):
                    errs.append(f"{ws.Name}!R{r_idx}C{c_idx}={val}")
    except Exception as exc:
        errs.append(f"{ws.Name}: error scan failed: {exc}")
    return errs


def assert_vbproject_access(wb):
    log_info("Checking VBProject access...")
    try:
        vbproj = wb.VBProject
        _ = vbproj.VBComponents.Count
        return vbproj
    except Exception as exc:
        raise RuntimeError(
            "Excel blocked programmatic VBA project access. "
            "Enable Excel > Options > Trust Center > Trust Center Settings > Macro Settings > "
            "'Trust access to the VBA project object model', then rerun."
        ) from exc


def seed_materials(ws):
    materials_seed = [
        "1018", "12L14", "4140", "4340", "8620", "17-4PH", "304", "316", "O1", "D2",
        "H13", "A2", "Brass360", "Al6061", "Al7075", "Ti-6Al-4V", "Monel400", "Inconel625", "UHMW", "Acetal"
    ]
    ws.Range("G2").Value = "Materials"
    ws.Range("G3:G102").ClearContents()
    for idx, val in enumerate(materials_seed, start=3):
        ws.Cells(idx, 7).Value = val


def build_workbook(out_path: Path, bas_path: Path, validation_path: Path):
    warnings = []
    result = {
        "build": {
            "final_saved": False,
            "warnings": warnings,
            "traceback": None,
        },
        "selftests": [],
        "reconciliation": {
            "executed": False,
            "tolerance": TOL,
            "details": [],
            "all_passed": False,
        },
    }

    if win32com is None:
        msg = "pywin32 / win32com is not installed."
        warnings.append(msg)
        write_validation_json(validation_path, result)
        raise RuntimeError(msg)

    if not bas_path.exists():
        msg = f"Missing BAS file: {bas_path}"
        warnings.append(msg)
        write_validation_json(validation_path, result)
        raise FileNotFoundError(msg)

    xl = None
    wb = None

    try:
        log_info("build_xlsm.py starting")
        log_info(f"Working folder: {HERE}")
        log_info(f"Output XLSM: {out_path}")
        log_info(f"BAS file: {bas_path}")
        log_info(f"Validation JSON: {validation_path}")

        log_info("DispatchEx Excel.Application")
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        xl.EnableEvents = False

        try:
            xl.AutomationSecurity = 1
        except Exception as exc:
            warnings.append(f"Could not set AutomationSecurity: {exc}")
            log_warn(f"Could not set AutomationSecurity: {exc}")

        log_info("Creating new workbook from scratch")
        wb = xl.Workbooks.Add()

        log_info("Resizing sheet count to exactly 6")
        while wb.Worksheets.Count < 6:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        while wb.Worksheets.Count > 6:
            wb.Worksheets(wb.Worksheets.Count).Delete()

        log_info("Renaming sheets to required order")
        for i, name in enumerate(SHEETS, start=1):
            wb.Worksheets(i).Name = name

        # INPUT_PARTS
        log_info("Building INPUT_PARTS")
        ws = wb.Worksheets("INPUT_PARTS")

        ws.Range("A1").Value = "ROUND BAR NESTING INPUT"
        ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "Setup values and materials list are above."
        ws.Range("A3").Font.Italic = True

        seed_materials(ws)

        ws.Range("I2").Value = "DefaultKerf_in"
        ws.Range("J2").Value = 0.125
        ws.Range("I3").Value = "DefaultFaceAllow_in"
        ws.Range("J3").Value = 0.125
        ws.Range("I4").Value = "MinRemnantKeep_in"
        ws.Range("J4").Value = 12
        ws.Range("I5").Value = "MaxPartRows"
        ws.Range("J5").Value = 500
        ws.Range("I6").Value = "MaxPieces"
        ws.Range("J6").Value = 20000
        ws.Range("I7").Value = "LastFinalRun"
        ws.Range("J7").Value = ""
        ws.Range("I8").Value = "LastCutlistGen"
        ws.Range("J8").Value = ""

        log_info("Creating named ranges")
        set_named_range(wb, "DefaultKerf_in", "INPUT_PARTS", "$J$2")
        set_named_range(wb, "DefaultFaceAllow_in", "INPUT_PARTS", "$J$3")
        set_named_range(wb, "MinRemnantKeep_in", "INPUT_PARTS", "$J$4")
        set_named_range(wb, "MaxPartRows", "INPUT_PARTS", "$J$5")
        set_named_range(wb, "MaxPieces", "INPUT_PARTS", "$J$6")
        set_named_range(wb, "LastFinalRun", "INPUT_PARTS", "$J$7")
        set_named_range(wb, "LastCutlistGen", "INPUT_PARTS", "$J$8")
        set_named_range(wb, "Materials", "INPUT_PARTS", "$G$3:$G$102")

        log_info(f"Creating tblParts starting at row {INPUT_PARTS_TABLE_START_ROW}")
        lo_parts = create_table(
            ws,
            INPUT_PARTS_TABLE_START_ROW,
            1,
            TABLES["INPUT_PARTS"][1],
            TABLES["INPUT_PARTS"][0],
        )

        log_info("Applying safe formulas to tblParts")
        lo_parts.ListColumns("RequiredCutLen_in").DataBodyRange.Formula = (
            '=IFERROR(N([@[FinishLen_in]])+N([@[ExtraStock_in]])+'
            'IF(OR([@[Kerf_in]]="",ISBLANK([@[Kerf_in]])),DefaultKerf_in,N([@[Kerf_in]]))+'
            'IF(OR([@[FaceAllow_in]]="",ISBLANK([@[FaceAllow_in]])),DefaultFaceAllow_in,N([@[FaceAllow_in]])),0)'
        )
        lo_parts.ListColumns("TotalLenReq_in").DataBodyRange.Formula = '=IFERROR(N([@[Qty]])*N([@[RequiredCutLen_in]]),0)'
        lo_parts.ListColumns("StockKey").DataBodyRange.Formula = '=TRIM([@[Material]])&"|"&TEXT([@[Diameter_in]],"0.000")'

        log_info("Attempting input validations")
        max_rows = 500
        first_data_row = lo_parts.HeaderRowRange.Row + 1
        try_validation_decimal(
            ws.Range(
                ws.Cells(first_data_row, lo_parts.ListColumns("Diameter_in").Range.Column),
                ws.Cells(first_data_row + max_rows - 1, lo_parts.ListColumns("Diameter_in").Range.Column),
            ),
            warnings,
            "Diameter_in",
        )
        try_validation_whole(
            ws.Range(
                ws.Cells(first_data_row, lo_parts.ListColumns("Qty").Range.Column),
                ws.Cells(first_data_row + max_rows - 1, lo_parts.ListColumns("Qty").Range.Column),
            ),
            warnings,
            "Qty",
        )
        try_validation_list_allow_free_text(
            ws.Range(
                ws.Cells(first_data_row, lo_parts.ListColumns("Material").Range.Column),
                ws.Cells(first_data_row + max_rows - 1, lo_parts.ListColumns("Material").Range.Column),
            ),
            warnings,
            "=Materials",
            "Material",
        )

        # MATERIAL_SUMMARY
        log_info("Building MATERIAL_SUMMARY")
        ws = wb.Worksheets("MATERIAL_SUMMARY")
        create_table(ws, 1, 1, TABLES["MATERIAL_SUMMARY"][1], TABLES["MATERIAL_SUMMARY"][0])

        # PURCHASE_PLAN
        log_info("Building PURCHASE_PLAN")
        ws = wb.Worksheets("PURCHASE_PLAN")
        create_table(ws, 1, 1, TABLES["PURCHASE_PLAN"][1], TABLES["PURCHASE_PLAN"][0])

        # NEST_BASIC
        log_info("Building NEST_BASIC plain sheet")
        ws = wb.Worksheets("NEST_BASIC")
        make_headers_row(ws, BASIC_HEADERS, 1, 1)

        # NEST_FINAL
        log_info("Building NEST_FINAL")
        ws = wb.Worksheets("NEST_FINAL")
        create_table(ws, 1, 1, TABLES["NEST_FINAL"][1], TABLES["NEST_FINAL"][0])

        # CUTLIST
        log_info("Building CUTLIST")
        ws = wb.Worksheets("CUTLIST")
        create_table(ws, 1, 1, TABLES["CUTLIST"][1], TABLES["CUTLIST"][0])

        log_info("Auto-fitting columns")
        for ws in wb.Worksheets:
            ws.Rows(1).Font.Bold = True
            ws.Columns.AutoFit()

        log_info("Checking VBA project access before import")
        vbproj = assert_vbproject_access(wb)

        log_info("Importing BAS module into VBProject")
        vbproj.VBComponents.Import(str(bas_path))

        log_info("Injecting ThisWorkbook.Workbook_Open")
        this_wb = vbproj.VBComponents("ThisWorkbook")
        code_mod = this_wb.CodeModule
        if code_mod.CountOfLines > 0:
            code_mod.DeleteLines(1, code_mod.CountOfLines)
        code_mod.AddFromString(
            "Option Explicit\n"
            "Private Sub Workbook_Open()\n"
            "    On Error Resume Next\n"
            "    EnsureInputValidations\n"
            "    EnsurePrintButton\n"
            "End Sub\n"
        )

        if out_path.exists():
            log_info(f"Deleting existing output file: {out_path}")
            out_path.unlink()

        log_info("Saving workbook as XLSM")
        wb.SaveAs(str(out_path), FileFormat=XL_OPENXML_WORKBOOK_MACRO_ENABLED)
        result["build"]["final_saved"] = True

        log_info("Running seed self-tests")
        workbook_macro_prefix = f"'{out_path.name}'!"
        all_passed = True

        for seed in SEEDS:
            entry = {
                "seed": seed,
                "executed": False,
                "runtime_error": None,
                "formula_errors": [],
                "reconciliation": {},
            }

            try:
                log_info(f"Self-test seed {seed}: ResetDemoData")
                xl.Run(workbook_macro_prefix + "ResetDemoData")

                log_info(f"Self-test seed {seed}: GenerateDemoData")
                xl.Run(workbook_macro_prefix + "GenerateDemoData", seed, 100, 20)

                log_info(f"Self-test seed {seed}: RunFinalNesting")
                xl.Run(workbook_macro_prefix + "RunFinalNesting")

                entry["executed"] = True

                formula_errors = []
                for ws in wb.Worksheets:
                    formula_errors.extend(scan_sheet_for_error_strings(ws))
                entry["formula_errors"] = formula_errors

                log_info(f"Self-test seed {seed}: reading final nesting and purchase plan")
                ws_final = wb.Worksheets("NEST_FINAL")
                ws_purchase = wb.Worksheets("PURCHASE_PLAN")
                lo_final = ws_final.ListObjects("tblNestFinal")
                lo_purchase = ws_purchase.ListObjects("tblPurchase")

                final_rows = get_table_rows(lo_final)
                purchase_rows = get_table_rows(lo_purchase)

                unique_bars = {}
                for r in final_rows:
                    stock_key = str(r[0]).strip()
                    if not stock_key:
                        continue
                    bar_no = str(r[4]).strip()
                    buy_len = float(r[1] or 0.0)
                    unique_bars[(stock_key, bar_no)] = buy_len

                sum_bar_buy = float(sum(unique_bars.values()))

                plan_purchased = 0.0
                plan_totaldrop_ok = True
                plan_rows_checked = 0
                plan_mismatches = []

                for r in purchase_rows:
                    stock_key = str(r[0]).strip()
                    if not stock_key:
                        continue

                    purchased = float(r[5] or 0.0)
                    total_req = float(r[6] or 0.0)
                    total_drop = float(r[7] or 0.0)

                    expected_total_drop = purchased - total_req
                    if abs(expected_total_drop - total_drop) > TOL:
                        plan_totaldrop_ok = False
                        plan_mismatches.append(
                            {
                                "stock_key": stock_key,
                                "purchased": purchased,
                                "total_req": total_req,
                                "total_drop": total_drop,
                                "expected_total_drop": expected_total_drop,
                            }
                        )

                    plan_purchased += purchased
                    plan_rows_checked += 1

                purchased_match = abs(sum_bar_buy - plan_purchased) <= TOL

                entry["reconciliation"] = {
                    "sum_bar_buy_lengths_from_final": sum_bar_buy,
                    "sum_purchased_len_from_plan": plan_purchased,
                    "purchased_match": purchased_match,
                    "totaldrop_match_all_rows": plan_totaldrop_ok,
                    "plan_rows_checked": plan_rows_checked,
                    "mismatches": plan_mismatches,
                    "tolerance": TOL,
                }

                if formula_errors or (not purchased_match) or (not plan_totaldrop_ok):
                    all_passed = False

                result["reconciliation"]["details"].append(
                    {
                        "seed": seed,
                        **entry["reconciliation"],
                    }
                )

                log_info(
                    f"Self-test seed {seed} complete. "
                    f"purchased_match={purchased_match}, "
                    f"totaldrop_match_all_rows={plan_totaldrop_ok}, "
                    f"formula_errors={len(formula_errors)}"
                )

            except Exception as exc:
                entry["runtime_error"] = f"{type(exc).__name__}: {exc}"
                all_passed = False
                log_error(f"Self-test seed {seed} failed: {type(exc).__name__}: {exc}")
                tb = traceback.format_exc()
                for line in tb.rstrip().splitlines():
                    log_error(line)

            result["selftests"].append(entry)

        result["reconciliation"]["executed"] = True
        result["reconciliation"]["all_passed"] = all_passed

        log_info("Writing validation JSON")
        write_validation_json(validation_path, result)

        log_info("Saving workbook after self-tests")
        wb.Save()

        if not all_passed:
            raise RuntimeError("One or more self-tests failed validation. See ROUND_BAR_Nesting_Template.validation.json")

        log_info("build_xlsm.py completed successfully")
        return result

    except Exception:
        result["build"]["traceback"] = traceback.format_exc()
        write_validation_json(validation_path, result)
        log_error("build_xlsm.py failed")
        for line in traceback.format_exc().rstrip().splitlines():
            log_error(line)
        raise

    finally:
        if wb is not None:
            try:
                log_info("Closing workbook")
                wb.Close(SaveChanges=False)
            except Exception as exc:
                log_warn(f"Workbook close warning: {exc}")

        if xl is not None:
            try:
                xl.ScreenUpdating = True
                xl.EnableEvents = True
            except Exception:
                pass
            try:
                log_info("Quitting Excel")
                xl.Quit()
            except Exception as exc:
                log_warn(f"Excel quit warning: {exc}")


def main():
    out_path = HERE / "ROUND_BAR_Nesting_Template.xlsm"
    bas_path = HERE / "modRoundBarNesting.bas"
    validation_path = HERE / "ROUND_BAR_Nesting_Template.validation.json"

    result = build_workbook(out_path, bas_path, validation_path)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()