"""Opt-in Windows/Excel smoke test. Never open supplied files or execute VBA macros."""
from __future__ import annotations
import argparse
import importlib.util
import json
import platform
import sys
import tempfile
from pathlib import Path
from common import write_new_json, validate_new_output

def dependencies() -> dict:
    windows = platform.system() == "Windows"
    return {"platform": platform.system(), "python": platform.python_version(),
            "windows": windows, "pywin32_available": windows and importlib.util.find_spec("win32com") is not None,
            "excel_installation": "not_probed", "excel_started": False}

def execute_smoke(factory, directory: Path) -> dict:
    """Inject a factory for unit tests. Real factory must return a NEW owned instance."""
    app = book = sheet = target = source = chart_object = chart = None
    results, cleanup = [], []
    error = None
    output = directory / "smoke-result.xlsx"
    if output.exists():
        raise FileExistsError(output)
    try:
        app = factory()
        # Do not suppress prompts or security. New blank workbook, no input file.
        app.Visible = True
        book = app.Workbooks.Add(Template=-4167)  # xlWBATWorksheet; avoid a user default Book.xltm template.
        sheet = book.Worksheets(1)
        sheet.Name = "Smoke"
        target = sheet.Range("A1:B3")
        target.Value2 = (("Label", "Value"), ("A", 3), ("B", 5))
        target = None
        target = sheet.Range("D1")
        target.Formula = "=SUM(B2:B3)"
        app.Calculate()
        observed = target.Value2
        if observed != 8:
            raise AssertionError(f"Expected D1=8, observed {observed!r}")
        results.append({"test": "formula_value", "status": "PASS", "observed": observed})
        target = None
        source = sheet.Range("A1:B3")
        chart_object = sheet.ChartObjects().Add(200, 20, 320, 200)
        chart = chart_object.Chart
        chart.SetSourceData(source)
        chart.ChartType = 51  # xlColumnClustered; not the file format despite same numeric value.
        if chart.SeriesCollection().Count != 1:
            raise AssertionError("Expected one chart series")
        results.append({"test": "chart_series", "status": "PASS"})
        chart = chart_object = source = target = sheet = None
        book.SaveAs(str(output.resolve()), FileFormat=51)  # xlOpenXMLWorkbook
        if not output.is_file():
            raise AssertionError("SaveAs did not produce the expected file")
        results.append({"test": "save_new_xlsx", "status": "PASS"})
    except Exception as exc:
        error = {"type": type(exc).__name__, "message": str(exc)}
    finally:
        chart = chart_object = source = target = sheet = None
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception as exc:
                cleanup.append({"operation": "Close", "error": str(exc)})
            book = None
        if app is not None:
            try:
                app.Quit()
            except Exception as exc:
                cleanup.append({"operation": "Quit", "error": str(exc)})
            app = None
    return {"status": "PASS" if error is None and not cleanup else "FAIL", "tests": results,
            "primary_error": error, "cleanup_errors": cleanup, "output": str(output),
            "scope": "New blank workbook; formula, chart and save only. No VBA compile, macro, UI or process-exit assertion."}

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--run-local-excel", action="store_true", help="Explicitly authorize starting an isolated Excel application")
    p.add_argument("--out", type=Path, help="Write NEW JSON report; existing files are not replaced")
    a = p.parse_args()
    try:
        validate_new_output(a.out)
    except OSError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    deps = dependencies()
    if not a.run_local_excel:
        report = {"status": "NOT_RUN", "dependencies": deps, "instruction": "Read tests/WINDOWS-ACCEPTANCE.md before opting in"}
    elif not deps["windows"] or not deps["pywin32_available"]:
        report = {"status": "BLOCKED_DEPENDENCY", "dependencies": deps}
    else:
        import pythoncom
        from win32com.client import DispatchEx
        directory = Path(tempfile.mkdtemp(prefix="excel-vba-expert-"))
        pythoncom.CoInitialize()
        try:
            report = execute_smoke(lambda: DispatchEx("Excel.Application"), directory)
            report["evidence"] = "REAL_EXCEL_COM"
            report["dependencies"] = deps
        finally:
            pythoncom.CoUninitialize()
    try:
        if a.out:
            write_new_json(a.out, report)
        print(json.dumps(report, indent=2))
        return 0 if report["status"] in {"PASS", "NOT_RUN"} else 1
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
