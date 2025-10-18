import argparse
import os
import sys
import time

import pythoncom
import win32com.client as win32
from win32com.client import constants


def wait_for_refresh(app, wb, timeout_sec=300):
    start = time.time()

    # Let Excel finish async queries if supported
    try:
        app.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    # Wait for calculation completion
    try:
        while getattr(app, "CalculationState", constants.xlDone) != constants.xlDone:
            if time.time() - start > timeout_sec:
                raise TimeoutError(
                    "Timed out waiting for Excel calculations to finish."
                )
            time.sleep(0.2)
    except Exception:
        # Fallback to small wait
        time.sleep(1.0)

    # Explicitly wait for any QueryTables/ListObject queries still refreshing
    def any_refreshing():
        try:
            for ws in wb.Worksheets:
                # Worksheet.QueryTables
                try:
                    for qt in ws.QueryTables:
                        if getattr(qt, "Refreshing", False):
                            return True
                except Exception:
                    pass
                # ListObjects with QueryTable
                try:
                    for lo in ws.ListObjects:
                        try:
                            qt = lo.QueryTable
                            if getattr(qt, "Refreshing", False):
                                return True
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass
        return False

    while any_refreshing():
        if time.time() - start > timeout_sec:
            raise TimeoutError("Timed out waiting for query/table refresh.")
        time.sleep(0.25)


def export_range_as_image(ws, range_address, out_path, copy_format="bitmap"):
    rng = ws.Range(range_address)

    # Copy as picture
    fmt = constants.xlBitmap if copy_format.lower() == "bitmap" else constants.xlPicture
    rng.CopyPicture(Appearance=constants.xlScreen, Format=fmt)

    # Add a temporary chart sized to the range and paste the picture
    ch = ws.ChartObjects().Add(
        Left=rng.Left, Top=rng.Top, Width=rng.Width, Height=rng.Height
    )
    try:
        chart = ch.Chart
        # Sometimes paste needs a short delay
        for _ in range(10):
            try:
                chart.Paste()
                break
            except Exception:
                time.sleep(0.1)
        # Export infers format from file extension (png, jpg, gif, bmp)
        chart.Export(Filename=os.path.abspath(out_path))
    finally:
        # Clean up the temporary chart
        try:
            ch.Delete()
        except Exception:
            pass


def open_workbook(app, path):
    # Open without UI prompts
    wb = app.Workbooks.Open(
        os.path.abspath(path),
        UpdateLinks=0,
        ReadOnly=True,
        AddToMru=False,
        CorruptLoad=0,
    )
    return wb


def main():
    parser = argparse.ArgumentParser(
        description="Open Excel, refresh all, copy range to image."
    )
    parser.add_argument(
        "--workbook", "-w", required=True, help="Path to the Excel workbook (.xlsx)"
    )
    parser.add_argument("--sheet", "-s", required=True, help="Worksheet name")
    parser.add_argument(
        "--range", "-r", required=True, help='Range address, e.g. "A1:D10"'
    )
    parser.add_argument(
        "--out", "-o", required=True, help="Output image path (e.g. out.png)"
    )
    parser.add_argument("--visible", action="store_true", help="Show Excel window")
    parser.add_argument(
        "--timeout", type=int, default=300, help="Timeout in seconds for refresh"
    )
    args = parser.parse_args()

    if not os.path.exists(args.workbook):
        print(f"Workbook not found: {args.workbook}", file=sys.stderr)
        sys.exit(1)

    out_dir = os.path.dirname(os.path.abspath(args.out)) or "."
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        app = win32.gencache.EnsureDispatch("Excel.Application")
        app.Visible = bool(args.visible)
        app.DisplayAlerts = False
        # Optional: speed tweaks
        try:
            app.ScreenUpdating = False
            app.EnableEvents = False
        except Exception:
            pass

        wb = open_workbook(app, args.workbook)

        # Refresh all queries/pivots/connections
        wb.RefreshAll()
        wait_for_refresh(app, wb, timeout_sec=args.timeout)

        # Export the specified range from the specified sheet
        try:
            ws = wb.Worksheets(args.sheet)
        except Exception:
            print(f"Worksheet not found: {args.sheet}", file=sys.stderr)
            sys.exit(2)

        export_range_as_image(ws, args.range, args.out, copy_format="bitmap")
        print(f"Saved image to: {os.path.abspath(args.out)}")

    finally:
        # Restore and clean up
        if app:
            try:
                app.DisplayAlerts = False
                try:
                    app.ScreenUpdating = True
                    app.EnableEvents = True
                except Exception:
                    pass
            except Exception:
                pass
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
