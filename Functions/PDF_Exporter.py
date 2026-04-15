import xlwings as xw
import os
from pathlib import Path
from datetime import datetime
import shutil

def MeaningfulUsedRange(sht) -> bool:
    try:
        used = sht.used_range
        addr = used.address
        a1 = sht.range("A1").value
        return bool(used and addr and (addr != "$A$1" or a1 not in (None, "")))
    except Exception:
        return False

def ClearHeadersFooters(PageSetup):
    try:
        PageSetup.api.left_header.set("")
        PageSetup.api.center_header.set("")
        PageSetup.api.right_header.set("")
        PageSetup.api.left_footer.set("")
        PageSetup.api.center_footer.set("")
        PageSetup.api.right_footer.set("")
    except Exception:
        pass

    try:
        PageSetup.api.header_margin.set(0)
        PageSetup.api.footer_margin.set(0)
    except Exception:
        pass

def ApplyTightMargins(PageSetup):
    try:
        PageSetup.api.left_margin.set(15)
        PageSetup.api.right_margin.set(15)
        PageSetup.api.top_margin.set(10)
        PageSetup.api.bottom_margin.set(10)
        PageSetup.api.header_margin.set(0)
        PageSetup.api.footer_margin.set(0)
    except Exception:
        pass

def ApplyOnePageSetup(sht, print_range=None):
    PageSetup = sht.page_setup

    if print_range:
        PageSetup.print_area = print_range
    elif not PageSetup.print_area and MeaningfulUsedRange(sht):
        PageSetup.print_area = sht.used_range.address

    try:
        PageSetup.api.zoom.set(False)
    except Exception:
        pass

    try:
        PageSetup.api.fit_to_pages_wide.set(1)
        PageSetup.api.fit_to_pages_tall.set(1)
    except Exception:
        pass

    try:
        PageSetup.api.center_horizontally.set(True)
    except Exception:
        pass

    try:
        PageSetup.api.center_vertically.set(False)
    except Exception:
        pass

    ClearHeadersFooters(PageSetup)
    ApplyTightMargins(PageSetup)

def ExportWeeklySnapshot(workbook_path, output_folder, username):
    SKIP_SHEETS = {"Time_Series_Indices", "FX_Series", "Annual_Returns_Indices", "Annual_Returns_FX", "1M_Time_Series_Indices", 
               "1Y_Time_Series_Indices", "1M_Time_Series_FX", "Sector_Returns"}

    PRINT_RANGES = {
        "Dashboard": "C1:I88",
    }

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    wb = None

    try:
        workbook_path = str(Path(workbook_path).expanduser().resolve())
        output_dir = Path(output_folder).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        wb = app.books.open(workbook_path)

        export_sheets = []

        for sht in wb.sheets:
            if sht.name in SKIP_SHEETS:
                continue

            forced_range = PRINT_RANGES.get(sht.name)
            ApplyOnePageSetup(sht, forced_range)

            try:
                current_print_area = sht.page_setup.print_area
            except Exception:
                current_print_area = None

            if current_print_area:
                export_sheets.append(sht.name)
                print(f"Included: {sht.name} | Range: {current_print_area}")

        if not export_sheets:
            raise ValueError("No printable sheets found.")

        stamp = datetime.now().strftime("%Y%m%d")
        final_pdf = output_dir / f"{stamp}_Weekly_Benchmarks_Snapshot_{username}.pdf"

        wb.to_pdf(path=str(final_pdf), include=export_sheets)

        return str(final_pdf)

    finally:
        if wb is not None:
            wb.close()
        app.quit()

def SharePointUploadPDF(PDF_Path, Sharepoint_Folder, username):
    if not os.path.exists(PDF_Path):
        raise FileNotFoundError(f"PDF file not found: {PDF_Path}")

    if not os.path.exists(Sharepoint_Folder):
        raise FileNotFoundError(f"SharePoint synced folder not found: {Sharepoint_Folder}")

    timestamp = datetime.now().strftime("%Y%m%d")
    file_name = f"{timestamp}_Weekly_Benchmarks_Snapshot_{username}.pdf"
    target_path = os.path.join(Sharepoint_Folder, file_name)

    shutil.copy2(PDF_Path, target_path)

    print(f"Copied PDF to synced folder: {target_path}")
    return target_path