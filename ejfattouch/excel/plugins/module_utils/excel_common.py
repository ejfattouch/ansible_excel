import os
import platform

def check_excel_installation():
    """Check if Excel is installed on Windows or macOS."""
    if platform.system() == 'Windows':
        try:
            excel_path = os.path.join(os.environ["ProgramFiles"], "Microsoft Office", "root", "Office16", "EXCEL.EXE")
            excel_path_x86 = os.path.join(os.environ["ProgramFiles(x86)"], "Microsoft Office", "root", "Office16",
                                          "EXCEL.EXE")
            return os.path.exists(excel_path) or os.path.exists(excel_path_x86)
        except KeyError:
            return False
    elif platform.system() == 'Darwin':
        return os.path.exists("/Applications/Microsoft Excel.app")
    else:
        raise RuntimeError("Only Windows and MacOS are supported for evaluation")


def evaluate_workbook_formulas(path):
    """Use xlwings to force Excel to recalculate formulas."""
    import xlwings as xw
    excel_app = xw.App(visible=False)
    excel_book = excel_app.books.open(path)
    excel_book.save()
    excel_book.close()
    excel_app.quit()


def read_sheet_data(sheet):
    """Read all data from a worksheet, returns list of rows."""
    data = [list(row) for row in sheet.iter_rows(values_only=True)]

    # Remove rows from the bottom that contain only None
    while data and all(cell is None for cell in data[-1]):
        data.pop()

    return data
