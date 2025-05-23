import xlwings as xw  # type: ignore


def get_sheet_names():
    """Get all sheet names from the active Excel workbook."""
    # Connect to the active Excel application
    app = xw.apps.active

    # Get the active workbook
    wb = app.books.active

    # Get all sheet names
    sheet_names = [sheet.name for sheet in wb.sheets]

    return sheet_names


if __name__ == "__main__":
    try:
        sheet_names = get_sheet_names()
        print("Sheet names in the active workbook:")
        for i, name in enumerate(sheet_names, 1):
            print(f"{i}. {name}")
    except Exception as e:
        print(f"Error: {e}")
        print("Make sure Excel is open with a workbook.")
