from mcp.server.fastmcp import FastMCP
import xlwings as xw  # type: ignore
from tabulate import tabulate
from pathlib import Path

# Create an MCP server
mcp = FastMCP("Excel API")


@mcp.tool()
def get_sheet_names() -> list[str]:
    """Get all sheet names from the active Excel workbook."""
    app = xw.apps.active
    wb = app.books.active
    return [sheet.name for sheet in wb.sheets]


@mcp.tool()
def read_cell(sheet_name: str, cell_address: str, get_formula: bool = False) -> str:
    """Read a single cell value or formula from Excel.

    Args:
        sheet_name: Name of the sheet
        cell_address: Cell address like 'A1', 'B5', etc.
        get_formula: If True, return formula; if False, return calculated value
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]
    cell = sheet.range(cell_address)

    if get_formula:
        formula = cell.formula
        return (
            formula
            if formula is not None
            else str(cell.value) if cell.value is not None else ""
        )
    else:
        value = cell.value
        return str(value) if value is not None else ""


@mcp.tool()
def read_range(
    sheet_name: str, range_address: str, get_formulas: bool = False
) -> list[list]:
    """Read a range of cells from Excel.

    Args:
        sheet_name: Name of the sheet
        range_address: Range address like 'A1:C3', 'B2:D10', etc.
        get_formulas: If True, return formulas; if False, return calculated values
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]

    if get_formulas:
        formulas = sheet.range(range_address).formula

        # Handle single cell case
        if not isinstance(formulas, list):
            formula_value = (
                formulas if formulas is not None else sheet.range(range_address).value
            )
            return [[str(formula_value) if formula_value is not None else ""]]

        # Handle single row case
        if not isinstance(formulas[0], list):
            result = []
            for i, formula in enumerate(formulas):
                if formula is not None:
                    result.append(str(formula))
                else:
                    cell_value = sheet.range(range_address).value
                    if isinstance(cell_value, list):
                        result.append(
                            str(cell_value[i]) if cell_value[i] is not None else ""
                        )
                    else:
                        result.append(str(cell_value) if cell_value is not None else "")
            return [result]

        # Handle multi-row case
        return [
            [str(formula) if formula is not None else "" for formula in row]
            for row in formulas
        ]
    else:
        values = sheet.range(range_address).value

        # Handle single cell case
        if not isinstance(values, list):
            return [[str(values) if values is not None else ""]]

        # Handle single row case
        if not isinstance(values[0], list):
            return [[str(v) if v is not None else "" for v in values]]

        # Handle multi-row case
        return [
            [str(cell) if cell is not None else "" for cell in row] for row in values
        ]


@mcp.tool()
def read_expanded_range(
    sheet_name: str, start_cell: str, get_formulas: bool = False
) -> list[list]:
    """Read a dynamic range starting from a cell, expanding to find the full data region.

    Args:
        sheet_name: Name of the sheet
        start_cell: Starting cell address like 'A1', 'B5', etc.
        get_formulas: If True, return formulas; if False, return calculated values
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]

    if get_formulas:
        formulas = sheet.range(start_cell).expand().formula

        # Handle single cell case
        if not isinstance(formulas, list):
            formula_value = (
                formulas
                if formulas is not None
                else sheet.range(start_cell).expand().value
            )
            return [[str(formula_value) if formula_value is not None else ""]]

        # Handle single row case
        if not isinstance(formulas[0], list):
            return [[str(f) if f is not None else "" for f in formulas]]

        # Handle multi-row case
        return [
            [str(formula) if formula is not None else "" for formula in row]
            for row in formulas
        ]
    else:
        values = sheet.range(start_cell).expand().value

        # Handle single cell case
        if not isinstance(values, list):
            return [[str(values) if values is not None else ""]]

        # Handle single row case
        if not isinstance(values[0], list):
            return [[str(v) if v is not None else "" for v in values]]

        # Handle multi-row case
        return [
            [str(cell) if cell is not None else "" for cell in row] for row in values
        ]


@mcp.tool()
def read_expanded_range_table(
    sheet_name: str,
    start_cell: str,
    headers: bool = True,
    show_row_numbers: bool = True,
    show_col_addresses: bool = True,
    tablefmt: str = "plain",
) -> str:
    """Read a dynamic range starting from a cell and format as a table.

    Args:
        sheet_name: Name of the sheet
        start_cell: Starting cell address like 'A1', 'B5', etc.
        headers: Whether first row contains headers
        show_row_numbers: Add row numbers as first column
        show_col_addresses: Use Excel column addresses as headers
        tablefmt: Table format (plain, simple, grid, pipe, etc.)
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]
    values = sheet.range(start_cell).expand().value

    # Parse start cell to get position
    col_letters = "".join(c for c in start_cell if c.isalpha())
    row_num = int("".join(c for c in start_cell if c.isdigit()))

    # Handle single cell case
    if not isinstance(values, list):
        return str(values) if values is not None else ""

    # Handle single row case
    if not isinstance(values[0], list):
        values = [values]

    # Convert None values to empty strings
    clean_values = [
        [str(cell) if cell is not None else "" for cell in row] for row in values
    ]

    # Add row numbers if requested
    if show_row_numbers:
        for i, row in enumerate(clean_values):
            row.insert(0, str(row_num + i))

    # Generate column headers if requested
    if show_col_addresses:
        start_col = ord(col_letters[0]) - ord("A")
        col_headers = []
        if show_row_numbers:
            col_headers.append("Row")
        for i in range(len(clean_values[0]) - (1 if show_row_numbers else 0)):
            col_headers.append(chr(ord("A") + start_col + i))

        if headers and clean_values:
            return tabulate(clean_values[1:], headers=col_headers, tablefmt=tablefmt)
        else:
            return tabulate(clean_values, headers=col_headers, tablefmt=tablefmt)

    elif headers and clean_values:
        if show_row_numbers:
            row_headers = ["Row"] + clean_values[0][1:]
            return tabulate(clean_values[1:], headers=row_headers, tablefmt=tablefmt)
        else:
            return tabulate(
                clean_values[1:], headers=clean_values[0], tablefmt=tablefmt
            )
    else:
        return tabulate(clean_values, tablefmt=tablefmt)


@mcp.tool()
def write_cell(sheet_name: str, cell_address: str, value: str) -> str:
    """Write a value to a single cell in Excel.

    Args:
        sheet_name: Name of the sheet
        cell_address: Cell address like 'A1', 'B5', etc.
        value: Value to write to the cell
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]
    sheet.range(cell_address).value = value
    return f"Written '{value}' to {sheet_name}!{cell_address}"


@mcp.tool()
def read_range_table(
    sheet_name: str,
    range_address: str,
    headers: bool = True,
    show_row_numbers: bool = True,
    show_col_addresses: bool = True,
    tablefmt: str = "plain",
) -> str:
    """Read a range of cells from Excel and format as a table.

    Args:
        sheet_name: Name of the sheet
        range_address: Range address like 'A1:C10', 'B2:D20', etc.
        headers: Whether first row contains headers
        show_row_numbers: Add row numbers as first column
        show_col_addresses: Use Excel column addresses (A, B, C) as headers
        tablefmt: Table format (simple, plain, grid, pipe, etc.)
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]
    values = sheet.range(range_address).value

    # Parse range to get start position
    start_cell = range_address.split(":")[0]
    col_letters = "".join(c for c in start_cell if c.isalpha())
    row_num = int("".join(c for c in start_cell if c.isdigit()))

    # Handle single cell case
    if not isinstance(values, list):
        return str(values) if values is not None else ""

    # Handle single row case
    if not isinstance(values[0], list):
        values = [values]

    # Convert None values to empty strings
    clean_values = [
        [str(cell) if cell is not None else "" for cell in row] for row in values
    ]

    # Add row numbers if requested
    if show_row_numbers:
        for i, row in enumerate(clean_values):
            row.insert(0, str(row_num + i))

    # Generate column headers if requested
    if show_col_addresses:
        start_col = ord(col_letters[0]) - ord("A")
        col_headers = []
        if show_row_numbers:
            col_headers.append("Row")
        for i in range(len(clean_values[0]) - (1 if show_row_numbers else 0)):
            col_headers.append(chr(ord("A") + start_col + i))

        if headers and clean_values:
            return tabulate(clean_values[1:], headers=col_headers, tablefmt=tablefmt)
        else:
            return tabulate(clean_values, headers=col_headers, tablefmt=tablefmt)

    elif headers and clean_values:
        if show_row_numbers:
            # Adjust headers to include row number column
            row_headers = ["Row"] + clean_values[0][1:]
            return tabulate(clean_values[1:], headers=row_headers, tablefmt=tablefmt)
        else:
            return tabulate(
                clean_values[1:], headers=clean_values[0], tablefmt=tablefmt
            )
    else:
        return tabulate(clean_values, tablefmt=tablefmt)


@mcp.tool()
def write_range(sheet_name: str, start_cell: str, values: list[list[str]]) -> str:
    """Write values to a range of cells in Excel.

    Args:
        sheet_name: Name of the sheet
        start_cell: Starting cell address like 'A1', 'B5', etc.
        values: 2D list of values to write
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]
    sheet.range(start_cell).value = values
    rows = len(values)
    cols = len(values[0]) if values else 0
    return f"Written {rows}x{cols} range starting at {sheet_name}!{start_cell}"


@mcp.tool()
def open_excel_file(file_path: str, create_if_not_exists: bool = True) -> str:
    """Open an Excel file in a new workbook.

    Args:
        file_path: Path to the Excel file to open
        create_if_not_exists: If True, create the file if it doesn't exist
    """
    try:
        path = Path(file_path)

        if not path.exists() and create_if_not_exists:
            # Create a new workbook and save it
            # Ensure the directory exists
            path.parent.mkdir(parents=True, exist_ok=True)
            wb = xw.Book()
            wb.save(path)
            return f"Created and opened new workbook: {wb.name} with sheets: {[sheet.name for sheet in wb.sheets]}"
        elif not path.exists() and not create_if_not_exists:
            return f"Error: File {file_path} does not exist and create_if_not_exists is False"
        else:
            wb = xw.Book(file_path)
            return f"Opened workbook: {wb.name} with sheets: {[sheet.name for sheet in wb.sheets]}"
    except Exception as e:
        return f"Error opening file: {e}"


@mcp.tool()
def close_active_workbook() -> str:
    """Close the currently active Excel workbook."""
    try:
        app = xw.apps.active
        wb = app.books.active
        wb_name = wb.name
        wb.close()
        return f"Closed workbook: {wb_name}"
    except Exception as e:
        return f"Error closing workbook: {e}"


@mcp.tool()
def list_open_workbooks() -> list[str]:
    """List all currently open Excel workbooks."""
    try:
        app = xw.apps.active
        return [wb.name for wb in app.books]
    except Exception as e:
        return [f"Error listing workbooks: {e}"]


@mcp.tool()
def save_active_workbook() -> str:
    """Save the currently active Excel workbook."""
    try:
        app = xw.apps.active
        wb = app.books.active
        wb.save()
        return f"Saved workbook: {wb.name}"
    except Exception as e:
        return f"Error saving workbook: {e}"


@mcp.tool()
def find_excel_files_in_downloads() -> list[str]:
    """Find all Excel files in the Downloads folder, sorted by modification time."""
    downloads_path = Path.home() / "Downloads"
    excel_files = list(downloads_path.glob("**/*.xlsx"))

    if not excel_files:
        return ["No Excel files found in Downloads folder"]

    # Sort by modification time, most recent first
    sorted_files = sorted(excel_files, key=lambda x: x.stat().st_mtime, reverse=True)
    return [str(file_path) for file_path in sorted_files[:20]]  # Return top 20


@mcp.tool()
def open_recent_excel_file() -> str:
    """Open the most recently modified Excel file from Downloads folder."""
    downloads_path = Path.home() / "Downloads"
    excel_files = list(downloads_path.glob("**/*.xlsx"))

    if not excel_files:
        return "No Excel files found in Downloads folder"

    # Sort by modification time, most recent first
    most_recent = max(excel_files, key=lambda x: x.stat().st_mtime)

    try:
        wb = xw.Book(str(most_recent))
        return f"Opened most recent file: {wb.name} with sheets: {[sheet.name for sheet in wb.sheets]}"
    except Exception as e:
        return f"Error opening file: {e}"


if __name__ == "__main__":
    mcp.run()
