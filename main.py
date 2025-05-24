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
def read_table(
    sheet_name: str,
    cell_address: str,
    is_expanded_range: bool = False,
    headers: bool = True,
    tablefmt: str = "plain",
) -> str:
    """Read cells from Excel and format as a table.

    Args:
        sheet_name: Name of the sheet
        cell_address: Cell address or range (e.g., 'A1' for expanded, 'A1:C10' for specific range)
        is_expanded_range: If True, expand from the cell to find full data region
        headers: Whether first row contains headers
        tablefmt: Table format (plain, simple, grid, pipe, etc.)
    """
    app = xw.apps.active
    wb = app.books.active
    sheet = wb.sheets[sheet_name]

    # Get values based on whether it's expanded or specific range
    if is_expanded_range:
        values = sheet.range(cell_address).expand().value
        # Parse start cell to get position
        col_letters = "".join(c for c in cell_address if c.isalpha())
        row_num = int("".join(c for c in cell_address if c.isdigit()))
    else:
        values = sheet.range(cell_address).value
        # Parse range to get start position
        start_cell = cell_address.split(":")[0]
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

    # Always add row numbers
    for i, row in enumerate(clean_values):
        row.insert(0, str(row_num + i))

    # Always generate column headers
    start_col = ord(col_letters[0]) - ord("A")
    col_headers = ["Row"]
    for i in range(len(clean_values[0]) - 1):
        col_headers.append(chr(ord("A") + start_col + i))

    if headers and clean_values:
        return tabulate(clean_values[1:], headers=col_headers, tablefmt=tablefmt)
    else:
        return tabulate(clean_values, headers=col_headers, tablefmt=tablefmt)


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
def save_workbook(filename: str | None = None) -> str:
    """Save a workbook.

    Args:
        filename: Name of the workbook to save. If None, saves the active workbook.
    """
    try:
        app = xw.apps.active
        if filename is None:
            wb = app.books.active
            wb.save()
            return f"Saved active workbook: {wb.name}"
        else:
            # Find the workbook by name
            for wb in app.books:
                if wb.name == filename:
                    wb.save()
                    return f"Saved workbook: {wb.name}"
            return f"Error: Workbook '{filename}' not found"
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
