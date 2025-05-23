import xlwings as xw  # type: ignore


def demo_read_write():
    """Demonstrate reading from and writing to Excel cells."""
    # Connect to the active Excel application
    app = xw.apps.active
    wb = app.books.active

    # Get the first sheet
    sheet = wb.sheets[0]
    print(f"Working with sheet: {sheet.name}")

    # Read a single cell (A1)
    cell_a1 = sheet.range("A1").value
    print(f"Current value in A1: {cell_a1}")

    # Read a range of cells (A1:C3)
    range_values = sheet.range("A1:C3").value
    print("Values in A1:C3:")
    for i, row in enumerate(range_values or []):
        print(f"  Row {i+1}: {row}")

    # Write to a cell
    test_cell = "Z1"
    original_value = sheet.range(test_cell).value
    print(f"\nOriginal value in {test_cell}: {original_value}")

    # Write new value
    new_value = "xlwings test"
    sheet.range(test_cell).value = new_value
    print(f"Set {test_cell} to: {new_value}")

    # Read it back to confirm
    updated_value = sheet.range(test_cell).value
    print(f"Confirmed value in {test_cell}: {updated_value}")

    # Write multiple values at once
    sheet.range("Z2:Z4").value = [["Test 1"], ["Test 2"], ["Test 3"]]
    print("Set Z2:Z4 to test values")

    # Read back the range
    test_range = sheet.range("Z2:Z4").value
    print(f"Values in Z2:Z4: {test_range}")

    # Restore original value
    sheet.range(test_cell).value = original_value
    print(f"\nRestored {test_cell} to original value: {original_value}")


if __name__ == "__main__":
    try:
        demo_read_write()
    except Exception as e:
        print(f"Error: {e}")
        print("Make sure Excel is open with a workbook.")
