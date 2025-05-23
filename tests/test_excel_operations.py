import re


class TestExcelOperations:
    """Test Excel MCP operations using Claude CLI"""

    def test_write_and_read_cell(self, excel_tester, temp_workbook):
        """Test writing and reading a cell value"""
        workbook_path = temp_workbook("simple.xlsx")

        result = excel_tester.run_claude_test(
            "write 'Hello World' to cell B2, then read it back and confirm",
            workbook_path,
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert "Hello World" in result.claude_output

        # Validate workbook
        assert excel_tester.validate_workbook_cell(
            workbook_path, "Sheet1", "B2", "Hello World"
        )

    def test_add_sum_formula(self, excel_tester, temp_workbook):
        """Test adding a SUM formula"""
        workbook_path = temp_workbook("with_numbers.xlsx")

        result = excel_tester.run_claude_test(
            "add a SUM formula to cell C1 that sums range A1:A5, return the formula",
            workbook_path,
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert re.search(r"SUM\(A1:A5\)", result.claude_output, re.IGNORECASE)

        # Validate formula in workbook
        assert excel_tester.validate_workbook_formula(
            workbook_path, "Sheet", "C1", "SUM(A1:A5)"
        )

    def test_read_range_as_table(self, excel_tester, temp_workbook):
        """Test reading range and formatting as table"""
        workbook_path = temp_workbook("sample_data.xlsx")

        result = excel_tester.run_claude_test(
            "read range A1:C3 and format as a table", workbook_path
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert re.search(r"Name.*Age.*City", result.claude_output, re.IGNORECASE)

    def test_list_sheet_names(self, excel_tester, temp_workbook):
        """Test listing sheet names"""
        workbook_path = temp_workbook("multi_sheet.xlsx")

        result = excel_tester.run_claude_test(
            "list all sheet names in this workbook", workbook_path
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert re.search(
            r"Sheet1.*Sheet2.*Summary", result.claude_output, re.IGNORECASE
        )

    def test_write_range_of_data(self, excel_tester, temp_workbook):
        """Test writing a range of data"""
        workbook_path = temp_workbook("empty.xlsx")

        result = excel_tester.run_claude_test(
            "write a 2x2 grid with values 1,2,3,4 starting at cell A1", workbook_path
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert re.search(r"written.*A1", result.claude_output, re.IGNORECASE)

        # Validate all cells
        assert excel_tester.validate_workbook_cell(workbook_path, "Sheet", "A1", 1)
        assert excel_tester.validate_workbook_cell(workbook_path, "Sheet", "B1", 2)
        assert excel_tester.validate_workbook_cell(workbook_path, "Sheet", "A2", 3)
        assert excel_tester.validate_workbook_cell(workbook_path, "Sheet", "B2", 4)

    def test_read_expanded_range(self, excel_tester, temp_workbook):
        """Test reading expanded range and counting rows"""
        workbook_path = temp_workbook("data_table.xlsx")

        result = excel_tester.run_claude_test(
            "read expanded range starting from A1 and count total rows", workbook_path
        )

        assert result.success, f"Claude test failed: {result.error_message}"
        assert re.search(r"\d+.*rows", result.claude_output, re.IGNORECASE)

    def test_error_handling_invalid_cell(self, excel_tester, temp_workbook):
        """Test error handling for invalid cell"""
        workbook_path = temp_workbook("simple.xlsx")

        result = excel_tester.run_claude_test(
            "try to read cell ZZ999999", workbook_path
        )

        # Should either succeed with error message or fail gracefully
        if result.success:
            assert re.search(
                r"error|invalid|failed", result.claude_output, re.IGNORECASE
            )
        # If it fails, that's also acceptable error handling


class TestExcelValidation:
    """Test the validation functions themselves"""

    def test_validate_cell_value(self, excel_tester, temp_workbook):
        """Test cell validation works correctly"""
        workbook_path = temp_workbook("simple.xlsx")

        # Should validate existing cell
        assert excel_tester.validate_workbook_cell(
            workbook_path, "Sheet1", "A1", "Test"
        )

        # Should fail for wrong value
        assert not excel_tester.validate_workbook_cell(
            workbook_path, "Sheet1", "A1", "Wrong"
        )

    def test_validate_formula(self, excel_tester, temp_workbook):
        """Test formula validation works correctly"""
        workbook_path = temp_workbook("with_numbers.xlsx")

        # Add a formula first via MCP
        excel_tester.run_claude_test(
            "add a SUM formula to cell C1 that sums range A1:A5", workbook_path
        )

        # Then validate it
        assert excel_tester.validate_workbook_formula(
            workbook_path, "Sheet", "C1", "SUM(A1:A5)"
        )
