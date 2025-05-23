from pathlib import Path
from test_runner import ExcelMCPTester


def get_test_scenarios():
    """Define all test scenarios"""
    return [
        {
            "name": "Write and read cell value",
            "fixture": "simple.xlsx",
            "prompt": "write 'Hello World' to cell B2, then read it back and confirm",
            "expected_response": "Hello World",
            "workbook_validations": [
                {
                    "type": "cell_value",
                    "sheet": "Sheet1", 
                    "cell": "B2",
                    "expected": "Hello World"
                }
            ]
        },
        {
            "name": "Add SUM formula",
            "fixture": "with_numbers.xlsx", 
            "prompt": "add a SUM formula to cell C1 that sums range A1:A5, return the formula",
            "expected_response": "SUM\\(A1:A5\\)",
            "workbook_validations": [
                {
                    "type": "formula",
                    "sheet": "Sheet1",
                    "cell": "C1", 
                    "expected": "SUM(A1:A5)"
                }
            ]
        },
        {
            "name": "Read range as table",
            "fixture": "sample_data.xlsx",
            "prompt": "read range A1:C3 and format as a table",
            "expected_response": "Name.*Age.*City",
            "workbook_validations": []
        },
        {
            "name": "List sheet names",
            "fixture": "multi_sheet.xlsx",
            "prompt": "list all sheet names in this workbook",
            "expected_response": "Sheet1.*Sheet2.*Summary",
            "workbook_validations": []
        },
        {
            "name": "Write range of data",
            "fixture": "empty.xlsx",
            "prompt": "write a 2x2 grid with values 1,2,3,4 starting at cell A1",
            "expected_response": "written.*A1",
            "workbook_validations": [
                {"type": "cell_value", "sheet": "Sheet1", "cell": "A1", "expected": 1},
                {"type": "cell_value", "sheet": "Sheet1", "cell": "B1", "expected": 2},
                {"type": "cell_value", "sheet": "Sheet1", "cell": "A2", "expected": 3},
                {"type": "cell_value", "sheet": "Sheet1", "cell": "B2", "expected": 4}
            ]
        },
        {
            "name": "Read expanded range",
            "fixture": "data_table.xlsx",
            "prompt": "read expanded range starting from A1 and count total rows",
            "expected_response": "\\d+.*rows",
            "workbook_validations": []
        },
        {
            "name": "Error handling - invalid cell",
            "fixture": "simple.xlsx", 
            "prompt": "try to read cell ZZ999999",
            "expected_response": "error|invalid|failed",
            "workbook_validations": []
        }
    ]


def run_all_tests():
    """Run all test scenarios"""
    fixtures_dir = Path(__file__).parent.parent / "fixtures"
    tester = ExcelMCPTester(fixtures_dir)
    
    scenarios = get_test_scenarios()
    passed = 0
    total = len(scenarios)
    
    print(f"Running {total} Excel MCP tests...\n")
    
    for scenario in scenarios:
        if tester.run_test_scenario(scenario):
            passed += 1
        print()  # Empty line between tests
    
    print(f"Results: {passed}/{total} tests passed")
    return passed == total


if __name__ == "__main__":
    success = run_all_tests()
    exit(0 if success else 1)