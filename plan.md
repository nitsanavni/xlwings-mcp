# Excel MCP Testing Plan

## Overview
Test the Excel MCP server using Claude CLI with `-p` flag for automated, non-interactive testing.

## Test Architecture

### Directory Structure
```
tests/
├── integration/
│   ├── test_excel_operations.py
│   ├── test_workbook_management.py
│   └── test_runner.py
├── fixtures/
│   ├── sample.xlsx
│   ├── multi_sheet.xlsx
│   └── formula_test.xlsx
└── expected/
    └── (expected result snapshots)
```

## Test Methodology

### 1. Claude CLI Integration Tests
- Use `claude -p` with `--disallowedTools Bash` to prevent shell access
- Execute predefined prompts that perform Excel operations
- Validate both:
  - Claude's response text matches expected patterns
  - Resulting workbook changes are correct

### 2. Test Flow
1. **Setup**: Copy fixture workbook to temp location
2. **Execute**: Run `claude -p "operation prompt" --disallowedTools Bash`
3. **Validate Response**: Check Claude's output against expected pattern
4. **Validate Workbook**: Inspect Excel file for expected changes
5. **Reset**: Restore original fixture for next test

### 3. Test Scenarios

#### Basic Operations
- **Write Cell**: `"Write 'Test Value' to cell B2, confirm the write"`
- **Read Cell**: `"Read cell A1 and return its value"`
- **Formula**: `"Add SUM formula to C1 for range A1:A10, return formula text"`

#### Range Operations
- **Read Range**: `"Read range A1:C3 and format as table"`
- **Write Range**: `"Write 2x3 grid of sequential numbers starting at D1"`
- **Expanded Range**: `"Read expanded range from A1, count total rows"`

#### Sheet Management
- **List Sheets**: `"List all sheet names in workbook"`
- **Multi-sheet**: `"Copy A1:B2 from Sheet1 to Sheet2 at C3"`

#### Error Handling
- **Invalid Cell**: `"Try to read cell ZZ999999"`
- **Invalid Sheet**: `"Read from non-existent sheet 'Missing'"`

### 4. Implementation Details

#### Test Runner Function
```python
def run_claude_test(prompt: str, fixture_path: str, expected_pattern: str) -> TestResult:
    # Copy fixture to temp location
    # Run: claude -p "prompt" --disallowedTools Bash
    # Validate response and workbook state
    # Clean up temp files
```

#### Workbook Validation
- Use openpyxl or similar to inspect Excel files
- Compare cell values, formulas, sheet names
- Snapshot-based testing for complex operations

#### Reset Strategy
- Keep original fixtures read-only
- Work on temporary copies
- Automatic cleanup between tests

## Benefits
- End-to-end validation of MCP server
- Realistic usage scenarios
- Prevents regression in Excel operations
- Validates both MCP responses and actual Excel changes