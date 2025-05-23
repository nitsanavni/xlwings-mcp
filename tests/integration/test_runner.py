import subprocess
import shutil
import tempfile
import re
from pathlib import Path
from typing import Dict, Any, Optional
from dataclasses import dataclass

try:
    import openpyxl  # type: ignore
except ImportError:
    openpyxl = None


@dataclass
class TestResult:
    success: bool
    claude_output: str
    error_message: Optional[str] = None


class ExcelMCPTester:
    def __init__(self, fixtures_dir: Path):
        self.fixtures_dir = fixtures_dir
        self.temp_dir: Optional[Path] = None

    def setup_test(self, fixture_name: str) -> Path:
        """Copy fixture to temp location for testing"""
        self.temp_dir = Path(tempfile.mkdtemp())
        fixture_path = self.fixtures_dir / fixture_name
        temp_workbook = self.temp_dir / fixture_name
        shutil.copy(fixture_path, temp_workbook)
        return temp_workbook

    def cleanup_test(self):
        """Remove temp files"""
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)

    def run_claude_test(self, prompt: str, workbook_path: Path) -> TestResult:
        """Execute claude -p with Excel MCP operations"""
        try:
            # Change to workbook directory so MCP can access it
            cmd = [
                "claude",
                "-p",
                f"Open {workbook_path.name} and {prompt}",
                "--disallowedTools",
                "Bash",
            ]

            result = subprocess.run(
                cmd,
                cwd=workbook_path.parent,
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode != 0:
                return TestResult(
                    success=False,
                    claude_output=result.stdout,
                    error_message=result.stderr,
                )

            return TestResult(success=True, claude_output=result.stdout)

        except subprocess.TimeoutExpired:
            return TestResult(
                success=False, claude_output="", error_message="Test timed out"
            )
        except Exception as e:
            return TestResult(success=False, claude_output="", error_message=str(e))

    def validate_response(self, result: TestResult, expected_pattern: str) -> bool:
        """Check if Claude's response matches expected pattern"""
        if not result.success:
            return False
        return bool(re.search(expected_pattern, result.claude_output, re.IGNORECASE))

    def validate_workbook_cell(
        self,
        workbook_path: Path,
        sheet_name: str,
        cell_address: str,
        expected_value: Any,
    ) -> bool:
        """Validate specific cell value in workbook"""
        try:
            wb = openpyxl.load_workbook(workbook_path)
            sheet = wb[sheet_name]
            actual_value = sheet[cell_address].value
            return actual_value == expected_value
        except Exception:
            return False

    def validate_workbook_formula(
        self,
        workbook_path: Path,
        sheet_name: str,
        cell_address: str,
        expected_formula: str,
    ) -> bool:
        """Validate formula in specific cell"""
        try:
            wb = openpyxl.load_workbook(workbook_path)
            sheet = wb[sheet_name]
            cell = sheet[cell_address]
            # Remove leading = if present in expected
            expected = expected_formula.lstrip("=")
            actual = str(cell.value) if cell.value else ""
            if hasattr(cell, "_value") and str(cell._value).startswith("="):
                actual = str(cell._value)[1:]  # Remove =
            return expected.lower() in actual.lower()
        except Exception:
            return False

    def run_test_scenario(self, scenario: Dict[str, Any]) -> bool:
        """Run complete test scenario with setup, execution, and validation"""
        fixture_name = scenario["fixture"]
        prompt = scenario["prompt"]
        expected_response = scenario.get("expected_response", "")
        workbook_validations = scenario.get("workbook_validations", [])

        try:
            # Setup
            workbook_path = self.setup_test(fixture_name)

            # Execute
            result = self.run_claude_test(prompt, workbook_path)

            # Validate response
            if expected_response and not self.validate_response(
                result, expected_response
            ):
                print(f"❌ Response validation failed for: {scenario['name']}")
                print(f"Expected pattern: {expected_response}")
                print(f"Actual output: {result.claude_output}")
                return False

            # Validate workbook changes
            for validation in workbook_validations:
                if validation["type"] == "cell_value":
                    if not self.validate_workbook_cell(
                        workbook_path,
                        validation["sheet"],
                        validation["cell"],
                        validation["expected"],
                    ):
                        print(f"❌ Workbook validation failed: {validation}")
                        return False

                elif validation["type"] == "formula":
                    if not self.validate_workbook_formula(
                        workbook_path,
                        validation["sheet"],
                        validation["cell"],
                        validation["expected"],
                    ):
                        print(f"❌ Formula validation failed: {validation}")
                        return False

            print(f"✅ Test passed: {scenario['name']}")
            return True

        except Exception as e:
            print(f"❌ Test error: {scenario['name']} - {e}")
            return False
        finally:
            self.cleanup_test()
