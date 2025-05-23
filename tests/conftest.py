import pytest
import tempfile
import shutil
from pathlib import Path
from integration.test_runner import ExcelMCPTester


@pytest.fixture
def fixtures_dir():
    """Provide path to test fixtures directory"""
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def excel_tester(fixtures_dir):
    """Provide configured ExcelMCPTester instance"""
    return ExcelMCPTester(fixtures_dir)


@pytest.fixture
def temp_workbook(fixtures_dir):
    """Create temporary copy of fixture for testing"""
    def _create_temp(fixture_name: str) -> Path:
        temp_dir = Path(tempfile.mkdtemp())
        fixture_path = fixtures_dir / fixture_name
        temp_workbook = temp_dir / fixture_name
        shutil.copy(fixture_path, temp_workbook)
        return temp_workbook
    
    temp_dirs = []
    
    def cleanup():
        for temp_dir in temp_dirs:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
    
    yield _create_temp
    cleanup()