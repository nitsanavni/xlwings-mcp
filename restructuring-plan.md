# PyPI Publishing Restructuring Plan

## Current State
- Single `main.py` file with MCP server implementation
- Basic `pyproject.toml` with minimal metadata
- No proper package structure

## Required Changes for PyPI Publishing

### 1. Project Structure (src layout)
```
excel-mcp/
├── pyproject.toml
├── uv.lock
├── CLAUDE.md
├── lint.sh
├── plan.md
├── README.md (create)
└── src/
    └── xlwings_mcp/
        ├── __init__.py
        └── server.py (renamed from main.py)
```

### 2. pyproject.toml Updates
- Add proper package metadata (description, author, license)
- Configure build system (hatchling)
- Add entry points for CLI usage
- Add classifiers for PyPI

### 3. Package Entry Point
- Create console script entry point for `xlwings-mcp` command
- Move main execution logic to a proper entry point function

### 4. Build and Test
- Use `uv build` to create sdist and wheel
- Test package installation locally
- Verify entry points work correctly

## Implementation Steps
1. Create src/xlwings_mcp/ directory structure
2. Move and refactor main.py → src/xlwings_mcp/server.py
3. Create __init__.py with proper exports
4. Update pyproject.toml with complete metadata and entry points
5. Test build process with `uv build`
6. Test installation from built wheel

## Benefits
- Proper package isolation with src layout
- Clean PyPI distribution
- CLI tool availability via `xlwings-mcp` command
- Standard Python packaging conventions