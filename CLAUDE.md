Meta learn: keep CLAUDE.md up-to-date; When we discover a new way of working, a workflow, a tool to use, learning about the domain, the project, the environment, etc. please capture it in CLAUDE.md ASAP, commit and push it.
important! lint, commit, push very often
prefer `uv run` over running python directly
`./lint.sh` runs the linter, do it often please

# Notes

- list_open_workbooks returns only filenames - consider if we need absolute paths for better workbook identification

# VS Code Python Environment Setup

To sync VS Code with uv dependencies:

```bash
uv venv .venv
. .venv/bin/activate
uv pip install -r <(uv pip compile pyproject.toml)
```

Then select the `.venv/bin/python` interpreter in VS Code to resolve import warnings.
