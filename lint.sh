#!/bin/bash

# Format code with black
uv run black .

# Fix linting issues with ruff
uv run ruff check --fix .

# Type check with mypy
uv run mypy .

# Format with prettier
bunx prettier --write .