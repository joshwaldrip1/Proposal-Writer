# Proposal Writer

## Project Overview

A Python desktop application for writing and managing business proposals (sales proposals, consulting proposals, SOWs, etc.).

## Tech Stack

- **Language:** Python 3.12+
- **Application type:** Desktop GUI
- **Platform:** Windows 11

## Project Structure

```
proposal_writer/        # Main application package
  __init__.py
  main.py              # Application entry point
  ui/                  # GUI components
  models/              # Data models
  services/            # Business logic
  templates/           # Proposal templates
  utils/               # Shared utilities
tests/                 # Test suite
  conftest.py
  test_*.py
requirements.txt       # Dependencies
CLAUDE.md              # This file
```

## Development Guidelines

### Code Style
- Follow PEP 8
- Use type hints for all function signatures
- Use `pathlib.Path` for file paths, not `os.path`
- Prefer dataclasses or Pydantic models for structured data

### Testing
- Use `pytest` for testing
- Run tests: `python -m pytest tests/`
- Run single test: `python -m pytest tests/test_foo.py::test_bar -v`

### Running the App
- Entry point: `python -m proposal_writer.main`

### Git Workflow
- Main branch: `main`
- Create feature branches for new work
- Keep commits focused and descriptive

## Conventions

- Keep modules small and focused (< 300 lines)
- Business logic lives in `services/`, not in UI code
- UI code should only handle presentation and user interaction
- All file I/O goes through utility functions, not scattered inline
- Use logging (`logging` module) instead of print statements
