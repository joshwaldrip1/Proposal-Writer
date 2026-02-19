# Proposal Writer

## Project Overview

A Python desktop GUI application for writing and managing business proposals (sales proposals, consulting proposals, SOWs, etc.). It reads client/scope data from Excel exports and generates Word documents from a `.dotx` template.

## Tech Stack

- **Language:** Python 3.13
- **GUI framework:** tkinter
- **Packaging:** PyInstaller (builds to single `.exe`)
- **Platform:** Windows 11
- **Key libraries:** openpyxl (Excel), python-docx (Word), PIL/Pillow (images), lxml, numpy

## Project Structure

```
proposal_gui.py                           # Main application (single-file GUI app)
proposal_gui.spec                         # PyInstaller build spec
PROPOSAL FOR PROFESSIONAL SERVICES.dotx   # Word template for proposals
Civil Engineering Services.xlsx           # Services reference data
Monday Export.xlsx                        # Monday.com project data export
Proposal Creator Key.xlsx                 # Key/mapping data for proposal fields
export_cache/                             # Cached exports (data.xlsx, scope_data.xlsx, etc.)
```

## Development Guidelines

### Code Style
- Follow PEP 8
- Use type hints for all function signatures
- Use `pathlib.Path` for file paths, not `os.path`
- Prefer dataclasses or Pydantic models for structured data

### Running the App
- Direct: `python proposal_gui.py`
- Built exe: `dist/proposal_gui.exe`

### Building
- `pyinstaller proposal_gui.spec`

### Testing
- Use `pytest` for testing
- Run tests: `python -m pytest tests/`
- Run single test: `python -m pytest tests/test_foo.py::test_bar -v`

### Git Workflow
- Main branch: `main`
- Create feature branches for new work
- Keep commits focused and descriptive

## Conventions

- Business logic should be separated from UI code
- All file I/O goes through utility functions, not scattered inline
- Use logging (`logging` module) instead of print statements
