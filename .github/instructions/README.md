# Project Instructions

## Project Overview

**Project Name**: ezxl
**Tech Stack**: Python 3.11+, win32com.client, openpyxl, python-calamine, ruff, pyright, ty, pytest
**Environment**: Corporate Windows, network proxy, no admin rights, local wheel distribution

## Project Description

EzXl is a generic Python library for Excel automation via COM (Component Object Model). It serves as a **foundation layer** for specialized consumer libraries (e.g., a future SAP-specific lib). EzXl contains no add-in-specific business logic whatsoever.

Scope:

- Opening and closing Excel files
- Attaching (hooking) to an already-running Excel instance
- Converting between formats (.xlsx to .csv, etc.)
- Simplified navigation through Excel menus and ribbon
- Generic cell, range, and sheet manipulation

EzXl does NOT contain SAP logic. Any SAP-specific behavior belongs in a separate consumer library.

## Architecture

```text
G:\.root\dev\.python\.libraries\ezxl\
├── src/
│   └── ezxl/                # Source package (src-layout)
│       ├── __init__.py      # Public API exports, Python version guard
│       ├── version.py       # Single source of truth for __version__
│       ├── exceptions.py    # EzXlError hierarchy (create this first)
│       ├── _excel_app.py    # ExcelApp: COM dispatch/attach, lifecycle
│       ├── _workbook.py     # WorkbookProxy
│       ├── _sheet.py        # SheetProxy
│       ├── _com_utils.py    # Internal COM utilities
│       └── _converters.py   # Format conversion logic
├── tests/                   # Test files
├── docs/                    # MkDocs source
├── .github/
│   ├── instructions/        # AI agent instructions
│   └── workflows/           # CI/CD workflows
├── .hooks/                  # Git hooks
├── .scripts/                # Dev and build scripts
├── pyproject.toml           # Project configuration
└── README.md                # Project documentation
```

Modules prefixed with `_` are internal and must not be exposed in the public API. Public symbols are declared in `__init__.py` via `__all__`.

## Key Conventions

- **src-layout**: all source code lives under `src/ezxl/`. No flat-layout exceptions.
- **Type hints**: mandatory on all public API functions, methods, and class signatures.
- **Docstrings**: Google style on all public symbols.
- **Logging**: use the standard `logging` module (not loguru). Named logger: `logging.getLogger("ezxl")`. No `print()` for diagnostic output.
- **COM threading**: `ExcelApp` is NOT thread-safe (COM STA model). This must be documented explicitly on the class and enforced with a thread identity assertion at construction time. Do not abstract this away.
- **Context manager**: `ExcelApp` implements `__enter__` / `__exit__` for explicit lifecycle management.
- **Exception wrapping**: all `pywintypes.com_error` exceptions must be caught at the COM boundary and re-raised as `EzXlError` subclasses. Never let raw COM errors propagate to callers.
- **`exceptions.py` first**: this module must exist before any other module is written or modified.
- **No menu hierarchy**: do not implement `xl.menu.fichier.ouvrir()` style trees. Expose flat methods directly on `ExcelApp` and proxy classes.
- **pyproject.toml version as source of truth**: `.scripts/dev/update_version.py` syncs `version.py` and the README badge. Do not edit `version.py` manually.

## Environment Setup

```bash
# Create virtual environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\activate

# Install dependencies (wheels only — no PyPI access)
pip install -e ".[dev]"

# Post-install: pywin32 requires a post-install script
python .venv/Scripts/pywin32_postinstall.py -install

# Set up git hooks
git config core.hooksPath .hooks
```

**Bitness**: ensure the Python interpreter bitness (32 or 64-bit) matches the installed Excel. COM dispatch will fail silently or raise on mismatch. Verify with `python -c "import struct; print(struct.calcsize('P') * 8)"`.

## Dependency Roles

| Package           | Role                                                              |
| ----------------- | ----------------------------------------------------------------- |
| `pywin32`         | Primary COM driver (`win32com.client`). Windows-only.             |
| `openpyxl`        | Post-processing Excel files when COM is not active (file closed). |
| `python-calamine` | Fast read-only data extraction via Rust binding.                  |
| `pywinauto`       | UI-level fallback for generic Windows automation (limited use).   |
| `uiautomation`    | Secondary UI fallback (limited use).                              |
| `rich`            | Terminal output in dev scripts only. Not used in library code.    |

## Testing Conventions

- Unit tests must not require Excel to be installed. Use `Protocol`-based fakes to stub COM objects.
- Integration tests that require a live Excel instance are marked `@pytest.mark.excel`.
- `@pytest.mark.excel` tests are excluded from CI/CD and run locally only.
- Test naming: `test_should_<expected_behavior>_when_<condition>`.

## Instruction Files

| File                                                            | Purpose                                                          |
| --------------------------------------------------------------- | ---------------------------------------------------------------- |
| `core/advanced-cognitive-conduct.instructions.md`               | Core reasoning framework                                         |
| `core/commit-standards.instructions.md`                         | Git commit conventions                                           |
| `core/hexagonal-architecture-standards.instructions.md`         | Hexagonal architecture reference (NOT used here — see overrides) |
| `languages/python/python-development-standards.instructions.md` | Python coding standards                                          |
| `languages/python/python-formatting-standards.instructions.md`  | Code formatting rules                                            |
| `languages/python/pyproject-standards.instructions.md`          | pyproject.toml conventions                                       |

## Project-Specific Overrides

- **Flat architecture, not hexagonal**: the Excel proxy hierarchy (`ExcelApp` > `WorkbookProxy` > `SheetProxy`) does not meet the criteria for Ports & Adapters. `core/hexagonal-architecture-standards.instructions.md` is available for reference but does not apply to this project.
- **No JavaScript**: ignore all `languages/javascript/` instruction files. This project is Python-only.
- **No loguru**: generic Python standards may reference loguru. This project uses the standard `logging` module exclusively to avoid imposing a runtime dependency on consumers.
- **COM tests excluded from CI**: `@pytest.mark.excel` tests require a local Excel installation and are never run in CI/CD pipelines.

## Hexagonal Architecture — Decision Record

**Decision**: Do not adopt hexagonal architecture at the project level.
**Date**: 2026-03-20
**Status**: Accepted

### Rationale

#### Current state

The GUI layer already implements a Ports & Adapters pattern at its own scope:

- `_protocols.py` defines four abstract protocols (`AbstractRibbonBackend`,
  `AbstractMenuBackend`, `AbstractDialogBackend`, `AbstractKeysBackend`) — these
  are the **ports**.
- The COM backends (`RibbonProxy`, `MenuProxy`, `DialogProxy`, `_COMKeysBackend`)
  and the pywinauto backends (`PywinautoRibbonBackend`, etc.) are the **adapters**.
- `GUIProxy` performs dependency injection, accepting any conforming backend at
  construction time.

This is a clean, correct application of Ports & Adapters — scoped to the layer
where it adds real value.

#### What full hexagonal architecture would add

Formal hexagonal at the project level would require:

- Renaming `core/` to `domain/` and enforcing strict "no infrastructure imports"
  rules inside it.
- Introducing adapter packages per technology boundary (COM, pywinauto, file I/O).
- Defining application-level ports (use-case interfaces) that the `ExcelApp`
  facade would implement.
- Mapping all entry points (scripts, tests, consumers) through those ports.

#### Why it is not worth it here

EzXl is a **thin automation toolbox**, not a domain-rich application.  There is
no domain model to protect from infrastructure contamination.  The core
abstraction IS the infrastructure — COM dispatch and Win32 window handles are
the primitives.

The proxy hierarchy (`ExcelApp` → `WorkbookProxy` → `SheetProxy`) already acts
as the adapter layer: it wraps raw COM objects behind a clean Python API.
Wrapping that adapter layer in a second hexagonal shell would be pure ceremony
with no protective benefit.

Adding a formal port/adapter split at the top level would increase indirection,
complicate the import graph, and make the library harder to understand for new
contributors — all with zero gain in testability or replaceability, since the
COM dependency is fundamental and irreplaceable.

#### What IS already hexagonal (keep this)

The `_protocols.py` + COM/pywinauto backend split inside `ezxl/gui/` is the
right scope for Ports & Adapters in this project.  It delivers real value:
backends can be swapped in tests and in production without changing callers.
This pattern should be preserved and extended to new GUI surfaces as needed.

#### Conclusion

Keep the current architecture.  The GUI-level Ports & Adapters pattern is
sufficient and correctly scoped.  Future contributors should not re-open the
debate about project-level hexagonal architecture — this decision record
documents why it was evaluated and rejected.

The `core/hexagonal-architecture-standards.instructions.md` file is available
for reference but does not apply to this project (see Project-Specific Overrides
below).

---

## Important Notes

- **SAP boundary**: EzXl must never contain SAP-specific logic. Any interaction with SAP add-ins belongs in a separate consumer library. If a feature request implies SAP knowledge, reject it at the EzXl level.
- **Design audit**: read `.tmp/rapport-audit-initial.md` for the full rationale behind architectural decisions made at project inception.
- **Module creation order**: `exceptions.py` must be created before any proxy module (`_excel_app.py`, `_workbook.py`, `_sheet.py`). The entire error contract depends on it.
- **COM errors**: never let `pywintypes.com_error` propagate out of `src/ezxl/`. Always wrap at the COM call site.
- **`core/hexagonal-architecture-standards.instructions.md`** contains intentional `{{my_project}}` placeholders — these are examples in a shared standards file. Do not replace them.
- **`cliff.toml`** uses Jinja2 `{{ variable }}` syntax — do not confuse with project template placeholders.
