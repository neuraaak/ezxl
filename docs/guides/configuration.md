# Configuration

This guide covers environment setup for both standard and corporate (offline) environments.

---

## Python version

`ezxl` requires Python 3.11 or later. This is enforced at import time: attempting to import the library on an older interpreter raises `RuntimeError` immediately.

```bash
python --version
# Python 3.11.x or later required
```

---

## Bitness verification

Excel COM dispatch requires that the Python interpreter bitness matches the installed Excel bitness. Microsoft 365 defaults to 64-bit. Mixing bitnesses causes `pythoncom` registration errors that are difficult to diagnose without this check.

```bash
python -c "import struct; print(struct.calcsize('P') * 8, 'bit')"
# Expected on a standard workstation: 64 bit
```

If the output does not match your Excel installation, use the matching Python installer from [python.org](https://www.python.org/downloads/windows/).

---

## Virtual environment

Create and activate a virtual environment before installing any dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate
```

---

## Standard installation

```bash
pip install -e ".[dev]"
```

This installs `ezxl` in editable mode with all development dependencies from `pyproject.toml`.

---

## Corporate offline installation

In environments with no PyPI access, all packages must be available as `.whl` files in a local directory. Build or obtain the wheels on a machine with internet access, then transfer them to the target environment.

```bash
# On a machine with internet access, download all wheels
pip download ezxl[dev] -d C:\wheels

# On the target machine (no internet), install from local wheels
pip install --no-index --find-links C:\wheels -e ".[dev]"
```

If the corporate proxy must be used for some network operations:

```bash
set HTTPS_PROXY=http://proxy.corp.example.com:8080
pip install --proxy http://proxy.corp.example.com:8080 -e ".[dev]"
```

!!! tip "Proxy in pip.ini"
To avoid setting the proxy on every command, add it to `%APPDATA%\pip\pip.ini`:

```ini
    [global]
    proxy = http://proxy.corp.example.com:8080
```

---

## pywin32 post-install step

`pywin32` ships COM registration scripts that must be run once after installation. Skipping this step causes `ImportError: No module named 'pywintypes'` or silent COM failures at runtime.

```bash
python .venv/Scripts/pywin32_postinstall.py -install
```

Run this command once per virtual environment. It does not need to be repeated on the same machine unless the virtual environment is recreated.

---

## Optional dependencies

### pywinauto

Required only if you intend to use the pywinauto GUI backends (`PywinautoRibbonBackend`, `PywinautoMenuBackend`, `PywinautoDialogBackend`, `PywinautoKeysBackend`). The main COM layer has no dependency on pywinauto.

```bash
pip install pywinauto
```

In a corporate offline environment:

```bash
pip install --no-index --find-links C:\wheels pywinauto
```

### Documentation dependencies

To build the documentation locally, install the `docs` optional group:

```bash
pip install -e ".[docs]"
mkdocs serve
```

The docs group includes `mkdocs`, `mkdocs-material`, `mkdocstrings[python]`, and related plugins.

---

## pyproject.toml optional dependency groups

The full optional dependency configuration in `pyproject.toml`:

```toml
[project.optional-dependencies]
dev = [
    "ruff>=0.1.0",
    "ty>=0.0.13",
    "pyright>=1.1.0",
    "pre-commit>=3.0.0",
    "import-linter>=2.0.0",
    "bandit>=1.7.0",
    "pytest>=7.0.0",
    "pytest-cov>=4.0.0",
    "pytest-mock>=3.10.0",
    "pytest-xdist>=3.0.0",
    "build>=1.0.0",
    "twine>=4.0.0",
    "rich>=13.0.0",
]
test = [
    "pytest>=7.0.0",
    "pytest-cov>=4.0.0",
    "pytest-mock>=3.10.0",
    "pytest-xdist>=3.0.0",
    "import-linter>=2.0.0",
]
docs = [
    "mkdocs>=1.6.0",
    "mkdocs-material>=9.5.0",
    "mkdocstrings[python]>=0.27.0",
    "mkdocs-section-index>=0.3.0",
    "mkdocs-coverage>=1.1.0",
    "git-cliff>=2.7.0",
]
```

---

## Git hooks

Pre-commit hooks are stored in `.hooks/`. Register them with git:

```bash
git config core.hooksPath .hooks
```

The hooks enforce commit message formatting (Conventional Commits) and run the linter before each commit.
