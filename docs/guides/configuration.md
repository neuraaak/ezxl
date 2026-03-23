# How to install ezxl

Use this guide to install EzXl in a local Windows environment, an offline corporate environment, or a documentation workspace.

## 🔧 Prerequisites

- Python 3.11 or later
- Windows if you plan to use COM automation
- Excel installed if you plan to automate a live session

## 📝 Steps

1. Verify interpreter bitness before any COM usage.

   ```bash
   python -c "import struct; print(struct.calcsize('P') * 8, 'bit')"
   ```

2. Create and activate a virtual environment.

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. Pick the installation flow that matches your environment.

    === "Runtime package"

        ```bash
        pip install ezxl
        ```

    === "Contributor workspace"

        ```bash
        uv sync --extra dev --extra test --extra docs
        ```

    === "Offline wheels"

        ```bash
        pip install --no-index --find-links C:\wheels ezxl
        ```

4. Register the pywin32 COM components for the virtual environment.

   ```bash
   python .venv/Scripts/pywin32_postinstall.py -install
   ```

5. Verify that the package imports cleanly.

   ```bash
   python -c "import ezxl; print(ezxl.__version__)"
   ```

## Variations

If you need a corporate proxy for download operations:

```bash
set HTTPS_PROXY=http://proxy.corp.example.com:8080
pip install --proxy http://proxy.corp.example.com:8080 ezxl
```

!!! tip "🔧 Persist the proxy in pip.ini"
    Add the proxy once in `%APPDATA%\pip\pip.ini` if your environment requires it for every installation.

If you only need to build the documentation locally:

```bash
uv sync --extra docs
uv run mkdocs serve
```

## ✅ Result

You have a virtual environment with EzXl installed, the COM registration step completed, and a command you can use to verify the package import immediately.
