# Project Instructions

## Project Overview

**Project Name**: ezxl
**Tech Stack**: [e.g., Python 3.11+, FastAPI, PostgreSQL]
**Environment**: [e.g., Corporate environment with proxy restrictions]

## Project Description

[Brief description of the project's purpose, scope, and key features.]

## Architecture

```text
project/
├── ezxl/   # Source code (Python package)
├── tests/               # Test files
├── docs/                # Documentation
├── .github/
│   ├── instructions/    # AI agent instructions
│   └── workflows/       # CI/CD workflows
├── .hooks/              # Git hooks
├── .scripts/            # Development & build scripts
├── pyproject.toml       # Project configuration
└── README.md            # Project documentation
```

## Key Conventions

- [List project-specific conventions]
- [e.g., "All API endpoints must return JSON responses"]
- [e.g., "Database migrations use Alembic"]

## Environment Setup

```bash
# Create virtual environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\activate

# Install dependencies
pip install -e ".[dev]"

# Set up git hooks
git config core.hooksPath .hooks
```

## Instruction Files

| File                                                                    | Purpose                                              |
| ----------------------------------------------------------------------- | ---------------------------------------------------- |
| `core/advanced-cognitive-conduct.instructions.md`                       | Core reasoning framework                             |
| `core/commit-standards.instructions.md`                                 | Git commit conventions                               |
| `core/hexagonal-architecture-standards.instructions.md`                 | Hexagonal Architecture (Ports & Adapters) guidelines |
| `languages/python/python-development-standards.instructions.md`         | Python coding standards                              |
| `languages/python/python-formatting-standards.instructions.md`          | Code formatting rules                                |
| `languages/python/pyproject-standards.instructions.md`                  | pyproject.toml conventions                           |
| `languages/javascript/javascript-development-standards.instructions.md` | JavaScript coding standards                          |
| `languages/javascript/javascript-formatting-standards.instructions.md`  | JavaScript formatting rules                          |

## Project-Specific Overrides

[Document any project-specific rules that override the general instructions above.]

- [e.g., "Use SQLAlchemy 2.0 style queries exclusively"]
- [e.g., "All CLI commands use Click, not argparse"]

## Important Notes

- [Any critical information that AI agents must know]
- [e.g., "This project uses a custom authentication system"]
- [e.g., "Database schema changes require a migration file"]
