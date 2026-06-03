# ///////////////////////////////////////////////////////////////
# CHECK_YAML - Pre-commit hook: YAML syntax validator
# Project: ezxl
# ///////////////////////////////////////////////////////////////

"""Pre-commit hook: validate YAML syntax."""

# ///////////////////////////////////////////////////////////////
# IMPORTS
# ///////////////////////////////////////////////////////////////
# Standard library
import sys

# Third-party
import yaml

# ///////////////////////////////////////////////////////////////

rc = 0
for path in sys.argv[1:]:
    try:
        with open(path, encoding="utf-8") as f:
            # yaml.Loader mirrors the --unsafe flag of pre-commit-hooks check-yaml,
            # required for MkDocs !!python/name: tags.
            yaml.load(f, Loader=yaml.Loader)  # noqa: S506
    except yaml.YAMLError as exc:
        print(f"{path}: {exc}")
        rc = 1
    except OSError as exc:
        print(f"{path}: {exc}")
        rc = 1

sys.exit(rc)
