# Publish to PyPI Workflow

`publish-pypi.yml` builds the package and publishes it to PyPI. It is normally
triggered by `auto-tag.yml` after a new version tag is created, but can also
be run manually.

## Triggers

| Event               | How                                                    |
| ------------------- | ------------------------------------------------------ |
| `workflow_call`     | Called by `auto-tag.yml` with `version` + `tag` inputs |
| `workflow_dispatch` | Manual run from the Actions tab                        |

Manual runs have a `skip_tests` option (default `false`, not recommended).

## Authentication

Publishing uses **OIDC trusted publishing** — no stored secret. The workflow
mints a short-lived token that PyPI trusts directly via the `id-token: write`
permission. The trusted publisher must be configured on pypi.org under the
ezplog project settings:

```text
pypi.org → Manage project ezplog → Publishing → Add a new publisher
  Owner:     neuraaak
  Repo:      ezplog
  Workflow:  publish-pypi.yml
  Environment: pypi
```

## Jobs

```text
validate (build + test) ──► publish (OIDC deploy)
```

### `validate` — Validate & Build

1. Resolve version: uses `inputs.version` if provided (workflow_call), otherwise reads from `pyproject.toml`
2. Preview mode summary — printed if triggered by push/PR or preview tag (no failure, informational only)
3. `uv sync --frozen --extra dev`
4. `PYTHONPATH=src uv run lint-imports` — import contracts
5. `uv run pytest` — full suite (skipped if `skip_tests=true`)
6. `uv build` — produces `dist/*.whl` and `dist/*.tar.gz`
7. `uv run twine check dist/*` — validates package metadata
8. Wheel integrity check: `python -m zipfile -t dist/*.whl` + inline script verifying `ezplog/` is present in the wheel
9. Upload `dist/` as a job artifact (retained 1 day)

### `publish` — Publish to PyPI (OIDC)

Runs only if `validate` succeeds **and** the event is `workflow_dispatch` or a
`workflow_call` whose `tag` input does not start with `preview-`. Downloads the
artifact built upstream (no rebuild). Steps:

1. `actions/download-artifact` — retrieves `dist/`
2. `pypa/gh-action-pypi-publish@v1.14.0` — publishes via OIDC, no password

> **Note — attestations disabled:** when called as a reusable workflow from
> `auto-tag.yml`, the OIDC token's Build Config URI reflects the caller
> (`auto-tag.yml`) rather than `publish-pypi.yml`, causing PyPI Trusted
> Publisher verification to fail. `attestations: false` works around this.

The job runs in the `pypi` environment. Configure required reviewers there
in Repository → Settings → Environments if you want a manual gate.

## Concurrency

A `publish` concurrency group with `cancel-in-progress: false` prevents two
simultaneous publish runs. A run in progress is never interrupted.

## Release workflow

```text
1. Bump version in pyproject.toml
2. Commit & push to main
   → auto-tag.yml detects a new version, creates vX.Y.Z tag
   → triggers publish-pypi.yml automatically
3. Verify on https://pypi.org/project/ezplog/
```

Manual trigger via gh CLI:

```bash
gh workflow run publish-pypi.yml
```

## Troubleshooting

**Version mismatch** — the `version` input from auto-tag differs from
`pyproject.toml`. Bump the version, commit, and push again.

**OIDC failure** — check that the trusted publisher on pypi.org matches
the repo name, workflow filename (`publish-pypi.yml`), and environment name
(`pypi`) exactly.

**Tests fail** — run `uv run pytest` locally, fix, push. Do not use
`skip_tests=true` to work around failures.
