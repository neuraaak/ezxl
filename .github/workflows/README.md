# CI/CD Pipeline

Numbered reusable workflows, one responsibility each, numbered by dependency
order. The number sorts them in the file tree and documents the cascade.

```text
00-install-deps  ← primes the uv cache (keyed on uv.lock)
01-ci            ← lint ∥ type-check ∥ security → test matrix; the quality gate
                   (the matrix runs Ubuntu + Windows; security gates the
                    workflow, not the matrix)
02-tag-sync      ← orchestrator: CI gate → tag_action → cascade 03 then 04
03-publish       ← build always; upload to PyPI on `create` only (OIDC)
04-docs          ← build always; deploy latest / dev / none by tag_action
```

## The three-state `tag_action`

`02-tag-sync` does **not** bump the version — a human writes it in
`pyproject.toml`. On each run it derives one signal, `tag_action`:

| `tag_action` | When                                        | Tag                           | 03 Publish            | 04 Docs                   |
| ------------ | ------------------------------------------- | ----------------------------- | --------------------- | ------------------------- |
| `create`     | push to `main`, `vX.Y.Z` doesn't exist yet  | create `vX.Y.Z` + `vX-latest` | upload (OIDC)         | deploy `X.Y.Z` + `latest` |
| `skip`       | push to `main`, `vX.Y.Z` already exists     | none                          | **not called**        | deploy `dev` alias        |
| `preview`    | `workflow_dispatch` off a non-`main` branch | none                          | build only, no upload | build only, no push       |

`create` is **release mode**; `skip` and `preview` are **validation mode**
(nothing mutating is uploaded). `preview` exercises the whole job graph from a
branch before it ever runs for real on `main`.

## Tag immutability

- **`vX.Y.Z`** — created **once, never force-pushed**. Re-running `02` on the
  same version yields `tag_action=skip` (the gate finds the tag). This is the
  reproducibility guarantee.
- **`vX-latest`** — floating major alias, force-pushed to the latest **release**
  only (on `create`), never on every commit.

## Why a single orchestrator (`02`)

A tag pushed by `GITHUB_TOKEN` does **not** trigger another workflow (loop
protection). So `02-tag-sync` calls `03` and `04` directly via `uses:` rather
than relying on a `push: tags` event that would never fire.

## Gating is keyed on inputs, not `github.event_name`

In a reusable workflow the `github` context is **inherited from the caller**, so
`github.event_name` inside `03`/`04` is the caller's event (`push`), never
`workflow_call`. Publish/deploy steps therefore gate on the `tag_action` input.

## No `secrets: inherit`

The cascade passes typed `with:` inputs only and relies on **OIDC**
(`id-token: write`) for PyPI trusted publishing — there is no token to pass.

## Local equivalents

Every CI step runs a **poe task** from `[tool.poe.tasks]`, so the local
command and the CI command are the same string:

```bash
# 01-ci quality gate
uv run poe check            # format-check → lint → typecheck → imports → test
uv run poe security         # bandit + pip-audit
uv run poe test-ci          # what CI runs: deselects `excel`, coverage gate on

# 03-publish build
uv run poe build
uv run twine check dist/*

# 04-docs build
uv run poe docs-coverage
uv run poe graph
uv run mike deploy <version> latest   # add --push to deploy
```

The release workflows are intentionally not run locally: tag creation, PyPI
publication and documentation deployment are remote side effects.

## The `excel` marker

Tests marked `excel` drive a live Excel instance over COM: they need Windows and
an Excel installation, so every CI job deselects them (`poe test-ci`,
`poe docs-coverage`) and the COM/GUI modules they cover are omitted from the
coverage gate. Run them locally on Windows with `uv run poe test`.
