# Deploy Documentation Workflow

`docs.yml` builds and deploys the MkDocs Material documentation to GitHub Pages
(`https://neuraaak.github.io/ezplog/`). It is normally called by `auto-tag.yml`
after a tag is created or moved, but can also be run manually.

## Triggers

| Event               | How                                                  |
| ------------------- | ---------------------------------------------------- |
| `workflow_call`     | Called by `auto-tag.yml` with boolean input `deploy` |
| `workflow_dispatch` | Manual run from the Actions tab (always deploys)     |

`workflow_call` input `deploy` (boolean, default `false`):

- `true` → full deploy to GitHub Pages (release mode)
- `false` → build + checks only, no push (preview mode)

## Job: `deploy`

Single job, runs on `ubuntu-24.04`. Steps in order:

| Step               | Command / Action                                                                                           |
| ------------------ | ---------------------------------------------------------------------------------------------------------- |
| Checkout           | Full history (`fetch-depth: 0`) required for git-cliff                                                     |
| Install            | `uv sync --frozen` with `docs` and `test` extras                                                           |
| Tests              | `uv run pytest --cov=src/ezplog --cov-report=html --cov-report=xml` — produces `htmlcov/` + `coverage.xml` |
| Import contracts   | `PYTHONPATH=src uv run lint-imports`                                                                       |
| Architecture graph | Runs `.scripts/dev/generate_architecture_graph.py` if present                                              |
| Changelog          | `git-cliff` writes `docs/changelog.md` from conventional commits                                           |
| Resolve version    | Reads `version` from `pyproject.toml`; fails if absent                                                     |
| Wait for PyPI      | **Deploy mode only** — polls `pypi.org/pypi/ezplog/<version>/json` every 15 s, timeout 900 s               |
| Git config         | **Deploy mode only** — configures `github-actions[bot]` user for mike                                      |
| Deploy             | **Deploy mode only** — `mike deploy --push --update-aliases <version> latest` + `mike set-default latest`  |
| Preview summary    | **Preview mode only** — writes a step summary, no Pages push                                               |

Deploy mode = `inputs.deploy == true` or `github.event_name == 'workflow_dispatch'`.

## Concurrency

Group `docs-$ref` with `cancel-in-progress: true` — a new docs run on the same ref
cancels the previous one.

## Permissions

`contents: write` is required so mike can push to the `gh-pages` branch.

## GitHub Pages setup

Pages must be configured to serve from the `gh-pages` branch:

```text
Repository → Settings → Pages → Source: Deploy from a branch → gh-pages / root
```

`mike` manages versioned aliases (`latest`) and creates/updates this branch automatically.

## Local preview

```bash
uv sync --extra docs
uv run mkdocs serve
# Open http://127.0.0.1:8000
```

Strict build (fails on warnings):

```bash
uv run mkdocs build --strict
```

## Manual trigger

```bash
gh workflow run docs.yml
```

## Troubleshooting

**Tests fail** — coverage generation fails if tests are broken. Fix the tests
and re-run. Do not add `|| true` to mask failures.

**mkdocstrings "Module not found"** — verify the module path in `docs/api/*.md`
matches the installed package. Run `uv run python -c "import ezplog"` to
confirm the package is importable.

**Permission denied on gh-pages** — check that `contents: write` is present and
that no branch-protection rule blocks force-pushes to `gh-pages`.

**Changelog not updating** — ensure `cliff.toml` exists at the repo root and
that commits follow the conventional commit format.
