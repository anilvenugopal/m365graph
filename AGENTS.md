# m365graph — Agent Instructions

A small Microsoft Graph client used by other projects (Lattice's SharePoint store, a mail
engine, spcopy). It **knows nothing of the projects that use it**: no Lattice, no app code.

- `m365graph/client.py`: app-only sign-in (MSAL), retries, throttling, paging, delta, streams.
- `m365graph/drives.py`: sites and libraries; `DriveStore`, a library as folders and files.
- `docs/public_contract.md`: every name other projects may rely on; keep it true.

## Running

```bash
uv sync
uv run pytest        # fakes only; never calls Microsoft
uv run ruff check .
```

`.env` holds a tenant's app credentials (`.env.example`). Never read it, print it or commit
it. Code loads it with `GraphClient.from_env()`.

## Git

No AI attribution in git history: never add `Co-Authored-By` lines or "Generated with" lines
to commits or PR bodies.
