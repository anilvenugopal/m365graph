# m365graph

A small Microsoft Graph client for app-only use (an Azure app registration with a client
secret), shared by projects that work with Microsoft 365:

- **`GraphClient`**: sign-in with MSAL, retries on network errors, and throttling handled for
  every thread at once (a 429's `Retry-After` pauses all callers); `json`, `pages` (paging and
  delta queries), `all`, `stream` (downloads in chunks) and `upload` (upload sessions).
- **`DriveStore`**: one SharePoint document library, under a root folder, seen as folders and
  files: `list`, `exists`, `make_folder`, `put`, `open`, `delete`. A Lattice host uses it as a
  folder store (the folder picker) or an attachment store, without either package importing
  the other.

```python
from m365graph import DriveStore, GraphClient

graph = GraphClient.from_env()       # TENANT_ID, CLIENT_ID, CLIENT_SECRET (+ AUTHORITY)
graph.roles()                        # ['Sites.Read.All', …]
store = DriveStore.for_site(graph, "contoso.sharepoint.com:/sites/claims", "Documents",
                            root="Claim files")
store.list("")
```

Permissions (application): `Sites.Read.All` to browse and read; `Sites.ReadWrite.All` (or
`Sites.Selected` with write granted on the site) to create folders and upload.

Everything other projects may use is in [docs/public_contract.md](docs/public_contract.md).

```bash
uv sync && uv run pytest      # tests use a fake Graph; nothing reaches Microsoft
```
