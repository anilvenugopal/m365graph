# m365graph: public contract

What other projects may rely on. Anything not listed may change in any release. Semantic
versioning from the package version; before 1.0 a minor release may break the contract.

| Name | Kind | Signature |
|---|---|---|
| `GraphClient` | class | `(tenant_id, client_id, client_secret, authority=None, *, timeout=60, max_retries=5, pool_size=32, session=None, token_provider=None)` |
| `GraphClient.from_env` | classmethod | `(env_file=".env", **kwargs)` |
| `GraphClient.token`, `.roles` | methods | `()` |
| `GraphClient.send` | method | `(method, url, *, auth=True, headers=None, **requests_kwargs) -> Response` |
| `GraphClient.json` | method | `(method, url, *, none_on_404=False, **kwargs) -> dict \| None` |
| `GraphClient.pages`, `.all` | methods | `(url, **kwargs)`: pages / every item, following `@odata.nextLink` |
| `GraphClient.stream` | method | `(url, chunk_size=1 MiB) -> Iterator[bytes]` |
| `GraphClient.upload` | method | `(session_url, chunks, size) -> dict` |
| `GraphClient.error` | staticmethod | `(response) -> GraphError` |
| `GraphError` | exception | `.status`, `.code`; the message never holds a token |
| `find_site`, `libraries` | functions | `(graph, site)`, `(graph, site_id)` |
| `DriveStore` | class | `(graph, drive_id, label, root="")`; `DriveStore.for_site(graph, site, library="", *, root="", label="")` |
| `DriveStore.list`, `.exists`, `.make_folder` | methods | `(path="")`, `(path)`, `(path, name) -> Folder` |
| `DriveStore.put`, `.open`, `.delete` | methods | `(key, chunks, max_bytes) -> (size, sha256)`, `(key) -> file`, `(key)` |
| `DriveStore.label`, `.root` | attributes | |
| `Folder` | dataclass | `name`, `path` |
| `PathError` | exception | a `ValueError`: a path outside the root, or not a folder |
| `TooLarge` | exception | a file passed `max_bytes` |
