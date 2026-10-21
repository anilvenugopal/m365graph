"""SharePoint sites and document libraries, and DriveStore: one library seen as folders and
files under a root folder.

    graph = GraphClient.from_env()
    store = DriveStore.for_site(graph, "contoso.sharepoint.com:/sites/claims", "Documents",
                                root="Claim files", label="Claims library")
    store.list("2026")                     # [Folder(name="CLM-26-004871", path="2026/CLM-26-004871"), …]
    store.make_folder("2026", "CLM-26-004999")
    size, sha256 = store.put("2026/x/estimate.pdf", chunks, max_bytes=25 * 2**20)
    with store.open("2026/x/estimate.pdf") as f: f.read()

Paths are relative to the root and use "/"; a ".." segment is refused. DriveStore has the same
methods as Lattice's local file store, so a Lattice host can use it as a folder store or an
attachment store without either package importing the other.
"""

from __future__ import annotations

import hashlib
import io
import tempfile
from collections.abc import Iterable, Iterator
from dataclasses import dataclass

from m365graph.client import CHUNK, GraphClient, GraphError, quote_path

SMALL_UPLOAD = 4 * 1024 * 1024         # Graph takes a simple PUT up to 4 MB; larger needs a session


class PathError(ValueError):
    """A path that leaves the store's root, names nothing, or is not a folder."""


class TooLarge(Exception):
    """A file passed max_bytes while it was being written."""


@dataclass(frozen=True)
class Folder:
    name: str
    path: str                           # relative to the store's root


def find_site(graph: GraphClient, site: str) -> dict:
    """A site by id ("contoso.sharepoint.com,guid,guid") or by host and path
    ("contoso.sharepoint.com:/sites/claims")."""
    found = graph.json("GET", f"/sites/{site}?$select=id,displayName,webUrl", none_on_404=True)
    if not found:
        raise GraphError(404, "SiteNotFound", f"No site {site!r}, or the app may not read it")
    return found


def libraries(graph: GraphClient, site_id: str) -> list[dict]:
    """A site's document libraries: id, name, webUrl."""
    return graph.all(f"/sites/{site_id}/drives?$select=id,name,driveType,webUrl")


class _Download(io.RawIOBase):
    """A Graph download read like a file, a chunk at a time."""

    def __init__(self, chunks: Iterator[bytes]):
        self._chunks, self._pending = chunks, b""

    def readable(self) -> bool:
        return True

    def readinto(self, buffer) -> int:
        while not self._pending:
            try:
                self._pending = next(self._chunks)
            except StopIteration:
                return 0
        n = min(len(buffer), len(self._pending))
        buffer[:n], self._pending = self._pending[:n], self._pending[n:]
        return n

    def close(self) -> None:
        close = getattr(self._chunks, "close", None)
        if close:
            close()
        super().close()


class DriveStore:
    def __init__(self, graph: GraphClient, drive_id: str, label: str, root: str = ""):
        self.graph, self.drive_id, self.label = graph, drive_id, label
        self.root = self._clean(root)

    @classmethod
    def for_site(cls, graph: GraphClient, site: str, library: str = "", *, root: str = "",
                 label: str = "") -> DriveStore:
        """The store for a site's library by name (default: the site's default library)."""
        site_row = find_site(graph, site)
        if library:
            drive = next((d for d in libraries(graph, site_row["id"]) if d["name"] == library), None)
            if drive is None:
                raise GraphError(404, "LibraryNotFound", f"No library {library!r} in {site_row['displayName']}")
        else:
            drive = graph.json("GET", f"/sites/{site_row['id']}/drive?$select=id,name")
        return cls(graph, drive["id"], label or f"{site_row['displayName']} › {drive['name']}", root)

    # ── Paths ────────────────────────────────────────────────────────────────

    @staticmethod
    def _clean(path: str) -> str:
        parts = [p for p in (path or "").replace("\\", "/").split("/") if p and p != "."]
        if ".." in parts:
            raise PathError(f"{path!r} leaves the store")
        return "/".join(parts)

    def _full(self, path: str) -> str:
        return "/".join(p for p in (self.root, self._clean(path)) if p)

    def _item_url(self, path: str) -> str:
        full = self._full(path)
        return f"/drives/{self.drive_id}/root:/{quote_path(full)}:" if full else f"/drives/{self.drive_id}/root"

    def _item(self, path: str) -> dict | None:
        url = self._item_url(path)
        return self.graph.json("GET", url.rstrip(":") + "?$select=id,name,folder,file,size",
                               none_on_404=True)

    # ── Folders ──────────────────────────────────────────────────────────────

    def list(self, path: str = "") -> list[Folder]:
        """The folders directly inside `path`, by name."""
        item = self._item(path)
        if not item or "folder" not in item:
            raise PathError(f"{path or '/'} is not a folder in {self.label}")
        base = self._clean(path)
        children = self.graph.all(f"/drives/{self.drive_id}/items/{item['id']}/children"
                                  "?$select=id,name,folder&$top=999")
        return sorted((Folder(c["name"], f"{base}/{c['name']}" if base else c["name"])
                       for c in children if "folder" in c), key=lambda f: f.name.lower())

    def exists(self, path: str) -> bool:
        try:
            item = self._item(path)
        except PathError:
            return False
        return bool(item and "folder" in item)

    def make_folder(self, path: str, name: str) -> Folder:
        name = (name or "").strip()
        if not name or "/" in name or "\\" in name or name in (".", ".."):
            raise PathError(f"{name!r} is not a folder name")
        parent = self._item(path)
        if not parent or "folder" not in parent:
            raise PathError(f"{path or '/'} is not a folder in {self.label}")
        resp = self.graph.send("POST", f"/drives/{self.drive_id}/items/{parent['id']}/children",
                               json={"name": name, "folder": {},
                                     "@microsoft.graph.conflictBehavior": "fail"})
        if not resp.ok and resp.status_code != 409:      # 409: it is already there
            raise self.graph.error(resp)
        base = self._clean(path)
        return Folder(name, f"{base}/{name}" if base else name)

    def _make_parents(self, path: str) -> None:
        parts = self._clean(path).split("/")[:-1]
        for i, part in enumerate(parts):
            here = "/".join(parts[:i])
            if not self.exists(f"{here}/{part}" if here else part):
                self.make_folder(here, part)

    # ── Files ────────────────────────────────────────────────────────────────

    def put(self, key: str, chunks: Iterable[bytes], max_bytes: int) -> tuple[int, str]:
        """Upload a file, creating its folders; returns its size and SHA-256. Nothing is uploaded
        if it passes max_bytes."""
        size, digest = 0, hashlib.sha256()
        with tempfile.SpooledTemporaryFile(max_size=SMALL_UPLOAD) as spool:
            for chunk in chunks:
                size += len(chunk)
                if size > max_bytes:
                    raise TooLarge
                digest.update(chunk)
                spool.write(chunk)
            spool.seek(0)
            self._make_parents(key)
            url = self._item_url(key)
            if size <= SMALL_UPLOAD:
                resp = self.graph.send("PUT", f"{url}/content", data=spool.read(),
                                       headers={"Content-Type": "application/octet-stream"})
                if not resp.ok:
                    raise self.graph.error(resp)
            else:
                session = self.graph.json("POST", f"{url}/createUploadSession",
                                          json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
                self.graph.upload(session["uploadUrl"], iter(lambda: spool.read(CHUNK), b""), size)
        return size, digest.hexdigest()

    def open(self, key: str) -> io.BufferedReader:
        """A file, read like a local one; raises FileNotFoundError when it is not there."""
        if not self._item(key):
            raise FileNotFoundError(key)
        return io.BufferedReader(_Download(self.graph.stream(f"{self._item_url(key)}/content")))

    def delete(self, key: str) -> None:
        resp = self.graph.send("DELETE", self._item_url(key).rstrip(":"))
        if not resp.ok and resp.status_code != 404:
            raise self.graph.error(resp)
