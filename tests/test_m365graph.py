"""GraphClient and DriveStore against a fake Graph served through a requests adapter."""

import io
import json
import re
from urllib.parse import unquote, urlsplit

import pytest
import requests
from requests.adapters import BaseAdapter

from m365graph import DriveStore, GraphClient, GraphError, PathError, TooLarge
from m365graph import drives as drives_module


class FakeGraph(BaseAdapter):
    """One drive: folders and files by path; answers the calls DriveStore makes."""

    def __init__(self):
        super().__init__()
        self.tree = {"": {"id": "root", "folder": {}}}        # path -> item
        self.files: dict[str, bytes] = {}
        self.calls: list[tuple[str, str]] = []
        self.throttle_next = 0
        self.uploads: dict[str, dict] = {}

    def add_folder(self, path):
        parts = path.split("/")
        for i in range(1, len(parts) + 1):
            self.tree.setdefault("/".join(parts[:i]), {"id": f"i{len(self.tree)}", "folder": {}})

    # requests' adapter interface
    def close(self):
        pass

    def send(self, request, **kwargs):
        url = urlsplit(request.url)
        path = unquote(url.path).removeprefix("/v1.0")
        self.calls.append((request.method, path))
        if self.throttle_next:
            self.throttle_next -= 1
            return self._resp(request, 429, {"error": {"code": "TooMany"}}, {"Retry-After": "0.01"})
        if path.startswith("/upload/"):
            return self._upload(request, path)
        m = re.fullmatch(r"/drives/d1/root(?::/(.*?):?)?(/children|/content|/createUploadSession)?", path)
        if m:
            item_path, tail = (m.group(1) or "").strip("/"), m.group(2) or ""
            return self._item(request, item_path, tail)
        m = re.fullmatch(r"/drives/d1/items/([^/]+)/children", path)
        if m:
            parent = next(p for p, item in self.tree.items() if item["id"] == m.group(1))
            return self._item(request, parent, "/children")
        return self._resp(request, 404, {"error": {"code": "itemNotFound", "message": path}})

    def _item(self, request, item_path, tail):
        if request.method == "GET" and tail == "":
            item = self.tree.get(item_path)
            if item is None and item_path in self.files:
                item = {"id": "f", "file": {}, "size": len(self.files[item_path])}
            if item is None:
                return self._resp(request, 404, {"error": {"code": "itemNotFound", "message": item_path}})
            return self._resp(request, 200, {**item, "name": item_path.rsplit("/", 1)[-1]})
        if request.method == "GET" and tail == "/children":
            prefix = f"{item_path}/" if item_path else ""
            kids = [{"id": i["id"], "name": p[len(prefix):], "folder": {}} for p, i in self.tree.items()
                    if p.startswith(prefix) and p and "/" not in p[len(prefix):]]
            kids += [{"id": "f", "name": p[len(prefix):], "file": {}} for p in self.files
                     if p.startswith(prefix) and "/" not in p[len(prefix):]]
            return self._resp(request, 200, {"value": kids})
        if request.method == "POST" and tail == "/children":
            name = json.loads(request.body)["name"]
            target = f"{item_path}/{name}" if item_path else name
            if target in self.tree:
                return self._resp(request, 409, {"error": {"code": "nameAlreadyExists"}})
            self.add_folder(target)
            return self._resp(request, 201, {"id": self.tree[target]["id"], "name": name})
        if request.method == "PUT" and tail == "/content":
            self.files[item_path] = request.body
            return self._resp(request, 201, {"id": "f", "size": len(request.body)})
        if request.method == "POST" and tail == "/createUploadSession":
            self.uploads[item_path] = {"data": b""}
            return self._resp(request, 200, {"uploadUrl": f"https://up.example/upload/{item_path}"})
        if request.method == "GET" and tail == "/content":
            if item_path not in self.files:
                return self._resp(request, 404, {"error": {"code": "itemNotFound"}})
            return self._resp(request, 200, None, raw=self.files[item_path])
        if request.method == "DELETE":
            if self.files.pop(item_path, None) is None:
                return self._resp(request, 404, {"error": {"code": "itemNotFound"}})
            return self._resp(request, 204, None)
        return self._resp(request, 400, {"error": {"code": "unexpected"}})

    def _upload(self, request, path):
        assert "Authorization" not in request.headers            # the session URL is pre-authorised
        target = path.removeprefix("/upload/")
        self.uploads[target]["data"] += request.body
        total = int(request.headers["Content-Range"].rsplit("/", 1)[1])
        if len(self.uploads[target]["data"]) == total:
            self.files[target] = self.uploads[target]["data"]
            return self._resp(request, 201, {"id": "f", "size": total})
        return self._resp(request, 202, {"nextExpectedRanges": []})

    @staticmethod
    def _resp(request, status, body, headers=None, raw=None):
        resp = requests.Response()
        resp.status_code, resp.request, resp.url = status, request, request.url
        resp.headers.update(headers or {})
        data = raw if raw is not None else (json.dumps(body).encode() if body is not None else b"")
        resp.raw = io.BytesIO(data)              # read by iter_content when streaming
        resp._content = False if raw is not None else data
        resp.reason = {200: "OK", 404: "Not Found", 429: "Too Many Requests"}.get(status, "")
        return resp


@pytest.fixture
def fake():
    graph = FakeGraph()
    graph.add_folder("Claims/2026/CLM-1")
    graph.add_folder("Claims/2026/CLM-2")
    graph.add_folder("Other")
    return graph


@pytest.fixture
def client(fake):
    session = requests.Session()
    session.mount("https://", fake)
    return GraphClient("t", "c", "s", session=session, token_provider=lambda: "token")


@pytest.fixture
def store(client):
    return DriveStore(client, "d1", "Claims library", root="Claims")


def test_folders_are_listed_under_the_root_and_paths_cannot_leave_it(store):
    assert [(f.name, f.path) for f in store.list("")] == [("2026", "2026")]
    assert [f.path for f in store.list("2026")] == ["2026/CLM-1", "2026/CLM-2"]
    assert store.exists("2026/CLM-1") and not store.exists("2026/nope")
    with pytest.raises(PathError):
        store.list("2026/nope")
    with pytest.raises(PathError):
        store.list("../Other")


def test_make_folder_is_idempotent(store, fake):
    made = store.make_folder("2026", "CLM-3")
    assert (made.name, made.path) == ("CLM-3", "2026/CLM-3") and "Claims/2026/CLM-3" in fake.tree
    assert store.make_folder("2026", "CLM-3").path == "2026/CLM-3"          # 409: already there
    for bad in ("", "a/b", ".."):
        with pytest.raises(PathError):
            store.make_folder("2026", bad)


def test_put_open_delete_small_and_large(store, fake, monkeypatch):
    size, digest = store.put("2026/CLM-1/notes.txt", [b"roof ", b"damage"], max_bytes=100)
    assert size == 11 and len(digest) == 64 and fake.files["Claims/2026/CLM-1/notes.txt"] == b"roof damage"
    with store.open("2026/CLM-1/notes.txt") as f:
        assert f.read() == b"roof damage"
    store.put("new/deeper/a.txt", [b"x"], max_bytes=10)                    # parents are made
    assert "Claims/new/deeper" in fake.tree

    monkeypatch.setattr(drives_module, "SMALL_UPLOAD", 4)                  # force an upload session
    store.put("2026/CLM-2/big.bin", [b"0123456789"], max_bytes=100)
    assert fake.files["Claims/2026/CLM-2/big.bin"] == b"0123456789"

    with pytest.raises(TooLarge):
        store.put("2026/x.bin", [b"x" * 60, b"x" * 60], max_bytes=100)
    assert "Claims/2026/x.bin" not in fake.files
    store.delete("2026/CLM-1/notes.txt")
    store.delete("2026/CLM-1/notes.txt")                                   # gone already: fine
    with pytest.raises(FileNotFoundError):
        store.open("2026/CLM-1/notes.txt")


def test_throttling_waits_then_succeeds_and_errors_are_clear(client, fake):
    fake.throttle_next = 2
    assert client.json("GET", "/drives/d1/root")["id"] == "root"
    assert [c for c in fake.calls if c[1] == "/drives/d1/root"] == [("GET", "/drives/d1/root")] * 3
    with pytest.raises(GraphError) as exc:
        client.json("GET", "/drives/d1/nowhere")
    assert exc.value.status == 404 and exc.value.code == "itemNotFound"


def test_pages_follow_next_links(client, monkeypatch):
    pages = iter([{"value": [1, 2], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"},
                  {"value": [3], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/delta?token=x"}])
    monkeypatch.setattr(client, "json", lambda method, url, **kw: next(pages))
    got = list(client.pages("/me/messages/delta"))
    assert [p["value"] for p in got] == [[1, 2], [3]] and got[-1]["@odata.deltaLink"].endswith("token=x")
