"""Microsoft Graph client: app-only sign-in, retries and throttling, paging, delta queries,
streamed downloads and chunked uploads.

    graph = GraphClient.from_env()                      # TENANT_ID, CLIENT_ID, CLIENT_SECRET
    site = graph.json("GET", "/sites/contoso.sharepoint.com:/sites/claims")
    for page in graph.pages("/users/{id}/mailFolders/inbox/messages/delta"): …

When Graph throttles (429, or 502/503/504), every thread using the client waits for the
Retry-After period before its next call, so one busy worker does not get the others blocked.
"""

from __future__ import annotations

import base64
import json
import logging
import os
import threading
import time
from collections.abc import Callable, Iterator
from urllib.parse import quote

import msal
import requests
from requests.adapters import HTTPAdapter

logger = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]
RETRY_STATUS = {429, 502, 503, 504}
CHUNK = 4 * 1024 * 1024            # upload sessions take multiples of 320 KiB; 4 MiB is one


class GraphError(Exception):
    """A Graph call failed. The message is safe to show: it never holds a token or secret."""

    def __init__(self, status: int | None, code: str, message: str):
        self.status = status
        self.code = code
        super().__init__(f"{status} {code}: {message}" if status else f"{code}: {message}")


def quote_path(path: str) -> str:
    return quote(path, safe="/")


class GraphClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, authority: str | None = None,
                 *, timeout: float = 60, max_retries: int = 5, pool_size: int = 32,
                 session: requests.Session | None = None,
                 token_provider: Callable[[], str] | None = None):
        """token_provider replaces MSAL (tests); session replaces the HTTP session (tests)."""
        self._app = None if token_provider else msal.ConfidentialClientApplication(
            client_id, authority=authority or f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret)
        self._token_provider = token_provider
        if session is None:
            session = requests.Session()
            adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
            session.mount("https://", adapter)
        self._session = session
        self._timeout = timeout
        self._max_retries = max_retries
        self._pause_until = 0.0
        self._lock = threading.Lock()

    @classmethod
    def from_env(cls, env_file: str | None = ".env", **kwargs) -> GraphClient:
        """A client from TENANT_ID, CLIENT_ID, CLIENT_SECRET (and optional AUTHORITY), read from
        the environment, after loading env_file if python-dotenv is installed."""
        if env_file:
            try:
                from dotenv import load_dotenv
                load_dotenv(env_file)
            except ImportError:
                pass
        missing = [k for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET") if not os.getenv(k)]
        if missing:
            raise GraphError(None, "Config", f"Missing {', '.join(missing)} (environment or {env_file})")
        return cls(os.environ["TENANT_ID"], os.environ["CLIENT_ID"], os.environ["CLIENT_SECRET"],
                   os.getenv("AUTHORITY") or None, **kwargs)

    # ── Sign-in ──────────────────────────────────────────────────────────────

    def token(self) -> str:
        if self._token_provider:
            return self._token_provider()
        # MSAL serves app-only tokens from its cache until shortly before they expire.
        result = self._app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            raise GraphError(None, result.get("error", "AuthFailed"),
                             result.get("error_description", "Could not sign in"))
        return result["access_token"]

    def roles(self) -> list[str]:
        """The application permissions the token carries, e.g. ['Sites.Read.All']."""
        payload = self.token().split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload)).get("roles", [])

    # ── Transport ────────────────────────────────────────────────────────────

    def send(self, method: str, url: str, *, auth: bool = True, headers: dict | None = None,
             **kwargs) -> requests.Response:
        """One request with retries: network errors and throttling are retried with backoff."""
        if not url.startswith("http"):
            url = GRAPH + url
        for attempt in range(self._max_retries + 1):
            with self._lock:
                delay = self._pause_until - time.monotonic()
            if delay > 0:
                time.sleep(delay)
            sent = dict(headers or {})
            if auth:
                sent["Authorization"] = f"Bearer {self.token()}"
            try:
                resp = self._session.request(method, url, headers=sent, timeout=self._timeout, **kwargs)
            except (requests.ConnectionError, requests.Timeout) as exc:
                if attempt == self._max_retries:
                    raise GraphError(None, "Network", str(exc)) from exc
                time.sleep(2 ** attempt)
                continue
            if resp.status_code in RETRY_STATUS and attempt < self._max_retries:
                wait = float(resp.headers.get("Retry-After") or 0) or 2 ** attempt
                logger.warning("Graph %s on %s %s; every caller waits %ss", resp.status_code, method,
                               url.split("?")[0], wait)
                with self._lock:
                    self._pause_until = max(self._pause_until, time.monotonic() + wait)
                continue
            return resp
        raise AssertionError("unreachable")

    @staticmethod
    def error(resp: requests.Response) -> GraphError:
        try:
            err = resp.json()["error"]
            return GraphError(resp.status_code, err.get("code", "Error"), err.get("message", resp.reason))
        except (ValueError, KeyError, TypeError):
            return GraphError(resp.status_code, "Error", resp.reason or resp.text[:200])

    def json(self, method: str, url: str, *, none_on_404: bool = False, **kwargs) -> dict | None:
        resp = self.send(method, url, **kwargs)
        if resp.status_code == 404 and none_on_404:
            return None
        if not resp.ok:
            raise self.error(resp)
        return resp.json() if resp.content else {}

    def pages(self, url: str, **kwargs) -> Iterator[dict]:
        """Every page of a collection, following @odata.nextLink. The last page of a delta
        query carries @odata.deltaLink: keep it, and pass it here next time."""
        while url:
            page = self.json("GET", url, **kwargs)
            yield page
            url = page.get("@odata.nextLink")

    def all(self, url: str, **kwargs) -> list[dict]:
        return [item for page in self.pages(url, **kwargs) for item in page.get("value", [])]

    def stream(self, url: str, chunk_size: int = 1024 * 1024) -> Iterator[bytes]:
        """A download as chunks, never held whole in memory (a message's /$value, a file's
        /content)."""
        resp = self.send("GET", url, stream=True)
        if not resp.ok:
            raise self.error(resp)
        try:
            yield from resp.iter_content(chunk_size)
        finally:
            resp.close()

    def upload(self, session_url: str, chunks: Iterator[bytes], size: int) -> dict:
        """Upload `size` bytes to an upload session (from createUploadSession), in 4 MiB pieces.
        The session URL is pre-authorised: no bearer token is sent to it."""
        sent, result = 0, {}
        for piece in _rechunk(chunks, CHUNK):
            end = sent + len(piece) - 1
            resp = self.send("PUT", session_url, auth=False, data=piece,
                             headers={"Content-Length": str(len(piece)),
                                      "Content-Range": f"bytes {sent}-{end}/{size}"})
            if not resp.ok:
                raise self.error(resp)
            sent = end + 1
            result = resp.json() if resp.content else {}
        return result


def _rechunk(chunks: Iterator[bytes], size: int) -> Iterator[bytes]:
    """Chunks regrouped into pieces of exactly `size` bytes (the last may be shorter)."""
    pending = b""
    for chunk in chunks:
        pending += chunk
        while len(pending) >= size:
            yield pending[:size]
            pending = pending[size:]
    if pending:
        yield pending
