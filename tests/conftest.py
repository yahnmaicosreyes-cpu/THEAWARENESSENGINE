import base64
import os
import sys

# Ensure the project root is importable
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# Set test credentials before app is imported so _check_basic_auth() passes.
os.environ.setdefault("SECRET_KEY",          "test-secret-key-for-pytest")
os.environ.setdefault("BASIC_AUTH_USERNAME", "testuser")
os.environ.setdefault("BASIC_AUTH_PASSWORD", "testpass")

import pytest


# ── Auth helper ───────────────────────────────────────────────────────────────

class _AuthClient:
    """Wraps the Flask test client so every request carries valid Basic Auth
    headers automatically. Tests stay clean — no per-request boilerplate."""

    def __init__(self, client, username, password):
        self._c = client
        creds = base64.b64encode(f"{username}:{password}".encode()).decode()
        self._auth_header = {"Authorization": f"Basic {creds}"}

    def _inject(self, kwargs):
        # Merge auth header with any caller-supplied headers without clobbering.
        headers = {**self._auth_header, **kwargs.pop("headers", {})}
        kwargs["headers"] = headers
        return kwargs

    def get(self,    *a, **kw): return self._c.get(   *a, **self._inject(kw))
    def post(self,   *a, **kw): return self._c.post(  *a, **self._inject(kw))
    def put(self,    *a, **kw): return self._c.put(   *a, **self._inject(kw))
    def delete(self, *a, **kw): return self._c.delete(*a, **self._inject(kw))
    # Pass-through for helpers that don't send HTTP requests.
    def get_cookie(self,        *a, **kw): return self._c.get_cookie(*a, **kw)
    def session_transaction(self, *a, **kw): return self._c.session_transaction(*a, **kw)


# ── Fixtures ──────────────────────────────────────────────────────────────────

@pytest.fixture
def app():
    from app import app as flask_app
    flask_app.config.update({
        "TESTING":             True,
        "SECRET_KEY":          "test-secret-key-for-pytest",
        "SESSION_COOKIE_SECURE": False,
        "RATELIMIT_ENABLED":   False,
    })
    yield flask_app


@pytest.fixture
def client(app):
    """Authenticated test client — all requests include valid Basic Auth."""
    return _AuthClient(app.test_client(), "testuser", "testpass")


@pytest.fixture
def unauthed_client(app):
    """Raw test client with no credentials — used to verify 401 behaviour."""
    return app.test_client()
