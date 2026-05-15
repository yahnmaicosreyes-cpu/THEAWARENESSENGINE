import os
import sys

# Ensure the project root is importable
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# Prevent the SECRET_KEY warning from firing during tests
os.environ.setdefault("SECRET_KEY", "test-secret-key-for-pytest")

import pytest


@pytest.fixture
def app():
    from app import app as flask_app
    flask_app.config.update({
        "TESTING": True,
        "SECRET_KEY": "test-secret-key-for-pytest",
        "SESSION_COOKIE_SECURE": False,
        "RATELIMIT_ENABLED": False,
    })
    yield flask_app


@pytest.fixture
def client(app):
    return app.test_client()
