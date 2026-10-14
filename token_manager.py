import os
import json
import time
import requests
from config import TOKEN_API_URL, TOKEN_API_KEY, TOKEN_SCOPES

# Cache file lives in /tmp so it's ephemeral and never committed
_CACHE_FILE = "/tmp/shopify_token_cache.json"
# Shopify tokens are valid for ~24h; refresh after 22h to be safe
_TTL_SECONDS = 22 * 3600


def _load_cached_token() -> "str | None":
    try:
        with open(_CACHE_FILE) as f:
            data = json.load(f)
        if time.time() < data.get("expires_at", 0):
            return data["token"]
    except Exception:
        pass
    return None


def _save_token(token: str) -> None:
    try:
        with open(_CACHE_FILE, "w") as f:
            json.dump({"token": token, "expires_at": time.time() + _TTL_SECONDS}, f)
    except Exception:
        pass


def get_shopify_token() -> str:
    """Return a valid Shopify access token, using a cached one if still fresh."""
    cached = _load_cached_token()
    if cached:
        print("  Token loaded from cache.")
        return cached

    response = requests.post(
        TOKEN_API_URL,
        headers={
            "Authorization": TOKEN_API_KEY,
            "Content-Type": "application/json",
        },
        json={"scopes": TOKEN_SCOPES},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    if not data.get("ok"):
        raise RuntimeError(f"Token generation failed: {data}")
    token = data["data"]["token"]
    _save_token(token)
    return token
