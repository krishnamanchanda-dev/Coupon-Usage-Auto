import requests
from config import TOKEN_API_URL, TOKEN_API_KEY, TOKEN_SCOPES


def get_shopify_token() -> str:
    """Fetch a fresh Shopify access token from the background processor."""
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
    return data["data"]["token"]
