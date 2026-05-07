"""Pexels image fetcher with local disk cache."""
import hashlib
import os
from pathlib import Path

_CACHE = Path(__file__).resolve().parent.parent.parent.parent.parent / "projects" / ".img_cache"
_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"


def fetch_pexels(query: str, api_key: str = "") -> str | None:
    """Search Pexels and return a local path to a downloaded landscape photo.

    Returns None if key is missing, query returns nothing, or network fails.
    Results are cached by query so repeated renders skip the network request.
    """
    key = api_key or os.environ.get("PEXELS_API_KEY", "")
    if not key:
        return None

    try:
        import requests
    except ImportError:
        return None

    _CACHE.mkdir(parents=True, exist_ok=True)
    slug = hashlib.md5(query.lower().encode()).hexdigest()[:12]
    cached = _CACHE / f"{slug}.jpg"
    if cached.exists():
        return str(cached)

    try:
        headers = {"Authorization": key, "User-Agent": _UA}
        resp = requests.get(
            "https://api.pexels.com/v1/search",
            params={"query": query, "orientation": "landscape", "per_page": 5},
            headers=headers, timeout=10
        )
        resp.raise_for_status()
        photos = resp.json().get("photos", [])
        if not photos:
            return None
        img_url = photos[0]["src"]["large2x"]
        img_resp = requests.get(img_url, headers={"User-Agent": _UA}, timeout=20, stream=True)
        img_resp.raise_for_status()
        with open(cached, "wb") as f:
            for chunk in img_resp.iter_content(65536):
                f.write(chunk)
        return str(cached)
    except Exception:
        return None
