"""
Schwab_Auth — OAuth helper and client factory for Schwab Trader API (sandbox or live).

Usage:
- Copy schwab_config.example.json to schwab_config.json at project root.
- Fill in:
    - api_key: your Schwab app key
    - app_secret: your Schwab app secret
    - callback_url: your registered callback URL (e.g. https://127.0.0.1:8182)
    - token_path: where to store OAuth tokens (e.g. schwab_tokens.json)
    - account_id: Schwab account number/hash to trade against
- Run this script once to complete OAuth and cache tokens:
    python scripts/Schwab_Auth.py

Other scripts (e.g. Open_Trades_ToS.py) can import create_client()
to reuse the same credentials and token store.
"""

import json
import os
import asyncio
from typing import Any, Dict, Tuple

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

CONFIG_FILENAME = "schwab_config.json"

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
CONFIG_PATH = os.path.join(_BASE_DIR, CONFIG_FILENAME)

SCHWAB_PY_IMPORT_ERROR = None
try:
    from schwab.auth import easy_client
except Exception as e:  # pragma: no cover - import-time failure
    easy_client = None  # type: ignore[assignment]
    SCHWAB_PY_IMPORT_ERROR = e


def load_config(path: str = CONFIG_PATH) -> Dict[str, Any]:
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Schwab config file not found: {path}\n"
            "Copy schwab_config.example.json to schwab_config.json"
            " and fill in your api_key, app_secret, callback_url, token_path, account_id."
        )
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    required = ["api_key", "app_secret", "callback_url", "token_path", "account_id"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(f"Missing Schwab config keys: {', '.join(missing)} in {path}")
    return cfg


def create_client() -> Tuple[Any, Dict[str, Any]]:
    """
    Create and return a Schwab API client plus the loaded config.

    Returns:
        (client, cfg) where client is the easy_client result, cfg is the config dict.
    """
    if easy_client is None:
        raise ImportError(
            "Could not import schwab-py. Install it with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_PY_IMPORT_ERROR}"
        )
    cfg = load_config()
    client = easy_client(
        api_key=cfg["api_key"],
        app_secret=cfg["app_secret"],
        callback_url=cfg["callback_url"],
        token_path=os.path.join(_BASE_DIR, cfg["token_path"]),
        requested_browser="windows-default",
        interactive=False,
    )
    return client, cfg


def main() -> None:
    """
    Simple sanity check: create a client and print basic account info.
    This will run the OAuth flow in your browser the first time.
    """
    try:
        client, cfg = create_client()
    except Exception as e:
        print(e)
        return

    try:
        resp = client.get_accounts()
        try:
            data = resp.json()
        except Exception:
            data = None
        if isinstance(data, list):
            count = len(data)
        elif data:
            count = 1
        else:
            count = 0
    except Exception as e:
        print(f"Authenticated but failed to fetch accounts: {e}")
        return

    print("Schwab authentication successful.")
    print(f"Loaded {count} account record(s). Configured account_id: {cfg.get('account_id')}")


if __name__ == "__main__":
    main()
