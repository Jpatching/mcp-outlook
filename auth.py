import os
import sys
import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any, List
import msal

logger = logging.getLogger("mcp-outlook.auth")

# Microsoft First-Party Pre-Approved Client ID (Pre-authorized across Microsoft 365 tenants)
DEFAULT_CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"

# Note: Do NOT include 'offline_access', 'openid', or 'profile' - MSAL adds them automatically.
DEFAULT_SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Calendars.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]

def get_app_dir() -> Path:
    """Returns cross-platform directory for config and token cache."""
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    
    app_dir = base / "mcp-outlook"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir

def get_config_path() -> Path:
    return get_app_dir() / "config.json"

def get_cache_path() -> Path:
    return get_app_dir() / "token_cache.bin"

def load_config() -> Dict[str, Any]:
    config_file = get_config_path()
    cfg = {}
    if config_file.exists():
        try:
            cfg = json.loads(config_file.read_text(encoding="utf-8"))
        except Exception as e:
            logger.warning(f"Failed to read config file: {e}")
            
    client_id = os.environ.get("AZURE_CLIENT_ID") or cfg.get("client_id") or DEFAULT_CLIENT_ID
    tenant_id = os.environ.get("AZURE_TENANT_ID") or cfg.get("tenant_id") or "common"
    
    # Filter out reserved scopes if previously saved
    raw_scopes = cfg.get("scopes", DEFAULT_SCOPES)
    clean_scopes = [s for s in raw_scopes if s not in ("offline_access", "openid", "profile")]
    
    return {
        "client_id": client_id.strip() or DEFAULT_CLIENT_ID,
        "tenant_id": tenant_id.strip() or "common",
        "scopes": clean_scopes or DEFAULT_SCOPES
    }

def save_config(client_id: str, tenant_id: str = "common"):
    config_file = get_config_path()
    data = {
        "client_id": client_id.strip() or DEFAULT_CLIENT_ID,
        "tenant_id": tenant_id.strip() or "common",
        "scopes": DEFAULT_SCOPES
    }
    config_file.write_text(json.dumps(data, indent=2), encoding="utf-8")

class OutlookAuth:
    def __init__(self, client_id: Optional[str] = None, tenant_id: Optional[str] = None):
        cfg = load_config()
        self.client_id = (client_id or cfg.get("client_id") or DEFAULT_CLIENT_ID).strip()
        self.tenant_id = (tenant_id or cfg.get("tenant_id") or "common").strip()
        self.scopes = [s for s in cfg.get("scopes", DEFAULT_SCOPES) if s not in ("offline_access", "openid", "profile")]
        self.cache_path = get_cache_path()
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app: Optional[msal.PublicClientApplication] = None

    def _load_cache(self):
        if self.cache_path.exists():
            try:
                self._cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
            except Exception as e:
                logger.warning(f"Error loading token cache: {e}")

    def _save_cache(self):
        if self._cache.has_state_changed:
            try:
                self.cache_path.write_text(self._cache.serialize(), encoding="utf-8")
            except Exception as e:
                logger.warning(f"Error saving token cache: {e}")

    def get_msal_app(self) -> msal.PublicClientApplication:
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self._app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=authority,
            token_cache=self._cache
        )
        return self._app

    def get_access_token(self) -> str:
        """Acquires a valid access token, using cache and refreshing if necessary."""
        app = self.get_msal_app()
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(self.scopes, account=accounts[0])
            self._save_cache()
            
        if result and "access_token" in result:
            return result["access_token"]

        error_msg = result.get("error_description") if result else "No accounts found in cache."
        raise PermissionError(
            f"Authentication required. Please run `./run.sh --login` (Linux) or `run.bat --login` (Windows) to sign in.\nDetails: {error_msg}"
        )

    def login_device_flow(self) -> Dict[str, Any]:
        """Runs the interactive Device Code login flow using Microsoft's pre-approved client ID."""
        app = self.get_msal_app()
        flow = app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to create device flow: {flow.get('error_description', flow)}")

        print("\n" + "=" * 60)
        print("MICROSOFT 365 WORK OUTLOOK LOGIN")
        print("=" * 60)
        print(flow["message"])
        print("=" * 60 + "\n")
        sys.stdout.flush()

        result = app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            self._save_cache()
            print("Authentication successful! Token saved locally.")
            return result
        else:
            raise RuntimeError(f"Authentication failed: {result.get('error_description', result)}")

    def logout(self):
        """Clears cached tokens."""
        if self.cache_path.exists():
            self.cache_path.unlink(missing_ok=True)
        self._cache = msal.SerializableTokenCache()
        print("Logged out successfully. Token cache cleared.")

    def get_status(self) -> Dict[str, Any]:
        """Checks authentication status without triggering interactive login."""
        try:
            app = self.get_msal_app()
            accounts = app.get_accounts()
            if not accounts:
                return {
                    "authenticated": False,
                    "error": "No accounts logged in yet",
                    "client_id": self.client_id,
                    "tenant_id": self.tenant_id
                }
            
            result = app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                return {
                    "authenticated": True,
                    "account": accounts[0].get("username", "Unknown"),
                    "tenant_id": self.tenant_id,
                    "client_id": self.client_id
                }
            return {
                "authenticated": False,
                "account": accounts[0].get("username", "Unknown"),
                "error": "Token expired, refresh required"
            }
        except Exception as e:
            return {"authenticated": False, "error": str(e)}
