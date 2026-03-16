import os
import json
import logging
import webbrowser
from msal import PublicClientApplication, SerializableTokenCache

logger = logging.getLogger(__name__)


class MSGraphAuth:
    """
    Handles authentication with Microsoft Graph API using MSAL.
    Supports interactive browser login and token caching.
    """

    # Default Client ID for LANX (if not provided in config)
    # The user needs to register this in Azure Portal
    # For now, I'll use a common placeholder or ask the user to provide it.
    DEFAULT_CLIENT_ID = "YOUR_CLIENT_ID_HERE"
    DEFAULT_TENANT_ID = "common"
    SCOPES = ["Mail.Send", "User.Read"]

    def __init__(self, client_id=None, tenant_id=None):
        self.client_id = client_id or self.DEFAULT_CLIENT_ID
        self.tenant_id = tenant_id or self.DEFAULT_TENANT_ID
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        # Token cache file
        self.cache_file = os.path.join(os.getcwd(), "tmp", ".msal_cache.bin")
        self.token_cache = SerializableTokenCache()

        if os.path.exists(self.cache_file):
            with open(self.cache_file, "r") as f:
                self.token_cache.deserialize(f.read())

        self.app = PublicClientApplication(
            self.client_id, authority=self.authority, token_cache=self.token_cache
        )

    def _save_cache(self):
        if self.token_cache.has_state_changed:
            with open(self.cache_file, "w") as f:
                f.write(self.token_cache.serialize())

    def get_access_token(self):
        """
        Retrieves an access token. Tries silent flow first, then interactive.
        """
        accounts = self.app.get_accounts()
        result = None

        if accounts:
            # Try to get token silently from cache
            result = self.app.acquire_token_silent(self.SCOPES, account=accounts[0])

        if not result:
            # Fallback to interactive login via browser
            logger.info("No valid token in cache. Opening browser for login...")
            result = self.app.acquire_token_interactive(self.SCOPES)
            self._save_cache()

        if "access_token" in result:
            return result["access_token"]
        else:
            error_msg = result.get(
                "error_description", "Unknown error during authentication."
            )
            raise Exception(f"Failed to acquire token: {error_msg}")


if __name__ == "__main__":
    # Test authentication
    auth = MSGraphAuth()
    try:
        token = auth.get_access_token()
        print("Successfully acquired access token!")
    except Exception as e:
        print(f"Error: {e}")
