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

    # Default Client ID for LANX
    DEFAULT_CLIENT_ID = "15cd9ff3-25a9-4fb8-9433-6201eef53878"
    DEFAULT_TENANT_ID = "62c4daa3-2df5-40eb-9aa4-a0d5708ee0e7"
    SCOPES = ["Mail.Send", "User.Read"]

    def __init__(self, client_id=None, tenant_id=None):
        self.client_id = client_id or self.DEFAULT_CLIENT_ID
        self.tenant_id = tenant_id or self.DEFAULT_TENANT_ID
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        # Token cache file
        self.cache_file = os.path.join(os.getcwd(), "tmp", ".msal_cache.bin")
        self.token_cache = SerializableTokenCache()

        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, "r") as f:
                    self.token_cache.deserialize(f.read())
            except Exception as e:
                logger.error(f"Erro ao ler cache de tokens: {e}")

        self.app = PublicClientApplication(
            self.client_id, authority=self.authority, token_cache=self.token_cache
        )

    def _save_cache(self):
        if self.token_cache.has_state_changed:
            if not os.path.exists(os.path.dirname(self.cache_file)):
                os.makedirs(os.path.dirname(self.cache_file), exist_ok=True)
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
            return self.login()

        if "access_token" in result:
            return result["access_token"]
        else:
            # If silent fails, try interactive
            return self.login()

    def login(self):
        """
        Executa o login interativo via navegador.
        """
        logger.info("Iniciando login interativo no navegador...")
        # Note: Localhost port will be automatically assigned by MSAL if not specified
        result = self.app.acquire_token_interactive(self.SCOPES)
        self._save_cache()

        if "access_token" in result:
            logger.info("Login realizado com sucesso.")
            return result["access_token"]
        else:
            error_msg = result.get("error_description", "Login cancelado ou falhou.")
            raise Exception(f"Erro no Login: {error_msg}")

    def logout(self):
        """
        Remove as contas do cache (logout).
        """
        accounts = self.app.get_accounts()
        for account in accounts:
            self.app.remove_account(account)
        self._save_cache()
        if os.path.exists(self.cache_file):
            try:
                os.remove(self.cache_file)
            except:
                pass
        logger.info("Logout realizado. Cache de tokens limpo.")


if __name__ == "__main__":
    # Test authentication
    auth = MSGraphAuth()
    try:
        token = auth.get_access_token()
        print("Successfully acquired access token!")
    except Exception as e:
        print(f"Error: {e}")
