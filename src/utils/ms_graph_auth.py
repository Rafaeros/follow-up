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
        raw_tenant = tenant_id or self.DEFAULT_TENANT_ID

        # Sanitize tenant_id (prevents "co[UUID]mmon" errors if user pasted over "common")
        if "common" in raw_tenant and len(raw_tenant) > 10:
            # If it looks like a mix of common and a UUID
            import re

            uuids = re.findall(
                r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
                raw_tenant,
            )
            if uuids:
                self.tenant_id = uuids[0]
            else:
                self.tenant_id = "common"
        else:
            self.tenant_id = raw_tenant

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
        if self.client_id == "YOUR_CLIENT_ID_HERE":
            raise Exception("MS Graph Client ID não configurado. Vá em Configurações.")

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
            error_msg = result.get(
                "error_description", "Erro desconhecido durante a autenticação."
            )
            raise Exception(f"Falha ao adquirir token: {error_msg}")

    def login(self):
        """
        Executa o login interativo via navegador.
        """
        if self.client_id == "YOUR_CLIENT_ID_HERE":
            raise Exception("MS Graph Client ID não configurado.")

        logger.info("Iniciando login interativo no navegador...")
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
        logger.info("Logout realizado. Cache de tokens limpo.")


if __name__ == "__main__":
    # Test authentication
    auth = MSGraphAuth()
    try:
        token = auth.get_access_token()
        print("Successfully acquired access token!")
    except Exception as e:
        print(f"Error: {e}")
