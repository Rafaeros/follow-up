import json
import os
from typing import Any, Dict


class ConfigManager:
    """
    Global system configuration manager.
    Responsible for reading, validating, and saving the JSON configuration file.
    """

    def __init__(self, config_path: str = "configs.json") -> None:
        self.config_path = config_path
        self.is_new_install = False
        self.config = self._initialize_config()

    def _initialize_config(self) -> Dict[str, Any]:
        """
        Validates the existence of the config file.
        If it doesn't exist, creates a default template and flags as a new install.

        Returns:
            Dict[str, Any]: Dictionary containing the configurations.
        """
        default_config = {
            "username": "",
            "password": "",
            "outlook_email": "",
            "outlook_password": "",
            "ms_graph_client_id": "15cd9ff3-25a9-4fb8-9433-6201eef53878",
            "ms_graph_tenant_id": "62c4daa3-2df5-40eb-9aa4-a0d5708ee0e7",
        }

        if not os.path.exists(self.config_path):
            self.is_new_install = True
            self._write_file(default_config)
            return default_config

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            self.is_new_install = True
            self._write_file(default_config)
            return default_config

    def get(self, key: str, default: Any = None) -> Any:
        """Gets a value from the configuration."""
        return self.config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """Sets a value in the configuration and saves to the file."""
        self.config[key] = value
        self.save()

    def save(self) -> None:
        """Persists the current dictionary to the JSON file."""
        self._write_file(self.config)

    def _write_file(self, data: Dict[str, Any]) -> None:
        """Internal utility method to write to the file."""
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def get_session_config(self) -> Dict[str, str]:
        """Retrieves the session credentials."""
        return {
            "username": self.get("username", ""),
            "password": self.get("password", ""),
        }

    def set_session_config(self, session_config: Dict[str, str]) -> None:
        """Updates the session credentials and saves them."""
        self.config["username"] = session_config.get("username", "")
        self.config["password"] = session_config.get("password", "")
        self.save()

    def get_outlook_config(self) -> Dict[str, str]:
        """Retrieves the Outlook SMTP credentials."""
        return {
            "outlook_email": self.get("outlook_email", ""),
            "outlook_password": self.get("outlook_password", ""),
        }

    def set_outlook_config(self, outlook_config: Dict[str, str]) -> None:
        """Updates the Outlook SMTP credentials and saves them."""
        self.config["outlook_email"] = outlook_config.get("outlook_email", "")
        self.config["outlook_password"] = outlook_config.get("outlook_password", "")
        self.save()

    def get_ms_graph_config(self) -> Dict[str, str]:
        """Retrieves the MS Graph API configuration."""
        return {
            "ms_graph_client_id": self.get(
                "ms_graph_client_id", "15cd9ff3-25a9-4fb8-9433-6201eef53878"
            ),
            "ms_graph_tenant_id": self.get(
                "ms_graph_tenant_id", "62c4daa3-2df5-40eb-9aa4-a0d5708ee0e7"
            ),
        }

    def set_ms_graph_config(self, ms_graph_config: Dict[str, str]) -> None:
        """Updates the MS Graph API configuration and saves them."""
        self.config["ms_graph_client_id"] = ms_graph_config.get(
            "ms_graph_client_id", "15cd9ff3-25a9-4fb8-9433-6201eef53878"
        )
        self.config["ms_graph_tenant_id"] = ms_graph_config.get(
            "ms_graph_tenant_id", "62c4daa3-2df5-40eb-9aa4-a0d5708ee0e7"
        )
        self.save()
