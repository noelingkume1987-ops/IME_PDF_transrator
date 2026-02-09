"""
Configuration management for PLC-linked Excel PDF Converter.
=============================================================
Persists settings (folder paths, PLC connection parameters, device
addresses, password) to a JSON file so they survive application restarts.
"""

import json
import os
import hashlib
from dataclasses import dataclass, field, asdict
from pathlib import Path

CONFIG_FILENAME = "plc_pdf_config.json"

# Default password hash (SHA-256 of "0000")
_DEFAULT_PW_HASH = hashlib.sha256("0000".encode()).hexdigest()


def hash_password(plain: str) -> str:
    """Return SHA-256 hex digest of *plain*."""
    return hashlib.sha256(plain.encode()).hexdigest()


@dataclass
class AppConfig:
    """All user-configurable settings with sensible defaults."""

    # --- Folder paths ---
    input_folder: str = ""      # 作業中フォルダ (Input)
    output_folder: str = ""     # PDF出力先フォルダ (Output)
    archive_folder: str = ""    # Excel保存用フォルダ (Archive)

    # --- PLC connection ---
    plc_ip: str = "192.168.1.10"
    plc_port: int = 5000

    # --- PLC device addresses ---
    command_device: str = "D0"
    complete_device: str = "D1"
    monitor_device: str = "D2"

    # --- Timing ---
    plc_poll_interval: float = 0.2   # seconds – D0 polling cycle
    heartbeat_interval: float = 0.5  # seconds – D2.0 toggle cycle

    # --- Security ---
    password_hash: str = _DEFAULT_PW_HASH   # SHA-256 hash of settings password

    # --- Auto-start ---
    auto_start: bool = True  # automatically start monitoring on launch

    # -----------------------------------------------------------------
    # Password helpers
    # -----------------------------------------------------------------

    def verify_password(self, plain: str) -> bool:
        """Return True if *plain* matches the stored password hash."""
        return hash_password(plain) == self.password_hash

    def change_password(self, new_plain: str) -> None:
        """Update the stored password hash."""
        self.password_hash = hash_password(new_plain)

    # -----------------------------------------------------------------
    # Serialisation helpers
    # -----------------------------------------------------------------

    def save(self, path: str | None = None) -> str:
        """Write current settings to *path* (defaults to CWD/config file)."""
        if path is None:
            path = os.path.join(os.getcwd(), CONFIG_FILENAME)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(asdict(self), fh, ensure_ascii=False, indent=2)
        return path

    @classmethod
    def load(cls, path: str | None = None) -> "AppConfig":
        """Load settings from *path*; return defaults if file is missing."""
        if path is None:
            path = os.path.join(os.getcwd(), CONFIG_FILENAME)
        if not os.path.isfile(path):
            return cls()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        # Only pick keys that are valid fields
        valid = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid}
        return cls(**filtered)
