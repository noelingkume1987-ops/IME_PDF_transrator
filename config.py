"""
Configuration management for PLC-linked Excel PDF Converter.
=============================================================
Persists settings (folder paths, PLC connection parameters, device
addresses) to a JSON file so they survive application restarts.
"""

import json
import os
from dataclasses import dataclass, field, asdict
from pathlib import Path

CONFIG_FILENAME = "plc_pdf_config.json"


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
    #   command_device  : D register read as 16-bit word (e.g. "D0")
    #   complete_device : D register for completion one-shot (e.g. "D1")
    #   monitor_device  : D register for heartbeat/ready/busy/error (e.g. "D2")
    command_device: str = "D0"
    complete_device: str = "D1"
    monitor_device: str = "D2"

    # --- Timing ---
    plc_poll_interval: float = 0.2   # seconds – D0 polling cycle
    heartbeat_interval: float = 0.5  # seconds – D2.0 toggle cycle

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
