"""
PLC communication module for Mitsubishi MC Protocol.
====================================================
Wraps *pymcprotocol* to provide a clean interface for reading/writing
D-register words and individual bits used by the PDF conversion system.

Device mapping (configurable):
    D0 (command)  : 16-bit word – each bit selects a sheet to convert
    D1 (complete) : bit 1 = one-shot completion pulse (100 ms)
    D2 (monitor)  : bit 0 = heartbeat, bit 1 = ready,
                    bit 2 = busy, bit 3 = error
"""

from __future__ import annotations

import logging
import threading
import time

logger = logging.getLogger(__name__)


class PLCConnection:
    """Thread-safe wrapper around a pymcprotocol Type3E connection.

    All public methods acquire ``_lock`` so the heartbeat thread and the
    main polling loop can share one connection safely.
    """

    def __init__(self, ip: str, port: int):
        self.ip = ip
        self.port = port
        self._pymc = None          # pymcprotocol.Type3E instance
        self._lock = threading.Lock()
        self._connected = False

    # -----------------------------------------------------------------
    # Connection lifecycle
    # -----------------------------------------------------------------

    def connect(self) -> None:
        """Open the MC-protocol TCP socket."""
        import pymcprotocol  # late import – not available on non-Windows dev

        with self._lock:
            self._pymc = pymcprotocol.Type3E()
            self._pymc.connect(self.ip, self.port)
            self._connected = True
            logger.info("PLC connected: %s:%d", self.ip, self.port)

    def disconnect(self) -> None:
        with self._lock:
            if self._pymc is not None:
                try:
                    self._pymc.close()
                except Exception:
                    pass
            self._connected = False
            self._pymc = None
            logger.info("PLC disconnected.")

    @property
    def is_connected(self) -> bool:
        return self._connected

    # -----------------------------------------------------------------
    # Word-level read / write  (D registers are 16-bit words)
    # -----------------------------------------------------------------

    def read_word(self, device: str) -> int:
        """Read a single 16-bit word from *device* (e.g. ``"D0"``)."""
        with self._lock:
            values = self._pymc.batchread_wordunits(headdevice=device, readsize=1)
        return values[0] & 0xFFFF

    def write_word(self, device: str, value: int) -> None:
        """Write a single 16-bit word to *device*."""
        with self._lock:
            self._pymc.batchwrite_wordunits(headdevice=device, values=[value & 0xFFFF])

    # -----------------------------------------------------------------
    # Bit-level helpers  (operate on a word via read-modify-write)
    # -----------------------------------------------------------------

    def set_bit(self, device: str, bit: int) -> None:
        """Set *bit* (0-15) inside the word at *device*."""
        with self._lock:
            val = self._pymc.batchread_wordunits(headdevice=device, readsize=1)[0]
            val |= (1 << bit)
            self._pymc.batchwrite_wordunits(headdevice=device, values=[val & 0xFFFF])

    def clear_bit(self, device: str, bit: int) -> None:
        """Clear *bit* (0-15) inside the word at *device*."""
        with self._lock:
            val = self._pymc.batchread_wordunits(headdevice=device, readsize=1)[0]
            val &= ~(1 << bit) & 0xFFFF
            self._pymc.batchwrite_wordunits(headdevice=device, values=[val & 0xFFFF])

    def write_bits(self, device: str, bits_to_set: list[int], bits_to_clear: list[int]) -> None:
        """Atomically set and clear multiple bits in one read-modify-write."""
        with self._lock:
            val = self._pymc.batchread_wordunits(headdevice=device, readsize=1)[0]
            for b in bits_to_set:
                val |= (1 << b)
            for b in bits_to_clear:
                val &= ~(1 << b) & 0xFFFF
            self._pymc.batchwrite_wordunits(headdevice=device, values=[val & 0xFFFF])


# =====================================================================
# Heartbeat helper
# =====================================================================

class HeartbeatThread:
    """Toggles bit 0 of the monitor device at a fixed interval to show
    the PC software is alive.
    """

    def __init__(self, plc: PLCConnection, device: str, interval: float = 0.5):
        self._plc = plc
        self._device = device
        self._interval = interval
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._state = False

    def start(self) -> None:
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=3)
        self._thread = None
        # Ensure heartbeat bit is OFF when stopped
        try:
            self._plc.clear_bit(self._device, 0)
        except Exception:
            pass

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                if self._state:
                    self._plc.clear_bit(self._device, 0)
                else:
                    self._plc.set_bit(self._device, 0)
                self._state = not self._state
            except Exception:
                logger.warning("Heartbeat write failed", exc_info=True)
            self._stop.wait(self._interval)


# =====================================================================
# One-shot pulse helper
# =====================================================================

def send_oneshot(plc: PLCConnection, device: str, bit: int, duration_s: float = 0.1) -> None:
    """Set *bit* ON in *device*, wait *duration_s*, then set it OFF.

    Used for the completion pulse on D1.1.
    """
    plc.set_bit(device, bit)
    time.sleep(duration_s)
    plc.clear_bit(device, bit)
