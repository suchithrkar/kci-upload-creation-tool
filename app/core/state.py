"""In-memory application state for single-user workflows."""

from __future__ import annotations


class AppState:
    """Singleton container for loaded data sets."""

    _instance: "AppState | None" = None

    def __new__(cls) -> "AppState":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self) -> None:
        if getattr(self, "_initialized", False):
            return
        self.dump = []
        self.wo = []
        self.mo = []
        self.mo_items = []
        self.so = []
        self.cso = []
        self.delivery = []
        self._initialized = True

    def clear(self) -> None:
        """Reset all stored data lists to empty."""

        self.dump = []
        self.wo = []
        self.mo = []
        self.mo_items = []
        self.so = []
        self.cso = []
        self.delivery = []


app_state = AppState()
