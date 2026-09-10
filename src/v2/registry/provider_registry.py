from __future__ import annotations

from utils.logger import get_logger


class ProviderRegistry:
    """
    Provider Registry (FS section 2.6.3) — one module per supported AI
    backend, registered by its PROVIDER_NAME. Built at Step 1 (FS 8.1),
    before configuration is even read: unlike the Function Registry, this
    has no dependency on config or the workbook, since provider modules
    ship with the framework's own source — same distribution model as
    built-in validation types (2.6.1), not scanned from a user-configurable
    directory the way custom functions are.
    """

    def __init__(self):
        self.logger = get_logger()
        self._modules: dict[str, object] = {}

    # -----------------------------------------------------------------------
    # Registration
    # -----------------------------------------------------------------------

    def register(self, module) -> None:
        name = module.PROVIDER_NAME.strip().lower()
        if name in self._modules:
            self.logger.warning(
                f"Provider '{name}' already registered — "
                f"keeping existing, skipping duplicate"
            )
            return
        self._modules[name] = module
        self.logger.debug(f"Registered AI provider '{name}'")

    def register_builtins(self) -> None:
        """Register all provider modules the framework ships with."""
        from providers import anthropic

        for module in (anthropic,):
            self.register(module)

        self.logger.info(
            f"AI provider registry ready — {len(self._modules)} provider(s): "
            f"{self.list_names()}"
        )

    # -----------------------------------------------------------------------
    # Lookup
    # -----------------------------------------------------------------------

    def get(self, name: str) -> object | None:
        return self._modules.get(name.strip().lower())

    def has(self, name: str) -> bool:
        return name.strip().lower() in self._modules

    def get_required_settings(self, name: str) -> list[str]:
        module = self.get(name)
        return list(getattr(module, "REQUIRED_SETTINGS", None) or [])

    def list_names(self) -> list[str]:
        return list(self._modules.keys())
