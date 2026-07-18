"""Provider-neutral deterministic configuration management."""

import os
from dataclasses import dataclass
from pathlib import Path


class ConfigurationError(Exception):
    """Raised when configuration is invalid."""


@dataclass
class Config:
    """Application configuration."""

    # Paths
    rules_dir: Path = Path("rules")
    cache_dir: Path = Path(".")

    # Deterministic feature flags
    enable_formula_repair: bool = True

    # Performance
    max_retries: int = 3
    timeout_seconds: int = 300

    @classmethod
    def from_env(cls) -> "Config":
        """Load configuration from environment variables.

        Returns:
            Config instance

        Raises:
            ConfigurationError: If required configuration is missing or invalid
        """
        config = cls()

        # Load optional overrides
        if rules_dir := os.environ.get("XLSLIBERATOR_RULES_DIR"):
            config.rules_dir = Path(rules_dir)

        if cache_dir := os.environ.get("XLSLIBERATOR_CACHE_DIR"):
            config.cache_dir = Path(cache_dir)

        return config

    def validate(self) -> None:
        """Validate configuration.

        Raises:
            ConfigurationError: If configuration is invalid
        """
        if self.max_retries < 0:
            raise ConfigurationError("max_retries must be >= 0")

        if self.timeout_seconds <= 0:
            raise ConfigurationError("timeout_seconds must be > 0")


# Global config instance
_config: Config | None = None


def get_config() -> Config:
    """Get global configuration instance.

    Returns:
        Config instance
    """
    global _config
    if _config is None:
        _config = Config.from_env()
    return _config


def reset_config() -> None:
    """Reset global configuration (mainly for testing)."""
    global _config
    _config = None
