"""Configuration management for xlsliberator."""

import os
from dataclasses import dataclass
from pathlib import Path

from loguru import logger


class ConfigurationError(Exception):
    """Raised when configuration is invalid."""


@dataclass
class Config:
    """Application configuration."""

    # API Keys
    anthropic_api_key: str | None = None

    # Paths
    rules_dir: Path = Path("rules")
    cache_dir: Path = Path(".")

    # Feature flags
    enable_llm: bool = True
    enable_vba_translation: bool = True
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

        # Load API key
        config.anthropic_api_key = os.environ.get("ANTHROPIC_API_KEY")

        # Validate API key if LLM is enabled
        if config.enable_llm and not config.anthropic_api_key:
            logger.warning(
                "ANTHROPIC_API_KEY not set - LLM features (VBA translation, formula repair) "
                "will be disabled or use fallback methods"
            )
            config.enable_llm = False

        if config.anthropic_api_key and len(config.anthropic_api_key) < 10:
            raise ConfigurationError(
                "ANTHROPIC_API_KEY appears to be invalid (too short). Please check your API key."
            )

        # Load optional overrides
        if rules_dir := os.environ.get("XLSLIBERATOR_RULES_DIR"):
            config.rules_dir = Path(rules_dir)

        if cache_dir := os.environ.get("XLSLIBERATOR_CACHE_DIR"):
            config.cache_dir = Path(cache_dir)

        # Validate paths
        if not config.rules_dir.exists():
            logger.warning(f"Rules directory not found: {config.rules_dir}")

        return config

    def validate(self) -> None:
        """Validate configuration.

        Raises:
            ConfigurationError: If configuration is invalid
        """
        if self.enable_llm and not self.anthropic_api_key:
            raise ConfigurationError(
                "LLM is enabled but ANTHROPIC_API_KEY is not set. "
                "Either set the API key or disable LLM features."
            )

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
