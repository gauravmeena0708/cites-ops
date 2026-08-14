import os
from pathlib import Path
from typing import Any, Dict, Optional
import yaml

class ConfigLoader:
    """Loads and caches YAML configurations with fallback defaults."""
    
    _CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
    _CACHE: Dict[str, Any] = {}

    @classmethod
    def load_yaml(cls, filename: str, custom_path: Optional[str] = None) -> Dict[str, Any]:
        cache_key = f"{filename}:{custom_path}"
        if cache_key in cls._CACHE:
            return cls._CACHE[cache_key]

        target_path: Path
        if custom_path and os.path.isfile(custom_path):
            target_path = Path(custom_path)
        else:
            target_path = cls._CONFIG_DIR / filename

        if not target_path.exists():
            raise FileNotFoundError(f"Configuration file not found: {target_path}")

        with open(target_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

        cls._CACHE[cache_key] = data
        return data

    @classmethod
    def get_rules(cls, custom_path: Optional[str] = None) -> Dict[str, Any]:
        return cls.load_yaml("rules.yaml", custom_path)

    @classmethod
    def get_entities(cls, custom_path: Optional[str] = None) -> Dict[str, Any]:
        return cls.load_yaml("entities.yaml", custom_path)

    @classmethod
    def get_routing(cls, custom_path: Optional[str] = None) -> Dict[str, Any]:
        return cls.load_yaml("routing.yaml", custom_path)

    @classmethod
    def get_default_config(cls, custom_path: Optional[str] = None) -> Dict[str, Any]:
        return cls.load_yaml("default_config.yaml", custom_path)

    @classmethod
    def clear_cache(cls) -> None:
        cls._CACHE.clear()
