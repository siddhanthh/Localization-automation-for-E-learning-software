import os
import json
import logging

logger = logging.getLogger("docx_translator.config")

DEFAULT_CONFIG = {
    "translation_backend": "google_translator",
    "gemini": {
        "api_key": "",
        "model": "gemini-1.5-flash",
        "endpoint": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    },
    "rate_limiting": {
        "max_requests_per_minute": 300,
        "max_tokens_per_minute": 40000,
        "max_batch_size": 20,
        "retry_backoff_factor": 2.0,
        "max_retries": 5
    },
    "concurrency": {
        "num_workers": 8
    }
}

class Config:
    def __init__(self, config_path=None):
        self.settings = DEFAULT_CONFIG.copy()
        
        # Load from config path if provided and exists
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    file_settings = json.load(f)
                    self._deep_update(self.settings, file_settings)
                logger.info(f"Loaded configuration from {config_path}")
            except Exception as e:
                logger.warning(f"Failed to load config from {config_path}: {e}. Using defaults.")

        # Environment variable overrides
        env_api_key = os.environ.get("GEMINI_API_KEY")
        if env_api_key:
            self.settings["gemini"]["api_key"] = env_api_key

        env_backend = os.environ.get("TRANSLATION_BACKEND")
        if env_backend:
            self.settings["translation_backend"] = env_backend

    def _deep_update(self, d, u):
        for k, v in u.items():
            if isinstance(v, dict) and k in d and isinstance(d[k], dict):
                self._deep_update(d[k], v)
            else:
                d[k] = v

    @property
    def translation_backend(self):
        return self.settings["translation_backend"]

    @property
    def gemini(self):
        return self.settings["gemini"]

    @property
    def rate_limiting(self):
        return self.settings["rate_limiting"]

    @property
    def concurrency(self):
        return self.settings["concurrency"]
