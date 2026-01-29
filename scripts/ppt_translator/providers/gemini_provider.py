"""Gemini provider implementation."""
from __future__ import annotations

import os

from google import genai

from .base import ProviderConfigurationError, TranslationProvider


class GeminiProvider(TranslationProvider):
    """Translate content using Google's Gemini API."""

    api_key_env = "GOOGLE_API_KEY"

    def __init__(self, model: str, *, api_key: str | None = None, temperature: float = 0.3) -> None:
        super().__init__(model, temperature=temperature)
        resolved_key = api_key or os.getenv(self.api_key_env)
        if not resolved_key:
            raise ProviderConfigurationError(
                "Missing API key for provider 'Gemini'. "
                f"Set the {self.api_key_env} environment variable."
            )
        self.client = genai.Client(api_key=resolved_key)

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        system_prompt = (
            "You are a translation assistant. Translate the user provided text "
            f"from {source_lang} to {target_lang} while preserving tone and formatting."
        )
        response = self.client.models.generate_content(
            model=self.model,
            contents=f"{system_prompt}\n\n{text}",
            config=genai.types.GenerateContentConfig(
                temperature=self.temperature,
            ),
        )
        return response.text.strip()
