"""
LLM Farm client wrapper.

LLM Farm is an OpenAI-compatible chat completions endpoint (your org's
internal LLM gateway). This module is the ONE place that talks to it, so any
future AI feature in this app (chatbot, summaries, or new features you add)
should route through here rather than instantiating its own client.

Configuration (never hardcode secrets):
  LLM_FARM_API_KEY    - required
  LLM_FARM_BASE_URL   - required, e.g. https://llm-farm.your-org.internal/v1
  LLM_FARM_MODEL      - optional, defaults to DEFAULT_MODEL below

Fill these in via `.env` (local dev) or `.streamlit/secrets.toml`
(Streamlit Cloud / shared deployments). See `.env.example` for the template.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from openai import APIError, APITimeoutError, AuthenticationError, OpenAI

# ---------------------------------------------------------------------------
# Bosch LLM Farm configuration
# ---------------------------------------------------------------------------
DEFAULT_MODEL = "gpt-4o"
DEFAULT_TIMEOUT_SECONDS = 30
API_KEY_HEADER = "genaiplatform-farm-subscription-key"

SYSTEM_PROMPT_CHATBOT = """You are the AI Champion Bot for the Gen AI Champions Initiative dashboard.
Answer questions ONLY using the dashboard data provided below. Be concise and factual.
If the data doesn't contain the answer, say so plainly instead of guessing.
Use simple, plain English — short sentences, no jargon, as if explaining out loud to a colleague.

DASHBOARD DATA:
{context}
"""

SYSTEM_PROMPT_SUMMARY = """You write short, plain-English status updates for a management audience.
Rules:
- Simple, everyday words. No jargon, no corporate buzzwords.
- Short sentences that would sound natural if read out loud.
- One tight paragraph (roughly 80-140 words). No headers, no bullet points.
- Base the summary ONLY on the data given below. Do not invent numbers or facts.

DASHBOARD DATA:
{context}
"""


class LLMFarmError(Exception):
    """Raised when the LLM Farm call fails, with a user-friendly message."""


@dataclass
class LLMFarmConfig:
    api_key: str | None
    base_url: str | None
    model: str

    @property
    def is_configured(self) -> bool:
        return bool(self.api_key and self.base_url)


def load_config(st_secrets: dict | None = None) -> LLMFarmConfig:
    """Read LLM Farm config from environment variables, falling back to st.secrets if provided.

    Call pattern from app.py:
        config = load_config(st.secrets if hasattr(st, "secrets") else None)
    """
    secrets = st_secrets or {}

    def _get(key: str, default: str | None = None) -> str | None:
        return os.environ.get(key) or secrets.get(key) or default

    return LLMFarmConfig(
        api_key=_get("genaiplatform-farm-subscription-key"),
        base_url=_get("openai_base_url"),
        model=_get("openai_model", DEFAULT_MODEL),
    )


def _client(config: LLMFarmConfig) -> OpenAI:
    return OpenAI(
        api_key=config.api_key, 
        base_url=config.base_url 
    )


def ask_chatbot(config: LLMFarmConfig, context: str, question: str, history: list[dict] | None = None) -> str:
    """Answer a natural-language question grounded in the dashboard data.

    `history` is an optional list of {"role": "user"|"assistant", "content": str}
    prior turns, for basic multi-turn context.
    """
    if not config.is_configured:
        raise LLMFarmError(
            "LLM Farm isn't configured yet. Set LLM_FARM_API_KEY and LLM_FARM_BASE_URL "
            "(see .env.example) to enable the chatbot."
        )

    messages = [{"role": "system", "content": SYSTEM_PROMPT_CHATBOT.format(context=context)}]
    messages.extend(history or [])
    messages.append({"role": "user", "content": question})

    try:
        client = _client(config)
        response = client.chat.completions.create(
            model=config.model,
            messages=messages,
            temperature=0.2,
        )
        return response.choices[0].message.content or ""
    except AuthenticationError as exc:
        raise LLMFarmError("LLM Farm rejected the API key. Double-check LLM_FARM_API_KEY.") from exc
    except APITimeoutError as exc:
        raise LLMFarmError("LLM Farm timed out. Please try again.") from exc
    except APIError as exc:
        raise LLMFarmError(f"LLM Farm returned an error: {exc}") from exc


def generate_monthly_summary(config: LLMFarmConfig, context: str) -> str:
    """Draft a short plain-English status paragraph from the current dashboard data."""
    if not config.is_configured:
        raise LLMFarmError(
            "LLM Farm isn't configured yet. Set LLM_FARM_API_KEY and LLM_FARM_BASE_URL "
            "(see .env.example) to enable summary generation."
        )

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT_SUMMARY.format(context=context)},
        {"role": "user", "content": "Draft this month's status update."},
    ]

    try:
        client = _client(config)
        response = client.chat.completions.create(
            model=config.model,
            messages=messages,
            temperature=0.4,
        )
        return response.choices[0].message.content or ""
    except AuthenticationError as exc:
        raise LLMFarmError("LLM Farm rejected the API key. Double-check LLM_FARM_API_KEY.") from exc
    except APITimeoutError as exc:
        raise LLMFarmError("LLM Farm timed out. Please try again.") from exc
    except APIError as exc:
        raise LLMFarmError(f"LLM Farm returned an error: {exc}") from exc
