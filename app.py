"""
Fallacy Finder: Fallacy detecting application for articles online or pasted text
by Jacques Laroche

Key technologies:
- Flask, Python, HTML, CSS, Javascript

This file is intended to be:
- run directly with: `python app.py` (dev)
- imported by a WSGI server

Important operational notes:
- For fallacy detecting functionality you must use one of two AI providers: either Ollama or OpenAI API

"""

import json
import os
import re
import hashlib
import secrets
import sqlite3
import uuid
import threading
import time
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from collections import Counter, defaultdict
from typing import Any

import requests
import trafilatura
from flask import Flask, render_template, request, make_response, jsonify, redirect, url_for, abort
from markupsafe import Markup, escape
from werkzeug.middleware.proxy_fix import ProxyFix

BASE_DIR = Path(__file__).resolve().parent
APP_VERSION = str(os.getenv("APP_VERSION") or "0.5.3.7").strip()
APP_AUTHOR_NAME = "Jacques Laroche"
APP_AUTHOR_URL = "https://www.digitalcuriosity.center/about-us/#"
APP_GITHUB_URL = "https://github.com/jlar0che/FallacyFinder"


def _instance_dir() -> Path:
    configured = str(os.getenv("INSTANCE_DIR") or "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    return BASE_DIR / "instance"


INSTANCE_DIR = _instance_dir()

app = Flask(__name__, instance_path=str(INSTANCE_DIR), instance_relative_config=True)
_APP_CONFIGURED = False

# -----------------------------
# UI localization
# -----------------------------
SUPPORTED_UI_LANGUAGES = ("en", "sr")
AVAILABLE_UI_LANGUAGES = (
    {"code": "en", "native_name": "English"},
    {"code": "sr", "native_name": "Srpski"},
)
_TRANSLATIONS_CACHE: dict[str, dict[str, str]] = {}
_FALLACY_LIBRARY_CACHE: dict[str, dict[str, Any]] = {}


def _env_bool(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return bool(default)
    return str(raw).strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except Exception:
        return int(default)


def _env_secret_key() -> str:
    configured = str(os.getenv("SECRET_KEY") or os.getenv("FLASK_SECRET_KEY") or "").strip()
    if configured:
        return configured
    if _env_bool("DEBUG", False) or _env_bool("FLASK_DEBUG", False):
        return "change-me"
    return secrets.token_hex(32)


def _translations_dir() -> str:
    configured = str(os.getenv("TRANSLATIONS_DIR") or "").strip()
    if configured:
        return str(Path(configured).expanduser().resolve())
    return str(BASE_DIR / "translations")


def _sanitize_ui_language(value: Any, default: str = "en") -> str:
    code = str(value or default).strip().lower()
    return code if code in SUPPORTED_UI_LANGUAGES else default


def _load_translation_catalog(locale: str) -> dict[str, str]:
    locale = _sanitize_ui_language(locale)
    cached = _TRANSLATIONS_CACHE.get(locale)
    if cached is not None:
        return cached
    data: dict[str, str] = {}
    path = os.path.join(_translations_dir(), f"{locale}.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if isinstance(loaded, dict):
            data = {str(k): str(v) for k, v in loaded.items()}
    except Exception:
        data = {}
    _TRANSLATIONS_CACHE[locale] = data
    return data


def _current_locale(settings: dict[str, Any] | None = None) -> str:
    if settings is None:
        try:
            settings = load_settings()
        except Exception:
            settings = {}
    return _sanitize_ui_language((settings or {}).get("UI_LANGUAGE"), default="en")


def _translate(key: str, locale: str | None = None, **kwargs: Any) -> str:
    key = str(key or "")
    active_locale = _sanitize_ui_language(locale or _current_locale(), default="en")
    value = _load_translation_catalog(active_locale).get(key)
    if value is None and active_locale != "en":
        value = _load_translation_catalog("en").get(key)
    if value is None:
        value = key
    if kwargs:
        try:
            value = value.format(**kwargs)
        except Exception:
            pass
    return value


def _translate_for_settings(key: str, settings: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _translate(key, locale=_current_locale(settings), **kwargs)


def _ui_language_name(locale: str | None = None) -> str:
    code = _sanitize_ui_language(locale or "en")
    english_names = {
        "en": "English",
        "sr": "Serbian",
    }
    if code in english_names:
        return english_names[code]
    for item in AVAILABLE_UI_LANGUAGES:
        if str(item.get("code") or "").strip().lower() == code:
            return str(item.get("english_name") or item.get("native_name") or code)
    return code


def _safe_next_url(target: str | None) -> str:
    candidate = (target or "").strip()
    if not candidate:
        return url_for("home")
    parsed = urlparse(candidate)
    if parsed.scheme or parsed.netloc:
        return url_for("home")
    if not candidate.startswith("/"):
        candidate = "/" + candidate
    return candidate


def language_switch_url(code: str) -> str:
    next_url = request.full_path if request else "/"
    next_url = (next_url or "/").rstrip("?") or "/"
    return url_for("set_ui_language", code=_sanitize_ui_language(code), next=next_url)


@app.context_processor
def inject_i18n_helpers():
    settings = {}
    try:
        settings = load_settings()
    except Exception:
        settings = {}
    locale = _current_locale(settings)

    def t(key: str, **kwargs: Any) -> str:
        return _translate(key, locale=locale, **kwargs)

    def localized_category(value: Any) -> str:
        normalized = str(value or "").strip().lower()
        if normalized == "formal":
            return t("common.formal")
        if normalized == "informal":
            return t("common.informal")
        return str(value or "").strip()

    def localized_relation_label(label: Any) -> str:
        raw_label = str(label or "").strip()
        direct_map = {
            "Logically close": "related.reason.logically_close",
            "Often paired": "related.reason.often_paired",
            "Same family": "related.reason.same_family",
            "Often confused": "related.reason.often_confused",
            "Similar wording": "related.reason.similar_wording",
            "Same category": "related.reason.same_category",
        }
        mapped_key = direct_map.get(raw_label)
        if mapped_key:
            return t(mapped_key)

        match = re.match(r"^Seen together in saved analyses \((\d+)\)$", raw_label)
        if match:
            return t("related.reason.seen_together", count=match.group(1))

        return raw_label

    def _localized_fallacy_payload(value: Any) -> dict[str, Any]:
        return _localized_fallacy_reference_for(value, locale=locale)

    def localized_fallacy_name(value: Any) -> str:
        payload = _localized_fallacy_payload(value)
        return str(payload.get("display_name") or payload.get("canonical_name") or payload.get("name") or payload.get("type") or value or "")

    def localized_fallacy_description(value: Any) -> str:
        payload = _localized_fallacy_payload(value)
        return str(payload.get("description") or "")

    def localized_fallacy_short_for(value: Any) -> str:
        payload = _localized_fallacy_payload(value)
        return str(payload.get("short_for") or "")

    def localized_fallacy_aliases(value: Any) -> list[str]:
        payload = _localized_fallacy_payload(value)
        return [str(item) for item in (payload.get("aliases") or []) if str(item or "").strip()]

    def localized_fallacy_keywords(value: Any) -> list[str]:
        payload = _localized_fallacy_payload(value)
        return [str(item) for item in (payload.get("keywords") or []) if str(item or "").strip()]

    def localized_fallacy_examples(value: Any) -> list[str]:
        payload = _localized_fallacy_payload(value)
        return [str(item) for item in (payload.get("examples") or []) if str(item or "").strip()]

    def localized_fallacy_explanation(value: Any) -> str:
        payload = _localized_fallacy_payload(value)
        return str(payload.get("explanation") or "")

    return {
        "t": t,
        "current_locale": locale,
        "available_ui_languages": AVAILABLE_UI_LANGUAGES,
        "language_switch_url": language_switch_url,
        "localized_category": localized_category,
        "localized_relation_label": localized_relation_label,
        "localized_fallacy_name": localized_fallacy_name,
        "localized_fallacy_description": localized_fallacy_description,
        "localized_fallacy_short_for": localized_fallacy_short_for,
        "localized_fallacy_aliases": localized_fallacy_aliases,
        "localized_fallacy_keywords": localized_fallacy_keywords,
        "localized_fallacy_examples": localized_fallacy_examples,
        "localized_fallacy_explanation": localized_fallacy_explanation,
        "app_version": APP_VERSION,
        "app_author_name": APP_AUTHOR_NAME,
        "app_author_url": APP_AUTHOR_URL,
        "app_github_url": APP_GITHUB_URL,
    }


@app.get("/language/<code>")
def set_ui_language(code: str):
    settings = load_settings()
    settings["UI_LANGUAGE"] = _sanitize_ui_language(code, default=settings.get("UI_LANGUAGE") or "en")
    save_settings(settings)
    return redirect(_safe_next_url(request.args.get("next")))

# -----------------------------
# In-memory analysis jobs
# -----------------------------
ANALYSIS_JOBS: dict[str, dict[str, Any]] = {}
ANALYSIS_JOBS_LOCK = threading.Lock()
ANALYSIS_JOB_TTL_SECONDS = int(os.getenv("ANALYSIS_JOB_TTL_SECONDS", "3600"))


def _format_elapsed_seconds(seconds: float | int | None) -> str:
    try:
        total_seconds = max(0, int(float(seconds or 0)))
    except Exception:
        total_seconds = 0
    hours, remainder = divmod(total_seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def _job_elapsed_seconds(job: dict[str, Any]) -> float:
    started_at = float(job.get("started_at") or job.get("created_at") or time.time())
    finished_at = job.get("finished_at")
    end_time = float(finished_at) if finished_at is not None else time.time()
    return max(0.0, end_time - started_at)


class AnalysisCancelled(Exception):
    pass


class ProviderConnectionError(Exception):
    pass


def _cleanup_jobs() -> None:
    cutoff = time.time() - ANALYSIS_JOB_TTL_SECONDS
    with ANALYSIS_JOBS_LOCK:
        stale_ids = [
            job_id for job_id, job in ANALYSIS_JOBS.items()
            if job.get("updated_at", 0) < cutoff and job.get("status") in {"done", "error", "cancelled"}
        ]
        for job_id in stale_ids:
            ANALYSIS_JOBS.pop(job_id, None)


def _create_job(url: str, *, source_kind: str = "url") -> str:
    _cleanup_jobs()
    settings = load_settings()
    job_id = str(uuid.uuid4())
    now = time.time()
    with ANALYSIS_JOBS_LOCK:
        ANALYSIS_JOBS[job_id] = {
            "job_id": job_id,
            "source_kind": source_kind,
            "source_label": "Pasted text" if source_kind == "pasted_text" else "Article URL",
            "url": url,
            "input_text": None,
            "phase": "extract" if source_kind == "url" else "analysis",
            "status": "queued",
            "stage": _translate_for_settings("job.queued", settings),
            "progress": 0,
            "error": None,
            "error_kind": None,
            "result_html": None,
            "title": None,
            "author": None,
            "date": None,
            "paragraphs": [],
            "extraction_stats": None,
            "info_message": None,
            "created_at": now,
            "started_at": now,
            "finished_at": None,
            "updated_at": now,
            "paragraph_current": 0,
            "paragraph_total": 0,
            "paragraphs_total": 0,
            "paragraphs_analyzed": 0,
            "paragraphs_skipped_short": 0,
            "paragraphs_timed_out": 0,
            "paragraphs_failed": 0,
            "cancel_requested": False,
            "cancel_message": None,
            "active_provider": None,
            "active_model": None,
            "active_response": None,
            "active_session": None,
            "history_overwrite_prompt": None,
            "pending_history_overwrite": None,
            "latest_findings": None,
            "latest_analysis_stats": None,
        }
    return job_id


def _update_job(job_id: str, **changes: Any) -> None:
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        if not job:
            return
        job.update(changes)
        job["updated_at"] = time.time()


def _get_job(job_id: str) -> dict[str, Any] | None:
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        return dict(job) if job else None


def _job_cancel_requested(job_id: str) -> bool:
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        return bool(job and job.get("cancel_requested"))


def _set_job_stream_handle(job_id: str, *, provider: str | None = None, model: str | None = None, response=None, session=None) -> None:
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        if not job:
            return
        if provider is not None:
            job["active_provider"] = provider
        if model is not None:
            job["active_model"] = model
        job["active_response"] = response
        job["active_session"] = session
        job["updated_at"] = time.time()


def _clear_job_stream_handle(job_id: str) -> None:
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        if not job:
            return
        job["active_response"] = None
        job["active_session"] = None
        job["updated_at"] = time.time()


def _abort_job_stream(job_id: str) -> tuple[bool, str | None, str | None]:
    response = None
    session = None
    provider = None
    model = None
    with ANALYSIS_JOBS_LOCK:
        job = ANALYSIS_JOBS.get(job_id)
        if not job:
            return False, None, None
        response = job.get("active_response")
        session = job.get("active_session")
        provider = job.get("active_provider")
        model = job.get("active_model")
        job["active_response"] = None
        job["active_session"] = None
        job["updated_at"] = time.time()
    aborted = False
    try:
        if response is not None:
            response.close()
            aborted = True
    except Exception:
        pass
    try:
        if session is not None:
            session.close()
            aborted = True or aborted
    except Exception:
        pass
    return aborted, provider, model


def ollama_unload_model(settings: dict[str, Any], model: str, timeout_seconds: float = 10.0) -> None:
    if not model:
        return
    url = f"{settings['OLLAMA_BASE_URL']}/api/generate"
    payload = {
        "model": model,
        "prompt": "",
        "stream": False,
        "keep_alive": 0,
    }
    try:
        requests.post(url, json=payload, timeout=(5, timeout_seconds))
    except Exception:
        pass


def _job_status_payload(job: dict[str, Any]) -> dict[str, Any]:
    elapsed_seconds = _job_elapsed_seconds(job)
    return {
        "ok": True,
        "job_id": job["job_id"],
        "source_kind": job.get("source_kind") or "url",
        "phase": job.get("phase"),
        "status": job.get("status"),
        "stage": job.get("stage"),
        "progress": job.get("progress", 0),
        "error": job.get("error"),
        "error_kind": job.get("error_kind"),
        "info_message": job.get("info_message"),
        "has_extraction": bool(job.get("paragraphs")),
        "started_at": job.get("started_at") or job.get("created_at"),
        "finished_at": job.get("finished_at"),
        "elapsed_seconds": elapsed_seconds,
        "elapsed_display": _format_elapsed_seconds(elapsed_seconds),
        "paragraph_current": job.get("paragraph_current", 0),
        "paragraph_total": job.get("paragraph_total", 0),
        "paragraphs_total": job.get("paragraphs_total", 0),
        "paragraphs_analyzed": job.get("paragraphs_analyzed", 0),
        "paragraphs_skipped_short": job.get("paragraphs_skipped_short", 0),
        "paragraphs_timed_out": job.get("paragraphs_timed_out", 0),
        "paragraphs_failed": job.get("paragraphs_failed", 0),
        "cancel_requested": bool(job.get("cancel_requested")),
        "cancel_message": job.get("cancel_message"),
    }


# -----------------------------
# Persisted settings
# -----------------------------
def ensure_dirs() -> None:
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    os.makedirs(_translations_dir(), exist_ok=True)
    init_history_db()
    settings_path = _settings_path()
    if not os.path.exists(settings_path):
        save_settings(dict(DEFAULT_SETTINGS))


DEFAULT_SETTINGS = {
    "AI_PROVIDER": os.getenv("AI_PROVIDER", "ollama"),
    "OLLAMA_BASE_URL": os.getenv("OLLAMA_BASE_URL", "http://localhost:11434").rstrip("/"),
    "OLLAMA_MODEL": os.getenv("OLLAMA_MODEL", "phi4:14b"),
    "OPENAI_BASE_URL": os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/"),
    "OPENAI_MODEL": os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
    "OLLAMA_TIMEOUT": float(os.getenv("OLLAMA_TIMEOUT", "1000")),
    "OLLAMA_TEMPERATURE": float(os.getenv("OLLAMA_TEMPERATURE", "0")),
    "OLLAMA_TOP_P": float(os.getenv("OLLAMA_TOP_P", "1")),
    "OLLAMA_TOP_K": int(os.getenv("OLLAMA_TOP_K", "40")),
    "OLLAMA_USE_STABLE_SEED": os.getenv("OLLAMA_USE_STABLE_SEED", "1") == "1",
    "FALLACY_PARAGRAPH_TIMEOUT": float(os.getenv("FALLACY_PARAGRAPH_TIMEOUT", "90")),
    "FALLACY_MIN_PARAGRAPH_CHARS": int(os.getenv("FALLACY_MIN_PARAGRAPH_CHARS", "80")),
    "FALLACY_CONTEXT_RADIUS": int(os.getenv("FALLACY_CONTEXT_RADIUS", "1")),
    "FALLACY_CONTEXT_PREVIEW_CHARS": int(os.getenv("FALLACY_CONTEXT_PREVIEW_CHARS", "450")),
    "SHOW_ANALYSIS_MODEL_BADGE": os.getenv("SHOW_ANALYSIS_MODEL_BADGE", "0") == "1",
    "SHOW_UNMATCHED_BADGE": os.getenv("SHOW_UNMATCHED_BADGE", "0") == "1",
    "SHOW_PARAGRAPHS_ANALYZED_BADGE": os.getenv("SHOW_PARAGRAPHS_ANALYZED_BADGE", "0") == "1",
    "SHOW_EXTRACTION_TIME_BADGE": os.getenv("SHOW_EXTRACTION_TIME_BADGE", "0") == "1",
    "SHOW_FALLACY_TYPE_BADGES": os.getenv("SHOW_FALLACY_TYPE_BADGES", "1") == "1",
    "INCLUDE_FALLACY_REASONING": os.getenv("INCLUDE_FALLACY_REASONING", "0") == "1",
    "RELATED_FALLACY_COUNT": int(os.getenv("RELATED_FALLACY_COUNT", "6")),
    "UI_LANGUAGE": os.getenv("UI_LANGUAGE", "en"),
}

ENABLE_FALLACY_ANALYSIS = os.getenv("ENABLE_FALLACY_ANALYSIS", "1") == "1"
DEBUG_HIGHLIGHT = os.getenv("DEBUG_HIGHLIGHT", "0") == "1"


def get_openai_api_key() -> str:
    return str(os.getenv("OPENAI_API_KEY") or "").strip()


def _settings_for_persistence(settings: dict[str, Any]) -> dict[str, Any]:
    persisted = dict(settings or {})
    persisted.pop("OPENAI_API_KEY", None)
    persisted.pop("OPENAI_API_KEY_CONFIGURED", None)
    return persisted



def _settings_path() -> str:
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    return str(INSTANCE_DIR / "settings.json")


def _apply_env_setting_overrides(settings: dict[str, Any]) -> dict[str, Any]:
    overridden = dict(settings or {})

    def apply_str(name: str) -> None:
        raw = os.getenv(name)
        if raw is not None:
            overridden[name] = str(raw).strip()

    def apply_int(name: str) -> None:
        raw = os.getenv(name)
        if raw is not None:
            try:
                overridden[name] = int(str(raw).strip())
            except Exception:
                pass

    def apply_float(name: str) -> None:
        raw = os.getenv(name)
        if raw is not None:
            try:
                overridden[name] = float(str(raw).strip())
            except Exception:
                pass

    def apply_bool(name: str) -> None:
        raw = os.getenv(name)
        if raw is not None:
            overridden[name] = str(raw).strip().lower() in {"1", "true", "yes", "on"}

    for key in ("AI_PROVIDER", "OLLAMA_BASE_URL", "OLLAMA_MODEL", "OPENAI_BASE_URL", "OPENAI_MODEL", "UI_LANGUAGE"):
        apply_str(key)

    for key in ("OLLAMA_TOP_K", "FALLACY_MIN_PARAGRAPH_CHARS", "FALLACY_CONTEXT_RADIUS", "FALLACY_CONTEXT_PREVIEW_CHARS", "RELATED_FALLACY_COUNT"):
        apply_int(key)

    for key in ("OLLAMA_TIMEOUT", "OLLAMA_TEMPERATURE", "OLLAMA_TOP_P", "FALLACY_PARAGRAPH_TIMEOUT"):
        apply_float(key)

    for key in (
        "OLLAMA_USE_STABLE_SEED",
        "SHOW_ANALYSIS_MODEL_BADGE",
        "SHOW_UNMATCHED_BADGE",
        "SHOW_PARAGRAPHS_ANALYZED_BADGE",
        "SHOW_EXTRACTION_TIME_BADGE",
        "SHOW_FALLACY_TYPE_BADGES",
        "INCLUDE_FALLACY_REASONING",
    ):
        apply_bool(key)

    return overridden


def load_settings() -> dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    path = _settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                saved = dict(saved)
                saved.pop("OPENAI_API_KEY", None)
                settings.update(saved)
        except Exception:
            pass
    settings = _apply_env_setting_overrides(settings)
    settings["OLLAMA_BASE_URL"] = str(settings.get("OLLAMA_BASE_URL") or DEFAULT_SETTINGS["OLLAMA_BASE_URL"]).rstrip("/")
    settings["OPENAI_BASE_URL"] = str(settings.get("OPENAI_BASE_URL") or DEFAULT_SETTINGS["OPENAI_BASE_URL"]).rstrip("/")
    settings["RELATED_FALLACY_COUNT"] = _sanitize_related_fallacy_count(settings.get("RELATED_FALLACY_COUNT"), default=DEFAULT_SETTINGS["RELATED_FALLACY_COUNT"])
    settings["OPENAI_API_KEY"] = ""
    settings["OPENAI_API_KEY_CONFIGURED"] = bool(get_openai_api_key())
    return settings


def save_settings(new_settings: dict[str, Any]) -> None:
    path = _settings_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(_settings_for_persistence(new_settings), f, indent=2)


def current_model_label(settings: dict[str, Any]) -> str:
    provider = (settings.get("AI_PROVIDER") or "ollama").lower()
    if provider == "openai":
        return str(settings.get("OPENAI_MODEL") or "OpenAI")
    return str(settings.get("OLLAMA_MODEL") or "Ollama")


def _sanitize_related_fallacy_count(value: Any, default: int = 6) -> int:
    try:
        count = int(value)
    except Exception:
        count = int(default)
    if count not in {3, 6, 9}:
        count = int(default)
    return count


def _related_fallacy_display_options(total_count: int) -> list[int]:
    safe_total = max(0, int(total_count or 0))
    if safe_total <= 0:
        return []
    if safe_total <= 3:
        return [3]
    if safe_total <= 6:
        return [3, 6]
    return [3, 6, 9]


def _resolve_related_fallacy_display_count(saved_value: Any, total_count: int) -> tuple[int, list[int]]:
    options = _related_fallacy_display_options(total_count)
    if not options:
        return 0, []
    preferred = _sanitize_related_fallacy_count(saved_value, default=DEFAULT_SETTINGS["RELATED_FALLACY_COUNT"])
    for option in reversed(options):
        if preferred >= option:
            return option, options
    return options[0], options


def parse_settings_form(form) -> dict[str, Any]:
    settings = load_settings()
    settings["AI_PROVIDER"] = (form.get("AI_PROVIDER") or "ollama").strip().lower()
    settings["OLLAMA_BASE_URL"] = (form.get("OLLAMA_BASE_URL") or settings["OLLAMA_BASE_URL"]).strip().rstrip("/")
    settings["OLLAMA_MODEL"] = (form.get("OLLAMA_MODEL") or settings["OLLAMA_MODEL"]).strip()
    settings["OPENAI_BASE_URL"] = (form.get("OPENAI_BASE_URL") or settings["OPENAI_BASE_URL"]).strip().rstrip("/")
    settings["OPENAI_MODEL"] = (form.get("OPENAI_MODEL") or settings["OPENAI_MODEL"]).strip()

    def get_float(name, fallback):
        raw = (form.get(name) or "").strip()
        return float(raw) if raw else float(fallback)

    def get_int(name, fallback):
        raw = (form.get(name) or "").strip()
        return int(raw) if raw else int(fallback)

    settings["OLLAMA_TIMEOUT"] = get_float("OLLAMA_TIMEOUT", settings["OLLAMA_TIMEOUT"])
    settings["OLLAMA_TEMPERATURE"] = get_float("OLLAMA_TEMPERATURE", settings["OLLAMA_TEMPERATURE"])
    settings["OLLAMA_TOP_P"] = get_float("OLLAMA_TOP_P", settings["OLLAMA_TOP_P"])
    settings["OLLAMA_TOP_K"] = get_int("OLLAMA_TOP_K", settings["OLLAMA_TOP_K"])

    def get_checkbox(name: str) -> bool:
        try:
            return "1" in form.getlist(name)
        except Exception:
            return form.get(name) == "1"

    settings["OLLAMA_USE_STABLE_SEED"] = get_checkbox("OLLAMA_USE_STABLE_SEED")
    settings["FALLACY_PARAGRAPH_TIMEOUT"] = max(5.0, get_float("FALLACY_PARAGRAPH_TIMEOUT", settings.get("FALLACY_PARAGRAPH_TIMEOUT", 90)))
    settings["FALLACY_MIN_PARAGRAPH_CHARS"] = max(1, get_int("FALLACY_MIN_PARAGRAPH_CHARS", settings["FALLACY_MIN_PARAGRAPH_CHARS"]))
    settings["FALLACY_CONTEXT_RADIUS"] = max(0, get_int("FALLACY_CONTEXT_RADIUS", settings["FALLACY_CONTEXT_RADIUS"]))
    settings["FALLACY_CONTEXT_PREVIEW_CHARS"] = max(50, get_int("FALLACY_CONTEXT_PREVIEW_CHARS", settings["FALLACY_CONTEXT_PREVIEW_CHARS"]))
    settings["SHOW_ANALYSIS_MODEL_BADGE"] = get_checkbox("SHOW_ANALYSIS_MODEL_BADGE")
    settings["SHOW_UNMATCHED_BADGE"] = get_checkbox("SHOW_UNMATCHED_BADGE")
    settings["SHOW_PARAGRAPHS_ANALYZED_BADGE"] = get_checkbox("SHOW_PARAGRAPHS_ANALYZED_BADGE")
    settings["SHOW_EXTRACTION_TIME_BADGE"] = get_checkbox("SHOW_EXTRACTION_TIME_BADGE")
    settings["SHOW_FALLACY_TYPE_BADGES"] = get_checkbox("SHOW_FALLACY_TYPE_BADGES")
    settings["INCLUDE_FALLACY_REASONING"] = get_checkbox("INCLUDE_FALLACY_REASONING")
    settings["RELATED_FALLACY_COUNT"] = _sanitize_related_fallacy_count(settings.get("RELATED_FALLACY_COUNT"), default=DEFAULT_SETTINGS["RELATED_FALLACY_COUNT"])
    settings.pop("FALLACY_CHUNK_SIZE", None)
    return settings


def list_ollama_models(base_url: str, timeout: float = 10.0, connect_timeout: float = 5.0) -> list[str]:
    url = f"{base_url.rstrip('/')}/api/tags"
    r = requests.get(url, timeout=(connect_timeout, timeout))
    r.raise_for_status()
    data = r.json()
    models = []
    for item in data.get("models") or []:
        name = item.get("name")
        if name:
            models.append(name)
    return models


def list_openai_models(base_url: str, api_key: str, timeout: float = 10.0, connect_timeout: float = 5.0) -> list[str]:
    url = f"{base_url.rstrip('/')}/models"
    headers = {
        "Content-Type": "application/json",
    }
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    r = requests.get(url, headers=headers, timeout=(connect_timeout, timeout))
    r.raise_for_status()
    data = r.json()
    models = []
    for item in data.get("data") or []:
        name = item.get("id")
        if name:
            models.append(name)
    return models


def validate_provider_connection(settings: dict[str, Any], timeout: float = 10.0, connect_timeout: float = 5.0) -> dict[str, Any]:
    provider = (settings.get("AI_PROVIDER") or "ollama").strip().lower()
    if provider == "openai":
        base_url = str(settings.get("OPENAI_BASE_URL") or "").strip().rstrip("/")
        api_key = get_openai_api_key()
        model_name = str(settings.get("OPENAI_MODEL") or "").strip()
        if not base_url:
            raise ValueError("Please enter an OpenAI-compatible base URL before testing the connection.")
        if not api_key:
            raise ValueError("OPENAI_API_KEY is not set in the server environment.")
        models = list_openai_models(base_url, api_key, timeout=timeout, connect_timeout=connect_timeout)
        if model_name and models and model_name not in models:
            raise ValueError(f'The configured model "{model_name}" is not available at the tested OpenAI-compatible endpoint.')
        return {
            "provider": "openai",
            "base_url": base_url,
            "models": models,
            "models_count": len(models),
        }

    base_url = str(settings.get("OLLAMA_BASE_URL") or "").strip().rstrip("/")
    model_name = str(settings.get("OLLAMA_MODEL") or "").strip()
    if not base_url:
        raise ValueError("Please enter an Ollama base URL before testing the connection.")
    models = list_ollama_models(base_url, timeout=timeout, connect_timeout=connect_timeout)
    if model_name and models and model_name not in models:
        raise ValueError(f'The configured Ollama model "{model_name}" is not available at the tested Ollama instance.')
    return {
        "provider": "ollama",
        "base_url": base_url,
        "models": models,
        "models_count": len(models),
    }


def _provider_connection_error_message(settings: dict[str, Any]) -> str:
    return _translate_for_settings("errors.provider_connection_required", settings)


def ensure_provider_connection(
    settings: dict[str, Any],
    *,
    attempts: int = 2,
    connect_timeout: float = 3.0,
    read_timeout: float = 3.0,
) -> None:
    last_exc: Exception | None = None
    total_attempts = max(1, int(attempts or 1))
    for attempt_index in range(total_attempts):
        try:
            validate_provider_connection(settings, timeout=read_timeout, connect_timeout=connect_timeout)
            return
        except Exception as exc:
            last_exc = exc
            if attempt_index + 1 < total_attempts:
                time.sleep(0.15)
    raise ProviderConnectionError(_provider_connection_error_message(settings)) from last_exc

# Remove tracking parameters from URLs when normalizing them for the history database. This ensures the same article accessed via different tracking URLs is stored as one entry and duplicate entries aren't created for the same content.
TRACKING_QUERY_PARAMS = {
    "fbclid",
    "gclid",
    "igshid",
    "mc_cid",
    "mc_eid",
    "mkt_tok",
    "ref_src",
    "spm",
    "utm_campaign",
    "utm_content",
    "utm_id",
    "utm_medium",
    "utm_name",
    "utm_reader",
    "utm_source",
    "utm_term",
}


def _history_db_path() -> str:
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    return str(INSTANCE_DIR / "analysis_history.sqlite3")


def _history_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(_history_db_path())
    conn.row_factory = sqlite3.Row
    return conn


def init_history_db() -> None:
    with _history_connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS saved_analyses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                url_original TEXT NOT NULL,
                url_normalized TEXT NOT NULL,
                provider TEXT NOT NULL,
                model_name TEXT NOT NULL,
                model_label TEXT NOT NULL,
                model_key TEXT NOT NULL,
                title TEXT,
                author TEXT,
                published_date TEXT,
                analyzed_at TEXT NOT NULL,
                created_at_epoch REAL NOT NULL,
                total_fallacies INTEGER NOT NULL DEFAULT 0,
                fallacy_breakdown_json TEXT NOT NULL,
                fallacy_sentences_json TEXT NOT NULL,
                findings_json TEXT NOT NULL,
                analysis_stats_json TEXT NOT NULL,
                summary_json TEXT NOT NULL,
                paragraphs_json TEXT NOT NULL,
                extraction_stats_json TEXT NOT NULL,
                scraped_content TEXT NOT NULL,
                UNIQUE(url_normalized, model_key)
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_saved_analyses_url_normalized ON saved_analyses(url_normalized)"
        )


def _json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False)


def _json_loads(value: str | None, default: Any):
    if not value:
        return default
    try:
        return json.loads(value)
    except Exception:
        return default


def _analysis_timestamp() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def _normalize_url_for_history(url: str) -> str:
    parsed = urlparse((url or "").strip())
    scheme = (parsed.scheme or "https").lower()
    netloc = (parsed.netloc or "").lower()
    path = parsed.path or "/"
    if path != "/":
        path = path.rstrip("/") or "/"
    query_items = []
    for key, value in parse_qsl(parsed.query, keep_blank_values=True):
        if key.lower() in TRACKING_QUERY_PARAMS:
            continue
        query_items.append((key, value))
    query = urlencode(sorted(query_items), doseq=True)
    return urlunparse((scheme, netloc, path, "", query, ""))


def _model_identity_from_settings(settings: dict[str, Any]) -> tuple[str, str, str, str]:
    provider = (settings.get("AI_PROVIDER") or "ollama").strip().lower()
    if provider == "openai":
        model_name = str(settings.get("OPENAI_MODEL") or "OpenAI")
    else:
        model_name = str(settings.get("OLLAMA_MODEL") or "Ollama")
    model_key = f"{provider}::{model_name}"
    provider_label = provider.upper() if provider else "LLM"
    model_label = f"{provider_label} · {model_name}"
    return provider, model_name, model_key, model_label

# When the AI identifies a fallacy and quotes a portion of text, users need to see the full sentence to understand the context and evaluate whether the analysis is correct. This function ensures that instead of showing just a fragmented quote, the user sees the complete sentence containing that quote.
# This function extracts and returns the complete sentence containing that quote (rather than just the partial excerpt).
def _sentence_from_quote(paragraph: str, quote: str) -> str:
    text = " ".join((paragraph or "").split())
    if not text:
        return ""
    quote = (quote or "").strip()
    if not quote:
        return text
    sentences = re.split(r"(?<=[.!?])\s+", text)
    for sentence in sentences:
        if quote in sentence:
            return sentence.strip()
    low_quote = quote.lower()
    for sentence in sentences:
        if low_quote and low_quote in sentence.lower():
            return sentence.strip()
    normalized_text = _norm_1to1(text)
    normalized_quote = _norm_1to1(quote)
    normalized_sentences = re.split(r"(?<=[.!?])\s+", normalized_text)
    for idx, sentence in enumerate(normalized_sentences):
        if normalized_quote and normalized_quote in sentence:
            try:
                return sentences[idx].strip()
            except Exception:
                return text
    return text


def _history_record_from_row(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    record = dict(row)
    record["fallacy_breakdown"] = _json_loads(record.pop("fallacy_breakdown_json", None), {})
    record["fallacy_sentences"] = _json_loads(record.pop("fallacy_sentences_json", None), [])
    record["findings"] = _json_loads(record.pop("findings_json", None), [])
    record["analysis_stats"] = _json_loads(record.pop("analysis_stats_json", None), {})
    record["summary"] = _json_loads(record.pop("summary_json", None), {})
    record["paragraphs"] = _json_loads(record.pop("paragraphs_json", None), [])
    record["extraction_stats"] = _json_loads(record.pop("extraction_stats_json", None), {})
    return record


def _history_brief_record(record: dict[str, Any] | None) -> dict[str, Any] | None:
    if not record:
        return None
    return {
        "record_id": record["id"],
        "url_original": record["url_original"],
        "provider": record["provider"],
        "model_name": record["model_name"],
        "model_label": record["model_label"],
        "model_key": record["model_key"],
        "analyzed_at": record["analyzed_at"],
        "total_fallacies": int(record.get("total_fallacies") or 0),
        "fallacy_breakdown": record.get("fallacy_breakdown") or {},
    }


def get_saved_analysis_record(record_id: int) -> dict[str, Any] | None:
    with _history_connect() as conn:
        row = conn.execute(
            "SELECT * FROM saved_analyses WHERE id = ?",
            (int(record_id),),
        ).fetchone()
    return _history_record_from_row(row)


def get_saved_analysis_record_for_model(url_normalized: str, model_key: str) -> dict[str, Any] | None:
    with _history_connect() as conn:
        row = conn.execute(
            "SELECT * FROM saved_analyses WHERE url_normalized = ? AND model_key = ? ORDER BY created_at_epoch DESC, id DESC LIMIT 1",
            (str(url_normalized or ""), str(model_key or "")),
        ).fetchone()
    return _history_record_from_row(row)


def _history_sortable_findings(findings: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for finding in findings or []:
        confidence_raw = finding.get("confidence")
        try:
            confidence_value: Any = round(float(confidence_raw), 6)
        except Exception:
            confidence_value = str(confidence_raw or "")
        normalized.append({
            "paragraph_index": int(finding.get("paragraph_index") or 0),
            "type": str(finding.get("type") or ""),
            "quote": str(finding.get("quote") or ""),
            "explanation": str(finding.get("explanation") or ""),
            "confidence": confidence_value,
        })
    normalized.sort(key=lambda item: (
        int(item.get("paragraph_index") or 0),
        str(item.get("type") or "").casefold(),
        str(item.get("quote") or "").casefold(),
        str(item.get("explanation") or "").casefold(),
        str(item.get("confidence") or ""),
    ))
    return normalized


def _history_sortable_sentences(sentences: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for sentence in sentences or []:
        normalized.append({
            "paragraph_index": int(sentence.get("paragraph_index") or 0),
            "type": str(sentence.get("type") or ""),
            "quote": str(sentence.get("quote") or ""),
            "sentence": str(sentence.get("sentence") or ""),
            "explanation": str(sentence.get("explanation") or ""),
        })
    normalized.sort(key=lambda item: (
        int(item.get("paragraph_index") or 0),
        str(item.get("type") or "").casefold(),
        str(item.get("quote") or "").casefold(),
        str(item.get("sentence") or "").casefold(),
        str(item.get("explanation") or "").casefold(),
    ))
    return normalized


def _history_comparison_bundle_from_payload(payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "provider": str(payload.get("provider") or ""),
        "model_key": str(payload.get("model_key") or ""),
        "title": str(payload.get("title") or ""),
        "author": str(payload.get("author") or ""),
        "published_date": str(payload.get("published_date") or ""),
        "paragraphs": [str(p or "") for p in (payload.get("paragraphs") or [])],
        "scraped_content": str(payload.get("scraped_content") or ""),
        "total_fallacies": int(payload.get("total_fallacies") or 0),
        "fallacy_breakdown": payload.get("fallacy_breakdown") or {},
        "fallacy_sentences": _history_sortable_sentences(payload.get("fallacy_sentences") or []),
        "findings": _history_sortable_findings(payload.get("findings") or []),
    }


def _history_comparison_bundle_from_record(record: dict[str, Any]) -> dict[str, Any]:
    return {
        "provider": str(record.get("provider") or ""),
        "model_key": str(record.get("model_key") or ""),
        "title": str(record.get("title") or ""),
        "author": str(record.get("author") or ""),
        "published_date": str(record.get("published_date") or ""),
        "paragraphs": [str(p or "") for p in (record.get("paragraphs") or [])],
        "scraped_content": str(record.get("scraped_content") or ""),
        "total_fallacies": int(record.get("total_fallacies") or 0),
        "fallacy_breakdown": record.get("fallacy_breakdown") or {},
        "fallacy_sentences": _history_sortable_sentences(record.get("fallacy_sentences") or []),
        "findings": _history_sortable_findings(record.get("findings") or []),
    }


def saved_analysis_differs(existing_record: dict[str, Any] | None, payload: dict[str, Any]) -> bool:
    if not existing_record:
        return True
    existing_bundle = _history_comparison_bundle_from_record(existing_record)
    new_bundle = _history_comparison_bundle_from_payload(payload)
    return json.dumps(existing_bundle, ensure_ascii=False, sort_keys=True) != json.dumps(new_bundle, ensure_ascii=False, sort_keys=True)


def overwrite_analysis_record(record_id: int, payload: dict[str, Any]) -> int:
    with _history_connect() as conn:
        conn.execute(
            """
            UPDATE saved_analyses
            SET url_original = ?,
                url_normalized = ?,
                provider = ?,
                model_name = ?,
                model_label = ?,
                model_key = ?,
                title = ?,
                author = ?,
                published_date = ?,
                analyzed_at = ?,
                created_at_epoch = ?,
                total_fallacies = ?,
                fallacy_breakdown_json = ?,
                fallacy_sentences_json = ?,
                findings_json = ?,
                analysis_stats_json = ?,
                summary_json = ?,
                paragraphs_json = ?,
                extraction_stats_json = ?,
                scraped_content = ?
            WHERE id = ?
            """,
            (
                payload["url_original"],
                payload["url_normalized"],
                payload["provider"],
                payload["model_name"],
                payload["model_label"],
                payload["model_key"],
                payload.get("title"),
                payload.get("author"),
                payload.get("published_date"),
                payload["analyzed_at"],
                float(payload.get("created_at_epoch") or time.time()),
                int(payload.get("total_fallacies") or 0),
                _json_dumps(payload.get("fallacy_breakdown") or {}),
                _json_dumps(payload.get("fallacy_sentences") or []),
                _json_dumps(payload.get("findings") or []),
                _json_dumps(payload.get("analysis_stats") or {}),
                _json_dumps(payload.get("summary") or {}),
                _json_dumps(payload.get("paragraphs") or []),
                _json_dumps(payload.get("extraction_stats") or {}),
                payload.get("scraped_content") or "",
                int(record_id),
            ),
        )
        conn.commit()
    return int(record_id)


def get_saved_history_summary(url: str, model_key: str | None = None) -> dict[str, Any]:
    normalized_url = _normalize_url_for_history(url)
    with _history_connect() as conn:
        rows = conn.execute(
            "SELECT * FROM saved_analyses WHERE url_normalized = ? ORDER BY created_at_epoch DESC, id DESC",
            (normalized_url,),
        ).fetchall()

    records = [_history_record_from_row(row) for row in rows]
    if not records:
        return {
            "exists": False,
            "url_original": url,
            "url_normalized": normalized_url,
            "saved_run_count": 0,
            "distinct_model_count": 0,
            "models_used": [],
            "current_model_has_saved_result": False,
            "latest_record": None,
            "preferred_record": None,
        }

    latest_record = records[0]
    preferred_record = next((record for record in records if model_key and record.get("model_key") == model_key), None) or latest_record
    models_used = []
    seen_model_keys = set()
    for record in records:
        record_model_key = record.get("model_key")
        if record_model_key in seen_model_keys:
            continue
        seen_model_keys.add(record_model_key)
        models_used.append({
            "provider": record.get("provider"),
            "model_name": record.get("model_name"),
            "model_label": record.get("model_label"),
            "model_key": record_model_key,
        })

    return {
        "exists": True,
        "url_original": latest_record.get("url_original") or url,
        "url_normalized": normalized_url,
        "saved_run_count": len(records),
        "distinct_model_count": len(models_used),
        "models_used": models_used,
        "current_model_has_saved_result": any(model_key and record.get("model_key") == model_key for record in records),
        "latest_record": _history_brief_record(latest_record),
        "preferred_record": _history_brief_record(preferred_record),
    }


def save_analysis_record(payload: dict[str, Any]) -> tuple[bool, int | None]:
    normalized_url = payload["url_normalized"]
    model_key = payload["model_key"]
    with _history_connect() as conn:
        existing = conn.execute(
            "SELECT id FROM saved_analyses WHERE url_normalized = ? AND model_key = ?",
            (normalized_url, model_key),
        ).fetchone()
        if existing:
            return False, int(existing["id"])
        cur = conn.execute(
            """
            INSERT INTO saved_analyses (
                url_original, url_normalized, provider, model_name, model_label, model_key,
                title, author, published_date, analyzed_at, created_at_epoch, total_fallacies,
                fallacy_breakdown_json, fallacy_sentences_json, findings_json, analysis_stats_json,
                summary_json, paragraphs_json, extraction_stats_json, scraped_content
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                payload["url_original"],
                normalized_url,
                payload["provider"],
                payload["model_name"],
                payload["model_label"],
                model_key,
                payload.get("title"),
                payload.get("author"),
                payload.get("published_date"),
                payload["analyzed_at"],
                float(payload.get("created_at_epoch") or time.time()),
                int(payload.get("total_fallacies") or 0),
                _json_dumps(payload.get("fallacy_breakdown") or {}),
                _json_dumps(payload.get("fallacy_sentences") or []),
                _json_dumps(payload.get("findings") or []),
                _json_dumps(payload.get("analysis_stats") or {}),
                _json_dumps(payload.get("summary") or {}),
                _json_dumps(payload.get("paragraphs") or []),
                _json_dumps(payload.get("extraction_stats") or {}),
                payload.get("scraped_content") or "",
            ),
        )
        conn.commit()
        return True, int(cur.lastrowid)


def _build_saved_analysis_payload(job: dict[str, Any], findings_all: list[dict[str, Any]], analysis_stats: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
    paragraphs = job.get("paragraphs") or []
    _, applied_findings, _, _ = apply_highlights(paragraphs, findings_all, include_reasoning=False)
    elapsed_seconds = _job_elapsed_seconds(job)
    summary = summarize_findings(applied_findings, analysis_stats, elapsed_seconds=elapsed_seconds)
    provider, model_name, model_key, model_label = _model_identity_from_settings(settings)
    fallacy_sentences = []
    for finding in applied_findings:
        paragraph_index = int(finding.get("paragraph_index") or 0)
        paragraph_text = paragraphs[paragraph_index] if 0 <= paragraph_index < len(paragraphs) else ""
        fallacy_sentences.append({
            "paragraph_index": paragraph_index,
            "type": finding.get("type") or "Fallacy",
            "quote": finding.get("quote") or "",
            "sentence": _sentence_from_quote(paragraph_text, finding.get("quote") or ""),
            "explanation": finding.get("explanation") or "",
            "confidence": finding.get("confidence") or 0.0,
        })

    return {
        "url_original": job.get("url") or "",
        "url_normalized": _normalize_url_for_history(job.get("url") or ""),
        "provider": provider,
        "model_name": model_name,
        "model_label": model_label,
        "model_key": model_key,
        "title": job.get("title") or "Extracted Article",
        "author": job.get("author"),
        "published_date": job.get("date"),
        "analyzed_at": _analysis_timestamp(),
        "created_at_epoch": float(job.get("finished_at") or time.time()),
        "total_fallacies": int(summary.get("total") or 0),
        "fallacy_breakdown": summary.get("by_type") or {},
        "fallacy_sentences": fallacy_sentences,
        "findings": findings_all,
        "analysis_stats": analysis_stats or {},
        "summary": summary,
        "paragraphs": paragraphs,
        "extraction_stats": job.get("extraction_stats") or {},
        "scraped_content": "\n\n".join(paragraphs),
    }


def load_saved_analysis_into_job(record_id: int, settings: dict[str, Any]) -> str:
    record = get_saved_analysis_record(record_id)
    if not record:
        raise ValueError("Saved analysis not found.")

    analysis_stats = record.get("analysis_stats") or {}
    elapsed_seconds = float((record.get("summary") or {}).get("elapsed_seconds") or analysis_stats.get("elapsed_seconds") or 0.0)
    finished_at = float(record.get("created_at_epoch") or time.time())
    started_at = finished_at - max(0.0, elapsed_seconds)

    job_id = _create_job(record.get("url_original") or "")
    job = _get_job(job_id) or {"job_id": job_id}
    job.update({
        "url": record.get("url_original") or "",
        "title": record.get("title") or "Extracted Article",
        "author": record.get("author"),
        "date": record.get("published_date"),
        "paragraphs": record.get("paragraphs") or [],
        "extraction_stats": record.get("extraction_stats") or {},
        "started_at": started_at,
        "finished_at": finished_at,
        "llm_model_label": record.get("model_label") or record.get("model_name") or "Saved analysis",
        "history_status_message": _translate_for_settings(
            "messages.saved_analysis_from",
            settings,
            analyzed_at=record.get("analyzed_at"),
            model_label=record.get("model_label"),
        ),
        "history_record_id": record.get("id"),
        "history_loaded": True,
    })
    result_html = _render_result_html_for_job(job, settings, record.get("findings") or [], analysis_stats)
    _update_job(
        job_id,
        phase="analysis",
        status="done",
        stage=_translate_for_settings("job.loaded_saved_analysis", settings),
        progress=100,
        title=job.get("title"),
        author=job.get("author"),
        date=job.get("date"),
        paragraphs=job.get("paragraphs"),
        extraction_stats=job.get("extraction_stats"),
        started_at=started_at,
        finished_at=finished_at,
        llm_model_label=job.get("llm_model_label"),
        history_status_message=job.get("history_status_message"),
        history_record_id=job.get("history_record_id"),
        history_loaded=True,
        result_html=result_html,
        paragraph_current=(analysis_stats or {}).get("paragraphs_analyzed", 0),
        paragraph_total=(analysis_stats or {}).get("paragraphs_eligible", 0),
        paragraphs_total=len(job.get("paragraphs") or []),
        paragraphs_analyzed=(analysis_stats or {}).get("paragraphs_analyzed", 0),
        paragraphs_skipped_short=(analysis_stats or {}).get("paragraphs_skipped_short", 0),
        paragraphs_timed_out=(analysis_stats or {}).get("paragraphs_timed_out", 0),
        paragraphs_failed=(analysis_stats or {}).get("paragraphs_failed", 0),
    )
    return job_id



# -----------------------------
# URL + paragraph utilities
# -----------------------------
def is_valid_http_url(url: str) -> bool:
    try:
        p = urlparse(url)
        return p.scheme in ("http", "https") and bool(p.netloc)
    except Exception:
        return False


def split_into_paragraphs(text: str) -> list[str]:
    cleaned = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not cleaned:
        return []
    chunks = re.split(r"\n\s*\n+", cleaned)
    chunks = [c.strip() for c in chunks if c and c.strip()]
    if len(chunks) <= 1 and "\n" in cleaned:
        chunks = re.split(r"\n+", cleaned)
        chunks = [c.strip() for c in chunks if c and c.strip()]
    return chunks


def normalize_pasted_text(text: str) -> str:
    return (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def validate_pasted_text(text: str) -> str | None:
    cleaned = normalize_pasted_text(text)
    if not cleaned:
        return "Please paste some text before starting the analysis."
    if len(cleaned) < 40:
        return "Please paste a bit more text so the analyzer has enough material to evaluate."
    alpha_count = sum(1 for ch in cleaned if ch.isalpha())
    if alpha_count < 20:
        return "The pasted text does not look substantial enough to analyze."
    return None


class _BlockTextParser(HTMLParser):
    BLOCK_TAGS = {"p", "li", "blockquote", "h1", "h2", "h3", "h4", "h5", "h6"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.paragraphs: list[str] = []
        self._in_block = False
        self._buf: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag in self.BLOCK_TAGS:
            self._flush()
            self._in_block = True

    def handle_endtag(self, tag):
        if tag in self.BLOCK_TAGS:
            self._flush()
            self._in_block = False

    def handle_data(self, data):
        if self._in_block and data:
            self._buf.append(data)

    def _flush(self):
        if not self._buf:
            return
        text = " ".join("".join(self._buf).split())
        if text:
            self.paragraphs.append(text)
        self._buf = []


def paragraphs_from_extracted_html(extracted_html: str | None) -> list[str]:
    if not extracted_html:
        return []
    parser = _BlockTextParser()
    parser.feed(extracted_html)
    parser.close()
    return [p for p in parser.paragraphs if p.strip()]

# -----------------------------
# LLM fallacy analysis
# -----------------------------
def build_fallacy_system_prompt(include_reasoning: bool = False, locale: str = "en") -> str:
    explanation_language = _ui_language_name(locale)
    lines = [
        "You are a careful logician. Identify potential logical fallacies in the provided focus paragraph.",
        "",
        "Rules:",
        "- Analyze ONLY the focus paragraph(s). Neighboring paragraphs are context only.",
        "- Only label something a fallacy if the quoted text supports the label.",
        "- Quotes MUST be exact substrings copied from the focus paragraph text.",
        "- If you cannot copy an exact substring, OMIT that finding.",
        "- Return ONLY valid JSON. No markdown, no extra commentary.",
        '- If none found, return {"fallacies": []}.',
        "",
        "Use this fallacy set (not exhaustive): Affirming the Consequent, Bad Reason Fallacy,",
        "Propositional Fallacies, Quantification Fallacies, Sunk Cost Fallacy, Syllogistic Fallacies,",
        "Ad Hominem, Ambiguity, Anecdotal, Appeal to Authority, Appeal to Emotion, Appeal to Nature,",
        "Appeal to Ridicule, Appeal to Tradition, Argument from Repetition, Argumentum ad Populum,",
        "Bandwagon, Begging the Question, Burden of Proof, Circular Reasoning, Continuum Fallacy,",
        "Equivocation, Etymological Fallacy, Fallacy Fallacy, Fallacy of Composition and Division,",
        "Fallacy of Quoting Out of Context, False Cause & False Attribution, False Dilemma, Furtive Fallacy,",
        "Gambler's Fallacy, Genetic Fallacy, Hasty Generalization, Ignoratio Elenchi, Incomplete Comparison,",
        "Inflation of Conflict, Kettle Logic, Loaded Question, Middle Ground, No True Scotsman,",
        "Personal Incredulity, Post Hoc, Proof by Verbosity, Proving Too Much, Red Herring, Reification,",
        "Retrospective Determinism, Shotgun Argumentation, Slippery Slope, Special Pleading, Strawman,",
        "Texas Sharpshooter, Tu Quoque.",
        "",
        (f"- Explanation text MUST be written in {explanation_language}." if include_reasoning else ""),
        "Output schema:",
        "{",
        '  "fallacies": [',
        "    {",
        '      "paragraph_index": 0,',
        '      "quote": "exact text span",',
        '      "type": "Strawman",',
        '      "confidence": 0.0' + (',' if include_reasoning else ''),
    ]
    if include_reasoning:
        lines.append('      "explanation": "1-2 sentences"')
    lines.extend([
        "    }",
        "  ]",
        "}",
        "",
        "When reasoning mode is OFF, keep the response minimal and do not include explanation text.",
    ])
    if include_reasoning:
        lines.append(f"When reasoning mode is ON, provide a short explanation for each finding in {explanation_language}.")
    return "\n".join(lines).strip()


def _analysis_int_setting(settings: dict[str, Any], name: str, default: int, minimum: int | None = None) -> int:
    try:
        value = int(settings.get(name, default))
    except Exception:
        value = default
    if minimum is not None:
        value = max(minimum, value)
    return value


def _analysis_float_setting(settings: dict[str, Any], name: str, default: float, minimum: float | None = None) -> float:
    try:
        value = float(settings.get(name, default))
    except Exception:
        value = default
    if minimum is not None:
        value = max(minimum, value)
    return value


def _stable_seed(*parts: str) -> int:
    h = hashlib.sha256()
    for p in parts:
        h.update((p or "").encode("utf-8", errors="ignore"))
        h.update(b"\x00")
    return int.from_bytes(h.digest()[:4], "big", signed=False)


def _extract_json_object(s: str) -> dict[str, Any] | None:
    if not s:
        return None
    s = s.strip()
    try:
        return json.loads(s)
    except Exception:
        pass
    start = s.find("{")
    end = s.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = s[start : end + 1]
        try:
            return json.loads(candidate)
        except Exception:
            return None
    return None


def ollama_chat(messages: list[dict[str, str]], settings: dict[str, Any], model: str | None = None, seed: int | None = None, timeout_seconds: float | None = None, job_id: str | None = None, cancel_check=None) -> str:
    url = f"{settings['OLLAMA_BASE_URL']}/api/chat"
    options: dict[str, Any] = {
        "temperature": settings["OLLAMA_TEMPERATURE"],
        "top_p": settings["OLLAMA_TOP_P"],
        "top_k": settings["OLLAMA_TOP_K"],
    }
    if seed is not None:
        options["seed"] = seed
    resolved_model = model or settings["OLLAMA_MODEL"]
    payload = {
        "model": resolved_model,
        "messages": messages,
        "stream": True,
        "format": "json",
        "options": options,
    }
    read_timeout = float(timeout_seconds if timeout_seconds is not None else settings["OLLAMA_TIMEOUT"])
    session = requests.Session()
    response = None
    chunks: list[str] = []
    try:
        if job_id:
            _set_job_stream_handle(job_id, provider="ollama", model=resolved_model, response=None, session=session)
        response = session.post(url, json=payload, stream=True, timeout=(3, read_timeout))
        response.raise_for_status()
        if job_id:
            _set_job_stream_handle(job_id, provider="ollama", model=resolved_model, response=response, session=session)
        for raw_line in response.iter_lines(decode_unicode=True):
            if cancel_check and cancel_check():
                raise AnalysisCancelled("Analysis cancelled by user.")
            if not raw_line:
                continue
            line = raw_line.strip()
            if not line:
                continue
            part = json.loads(line)
            message = part.get("message") or {}
            content_piece = message.get("content")
            if content_piece:
                chunks.append(content_piece)
            if part.get("done"):
                final_content = message.get("content")
                if final_content and not chunks:
                    chunks.append(final_content)
                break
        return "".join(chunks)
    except requests.RequestException as e:
        if cancel_check and cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.") from e
        raise
    finally:
        if job_id:
            _clear_job_stream_handle(job_id)
        try:
            if response is not None:
                response.close()
        except Exception:
            pass
        session.close()


def openai_chat(messages: list[dict[str, str]], settings: dict[str, Any], model: str | None = None, seed: int | None = None, timeout_seconds: float | None = None) -> str:
    api_key = get_openai_api_key()
    if not api_key:
        raise ValueError("OpenAI is selected, but OPENAI_API_KEY is not set in the server environment.")
    url = f"{settings['OPENAI_BASE_URL']}/chat/completions"
    payload: dict[str, Any] = {
        "model": model or settings["OPENAI_MODEL"],
        "messages": messages,
        "temperature": settings["OLLAMA_TEMPERATURE"],
        "top_p": settings["OLLAMA_TOP_P"],
        "response_format": {"type": "json_object"},
    }
    if seed is not None:
        payload["seed"] = int(seed)
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    read_timeout = float(timeout_seconds if timeout_seconds is not None else settings["OLLAMA_TIMEOUT"])
    r = requests.post(url, headers=headers, json=payload, timeout=(3, read_timeout))
    r.raise_for_status()
    data = r.json()
    return (((data.get("choices") or [{}])[0].get("message") or {}).get("content", "")) or ""


def provider_chat(messages: list[dict[str, str]], settings: dict[str, Any], seed: int | None = None, timeout_seconds: float | None = None, job_id: str | None = None, cancel_check=None) -> str:
    provider = (settings.get("AI_PROVIDER") or "ollama").lower()
    if provider == "openai":
        return openai_chat(messages, settings, seed=seed, timeout_seconds=timeout_seconds)
    return ollama_chat(messages, settings, seed=seed, timeout_seconds=timeout_seconds, job_id=job_id, cancel_check=cancel_check)


def _clean_and_dedupe_findings(raw: list[dict[str, Any]], paragraphs: list[str]) -> list[dict[str, Any]]:
    cleaned: list[dict[str, Any]] = []
    seen = set()
    for f in raw or []:
        try:
            pi = int(f.get("paragraph_index"))
        except Exception:
            continue
        quote = (f.get("quote") or "").strip()
        ftype = (f.get("type") or "").strip() or "Fallacy"
        if pi < 0 or pi >= len(paragraphs) or not quote:
            continue
        key = (pi, quote, ftype)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append({
            "paragraph_index": pi,
            "quote": quote,
            "type": ftype,
            "confidence": float(f.get("confidence") or 0.0),
            "explanation": (f.get("explanation") or "").strip(),
        })
    return cleaned


def _should_analyze_paragraph(paragraph: str, settings: dict[str, Any]) -> bool:
    text = (paragraph or "").strip()
    if not text:
        return False
    min_chars = _analysis_int_setting(settings, "FALLACY_MIN_PARAGRAPH_CHARS", 80, minimum=1)
    if len(text) < min_chars:
        return False
    alpha_count = sum(1 for ch in text if ch.isalpha())
    return alpha_count >= max(25, min_chars // 3)


def _context_preview(paragraph: str, settings: dict[str, Any]) -> str:
    preview_chars = _analysis_int_setting(settings, "FALLACY_CONTEXT_PREVIEW_CHARS", 450, minimum=50)
    text = " ".join((paragraph or "").split())
    if len(text) <= preview_chars:
        return text
    return text[:preview_chars].rstrip() + " …"


def _build_paragraph_analysis_prompt(paragraphs: list[str], focus_index: int, settings: dict[str, Any]) -> str:
    context_radius = _analysis_int_setting(settings, "FALLACY_CONTEXT_RADIUS", 1, minimum=0)
    start = max(0, focus_index - context_radius)
    end = min(len(paragraphs), focus_index + context_radius + 1)
    lines = [
        f"Analyze only paragraph {focus_index}. The surrounding paragraphs are for context only.",
        "",
    ]
    for idx in range(start, end):
        label = "FOCUS" if idx == focus_index else "CONTEXT"
        text = paragraphs[idx] if idx == focus_index else _context_preview(paragraphs[idx], settings)
        lines.append(f"{label} PARA {idx}: {text}")
        lines.append("")
    lines.extend([
        "Return JSON only.",
        f"If you find fallacies, every item must use paragraph_index {focus_index}.",
        "If none are present, return {\"fallacies\": []}.",
    ])
    return "\n".join(lines).strip()



def analyze_fallacies(
    paragraphs: list[str],
    url: str,
    settings: dict[str, Any],
    progress_callback=None,
    cancel_check=None,
    job_id: str | None = None,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if not paragraphs or not ENABLE_FALLACY_ANALYSIS:
        stats = {
            "paragraphs_total": len(paragraphs),
            "paragraphs_eligible": 0,
            "paragraphs_analyzed": 0,
            "paragraphs_skipped_short": 0,
            "paragraphs_timed_out": 0,
            "paragraphs_failed": 0,
            "effective_timeout_seconds": 0,
        }
        if progress_callback:
            progress_callback({
                "stage": _translate_for_settings("job.analysis_disabled", settings),
                "current": 0,
                "total": 0,
                "stats": stats,
            })
        return [], stats

    base_seed = None
    if settings.get("OLLAMA_USE_STABLE_SEED"):
        base_seed = _stable_seed(url, "\n\n".join(paragraphs))

    paragraph_timeout = _analysis_float_setting(settings, "FALLACY_PARAGRAPH_TIMEOUT", 90.0, minimum=5.0)
    try:
        global_timeout = float(settings.get("OLLAMA_TIMEOUT", paragraph_timeout))
    except Exception:
        global_timeout = paragraph_timeout
    effective_timeout = min(paragraph_timeout, global_timeout) if global_timeout > 0 else paragraph_timeout

    eligible_total = sum(1 for paragraph in paragraphs if _should_analyze_paragraph(paragraph, settings))
    include_reasoning = bool(settings.get("INCLUDE_FALLACY_REASONING"))
    locale = _current_locale(settings)
    system_prompt = build_fallacy_system_prompt(include_reasoning=include_reasoning, locale=locale)
    findings: list[dict[str, Any]] = []
    stats: dict[str, Any] = {
        "paragraphs_total": len(paragraphs),
        "paragraphs_eligible": eligible_total,
        "paragraphs_analyzed": 0,
        "paragraphs_skipped_short": 0,
        "paragraphs_timed_out": 0,
        "paragraphs_failed": 0,
        "effective_timeout_seconds": effective_timeout,
    }

    successful_provider_calls = 0
    consecutive_provider_failures = 0

    if progress_callback:
        progress_callback({
            "stage": _translate_for_settings("job.preparing_analysis", settings),
            "current": 0,
            "total": eligible_total,
            "stats": dict(stats),
        })

    def should_abort_for_provider_failure() -> bool:
        return successful_provider_calls == 0 or consecutive_provider_failures >= 2

    for idx, paragraph in enumerate(paragraphs):
        if cancel_check and cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.")

        if not _should_analyze_paragraph(paragraph, settings):
            stats["paragraphs_skipped_short"] += 1
            if progress_callback:
                progress_callback({
                    "stage": _translate_for_settings("job.skipping_short_paragraph", settings, current=idx + 1, total=len(paragraphs)),
                    "current": stats["paragraphs_analyzed"],
                    "total": eligible_total,
                    "paragraph_index": idx,
                    "stats": dict(stats),
                })
            continue

        seed = ((base_seed ^ idx) & 0xFFFFFFFF) if base_seed is not None else None
        prompt = _build_paragraph_analysis_prompt(paragraphs, idx, settings)

        if progress_callback:
            progress_callback({
                "stage": _translate_for_settings("job.analyzing_paragraph", settings, current=stats['paragraphs_analyzed'] + 1, total=eligible_total),
                "current": stats["paragraphs_analyzed"],
                "total": eligible_total,
                "paragraph_index": idx,
                "stats": dict(stats),
            })

        try:
            content = provider_chat([
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ], settings, seed=seed, timeout_seconds=effective_timeout, job_id=job_id, cancel_check=cancel_check)
        except AnalysisCancelled:
            raise
        except requests.Timeout as e:
            consecutive_provider_failures += 1
            if should_abort_for_provider_failure():
                raise ProviderConnectionError(_provider_connection_error_message(settings)) from e
            stats["paragraphs_timed_out"] += 1
            stats["paragraphs_analyzed"] += 1
            if DEBUG_HIGHLIGHT:
                print(f"[fallacy-analysis] Paragraph {idx} timed out after {effective_timeout}s: {e}")
            if progress_callback:
                progress_callback({
                    "stage": _translate_for_settings("job.timed_out_paragraph", settings, current=stats['paragraphs_analyzed'], total=eligible_total),
                    "current": stats["paragraphs_analyzed"],
                    "total": eligible_total,
                    "paragraph_index": idx,
                    "stats": dict(stats),
                })
            continue
        except requests.RequestException as e:
            consecutive_provider_failures += 1
            if should_abort_for_provider_failure():
                raise ProviderConnectionError(_provider_connection_error_message(settings)) from e
            stats["paragraphs_failed"] += 1
            stats["paragraphs_analyzed"] += 1
            if DEBUG_HIGHLIGHT:
                print(f"[fallacy-analysis] Paragraph {idx} request failed: {e}")
            if progress_callback:
                progress_callback({
                    "stage": _translate_for_settings("job.failed_paragraph", settings, current=stats['paragraphs_analyzed'], total=eligible_total),
                    "current": stats["paragraphs_analyzed"],
                    "total": eligible_total,
                    "paragraph_index": idx,
                    "stats": dict(stats),
                })
            continue
        except ValueError as e:
            raise ProviderConnectionError(_provider_connection_error_message(settings)) from e

        consecutive_provider_failures = 0
        successful_provider_calls += 1
        stats["paragraphs_analyzed"] += 1

        obj = _extract_json_object(content) or {"fallacies": []}
        cleaned = _clean_and_dedupe_findings(obj.get("fallacies") or [], paragraphs)
        findings.extend([f for f in cleaned if f["paragraph_index"] == idx])

        if progress_callback:
            progress_callback({
                "stage": _translate_for_settings("job.completed_paragraph", settings, current=stats['paragraphs_analyzed'], total=eligible_total),
                "current": stats["paragraphs_analyzed"],
                "total": eligible_total,
                "paragraph_index": idx,
                "stats": dict(stats),
            })

    final_findings = _clean_and_dedupe_findings(findings, paragraphs)
    if progress_callback:
        progress_callback({
            "stage": _translate_for_settings("job.analysis_complete", settings),
            "current": stats["paragraphs_analyzed"],
            "total": eligible_total,
            "stats": dict(stats),
        })
    return final_findings, stats


# -----------------------------
# Highlighting
# -----------------------------
_TRANSLATE_1TO1 = str.maketrans({
    "\u00A0": " ",
    "\u2009": " ",
    "\u202F": " ",
    "“": '"',
    "”": '"',
    "„": '"',
    "‟": '"',
    "’": "'",
    "‘": "'",
    "—": "-",
    "–": "-",
})


def _norm_1to1(s: str) -> str:
    return (s or "").translate(_TRANSLATE_1TO1)


def _find_quote_span(paragraph: str, quote: str) -> tuple[int, int] | None:
    if not paragraph or not quote:
        return None
    idx = paragraph.find(quote)
    if idx != -1:
        return (idx, idx + len(quote))
    pattern = re.escape(quote)
    pattern = re.sub(r"\\\s+", r"\\s+", pattern)
    try:
        m = re.search(pattern, paragraph, flags=re.IGNORECASE)
        if m:
            return (m.start(), m.end())
    except Exception:
        pass
    low_p = paragraph.lower()
    low_q = quote.lower()
    idx2 = low_p.find(low_q)
    if idx2 != -1:
        return (idx2, idx2 + len(quote))
    np = _norm_1to1(paragraph)
    nq = _norm_1to1(quote)
    idx3 = np.find(nq)
    if idx3 != -1:
        return (idx3, idx3 + len(nq))
    idx4 = np.lower().find(nq.lower())
    if idx4 != -1:
        return (idx4, idx4 + len(nq))
    pattern2 = re.escape(nq)
    pattern2 = re.sub(r"\\\s+", r"\\s+", pattern2)
    try:
        m2 = re.search(pattern2, np, flags=re.IGNORECASE)
        if m2:
            return (m2.start(), m2.end())
    except Exception:
        pass
    return None


def apply_highlights(
    paragraphs: list[str],
    findings: list[dict[str, Any]],
    include_reasoning: bool = False,
    locale: str | None = None,
) -> tuple[list[Markup], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    by_para: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for f in findings:
        by_para[f["paragraph_index"]].append(f)
    rendered: list[Markup] = []
    applied_findings: list[dict[str, Any]] = []
    skipped_findings: list[dict[str, Any]] = []
    paragraph_rows: list[dict[str, Any]] = []
    for i, p in enumerate(paragraphs):
        items = by_para.get(i, [])
        if not items:
            rendered_markup = Markup(escape(p))
            rendered.append(rendered_markup)
            paragraph_rows.append({"rendered": rendered_markup, "reasonings": []})
            continue
        spans: list[tuple[int, int, str, dict[str, Any]]] = []
        for f in items:
            span = _find_quote_span(p, f["quote"])
            if not span:
                skipped_findings.append(f)
                continue
            start_idx, end_idx = span
            spans.append((start_idx, end_idx, f["type"], f))
        if not spans:
            rendered_markup = Markup(escape(p))
            rendered.append(rendered_markup)
            paragraph_rows.append({"rendered": rendered_markup, "reasonings": []})
            continue
        spans.sort(key=lambda x: (x[0], -(x[1] - x[0])))
        accepted: list[tuple[int, int, str, dict[str, Any]]] = []
        last_end = -1
        for s, e, t, f in spans:
            if s < last_end:
                skipped_findings.append(f)
                continue
            accepted.append((s, e, t, f))
            last_end = e
        out = Markup("")
        cursor = 0
        reasonings: list[dict[str, Any]] = []
        local_index = 0
        for s, e, t, f in accepted:
            out += Markup(escape(p[cursor:s]))
            localized_label = str(
                _localized_fallacy_reference_for(t, locale=locale).get("display_name")
                or t
            ).strip() or t
            attrs = [
                f'class="fallacy-hl"',
                f'data-fallacy="{escape(localized_label)}"',
                f'title="{escape(localized_label)}"',
                f'aria-label="{escape(localized_label)}"',
            ]
            if include_reasoning:
                panel_id = f"reasoning-{i}-{local_index}"
                attrs.extend([
                    'tabindex="0"',
                    'role="button"',
                    f'aria-controls="{panel_id}"',
                    'aria-expanded="false"',
                    f'data-reasoning-target="{panel_id}"',
                ])
                confidence = float(f.get("confidence") or 0.0)
                reasonings.append({
                    "panel_id": panel_id,
                    "type": t,
                    "quote": f.get("quote") or "",
                    "explanation": (f.get("explanation") or "").strip() or "No additional reasoning was returned for this finding.",
                    "confidence_percent": max(0, min(100, int(round(confidence * 100)))) if confidence > 0 else 0,
                })
                local_index += 1
            out += Markup('<mark ' + ' '.join(attrs) + '>')
            out += Markup(escape(p[s:e]))
            out += Markup("</mark>")
            cursor = e
            applied_findings.append(f)
        out += Markup(escape(p[cursor:]))
        rendered.append(out)
        paragraph_rows.append({"rendered": out, "reasonings": reasonings})
    return rendered, applied_findings, skipped_findings, paragraph_rows



def summarize_findings(findings: list[dict[str, Any]], analysis_stats: dict[str, Any] | None = None, elapsed_seconds: float | None = None) -> dict[str, Any]:
    c = Counter([f["type"] for f in findings])
    summary = {"total": len(findings), "by_type": dict(c)}
    if analysis_stats:
        summary.update(analysis_stats)
    if elapsed_seconds is not None:
        summary["elapsed_seconds"] = max(0.0, float(elapsed_seconds))
        summary["elapsed_display"] = _format_elapsed_seconds(elapsed_seconds)
    return summary


# Fail-safe backup to ensure core functionality always works, even if localized JSON files (with complete Fallacy List) are missing or corrupted
DEFAULT_FALLACY_REFERENCE: dict[str, dict[str, str]] = {
    "Affirming the Consequent": {
        "icon": "🔁",
        "category": "Formal",
        "description": "Treats a result as proof of a particular cause, even though the same result could follow from other causes too."
    },
    "Denying the Antecedent": {
        "icon": "↪️",
        "category": "Formal",
        "description": "Assumes that if a condition is absent, the outcome must also be absent, even though the outcome could still arise in other ways."
    },
    "Undistributed Middle": {
        "icon": "🔺",
        "category": "Formal",
        "description": "Uses a shared middle term in a syllogism without properly connecting the categories, so the conclusion does not follow."
    },
    "Illicit Major": {
        "icon": "📏",
        "category": "Formal",
        "description": "Draws a syllogistic conclusion that distributes the major term more broadly than the premises allow."
    },
    "Illicit Minor": {
        "icon": "📐",
        "category": "Formal",
        "description": "Draws a syllogistic conclusion that distributes the minor term more broadly than the premises allow."
    },
    "Four-Term Fallacy": {
        "icon": "4️⃣",
        "category": "Formal",
        "description": "Uses a syllogism with four terms instead of three, often because a key term shifts meaning."
    },
    "Propositional Fallacies": {
        "icon": "∴",
        "category": "Formal",
        "description": "A broader label for invalid moves in propositional logic where the conclusion does not validly follow from the premises."
    },
    "Quantification Fallacies": {
        "icon": "🔢",
        "category": "Formal",
        "description": "Misuses words like all, some, or none so the conclusion goes beyond what the premises support."
    },
    "Syllogistic Fallacies": {
        "icon": "📚",
        "category": "Formal",
        "description": "Uses a broken syllogism or category relationship so the conclusion does not logically follow."
    },
    "Bad Reason Fallacy": {
        "icon": "🧩",
        "category": "Informal",
        "description": "Offers reasons that do not actually justify the conclusion being claimed."
    },
    "Sunk Cost Fallacy": {
        "icon": "🪙",
        "category": "Informal",
        "description": "Keeps defending or continuing something mainly because time, money, or effort has already been spent on it."
    },
    "Ad Hominem": {
        "icon": "👤",
        "category": "Informal",
        "description": "Attacks the person making the argument instead of addressing the argument itself."
    },
    "Ambiguity": {
        "icon": "🌫️",
        "category": "Informal",
        "description": "Relies on vague or slippery wording that can be taken in more than one way."
    },
    "Anecdotal": {
        "icon": "🗣️",
        "category": "Informal",
        "description": "Uses a personal story or isolated example as if it were enough to prove a broad claim."
    },
    "Appeal to Authority": {
        "icon": "🏛️",
        "category": "Informal",
        "description": "Treats a claim as true mainly because an authority figure or expert said it, without enough supporting evidence."
    },
    "Appeal to Emotion": {
        "icon": "💥",
        "category": "Informal",
        "description": "Pushes the audience toward a conclusion by provoking emotion rather than giving solid reasons."
    },
    "Appeal to Fear": {
        "icon": "😨",
        "category": "Informal",
        "description": "Leans on fear or alarm to drive acceptance of a claim instead of supporting it with sound reasoning."
    },
    "Appeal to Pity": {
        "icon": "🥺",
        "category": "Informal",
        "description": "Treats sympathy or compassion as if it were evidence that a claim is true or justified."
    },
    "Appeal to Force": {
        "icon": "🛑",
        "category": "Informal",
        "description": "Uses threats, coercion, or the prospect of punishment in place of actual reasons."
    },
    "Appeal to Ignorance": {
        "icon": "🌑",
        "category": "Informal",
        "description": "Claims something is true or false simply because it has not been proven otherwise."
    },
    "Appeal to Nature": {
        "icon": "🌿",
        "category": "Informal",
        "description": "Assumes something is good, right, or safer simply because it is called natural."
    },
    "Appeal to Probability": {
        "icon": "🎲",
        "category": "Informal",
        "description": "Assumes that because something could happen or seems likely, it therefore will happen."
    },
    "Appeal to Ridicule": {
        "icon": "😏",
        "category": "Informal",
        "description": "Mocks a position to make it seem false instead of actually refuting it."
    },
    "Appeal to Tradition": {
        "icon": "🏺",
        "category": "Informal",
        "description": "Claims something is correct or preferable mainly because it has long been done that way."
    },
    "Appeal to Consequences": {
        "icon": "🧭",
        "category": "Informal",
        "description": "Treats desirable or undesirable consequences as proof that a claim itself is true or false."
    },
    "Argument from Repetition": {
        "icon": "📣",
        "category": "Informal",
        "description": "Repeats a claim so often that the repetition itself is used to make it feel true."
    },
    "Argumentum ad Populum": {
        "icon": "👥",
        "category": "Informal",
        "description": "Treats popularity or public approval as evidence that a claim is true."
    },
    "Bandwagon": {
        "icon": "🚩",
        "category": "Informal",
        "description": "Urges acceptance of a claim because many people supposedly already accept it."
    },
    "Begging the Question": {
        "icon": "⭕",
        "category": "Informal",
        "description": "Builds the conclusion into the premise, so the argument assumes what it is trying to prove."
    },
    "Burden of Proof": {
        "icon": "⚖️",
        "category": "Informal",
        "description": "Shifts responsibility for proving a claim onto other people instead of supporting it directly."
    },
    "Cherry Picking": {
        "icon": "🍒",
        "category": "Informal",
        "description": "Selects only the evidence that supports a conclusion while ignoring relevant evidence that points the other way."
    },
    "Circular Reasoning": {
        "icon": "🔄",
        "category": "Informal",
        "description": "Restates the conclusion in different words instead of supplying independent support for it."
    },
    "Composition Fallacy": {
        "icon": "🧱",
        "category": "Informal",
        "description": "Assumes that what is true of the parts must also be true of the whole."
    },
    "Continuum Fallacy": {
        "icon": "📏",
        "category": "Informal",
        "description": "Claims a distinction is meaningless just because the boundary between cases is gradual or fuzzy."
    },
    "Division Fallacy": {
        "icon": "🧩",
        "category": "Informal",
        "description": "Assumes that what is true of the whole must also be true of each part."
    },
    "Equivocation": {
        "icon": "🪞",
        "category": "Informal",
        "description": "Switches the meaning of a key word or phrase in the middle of an argument."
    },
    "Etymological Fallacy": {
        "icon": "📚",
        "category": "Informal",
        "description": "Assumes a word’s current meaning or truth is determined by its historical origin."
    },
    "Fallacy Fallacy": {
        "icon": "🪤",
        "category": "Informal",
        "description": "Assumes a claim must be false merely because the argument for it was weak or fallacious."
    },
    "Fallacy of Composition and Division": {
        "icon": "🧱",
        "category": "Informal",
        "description": "Improperly transfers a property from parts to the whole, or from the whole to the parts."
    },
    "Fallacy of Quoting Out of Context": {
        "icon": "✂️",
        "category": "Informal",
        "description": "Pulls words from their surrounding context so they appear to mean something different."
    },
    "False Analogy": {
        "icon": "⚠️",
        "category": "Informal",
        "description": "Treats two things as more alike than they really are, using a weak comparison to support the conclusion."
    },
    "False Attribution": {
        "icon": "🏷️",
        "category": "Informal",
        "description": "Assigns a statement, motive, source, or cause to the wrong person or factor."
    },
    "False Cause": {
        "icon": "🔗",
        "category": "Informal",
        "description": "Treats two things as causally connected without enough evidence that one really caused the other."
    },
    "False Cause & False Attribution": {
        "icon": "🔗",
        "category": "Informal",
        "description": "Treats two things as causally connected, or assigns credit or blame, without enough evidence."
    },
    "False Dilemma": {
        "icon": "🚪",
        "category": "Informal",
        "description": "Frames a situation as if there were only two options when more possibilities exist."
    },
    "False Equivalence": {
        "icon": "⚖️",
        "category": "Informal",
        "description": "Treats two things as morally, logically, or practically equivalent when the differences matter."
    },
    "Faulty Generalization": {
        "icon": "📉",
        "category": "Informal",
        "description": "Draws a broad rule or conclusion from too little, too narrow, or unrepresentative evidence."
    },
    "Furtive Fallacy": {
        "icon": "🕵️",
        "category": "Informal",
        "description": "Assumes hidden intent, secret coordination, or deliberate design without sufficient evidence."
    },
    "Gambler's Fallacy": {
        "icon": "🎰",
        "category": "Informal",
        "description": "Assumes past random outcomes make a different future outcome more likely, even when events are independent."
    },
    "Genetic Fallacy": {
        "icon": "🧬",
        "category": "Informal",
        "description": "Judges a claim by where it came from rather than by the quality of the claim itself."
    },
    "Guilt by Association": {
        "icon": "🕸️",
        "category": "Informal",
        "description": "Dismisses a person or claim by linking it to someone or something disliked rather than addressing the reasoning itself."
    },
    "Hasty Generalization": {
        "icon": "📉",
        "category": "Informal",
        "description": "Jumps from a small or unrepresentative sample to a sweeping conclusion."
    },
    "Ignoratio Elenchi": {
        "icon": "🎯",
        "category": "Informal",
        "description": "Presents reasons that may sound relevant but do not actually answer the point at issue."
    },
    "Incomplete Comparison": {
        "icon": "⚗️",
        "category": "Informal",
        "description": "Makes a comparison without stating what the thing is being compared against or by which standard."
    },
    "Inflation of Conflict": {
        "icon": "🔥",
        "category": "Informal",
        "description": "Presents a disagreement as a dramatic clash when the actual gap may be smaller or more nuanced."
    },
    "Kettle Logic": {
        "icon": "🫖",
        "category": "Informal",
        "description": "Uses mutually inconsistent defenses all at once, hoping one of them will stick."
    },
    "Loaded Question": {
        "icon": "❓",
        "category": "Informal",
        "description": "Asks a question that smuggles in an unproven assumption."
    },
    "Middle Ground": {
        "icon": "🤝",
        "category": "Informal",
        "description": "Assumes the compromise position must be correct simply because it sits between two extremes."
    },
    "Motte and Bailey": {
        "icon": "🏰",
        "category": "Informal",
        "description": "Retreats to a safer, easier claim when challenged, then returns to the stronger controversial claim once the pressure passes."
    },
    "Moving the Goalposts": {
        "icon": "🥅",
        "category": "Informal",
        "description": "Changes the standard of proof or success after the original standard has been met."
    },
    "Nirvana Fallacy": {
        "icon": "✨",
        "category": "Informal",
        "description": "Rejects a realistic option because it is not perfect, comparing it unfairly to an idealized alternative."
    },
    "No True Scotsman": {
        "icon": "🛡️",
        "category": "Informal",
        "description": "Protects a general claim by redefining the group whenever a counterexample appears."
    },
    "Personal Incredulity": {
        "icon": "🤨",
        "category": "Informal",
        "description": "Treats something as false simply because it seems hard to imagine, understand, or believe."
    },
    "Poisoning the Well": {
        "icon": "☠️",
        "category": "Informal",
        "description": "Discredits a person in advance so that anything they say is dismissed before it is even considered."
    },
    "Post Hoc": {
        "icon": "⏱️",
        "category": "Informal",
        "description": "Assumes that because one thing happened after another, the first thing must have caused the second."
    },
    "Proof by Verbosity": {
        "icon": "🧵",
        "category": "Informal",
        "description": "Uses a flood of words, complexity, or detail to create the impression of proof without actually proving the point."
    },
    "Proving Too Much": {
        "icon": "📣",
        "category": "Informal",
        "description": "Uses reasoning that, if accepted, would justify far more than the speaker likely intends."
    },
    "Red Herring": {
        "icon": "🐟",
        "category": "Informal",
        "description": "Introduces a distraction that pulls attention away from the main issue."
    },
    "Reification": {
        "icon": "🗿",
        "category": "Informal",
        "description": "Treats an abstract idea or label as if it were a concrete thing with independent existence."
    },
    "Retrospective Determinism": {
        "icon": "🕰️",
        "category": "Informal",
        "description": "Looks back at past events as if the outcome had been inevitable all along."
    },
    "Shotgun Argumentation": {
        "icon": "💬",
        "category": "Informal",
        "description": "Throws out many weak points at once so they become difficult to answer individually."
    },
    "Slippery Slope": {
        "icon": "🛝",
        "category": "Informal",
        "description": "Claims a relatively small first step will inevitably trigger a chain of extreme consequences."
    },
    "Special Pleading": {
        "icon": "🎟️",
        "category": "Informal",
        "description": "Applies standards selectively, carving out an exception when the rule becomes inconvenient."
    },
    "Strawman": {
        "icon": "🌾",
        "category": "Informal",
        "description": "Misrepresents an opposing view so it becomes easier to dismiss."
    },
    "Texas Sharpshooter": {
        "icon": "🎯",
        "category": "Informal",
        "description": "Cherry-picks a pattern or cluster after the fact while ignoring the larger body of evidence."
    },
    "Tu Quoque": {
        "icon": "↩️",
        "category": "Informal",
        "description": "Responds to criticism by accusing the critic of similar behavior instead of addressing the criticism."
    }
}
DEFAULT_FALLACY_ALIASES: dict[str, str] = {
    "Ad Populum": "Argumentum ad Populum",
    "Appeal to Majority": "Argumentum ad Populum",
    "Appeal to Popularity": "Argumentum ad Populum",
    "Appeal to the Majority": "Argumentum ad Populum",
    "Appeal to the People": "Argumentum ad Populum",
    "Circular Argument": "Circular Reasoning",
    "Circular Logic": "Circular Reasoning",
    "Correlation Implies Causation": "False Cause",
    "Division": "Division Fallacy",
    "Fallacy of Composition": "Composition Fallacy",
    "Fallacy of Division": "Division Fallacy",
    "False Choice": "False Dilemma",
    "False Dichotomy": "False Dilemma",
    "Faulty Analogy": "False Analogy",
    "Gish Gallop": "Shotgun Argumentation",
    "Loaded Questions": "Loaded Question",
    "Post Hoc Ergo Propter Hoc": "Post Hoc",
    "Quoting Out of Context": "Fallacy of Quoting Out of Context",
    "Straw Man": "Strawman",
    "Weak Analogy": "False Analogy"
}
DEFAULT_FALLACY_FALLBACK: dict[str, str] = {
    "icon": "⚠️",
    "category": "",
    "description": "This detected label is not yet defined in the fallacy reference library. Review the highlighted passage directly instead of relying on a category badge."
}


def _fallacy_library_translation_path(locale: str = "en") -> str:
    os.makedirs(_translations_dir(), exist_ok=True)
    return os.path.join(_translations_dir(), f"fallacies_{_sanitize_ui_language(locale)}.json")


# Fail-safe search for legacy Fallacy list (pre-localization)
def _fallacy_library_instance_path() -> str:
    os.makedirs(INSTANCE_DIR, exist_ok=True)
    return str(INSTANCE_DIR / "fallacies.json")


# Fail-safe search for legacy Fallacy list (pre-localization)
def _fallacy_library_static_path() -> str:
    return os.path.join(app.root_path, "static", "fallacies.json")


# Fail-safe search for legacy Fallacy list (pre-localization)
def _fallacy_library_local_path() -> str:
    return os.path.join(app.root_path, "fallacies.json")


def _default_fallacy_library_payload() -> dict[str, Any]:
    return {
        "version": 1,
        "updated_at": "2026-03-30",
        "fallback": dict(DEFAULT_FALLACY_FALLBACK),
        "fallacies": {name: dict(meta) for name, meta in DEFAULT_FALLACY_REFERENCE.items()},
        "aliases": dict(DEFAULT_FALLACY_ALIASES),
    }


def _write_default_fallacy_library_if_missing(locale: str = "en") -> None:
    path = _fallacy_library_translation_path(locale)
    if os.path.exists(path):
        return
    with open(path, "w", encoding="utf-8") as f:
        json.dump(_default_fallacy_library_payload(), f, ensure_ascii=False, indent=2)


def _load_json_file(path: str) -> dict[str, Any] | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        return payload if isinstance(payload, dict) else None
    except Exception:
        return None


def _coerce_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, (list, tuple, set)):
        return []
    items: list[str] = []
    seen: set[str] = set()
    for item in value:
        clean_item = " ".join(str(item or "").split())
        if clean_item and clean_item not in seen:
            seen.add(clean_item)
            items.append(clean_item)
    return items


def _merge_unique_strings(*groups: Any) -> list[str]:
    merged: list[str] = []
    seen: set[str] = set()
    for group in groups:
        for item in _coerce_string_list(group):
            normalized = item.casefold()
            if normalized in seen:
                continue
            seen.add(normalized)
            merged.append(item)
    return merged


def _coerce_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value or "").strip().lower() in {"1", "true", "yes", "y", "on"}


def _slugify_fallacy_name(name: str) -> str:
    cleaned = _normalize_fallacy_name(name).casefold()
    cleaned = cleaned.replace("&", " and ")
    cleaned = re.sub(r"[^a-z0-9]+", "-", cleaned)
    return cleaned.strip("-") or "fallacy"


FALLACY_RELATION_STOPWORDS = {
    "the", "and", "that", "with", "from", "into", "when", "where", "what", "than", "then", "they",
    "them", "this", "those", "these", "because", "while", "does", "dont", "just", "have", "been",
    "being", "were", "their", "there", "would", "could", "should", "about", "through", "across",
    "using", "treats", "claim", "claims", "argument", "arguments", "conclusion", "conclusions",
    "reason", "reasons", "fallacy", "fallacies", "instead", "rather", "into", "over", "under",
    "more", "most", "very", "much", "some", "such", "only", "mainly", "really", "other", "others",
    "itself", "it", "its", "your", "ours", "theirs", "also", "make", "makes", "made", "same",
}


def _relation_tokens_from_text(text: str) -> set[str]:
    tokens = {
        token for token in re.findall(r"[a-z0-9']+", (text or "").casefold())
        if len(token) >= 4 and token not in FALLACY_RELATION_STOPWORDS
    }
    return tokens


def _explicit_relation_names(item: dict[str, Any], key: str) -> set[str]:
    names: set[str] = set()
    for raw_name in item.get(key) or []:
        canonical_name = _resolve_fallacy_name(str(raw_name or ""))
        if canonical_name:
            names.add(canonical_name)
    return names


def _text_similarity_score(current_item: dict[str, Any], candidate_item: dict[str, Any]) -> int:
    current_tokens = set(current_item.get("relation_tokens") or [])
    candidate_tokens = set(candidate_item.get("relation_tokens") or [])
    if not current_tokens or not candidate_tokens:
        return 0
    overlap = current_tokens & candidate_tokens
    if not overlap:
        return 0
    ratio = (2 * len(overlap)) / (len(current_tokens) + len(candidate_tokens))
    return max(0, min(20, int(round(ratio * 30))))


def _history_cooccurrence_scores(current_name: str) -> dict[str, dict[str, Any]]:
    current_canonical = _resolve_fallacy_name(current_name)
    if not current_canonical:
        return {}
    db_path = _history_db_path()
    if not os.path.exists(db_path):
        return {}

    counts: Counter[str] = Counter()
    current_runs = 0
    try:
        with _history_connect() as conn:
            rows = conn.execute("SELECT fallacy_breakdown_json FROM saved_analyses").fetchall()
    except Exception:
        return {}

    for row in rows:
        breakdown = _json_loads(row["fallacy_breakdown_json"], {})
        if not isinstance(breakdown, dict):
            continue
        present = {
            canonical for canonical in (
                _resolve_fallacy_name(str(name or "")) for name in breakdown.keys()
            )
            if canonical and canonical in FALLACY_REFERENCE
        }
        if current_canonical not in present:
            continue
        current_runs += 1
        for candidate_name in present:
            if candidate_name != current_canonical:
                counts[candidate_name] += 1

    if current_runs <= 0:
        return {}

    scored: dict[str, dict[str, Any]] = {}
    for candidate_name, count in counts.items():
        ratio = count / current_runs
        score = max(0, min(34, int(round((count * 6) + (ratio * 10)))))
        scored[candidate_name] = {
            "score": score,
            "count": count,
            "ratio": ratio,
            "label": f"Seen together in saved analyses ({count})",
        }
    return scored


def _relation_badge_class(reason: str) -> str:
    label = (reason or "").strip()
    if label == "Logically close":
        return "ff-badge-relation-logically-close"
    if label == "Often paired":
        return "ff-badge-relation-often-paired"
    if label == "Same family":
        return "ff-badge-relation-same-family"
    if label == "Often confused":
        return "ff-badge-relation-often-confused"
    if label == "Similar wording":
        return "ff-badge-relation-similar-wording"
    if label.startswith("Seen together in saved analyses"):
        return "ff-badge-relation-seen-together"
    if label == "Same category":
        return "ff-badge-relation-same-category"
    return "ff-badge-relation"


def _related_fallacy_candidates(current_item: dict[str, Any], limit: int | None = 6) -> list[dict[str, Any]]:
    current_name = current_item["name"]
    current_slug = current_item["slug"]
    current_category = (current_item.get("category") or "").strip()
    current_related = _explicit_relation_names(current_item, "related")
    current_often_with = _explicit_relation_names(current_item, "often_with")
    current_confused = _explicit_relation_names(current_item, "confused_with")
    current_family = {family.casefold() for family in (current_item.get("same_family") or []) if family}
    history_scores = _history_cooccurrence_scores(current_name)

    candidates: list[dict[str, Any]] = []
    for candidate in FALLACY_CATALOG:
        if candidate["slug"] == current_slug:
            continue

        candidate_name = candidate["name"]
        candidate_related = _explicit_relation_names(candidate, "related")
        candidate_often_with = _explicit_relation_names(candidate, "often_with")
        candidate_confused = _explicit_relation_names(candidate, "confused_with")
        candidate_family = {family.casefold() for family in (candidate.get("same_family") or []) if family}

        score = 0
        reasons: list[str] = []

        if candidate_name in current_related or current_name in candidate_related:
            score += 100
            reasons.append("Logically close")

        if candidate_name in current_often_with or current_name in candidate_often_with:
            score += 80
            reasons.append("Often paired")

        family_overlap = sorted(current_family & candidate_family)
        if family_overlap:
            score += min(70, 28 * len(family_overlap))
            reasons.append("Same family")

        if candidate_name in current_confused or current_name in candidate_confused:
            score += 24
            reasons.append("Often confused")

        text_score = _text_similarity_score(current_item, candidate)
        if text_score:
            score += text_score
            if text_score >= 8:
                reasons.append("Similar wording")

        history_meta = history_scores.get(candidate_name)
        if history_meta and history_meta.get("score"):
            score += int(history_meta["score"])
            reasons.append(history_meta["label"])

        if (candidate.get("category") or "").strip() == current_category:
            score += 6
            reasons.append("Same category")

        if score <= 0:
            continue

        unique_reasons: list[str] = []
        for reason in reasons:
            if reason and reason not in unique_reasons:
                unique_reasons.append(reason)

        relation_summary = " · ".join(unique_reasons[:3])
        relation_badges = [
            {"label": reason, "class_name": _relation_badge_class(reason)}
            for reason in unique_reasons
        ]
        candidate_entry = dict(candidate)
        candidate_entry.update({
            "relation_score": score,
            "relation_reasons": unique_reasons,
            "relation_badges": relation_badges,
            "relation_summary": relation_summary,
            "shared_family": family_overlap,
            "history_overlap_count": int((history_meta or {}).get("count") or 0),
        })
        candidates.append(candidate_entry)

    candidates.sort(
        key=lambda item: (
            -int(item.get("relation_score") or 0),
            not bool(item.get("common")),
            item["name"].casefold(),
        )
    )
    if limit is None:
        return candidates
    return candidates[:max(0, int(limit))]


def _fallback_example_for(name: str, category: str) -> str:
    if (category or "").strip().lower() == "formal":
        return f"A speaker uses the structure of an argument incorrectly and still presents the conclusion as if it necessarily followed. This is a typical {name.lower()} pattern."
    return f"A speaker makes a claim in a way that feels persuasive at first, but the reasoning shortcut matches {name.lower()} rather than a sound argument."


def _fallback_explanation_for(description: str) -> str:
    base = (description or FALLACY_FALLBACK["description"]).strip()
    if not base:
        base = FALLACY_FALLBACK["description"]
    return base + " When reviewing an argument, focus on whether the conclusion is actually supported by reasons or whether it depends on this shortcut instead."


def _localized_fallback_explanation_for(description: str, locale: str = "en") -> str:
    base = (description or FALLACY_FALLBACK["description"]).strip()
    if not base:
        base = FALLACY_FALLBACK["description"]
    locale = _sanitize_ui_language(locale)
    suffixes = {
        "sr": " Kada procenjujete argument, usredsredite se na to da li zaključak zaista proizlazi iz navedenih razloga ili se umesto toga oslanja na ovu prečicu."
    }
    return base + suffixes.get(locale, " When reviewing an argument, focus on whether the conclusion is actually supported by reasons or whether it depends on this shortcut instead.")


def _coerce_fallacy_library_payload(payload: dict[str, Any] | None) -> dict[str, Any]:
    payload = payload or {}

    fallacies_raw = payload.get("fallacies")
    fallacies: dict[str, dict[str, Any]] = {}
    if isinstance(fallacies_raw, dict):
        for name, meta in fallacies_raw.items():
            clean_name = " ".join(str(name or "").split())
            if not clean_name or not isinstance(meta, dict):
                continue
            description = str(meta.get("description") or DEFAULT_FALLACY_FALLBACK["description"]).strip() or DEFAULT_FALLACY_FALLBACK["description"]
            category = " ".join(str(meta.get("category") or "").split())
            fallacies[clean_name] = {
                "display_name": str(meta.get("name") or meta.get("display_name") or "").strip(),
                "icon": str(meta.get("icon") or DEFAULT_FALLACY_FALLBACK["icon"]).strip() or DEFAULT_FALLACY_FALLBACK["icon"],
                "category": category,
                "description": description,
                "slug": str(meta.get("slug") or _slugify_fallacy_name(clean_name)).strip() or _slugify_fallacy_name(clean_name),
                "common": _coerce_bool(meta.get("common")),
                "short_for": str(meta.get("short_for") or "").strip(),
                "examples": _coerce_string_list(meta.get("examples") if meta.get("examples") is not None else meta.get("example")),
                "explanation": str(meta.get("explanation") or "").strip(),
                "aliases": _coerce_string_list(meta.get("aliases")),
                "keywords": _coerce_string_list(meta.get("keywords") if meta.get("keywords") is not None else meta.get("keyword")),
                "same_family": _coerce_string_list(meta.get("same_family") if meta.get("same_family") is not None else meta.get("family")),
                "related": _coerce_string_list(meta.get("related")),
                "often_with": _coerce_string_list(meta.get("often_with")),
                "confused_with": _coerce_string_list(meta.get("confused_with")),
            }
    if not fallacies:
        for name, meta in DEFAULT_FALLACY_REFERENCE.items():
            fallacies[name] = {
                "display_name": name,
                "icon": meta.get("icon") or DEFAULT_FALLACY_FALLBACK["icon"],
                "category": meta.get("category") or "",
                "description": meta.get("description") or DEFAULT_FALLACY_FALLBACK["description"],
                "slug": _slugify_fallacy_name(name),
                "common": False,
                "short_for": "",
                "examples": [],
                "explanation": "",
                "aliases": [],
                "keywords": [],
                "same_family": [],
                "related": [],
                "often_with": [],
                "confused_with": [],
            }

    aliases_raw = payload.get("aliases")
    aliases: dict[str, str] = {}
    if isinstance(aliases_raw, dict):
        for alias, canonical in aliases_raw.items():
            clean_alias = " ".join(str(alias or "").split())
            clean_canonical = " ".join(str(canonical or "").split())
            if clean_alias and clean_canonical:
                aliases[clean_alias] = clean_canonical
    if not aliases:
        aliases = dict(DEFAULT_FALLACY_ALIASES)

    for canonical_name, meta in fallacies.items():
        for alias_name in meta.get("aliases") or []:
            aliases.setdefault(alias_name, canonical_name)

    fallback_raw = payload.get("fallback")
    fallback = {
        "icon": DEFAULT_FALLACY_FALLBACK["icon"],
        "category": DEFAULT_FALLACY_FALLBACK["category"],
        "description": DEFAULT_FALLACY_FALLBACK["description"],
    }
    if isinstance(fallback_raw, dict):
        fallback["icon"] = str(fallback_raw.get("icon") or fallback["icon"]).strip() or fallback["icon"]
        fallback["category"] = " ".join(str(fallback_raw.get("category") or fallback["category"]).split())
        fallback["description"] = str(fallback_raw.get("description") or fallback["description"]).strip() or fallback["description"]

    return {
        "source_path": payload.get("_source_path"),
        "fallacies": fallacies,
        "aliases": aliases,
        "fallback": fallback,
    }


def _load_fallacy_library(locale: str = "en") -> dict[str, Any]:
    locale = _sanitize_ui_language(locale)
    cached = _FALLACY_LIBRARY_CACHE.get(locale)
    if cached is not None:
        return cached

    primary_path = _fallacy_library_translation_path(locale)
    legacy_paths = (
        _fallacy_library_local_path(),
        _fallacy_library_static_path(),
        _fallacy_library_instance_path(),
    )
    if locale == "en":
        search_paths = (primary_path, *legacy_paths)
    else:
        search_paths = (primary_path, _fallacy_library_translation_path("en"), *legacy_paths)

    for path in search_paths:
        if not os.path.exists(path):
            continue
        payload = _load_json_file(path)
        if payload is None:
            continue
        payload["_source_path"] = path
        coerced = _coerce_fallacy_library_payload(payload)
        _FALLACY_LIBRARY_CACHE[locale] = coerced
        return coerced

    _write_default_fallacy_library_if_missing("en")
    default_payload = _default_fallacy_library_payload()
    default_payload["_source_path"] = _fallacy_library_translation_path("en")
    coerced = _coerce_fallacy_library_payload(default_payload)
    _FALLACY_LIBRARY_CACHE[locale] = coerced
    return coerced


def _normalize_fallacy_name(name: str) -> str:
    return " ".join((name or "").replace("’", "'").replace("“", '"').replace("”", '"').split())


def _normalize_fallacy_lookup_key(name: str) -> str:
    return _normalize_fallacy_name(name).casefold()


FALLACY_LIBRARY = _load_fallacy_library("en")
FALLACY_REFERENCE: dict[str, dict[str, Any]] = FALLACY_LIBRARY["fallacies"]
FALLACY_ALIASES: dict[str, str] = FALLACY_LIBRARY["aliases"]
FALLACY_FALLBACK: dict[str, str] = FALLACY_LIBRARY["fallback"]
FALLACY_REFERENCE_SOURCE: str | None = FALLACY_LIBRARY.get("source_path")


def _build_fallacy_reference_index() -> dict[str, str]:
    index: dict[str, str] = {}
    for canonical_name, meta in FALLACY_REFERENCE.items():
        normalized_key = _normalize_fallacy_lookup_key(canonical_name)
        if normalized_key:
            index[normalized_key] = canonical_name
        for alias_name in meta.get("aliases") or []:
            normalized_alias = _normalize_fallacy_lookup_key(alias_name)
            if normalized_alias:
                index[normalized_alias] = canonical_name
    for alias_name, canonical_name in FALLACY_ALIASES.items():
        normalized_alias = _normalize_fallacy_lookup_key(alias_name)
        if normalized_alias and canonical_name in FALLACY_REFERENCE:
            index[normalized_alias] = canonical_name
    return index


FALLACY_REFERENCE_INDEX: dict[str, str] = _build_fallacy_reference_index()


def _resolve_fallacy_name(name: str) -> str:
    cleaned_name = _normalize_fallacy_name(name)
    if not cleaned_name:
        return ""
    return FALLACY_REFERENCE_INDEX.get(_normalize_fallacy_lookup_key(cleaned_name), cleaned_name)


def _aliases_for_canonical_name(canonical_name: str) -> list[str]:
    ref = FALLACY_REFERENCE.get(canonical_name) or {}
    explicit_aliases = _coerce_string_list(ref.get("aliases"))
    mapped_aliases = [
        alias_name
        for alias_name, mapped_name in FALLACY_ALIASES.items()
        if _resolve_fallacy_name(mapped_name) == canonical_name
    ]
    return _merge_unique_strings(explicit_aliases, mapped_aliases)


def _fallacy_reference_for(name: str) -> dict[str, Any]:
    canonical_name = _resolve_fallacy_name(name)
    ref = FALLACY_REFERENCE.get(canonical_name)
    if ref:
        examples = _coerce_string_list(ref.get("examples"))
        description = ref.get("description") or FALLACY_FALLBACK["description"]
        aliases = _aliases_for_canonical_name(canonical_name)
        keywords = _merge_unique_strings(ref.get("keywords"), ref.get("short_for"))
        return {
            "canonical_name": canonical_name,
            "display_name": ref.get("display_name") or canonical_name,
            "icon": ref.get("icon") or FALLACY_FALLBACK["icon"],
            "category": ref.get("category") or "",
            "description": description,
            "slug": ref.get("slug") or _slugify_fallacy_name(canonical_name),
            "common": bool(ref.get("common")),
            "short_for": ref.get("short_for") or "",
            "aliases": aliases,
            "keywords": keywords,
            "examples": examples or [_fallback_example_for(canonical_name, ref.get("category") or "")],
            "explanation": (ref.get("explanation") or "").strip() or _fallback_explanation_for(description),
            "same_family": _coerce_string_list(ref.get("same_family")),
            "related": _coerce_string_list(ref.get("related")),
            "often_with": _coerce_string_list(ref.get("often_with")),
            "confused_with": _coerce_string_list(ref.get("confused_with")),
        }
    return {
        "canonical_name": canonical_name,
        "display_name": canonical_name,
        "icon": FALLACY_FALLBACK["icon"],
        "category": "",
        "description": FALLACY_FALLBACK["description"],
        "slug": _slugify_fallacy_name(canonical_name or "fallacy"),
        "common": False,
        "short_for": "",
        "aliases": [],
        "keywords": [],
        "examples": [_fallback_example_for(canonical_name or "fallacy", "")],
        "explanation": _fallback_explanation_for(FALLACY_FALLBACK["description"]),
        "same_family": [],
        "related": [],
        "often_with": [],
        "confused_with": [],
    }


def _canonical_fallacy_name_from_value(value: Any) -> str:
    if isinstance(value, dict):
        candidate = str(
            value.get("canonical_name")
            or value.get("name")
            or value.get("type")
            or value.get("display_name")
            or ""
        ).strip()
    else:
        candidate = str(value or "").strip()
    return _resolve_fallacy_name(candidate) if candidate else ""


def _localized_fallacy_reference_for(value: Any, locale: str | None = None) -> dict[str, Any]:
    active_locale = _sanitize_ui_language(locale or "en")
    canonical_name = _canonical_fallacy_name_from_value(value)

    payload: dict[str, Any]
    if isinstance(value, dict):
        payload = dict(value)
        if canonical_name:
            base_ref = _fallacy_reference_for(canonical_name)
            if base_ref:
                merged = dict(base_ref)
                merged.update(payload)
                payload = merged
    elif canonical_name:
        payload = _fallacy_reference_for(canonical_name)
    else:
        payload = {}

    if not payload:
        return payload

    payload.setdefault("canonical_name", canonical_name)
    payload.setdefault("display_name", canonical_name or str(value or "").strip())

    if active_locale == "en":
        return payload

    localized_library = _load_fallacy_library(active_locale)
    localized_ref = localized_library.get("fallacies", {}).get(canonical_name) if canonical_name else None
    localized_fallback = localized_library.get("fallback", {}) if isinstance(localized_library, dict) else {}

    merged = dict(payload)
    if isinstance(localized_ref, dict):
        if localized_ref.get("display_name"):
            merged["display_name"] = localized_ref.get("display_name")
        for field in ("icon", "description", "short_for", "explanation", "category"):
            localized_value = localized_ref.get(field)
            if str(localized_value or "").strip():
                merged[field] = localized_value
        for field in ("aliases", "keywords", "examples"):
            localized_list = _coerce_string_list(localized_ref.get(field))
            if localized_list:
                merged[field] = localized_list

    if canonical_name not in FALLACY_REFERENCE and str(localized_fallback.get("description") or "").strip():
        merged["description"] = str(localized_fallback.get("description")).strip()
    elif not str(merged.get("description") or "").strip() and str(localized_fallback.get("description") or "").strip():
        merged["description"] = str(localized_fallback.get("description")).strip()
    if not str(merged.get("explanation") or "").strip():
        merged["explanation"] = _localized_fallback_explanation_for(
            str(merged.get("description") or localized_fallback.get("description") or FALLACY_FALLBACK["description"]),
            locale=active_locale,
        )
    if not str(merged.get("display_name") or "").strip():
        merged["display_name"] = canonical_name or str(value or "").strip()
    return merged


def _build_fallacy_catalog() -> list[dict[str, Any]]:
    catalog: list[dict[str, Any]] = []
    for canonical_name in sorted(FALLACY_REFERENCE.keys(), key=lambda item: item.casefold()):
        ref = _fallacy_reference_for(canonical_name)
        relation_parts = [
            canonical_name,
            ref.get("category") or "",
            ref.get("description") or "",
            ref.get("explanation") or "",
            " ".join(ref.get("aliases") or []),
            " ".join(ref.get("keywords") or []),
            " ".join(ref.get("examples") or []),
            " ".join(ref.get("same_family") or []),
            " ".join(ref.get("related") or []),
            " ".join(ref.get("often_with") or []),
            " ".join(ref.get("confused_with") or []),
        ]
        relation_text = " ".join(part for part in relation_parts if part).casefold()
        library_search_name = canonical_name
        library_search_aliases = ref.get("aliases") or []
        library_search_category = ref.get("category") or ""
        library_search_keywords = ref.get("keywords") or []
        library_search_description = ref.get("description") or FALLACY_FALLBACK["description"]
        catalog.append({
            "name": canonical_name,
            "slug": ref.get("slug") or _slugify_fallacy_name(canonical_name),
            "icon": ref.get("icon") or FALLACY_FALLBACK["icon"],
            "category": ref.get("category") or "",
            "description": library_search_description,
            "examples": ref.get("examples") or [_fallback_example_for(canonical_name, ref.get("category") or "")],
            "explanation": ref.get("explanation") or _fallback_explanation_for(ref.get("description") or ""),
            "aliases": library_search_aliases,
            "keywords": library_search_keywords,
            "common": bool(ref.get("common")),
            "short_for": ref.get("short_for") or "",
            "same_family": ref.get("same_family") or [],
            "related": ref.get("related") or [],
            "often_with": ref.get("often_with") or [],
            "confused_with": ref.get("confused_with") or [],
            "library_search_name": library_search_name,
            "library_search_aliases": library_search_aliases,
            "library_search_category": library_search_category,
            "library_search_keywords": library_search_keywords,
            "library_search_description": library_search_description,
            "search_text": " ".join(part for part in [
                library_search_name,
                library_search_category,
                " ".join(library_search_aliases),
                " ".join(library_search_keywords),
                library_search_description,
            ] if part).casefold(),
            "relation_tokens": sorted(_relation_tokens_from_text(relation_text)),
        })
    return catalog


FALLACY_CATALOG: list[dict[str, Any]] = _build_fallacy_catalog()
FALLACY_BY_SLUG: dict[str, dict[str, Any]] = {item["slug"]: item for item in FALLACY_CATALOG}


def build_fallacy_cards(findings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    counts: Counter[str] = Counter()
    for finding in findings:
        raw_type = _normalize_fallacy_name(str(finding.get("type") or ""))
        if not raw_type:
            continue
        canonical_type = _resolve_fallacy_name(raw_type) or raw_type
        counts[canonical_type] += 1

    cards: list[dict[str, Any]] = []
    for fallacy_type, count in sorted(counts.items(), key=lambda item: (-item[1], item[0].lower())):
        ref = _fallacy_reference_for(fallacy_type)
        cards.append({
            "type": ref.get("canonical_name") or fallacy_type,
            "count": count,
            "icon": ref["icon"],
            "category": (ref.get("category") or "").strip(),
            "description": ref["description"],
            "slug": ref.get("slug") or _slugify_fallacy_name(fallacy_type),
            "common": bool(ref.get("common")),
        })
    return cards


def _related_fallacies_for(slug: str, category: str, limit: int | None = 6) -> list[dict[str, Any]]:
    current_item = FALLACY_BY_SLUG.get((slug or "").strip())
    if not current_item:
        return []
    return _related_fallacy_candidates(current_item, limit=limit)


def _fetch_and_extract_article(url: str) -> tuple[str, dict[str, Any], list[str]]:
    headers = {
        "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    resp = requests.get(url, headers=headers, timeout=20)
    resp.raise_for_status()
    html = resp.text
    text = trafilatura.extract(
        html,
        url=url,
        include_comments=False,
        include_tables=False,
        favor_recall=True,
        output_format="txt",
    )
    if not text:
        raise ValueError("Could not extract article text from that page (it may be paywalled, blocked, or heavily JS-rendered).")

    meta_doc = trafilatura.extract_metadata(html, default_url=url)
    meta = meta_doc.as_dict() if meta_doc else {}
    paragraphs = split_into_paragraphs(text)
    if len(paragraphs) <= 1:
        extracted_html = trafilatura.extract(
            html,
            url=url,
            include_comments=False,
            include_tables=False,
            favor_recall=True,
            output_format="html",
        )
        html_paras = paragraphs_from_extracted_html(extracted_html)
        if len(html_paras) > len(paragraphs):
            paragraphs = html_paras
    return html, meta, paragraphs


def _count_words(text: str) -> int:
    return len(re.findall(r"\S+", text or ""))


def _build_extraction_stats(paragraphs: list[str], elapsed_seconds: float) -> dict[str, Any]:
    text = "\n\n".join(paragraphs or [])
    stats = {
        "paragraph_count": len(paragraphs or []),
        "character_count": len(text),
        "word_count": _count_words(text),
        "elapsed_seconds": max(0.0, float(elapsed_seconds or 0.0)),
    }
    stats["elapsed_display"] = _format_elapsed_seconds(stats["elapsed_seconds"])
    return stats


def _render_review_html(
    job: dict[str, Any],
    settings: dict[str, Any],
    info: str | None = None,
    error: str | None = None,
    error_kind: str | None = None,
):
    return render_template(
        "review.html",
        job=job,
        url=job.get("url"),
        title=job.get("title") or "Extracted Article",
        author=job.get("author"),
        date=job.get("date"),
        paragraphs=job.get("paragraphs") or [],
        extraction_stats=job.get("extraction_stats") or {},
        settings=settings,
        info=info,
        error=error,
        error_kind=error_kind,
    )


def _render_result_html(job: dict[str, Any], settings: dict[str, Any], findings_all: list[dict[str, Any]], analysis_stats: dict[str, Any]):
    paragraphs = job.get("paragraphs") or []
    include_fallacy_reasoning = bool(settings.get("INCLUDE_FALLACY_REASONING"))
    rendered_paragraphs, applied_findings, skipped_findings, paragraph_rows = apply_highlights(
        paragraphs,
        findings_all,
        include_reasoning=include_fallacy_reasoning,
        locale=_current_locale(settings),
    )
    summary = summarize_findings(
        applied_findings,
        analysis_stats,
        elapsed_seconds=_job_elapsed_seconds(job),
    )
    fallacy_cards = build_fallacy_cards(applied_findings)
    llm_model = job.get("llm_model_label") or current_model_label(settings)
    history_status_message = job.get("history_status_message")
    source_kind = job.get("source_kind") or "url"
    template_name = "pasted.html" if source_kind == "pasted_text" else "article.html"
    title = job.get("title") or ("Pasted Text Analysis" if source_kind == "pasted_text" else "Extracted Article")
    return render_template(
        template_name,
        url=job.get("url"),
        title=title,
        author=job.get("author"),
        date=job.get("date"),
        paragraphs=paragraphs,
        rendered_paragraphs=rendered_paragraphs,
        paragraph_rows=paragraph_rows,
        findings=applied_findings,
        include_fallacy_reasoning=include_fallacy_reasoning,
        summary=summary,
        analysis_stats=analysis_stats,
        llm_model=llm_model,
        run_id=str(uuid.uuid4()),
        skipped_findings=skipped_findings,
        raw_findings=findings_all,
        fallacy_cards=fallacy_cards,
        history_status_message=history_status_message,
        history_overwrite_prompt=job.get("history_overwrite_prompt"),
        analysis_job_id=job.get("job_id"),
        extraction_stats=job.get("extraction_stats") or {},
        input_text=job.get("input_text") or "",
        settings=settings,
    )


def _render_result_html_for_job(job: dict[str, Any], settings: dict[str, Any], findings_all: list[dict[str, Any]], analysis_stats: dict[str, Any]) -> str:
    with app.app_context():
        with app.test_request_context("/"):
            return _render_result_html(job, settings, findings_all, analysis_stats)


def _run_extract_job(job_id: str, url: str) -> None:
    settings = load_settings()
    run_started_at = time.time()
    _update_job(
        job_id,
        phase="extract",
        started_at=run_started_at,
        finished_at=None,
        result_html=None,
        error=None,
        info_message=None,
        cancel_requested=False,
        cancel_message=None,
    )

    def cancel_check() -> bool:
        return _job_cancel_requested(job_id)

    def finalize_cancelled(stage: str | None = None) -> None:
        stage = stage or _translate_for_settings("job.cancelled", settings)
        _clear_job_stream_handle(job_id)
        _update_job(
            job_id,
            phase="extract",
            status="cancelled",
            stage=stage,
            progress=100,
            finished_at=time.time(),
            result_html=None,
        )

    try:
        if cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.")

        _update_job(job_id, phase="extract", status="running", stage=_translate_for_settings("job.fetching_article", settings), progress=12)
        _, meta, paragraphs = _fetch_and_extract_article(url)

        if cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.")

        extraction_elapsed_seconds = time.time() - run_started_at
        extraction_stats = _build_extraction_stats(paragraphs, extraction_elapsed_seconds)
        _update_job(
            job_id,
            phase="review",
            status="ready",
            stage=_translate_for_settings("job.extraction_complete_review", settings),
            progress=100,
            title=meta.get("title") or "Extracted Article",
            author=meta.get("author"),
            date=meta.get("date"),
            paragraphs=paragraphs,
            extraction_stats=extraction_stats,
            paragraphs_total=len(paragraphs),
            finished_at=time.time(),
        )
    except AnalysisCancelled:
        finalize_cancelled()
    except requests.RequestException as e:
        if cancel_check():
            finalize_cancelled()
        else:
            _update_job(job_id, phase="extract", status="error", stage=_translate_for_settings("job.fetch_failed", settings), progress=100, finished_at=time.time(), error=f"Fetch failed: {e}")
    except Exception as e:
        if cancel_check():
            finalize_cancelled()
        else:
            _update_job(job_id, phase="extract", status="error", stage=_translate_for_settings("job.extraction_failed", settings), progress=100, finished_at=time.time(), error=f"{e}")


def _run_analysis_job(job_id: str) -> None:
    settings = load_settings()
    job = _get_job(job_id)
    if not job:
        return
    source_kind = job.get("source_kind") or "url"
    url = str(job.get("url") or "")
    paragraphs = list(job.get("paragraphs") or [])
    if not paragraphs:
        error_phase = "analysis" if source_kind == "pasted_text" else "review"
        _update_job(job_id, phase=error_phase, status="error", stage=_translate_for_settings("job.analysis_failed", settings), progress=100, finished_at=time.time(), error="No extracted text is available for analysis.")
        return

    run_started_at = time.time()
    _update_job(
        job_id,
        phase="analysis",
        status="running",
        stage=_translate_for_settings("job.preparing_analysis", settings),
        progress=8,
        started_at=run_started_at,
        finished_at=None,
        result_html=None,
        error=None,
        error_kind=None,
        info_message=None,
        cancel_requested=False,
        cancel_message=None,
        paragraph_current=0,
        paragraph_total=0,
        paragraphs_analyzed=0,
        paragraphs_skipped_short=0,
        paragraphs_timed_out=0,
        paragraphs_failed=0,
    )

    def cancel_check() -> bool:
        return _job_cancel_requested(job_id)

    def finalize_cancelled(stage: str | None = None) -> None:
        stage = stage or _translate_for_settings("job.analysis_cancelled", settings)
        _clear_job_stream_handle(job_id)
        if source_kind == "pasted_text":
            _update_job(
                job_id,
                phase="analysis",
                status="cancelled",
                stage=stage,
                progress=100,
                finished_at=time.time(),
                result_html=None,
                cancel_requested=False,
                cancel_message=None,
                info_message=_translate_for_settings("job.analysis_cancelled", settings),
            )
            return
        _update_job(
            job_id,
            phase="review",
            status="ready",
            stage=_translate_for_settings("job.extraction_complete_review", settings),
            progress=100,
            finished_at=time.time(),
            result_html=None,
            cancel_requested=False,
            cancel_message=None,
            info_message=_translate_for_settings("job.analysis_cancelled", settings),
        )

    try:
        if cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.")

        _update_job(
            job_id,
            phase="analysis",
            status="running",
            stage=_translate_for_settings("job.checking_provider_connection", settings),
            progress=10,
        )
        ensure_provider_connection(settings, attempts=2, connect_timeout=3.0, read_timeout=3.0)

        def report_analysis_progress(info: dict[str, Any]) -> None:
            total = max(0, int(info.get("total") or 0))
            current = max(0, int(info.get("current") or 0))
            if total > 0:
                ratio = min(1.0, current / total)
                progress = 14 + int(ratio * 80)
            else:
                progress = 92
            stats = info.get("stats") or {}
            next_status = "cancelling" if cancel_check() else "running"
            next_stage = info.get("stage") or _translate_for_settings("job.preparing_analysis", settings)
            if cancel_check():
                next_stage = _translate_for_settings("job.cancelling_analysis", settings)
            _update_job(
                job_id,
                phase="analysis",
                status=next_status,
                stage=next_stage,
                progress=progress,
                paragraph_current=current,
                paragraph_total=total,
                paragraphs_total=len(paragraphs),
                paragraphs_analyzed=stats.get("paragraphs_analyzed", current),
                paragraphs_skipped_short=stats.get("paragraphs_skipped_short", 0),
                paragraphs_timed_out=stats.get("paragraphs_timed_out", 0),
                paragraphs_failed=stats.get("paragraphs_failed", 0),
            )

        findings_all, analysis_stats = analyze_fallacies(
            paragraphs,
            url=url,
            settings=settings,
            progress_callback=report_analysis_progress,
            cancel_check=cancel_check,
            job_id=job_id,
        )

        if cancel_check():
            raise AnalysisCancelled("Analysis cancelled by user.")

        saved_to_history = False
        saved_record_id = None
        history_status_message = None
        history_overwrite_prompt = None
        pending_history_overwrite = None
        if source_kind == "url" and url:
            history_payload = _build_saved_analysis_payload(job, findings_all, analysis_stats, settings)
            _, _, model_key, model_label = _model_identity_from_settings(settings)
            existing_record = get_saved_analysis_record_for_model(history_payload["url_normalized"], model_key)
            if existing_record is None:
                saved_to_history, saved_record_id = save_analysis_record(history_payload)
                if saved_to_history:
                    history_status_message = _translate_for_settings(
                        "messages.analysis_saved_to_history",
                        settings,
                        model_label=model_label,
                    )
            else:
                saved_record_id = int(existing_record.get("id") or 0) or None
                if saved_analysis_differs(existing_record, history_payload):
                    history_overwrite_prompt = {
                        "message": _translate_for_settings(
                            "messages.analysis_differs_saved_prompt",
                            settings,
                            model_label=model_label,
                        ),
                        "record_id": saved_record_id,
                    }
                    pending_history_overwrite = {
                        "record_id": saved_record_id,
                        "payload": history_payload,
                        "model_label": model_label,
                    }
                else:
                    history_status_message = _translate_for_settings(
                        "messages.analysis_not_added_duplicate",
                        settings,
                        model_label=model_label,
                    )

        _update_job(
            job_id,
            phase="analysis",
            status="running",
            stage=_translate_for_settings("job.finalizing_output", settings),
            progress=96,
            llm_model_label=current_model_label(settings),
            history_status_message=history_status_message,
            history_record_id=saved_record_id,
            history_overwrite_prompt=history_overwrite_prompt,
            pending_history_overwrite=pending_history_overwrite,
            latest_findings=findings_all,
            latest_analysis_stats=analysis_stats,
        )

        result_html = _render_result_html_for_job(_get_job(job_id) or job, settings, findings_all, analysis_stats)

        _update_job(
            job_id,
            phase="analysis",
            status="done",
            stage=_translate_for_settings("job.analysis_complete", settings),
            progress=100,
            finished_at=time.time(),
            result_html=result_html,
            paragraph_current=analysis_stats.get("paragraphs_analyzed", 0),
            paragraph_total=analysis_stats.get("paragraphs_eligible", 0),
            paragraphs_total=len(paragraphs),
            paragraphs_analyzed=analysis_stats.get("paragraphs_analyzed", 0),
            paragraphs_skipped_short=analysis_stats.get("paragraphs_skipped_short", 0),
            paragraphs_timed_out=analysis_stats.get("paragraphs_timed_out", 0),
            paragraphs_failed=analysis_stats.get("paragraphs_failed", 0),
            llm_model_label=current_model_label(settings),
            history_status_message=history_status_message,
            history_record_id=saved_record_id,
            history_overwrite_prompt=history_overwrite_prompt,
            pending_history_overwrite=pending_history_overwrite,
            latest_findings=findings_all,
            latest_analysis_stats=analysis_stats,
        )
    except AnalysisCancelled:
        finalize_cancelled()
    except ProviderConnectionError as e:
        if cancel_check():
            finalize_cancelled()
        else:
            error_phase = "analysis" if source_kind == "pasted_text" else "review"
            _update_job(
                job_id,
                phase=error_phase,
                status="error",
                stage=_translate_for_settings("job.analysis_failed", settings),
                progress=100,
                finished_at=time.time(),
                error=str(e),
                error_kind="provider_connection",
            )
    except Exception as e:
        if cancel_check():
            finalize_cancelled()
        else:
            error_phase = "analysis" if source_kind == "pasted_text" else "review"
            _update_job(job_id, phase=error_phase, status="error", stage=_translate_for_settings("job.analysis_failed", settings), progress=100, finished_at=time.time(), error=f"{e}", error_kind=None)


def _start_extract_job(url: str) -> str:
    job_id = _create_job(url)
    thread = threading.Thread(target=_run_extract_job, args=(job_id, url), daemon=True)
    thread.start()
    return job_id


def _start_analysis_job(job_id: str) -> None:
    thread = threading.Thread(target=_run_analysis_job, args=(job_id,), daemon=True)
    thread.start()


def _job_result_path(job_id: str) -> str:
    job = _get_job(job_id) or {}
    if (job.get("source_kind") or "url") == "pasted_text":
        return url_for("paste_result", job_id=job_id)
    return url_for("extract_result", job_id=job_id)


def _create_pasted_text_job(pasted_text: str) -> str:
    cleaned_text = normalize_pasted_text(pasted_text)
    paragraphs = split_into_paragraphs(cleaned_text)
    job_id = _create_job("", source_kind="pasted_text")
    _update_job(
        job_id,
        source_kind="pasted_text",
        source_label="Pasted text",
        input_text=cleaned_text,
        phase="analysis",
        status="queued",
        stage=_translate_for_settings("job.queued", load_settings()),
        title="Pasted Text Analysis",
        author=None,
        date=None,
        paragraphs=paragraphs,
        extraction_stats=_build_extraction_stats(paragraphs, 0.0),
        paragraphs_total=len(paragraphs),
    )
    return job_id


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def home():
    settings = load_settings()
    info = request.args.get("info")
    return render_template("index.html", settings=settings, info=info)


def _no_store(resp):
    resp.headers["Cache-Control"] = "no-store, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


@app.get("/paste")
def paste_text_page():
    settings = load_settings()
    info = request.args.get("info")
    error = request.args.get("error")
    pasted_text = request.args.get("text") or ""
    return render_template("paste.html", settings=settings, info=info, error=error, pasted_text=pasted_text)


@app.get("/settings")
def settings_page():
    settings = load_settings()
    resp = make_response(
        render_template(
            "settings.html",
            settings=settings,
            ollama_models=[],
            model_error=None,
            save_error=None,
            saved=request.args.get("saved"),
        )
    )
    return _no_store(resp)


@app.post("/settings")
def save_settings_route():
    settings = parse_settings_form(request.form)
    models: list[str] = []
    model_error = None
    save_error = None
    try:
        provider_info = validate_provider_connection(settings)
        if provider_info.get("provider") == "ollama":
            models = provider_info.get("models") or []
        save_settings(settings)
        resp = redirect(url_for("settings_page", saved="1"))
        return _no_store(resp)
    except Exception as e:
        save_error = str(e)
        if (settings.get("AI_PROVIDER") or "ollama").strip().lower() == "ollama":
            try:
                models = list_ollama_models(settings["OLLAMA_BASE_URL"])
            except Exception:
                models = []
        resp = make_response(render_template("settings.html", settings=settings, ollama_models=models, model_error=model_error, save_error=save_error, saved=None))
        return _no_store(resp)


def _request_base_url(default: str) -> str:
    base_url = ""
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        if isinstance(payload, dict):
            base_url = str(payload.get("base_url") or "").strip()
    else:
        base_url = str(request.form.get("base_url") or "").strip()
    return (base_url or default).rstrip("/")


@app.post("/api/ollama/models")
def api_ollama_models():
    base_url = _request_base_url(load_settings()["OLLAMA_BASE_URL"])
    try:
        models = list_ollama_models(base_url)
        return jsonify({"ok": True, "base_url": base_url, "models": models})
    except Exception as e:
        return jsonify({"ok": False, "base_url": base_url, "error": str(e)}), 400


@app.post("/api/ollama/test")
def api_ollama_test():
    base_url = _request_base_url(load_settings()["OLLAMA_BASE_URL"])
    try:
        models = list_ollama_models(base_url)
        return jsonify({"ok": True, "base_url": base_url, "models_count": len(models), "models": models[:25]})
    except Exception as e:
        return jsonify({"ok": False, "base_url": base_url, "error": str(e)}), 400


@app.post("/api/settings/related-fallacy-count")
def api_related_fallacy_count_setting():
    payload = request.get_json(silent=True) if request.is_json else request.form
    payload = payload or {}
    requested_count = _sanitize_related_fallacy_count(payload.get("count"), default=DEFAULT_SETTINGS["RELATED_FALLACY_COUNT"])
    settings = load_settings()
    settings["RELATED_FALLACY_COUNT"] = requested_count
    save_settings(settings)
    return jsonify({"ok": True, "related_fallacy_count": requested_count})


@app.post("/api/openai/test")
def api_openai_test():
    payload = request.get_json(silent=True) if request.is_json else request.form
    payload = payload or {}
    base_url = str(payload.get("base_url") or load_settings()["OPENAI_BASE_URL"]).strip().rstrip("/")
    model_name = str(payload.get("model") or "").strip()
    try:
        settings = {
            "AI_PROVIDER": "openai",
            "OPENAI_BASE_URL": base_url,
            "OPENAI_MODEL": model_name,
        }
        info = validate_provider_connection(settings)
        return jsonify({
            "ok": True,
            "base_url": info.get("base_url") or base_url,
            "models_count": info.get("models_count") or 0,
            "models": (info.get("models") or [])[:25],
        })
    except Exception as e:
        return jsonify({"ok": False, "base_url": base_url, "error": str(e)}), 400


@app.post("/api/history/check")
def api_history_check():
    url = (request.form.get("url") or "").strip()
    if not is_valid_http_url(url):
        return jsonify({"ok": False, "error": "Please enter a valid http(s) URL (including https://)."}), 400
    settings = load_settings()
    _, _, model_key, model_label = _model_identity_from_settings(settings)
    summary = get_saved_history_summary(url, model_key=model_key)
    summary["current_model_label"] = model_label
    return jsonify({"ok": True, **summary})


@app.post("/api/history/use/<int:record_id>")
def api_history_use(record_id: int):
    settings = load_settings()
    try:
        job_id = load_saved_analysis_into_job(record_id, settings)
        return jsonify({
            "ok": True,
            "job_id": job_id,
            "result_url": url_for("extract_result", job_id=job_id),
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 404


@app.post("/api/history/resolve/<job_id>")
def api_history_resolve(job_id: str):
    job = _get_job(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Job not found or expired."}), 404

    settings = load_settings()
    pending = job.get("pending_history_overwrite") or {}
    record_id = int(pending.get("record_id") or 0)
    payload = pending.get("payload") or {}
    model_label = pending.get("model_label") or current_model_label(settings)
    if not record_id or not payload:
        return jsonify({
            "ok": False,
            "error": _translate_for_settings("messages.history_resolution_not_required", settings),
        }), 409

    decision = str(request.form.get("decision") or (request.get_json(silent=True) or {}).get("decision") or "").strip().lower()
    if decision not in {"overwrite", "keep"}:
        return jsonify({"ok": False, "error": _translate_for_settings("errors.invalid_history_resolution", settings)}), 400

    if decision == "overwrite":
        overwrite_analysis_record(record_id, payload)
        message = _translate_for_settings("messages.analysis_saved_over_existing", settings, model_label=model_label)
    else:
        message = _translate_for_settings("messages.analysis_kept_existing_saved_history", settings, model_label=model_label)

    findings_all = job.get("latest_findings") or []
    analysis_stats = job.get("latest_analysis_stats") or {}
    _update_job(
        job_id,
        history_status_message=message,
        history_record_id=record_id,
        history_overwrite_prompt=None,
        pending_history_overwrite=None,
    )
    refreshed_job = _get_job(job_id) or job
    result_html = _render_result_html_for_job(refreshed_job, settings, findings_all, analysis_stats)
    _update_job(job_id, result_html=result_html)

    return jsonify({
        "ok": True,
        "message": message,
        "record_id": record_id,
        "decision": decision,
    })


@app.post("/api/extract/start")
def api_extract_start():
    url = (request.form.get("url") or "").strip()
    force_rerun = (request.form.get("force_rerun") or "").strip() == "1"
    if not is_valid_http_url(url):
        return jsonify({"ok": False, "error": "Please enter a valid http(s) URL (including https://)."}), 400

    settings = load_settings()
    _, _, model_key, model_label = _model_identity_from_settings(settings)
    if not force_rerun:
        history_summary = get_saved_history_summary(url, model_key=model_key)
        if history_summary.get("exists"):
            return jsonify({
                "ok": False,
                "duplicate": True,
                "error": "This URL has already been analyzed.",
                "current_model_label": model_label,
                **history_summary,
            }), 409

    job_id = _start_extract_job(url)
    job = _get_job(job_id)
    r = make_response(jsonify(_job_status_payload(job)))
    return _no_store(r)


@app.post("/api/paste/start")
def api_paste_start():
    pasted_text = request.form.get("pasted_text") or ""
    error = validate_pasted_text(pasted_text)
    if error:
        return jsonify({"ok": False, "error": error}), 400

    job_id = _create_pasted_text_job(pasted_text)
    _start_analysis_job(job_id)
    job = _get_job(job_id)
    r = make_response(jsonify(_job_status_payload(job)))
    return _no_store(r)


@app.get("/api/extract/status/<job_id>")
def api_extract_status(job_id: str):
    job = _get_job(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Job not found or expired."}), 404
    r = make_response(jsonify(_job_status_payload(job)))
    return _no_store(r)


@app.post("/api/extract/cancel/<job_id>")
def api_extract_cancel(job_id: str):
    job = _get_job(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Job not found or expired."}), 404

    if job.get("status") == "ready":
        return jsonify({
            "ok": True,
            "job_id": job_id,
            "status": job.get("status"),
            "message": _translate_for_settings("job.nothing_running", load_settings()),
        })

    terminal_statuses = {"done", "error", "cancelled"}
    if job.get("status") in terminal_statuses:
        return jsonify({
            "ok": True,
            "job_id": job_id,
            "status": job.get("status"),
            "message": _translate_for_settings("job.job_finished", load_settings()),
        })

    settings = load_settings()
    cancelling_message = _translate_for_settings("job.cancelling_extraction", settings) if job.get("phase") == "extract" else _translate_for_settings("job.cancelling_analysis", settings)

    _update_job(
        job_id,
        cancel_requested=True,
        cancel_message=cancelling_message,
        status="cancelling",
        stage=cancelling_message,
    )

    aborted, provider, model = _abort_job_stream(job_id)

    if provider == "ollama" and model:
        ollama_unload_model(settings, model)

    refreshed_job = _get_job(job_id) or job
    return jsonify({
        "ok": True,
        "job_id": job_id,
        "status": refreshed_job.get("status"),
        "cancel_requested": True,
        "ollama_request_aborted": aborted,
        "message": cancelling_message,
    })


@app.post("/api/analyze/start/<job_id>")
def api_analyze_start(job_id: str):
    job = _get_job(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Job not found or expired."}), 404
    if not job.get("paragraphs"):
        return jsonify({"ok": False, "error": "No extracted text is available for analysis."}), 409
    if job.get("status") == "done":
        return jsonify({"ok": True, "job_id": job_id, "status": "done", "phase": "analysis"})
    if job.get("status") in {"running", "cancelling"}:
        return jsonify({"ok": False, "error": "A job is already running for this article."}), 409

    _start_analysis_job(job_id)
    started_job = _get_job(job_id) or job
    r = make_response(jsonify(_job_status_payload(started_job)))
    return _no_store(r)


@app.get("/review/<job_id>")
def review_extraction(job_id: str):
    job = _get_job(job_id)
    if not job:
        return make_response("Review not found or expired.", 404)
    if not job.get("paragraphs"):
        return make_response("Extraction is not ready yet.", 409)
    settings = load_settings()
    info = request.args.get("info") or job.get("info_message")
    error = request.args.get("error") or (job.get("error") if job.get("status") == "error" else None)
    error_kind = request.args.get("error_kind") or (job.get("error_kind") if job.get("status") == "error" else None)
    r = make_response(_render_review_html(job, settings, info=info, error=error, error_kind=error_kind))
    return _no_store(r)


@app.get("/result/<job_id>")
def extract_result(job_id: str):
    job = _get_job(job_id)
    if not job:
        return make_response("Result not found or expired.", 404)
    if job.get("status") == "error":
        return make_response(job.get("error") or "Analysis failed.", 500)
    if job.get("status") != "done" or not job.get("result_html"):
        return make_response("Result is not ready yet.", 409)
    r = make_response(job["result_html"])
    return _no_store(r)


@app.get("/paste/result/<job_id>")
def paste_result(job_id: str):
    job = _get_job(job_id)
    if not job:
        return make_response("Result not found or expired.", 404)
    if (job.get("source_kind") or "url") != "pasted_text":
        return redirect(url_for("extract_result", job_id=job_id))
    if job.get("status") == "error":
        return make_response(job.get("error") or "Analysis failed.", 500)
    if job.get("status") != "done" or not job.get("result_html"):
        return make_response("Result is not ready yet.", 409)
    r = make_response(job["result_html"])
    return _no_store(r)


@app.post("/paste")
def paste_text_analyze():
    pasted_text = request.form.get("pasted_text") or ""
    settings = load_settings()
    error = validate_pasted_text(pasted_text)
    if error:
        resp = make_response(render_template("paste.html", settings=settings, error=error, pasted_text=pasted_text))
        return _no_store(resp)

    job_id = _create_pasted_text_job(pasted_text)
    _run_analysis_job(job_id)
    job = _get_job(job_id)
    if not job:
        resp = make_response(render_template("paste.html", settings=settings, error="The pasted-text analysis job could not be completed.", pasted_text=pasted_text))
        return _no_store(resp)
    if job.get("status") == "error":
        resp = make_response(render_template("paste.html", settings=settings, error=job.get("error") or "Analysis failed.", error_kind=job.get("error_kind"), pasted_text=pasted_text))
        return _no_store(resp)
    if job.get("status") != "done" or not job.get("result_html"):
        resp = make_response(render_template("paste.html", settings=settings, error="The pasted-text analysis did not finish successfully.", pasted_text=pasted_text))
        return _no_store(resp)
    resp = make_response(job["result_html"])
    return _no_store(resp)


@app.post("/extract")
def extract_article():
    start_time = time.time()
    url = (request.form.get("url") or "").strip()
    force_rerun = (request.form.get("force_rerun") or "").strip() == "1"
    settings = load_settings()
    if not is_valid_http_url(url):
        resp = make_response(render_template("index.html", settings=settings, error="Please enter a valid http(s) URL (including https://).", url=url))
        return _no_store(resp)

    if not force_rerun:
        _, _, model_key, _ = _model_identity_from_settings(settings)
        history_summary = get_saved_history_summary(url, model_key=model_key)
        if history_summary.get("exists"):
            latest = history_summary.get("latest_record") or {}
            message = _translate_for_settings(
                "errors.url_already_analyzed_on",
                settings,
                analyzed_at=latest.get("analyzed_at"),
                model_label=latest.get("model_label"),
            )
            resp = make_response(render_template("index.html", settings=settings, error=message, url=url))
            return _no_store(resp)
    try:
        _, meta, paragraphs = _fetch_and_extract_article(url)
        extraction_stats = _build_extraction_stats(paragraphs, time.time() - start_time)
        job_id = _create_job(url)
        _update_job(
            job_id,
            phase="review",
            status="ready",
            stage=_translate_for_settings("job.extraction_complete_review", settings),
            progress=100,
            title=meta.get("title") or "Extracted Article",
            author=meta.get("author"),
            date=meta.get("date"),
            paragraphs=paragraphs,
            extraction_stats=extraction_stats,
            finished_at=time.time(),
        )
        job = _get_job(job_id) or {"job_id": job_id, "url": url, "paragraphs": paragraphs, "extraction_stats": extraction_stats}
        r = make_response(_render_review_html(job, settings))
        return _no_store(r)
    except requests.RequestException as e:
        r = make_response(render_template("index.html", settings=settings, error=f"Fetch failed: {e}", url=url))
        return _no_store(r)
    except Exception as e:
        r = make_response(render_template("index.html", settings=settings, error=f"Unexpected error: {e}", url=url))
        return _no_store(r)


@app.post("/api/extract")
def api_extract():
    start_time = time.time()
    url = (request.form.get("url") or "").strip()
    if not is_valid_http_url(url):
        return jsonify({"ok": False, "error": "Please enter a valid http(s) URL (including https://)."}), 400
    _, meta, paragraphs = _fetch_and_extract_article(url)
    extraction_stats = _build_extraction_stats(paragraphs, time.time() - start_time)
    return jsonify({
        "ok": True,
        "url": url,
        "title": meta.get("title") or "Extracted Article",
        "author": meta.get("author"),
        "date": meta.get("date"),
        "extraction_stats": extraction_stats,
        "paragraphs": paragraphs,
    })


@app.get("/fallacies")
def fallacy_library_page():
    settings = load_settings()
    formal_fallacies = [item for item in FALLACY_CATALOG if (item.get("category") or "").strip().lower() == "formal"]
    informal_fallacies = [item for item in FALLACY_CATALOG if (item.get("category") or "").strip().lower() != "formal"]
    return render_template(
        "fallacies.html",
        settings=settings,
        formal_fallacies=formal_fallacies,
        informal_fallacies=informal_fallacies,
        total_fallacies=len(FALLACY_CATALOG),
        total_formal=len(formal_fallacies),
        total_informal=len(informal_fallacies),
        total_common=sum(1 for item in FALLACY_CATALOG if item.get("common")),
    )


@app.get("/fallacies/<slug>")
def fallacy_detail_redirect(slug: str):
    return redirect(url_for("fallacy_detail_page", slug=slug), code=301)


@app.get("/fallacies/<slug>.html", endpoint="fallacy_detail_page")
def fallacy_detail_page(slug: str):
    fallacy = FALLACY_BY_SLUG.get((slug or "").strip())
    if not fallacy:
        abort(404)
    settings = load_settings()
    all_related_fallacies = _related_fallacies_for(fallacy["slug"], fallacy.get("category") or "", limit=None)
    related_display_count, related_display_options = _resolve_related_fallacy_display_count(
        settings.get("RELATED_FALLACY_COUNT"),
        len(all_related_fallacies),
    )
    return render_template(
        "fallacy_detail.html",
        settings=settings,
        fallacy=fallacy,
        related_fallacies=all_related_fallacies,
        related_fallacy_count=len(all_related_fallacies),
        related_display_count=related_display_count,
        related_display_options=related_display_options,
    )



def create_app() -> Flask:
    global _APP_CONFIGURED
    if _APP_CONFIGURED:
        return app

    app.config.update(
        SECRET_KEY=_env_secret_key(),
        DEBUG=_env_bool("DEBUG", False) or _env_bool("FLASK_DEBUG", False),
        JSON_AS_ASCII=False,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE=os.getenv("SESSION_COOKIE_SAMESITE", "Lax"),
        SESSION_COOKIE_SECURE=_env_bool("SESSION_COOKIE_SECURE", False),
        PREFERRED_URL_SCHEME=os.getenv("PREFERRED_URL_SCHEME", "https"),
    )

    app.wsgi_app = ProxyFix(
        app.wsgi_app,
        x_for=_env_int("PROXY_FIX_X_FOR", 1),
        x_proto=_env_int("PROXY_FIX_X_PROTO", 1),
        x_host=_env_int("PROXY_FIX_X_HOST", 1),
        x_port=_env_int("PROXY_FIX_X_PORT", 1),
        x_prefix=_env_int("PROXY_FIX_X_PREFIX", 1),
    )

    @app.after_request
    def _add_security_headers(response):
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        response.headers.setdefault("Permissions-Policy", "geolocation=(), camera=(), microphone=()")
        return response

    ensure_dirs()
    _APP_CONFIGURED = True
    return app


app = create_app()


if __name__ == "__main__":
    app.run(
        host=os.getenv("HOST", "0.0.0.0"),
        port=_env_int("PORT", 8080),
        debug=bool(app.config.get("DEBUG")),
    )
