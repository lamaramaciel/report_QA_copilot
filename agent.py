"""Gemini URL-context evaluator for slide claims."""
from __future__ import annotations

import json
import re
import time
from typing import Any

import requests

DEFAULT_MODEL = "gemini-3.5-flash"
MAX_URLS = 20
URL_RE = re.compile(r'https?://[^\s|,\]>)"\']+')

MODEL_PRICES_USD_PER_MILLION = {
    # Update these values if the provider pricing changes.
    "gemini-3.5-flash": {"input": 1.50, "output": 9.00},
    "gemini-2.5-flash": {"input": 0.15, "output": 1.25},
}

RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "status": {"type": "string", "enum": ["CONFIRMED", "PARTIAL", "INCORRECT", "NOT_FOUND"]},
        "confidence": {"type": "integer", "minimum": 0, "maximum": 100},
        "verdict": {"type": "string"},
        "supported_claims": {"type": "array", "items": {"type": "string"}},
        "unsupported_claims": {"type": "array", "items": {"type": "string"}},
        "unverified_claims": {"type": "array", "items": {"type": "string"}},
        "direct_verbatims": {"type": "array", "items": {"type": "string"}},
        "evidence_by_source": {"type": "array", "items": {"type": "string"}},
        "qa_flags": {"type": "array", "items": {"type": "string"}},
        "suggested_fix": {"type": "string"},
    },
    "required": [
        "status", "confidence", "verdict", "supported_claims", "unsupported_claims",
        "unverified_claims", "direct_verbatims", "evidence_by_source", "qa_flags", "suggested_fix",
    ],
}


def extract_urls(raw: str) -> list[str]:
    out: list[str] = []
    for url in URL_RE.findall(raw or ""):
        cleaned = url.rstrip(".,;)")
        if cleaned and cleaned not in out:
            out.append(cleaned)
    return out[:MAX_URLS]


def _join(value: Any, limit: int = 1400) -> str:
    if isinstance(value, list):
        return "\n\n".join(str(x).strip()[:limit] for x in value if str(x).strip())
    return str(value or "").strip()[:limit]


def _response_text(data: dict) -> str:
    candidates = data.get("candidates") or []
    if not candidates:
        return ""
    parts = ((candidates[0].get("content") or {}).get("parts") or [])
    return "".join(str(p.get("text", "")) for p in parts).strip()


def _url_metadata(data: dict, requested: list[str]) -> list[dict]:
    candidates = data.get("candidates") or []
    metadata = {}
    if candidates:
        metadata = candidates[0].get("urlContextMetadata") or candidates[0].get("url_context_metadata") or {}
    items = metadata.get("urlMetadata") or metadata.get("url_metadata") or []
    output: list[dict] = []
    for idx, url in enumerate(requested):
        item = items[idx] if idx < len(items) and isinstance(items[idx], dict) else {}
        raw = str(item.get("urlRetrievalStatus") or item.get("url_retrieval_status") or "UNSPECIFIED")
        if "SUCCESS" in raw.upper():
            friendly = "Retrieved"
        elif any(x in raw.upper() for x in ("FAILED", "ERROR", "PAYWALL", "ROBOTS", "UNSAFE")):
            friendly = "Inaccessible"
        else:
            friendly = "Retrieval status unavailable"
        output.append({"source_number": idx + 1, "url": url, "retrieval_status": friendly, "raw_status": raw})
    return output


def _usage(data: dict, model: str) -> dict:
    usage = data.get("usageMetadata") or data.get("usage_metadata") or {}
    prompt = int(usage.get("promptTokenCount") or usage.get("prompt_token_count") or 0)
    candidates = int(usage.get("candidatesTokenCount") or usage.get("candidates_token_count") or 0)
    thoughts = int(usage.get("thoughtsTokenCount") or usage.get("thoughts_token_count") or 0)
    tool_input = int(usage.get("toolUsePromptTokenCount") or usage.get("tool_use_prompt_token_count") or 0)
    total = int(usage.get("totalTokenCount") or usage.get("total_token_count") or 0)
    prices = MODEL_PRICES_USD_PER_MILLION.get(model, {})
    estimated = None
    if prices:
        input_tokens = prompt + tool_input
        output_tokens = candidates + thoughts
        estimated = (input_tokens / 1_000_000) * prices["input"] + (output_tokens / 1_000_000) * prices["output"]
    return {
        "prompt_tokens": prompt,
        "tool_input_tokens": tool_input,
        "candidate_tokens": candidates,
        "thinking_tokens": thoughts,
        "total_tokens": total,
        "estimated_cost_usd": estimated,
    }


def _parse(text: str) -> tuple[dict, str]:
    try:
        value = json.loads(text)
        return value if isinstance(value, dict) else {}, ""
    except Exception:
        clean = re.sub(r"```json\s*|\s*```", "", text or "").strip()
        match = re.search(r"\{.*\}", clean, re.DOTALL)
        if not match:
            return {}, "Could not parse the model response"
        try:
            return json.loads(match.group()), ""
        except Exception as exc:
            return {}, f"JSON parse error: {str(exc)[:160]}"


def check_slide_claim(
    *,
    claim: str,
    slide_number: int,
    slide_title: str,
    surrounding_text: str,
    urls_raw: str,
    api_key: str,
    project_context: str = "",
    qa_instruction: str = "",
    model: str = DEFAULT_MODEL,
) -> dict:
    urls = extract_urls(urls_raw)
    if not claim.strip():
        return {"status": "⏭️ Skipped", "verdict": "No claim text", "sources": [], "usage": {}}
    if not urls:
        return {"status": "❓ No Reference", "verdict": "No valid source URL", "sources": [], "usage": {}}

    refs = "\n".join(f"[Source {i}] {u}" for i, u in enumerate(urls, start=1))
    prompt = f"""You are a strict evidence QA reviewer for client-facing consulting slides.
Evaluate only against the cited URLs. Do not use general knowledge to fill gaps.

Slide: {slide_number}
Slide title: {slide_title}
Claim to verify:
---
{claim[:3500]}
---
Surrounding slide text (context only; do not treat it as evidence):
---
{surrounding_text[:5000]}
---
Project context:
{project_context[:4000]}

Additional QA instruction:
{qa_instruction[:3000]}

Cited sources:
{refs}

Rules:
- CONFIRMED: every material part of the claim is directly supported by the combined cited evidence.
- PARTIAL: only part is supported, the wording is stronger/more specific than the evidence, or a reasonable inference is presented as fact.
- INCORRECT: the sources directly contradict a material part of the claim.
- NOT_FOUND: sources were retrieved but do not contain evidence for the material claim.
- Do not mark a claim confirmed merely because it is plausible.
- Treat source accessibility separately from factual correctness.
- Put a claim in unsupported_claims only when a retrieved source fails to support it or contradicts it.
- Put a claim in unverified_claims when it cannot be checked because the relevant source was not retrieved.
- If no cited source is retrieved, do not label the claim unsupported or incorrect.
- The suggested fix must preserve the slide's purpose and remain within what the sources support.
- Provide short exact verbatim excerpts from the retrieved pages where possible.
- In evidence_by_source, begin each item with [Source N].
"""

    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "tools": [{"url_context": {}}],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 8192,
            "responseMimeType": "application/json",
            "responseJsonSchema": RESPONSE_SCHEMA,
        },
    }
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}

    data: dict = {}
    error = ""
    for attempt in range(3):
        try:
            response = requests.post(endpoint, headers=headers, json=payload, timeout=150)
            if response.status_code == 200:
                data = response.json()
                break
            try:
                message = (response.json().get("error") or {}).get("message", "")
            except Exception:
                message = response.text[:250]
            error = f"Gemini HTTP {response.status_code}: {message}"
            if response.status_code not in {429, 500, 502, 503, 504}:
                break
            time.sleep(4 * (attempt + 1))
        except Exception as exc:
            error = f"Gemini request error: {str(exc)[:220]}"
            time.sleep(3 * (attempt + 1))

    if not data:
        return {
            "status": "🔒 Inaccessible",
            "confidence": "",
            "verdict": error or "Gemini request failed",
            "supported_claims": "",
            "unsupported_claims": "",
            "unverified_claims": "",
            "direct_verbatims": "",
            "evidence_by_source": "",
            "qa_flags": "technical issue",
            "suggested_fix": "",
            "sources": [{"source_number": i + 1, "url": u, "retrieval_status": "Not checked", "raw_status": ""} for i, u in enumerate(urls)],
            "usage": {},
        }

    parsed, parse_error = _parse(_response_text(data))
    if parse_error:
        return {
            "status": "🔒 Inaccessible",
            "confidence": "",
            "verdict": parse_error,
            "supported_claims": "",
            "unsupported_claims": "",
            "unverified_claims": "",
            "direct_verbatims": "",
            "evidence_by_source": "",
            "qa_flags": "parse issue",
            "suggested_fix": "",
            "sources": _url_metadata(data, urls),
            "usage": _usage(data, model),
        }

    status_map = {
        "CONFIRMED": "✅ Confirmed",
        "PARTIAL": "⚠️ Partial",
        "INCORRECT": "❌ Incorrect",
        "NOT_FOUND": "❓ Not Found",
    }
    sources = _url_metadata(data, urls)
    retrieved = [s for s in sources if s["retrieval_status"] == "Retrieved"]
    final_status = status_map.get(str(parsed.get("status", "")).upper(), "❓ Not Found")
    supported = _join(parsed.get("supported_claims"))
    unsupported = _join(parsed.get("unsupported_claims"))
    unverified = _join(parsed.get("unverified_claims"))
    verdict = str(parsed.get("verdict", "")).strip()

    if not retrieved:
        final_status = "🔒 Inaccessible"
        # Accessibility is not factual disproof. Preserve the item as unverified.
        if not unverified:
            unverified = unsupported or claim.strip()
        unsupported = ""
        supported = ""
        if not verdict:
            verdict = "No cited source could be retrieved, so the claim could not be verified."

    return {
        "status": final_status,
        "confidence": parsed.get("confidence", ""),
        "verdict": verdict,
        "supported_claims": supported,
        "unsupported_claims": unsupported,
        "unverified_claims": unverified,
        "direct_verbatims": _join(parsed.get("direct_verbatims")),
        "evidence_by_source": _join(parsed.get("evidence_by_source")),
        "qa_flags": _join(parsed.get("qa_flags"), 300),
        "suggested_fix": str(parsed.get("suggested_fix", "")).strip(),
        "sources": sources,
        "usage": _usage(data, model),
    }
