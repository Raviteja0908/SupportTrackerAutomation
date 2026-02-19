import re


_TEST_RE = re.compile(r"\btest(?:ing)?\b", re.IGNORECASE)
_NON_PROD_RE = re.compile(r"\bnon\s*-\s*prod\b|\bnon\s+prod\b", re.IGNORECASE)

_PROD_RE = re.compile(r"\b(prod|prd|production|fcp|bip)\b", re.IGNORECASE)
_UAT_RE = re.compile(r"\b(uat|fct|biu)\b", re.IGNORECASE)
_QA_RE = re.compile(r"\b(qa|fcq|biq)\b", re.IGNORECASE)
_DEV_RE = re.compile(r"\b(dev|development|fcd|bid)\b", re.IGNORECASE)

# In test-override branch, requested behavior keeps FCT as PROD override.
_TEST_OVERRIDE_PROD_RE = re.compile(r"\b(prod|prd|production|fcp|bip|fct)\b", re.IGNORECASE)
_TEST_OVERRIDE_UAT_RE = re.compile(r"\b(uat|biu)\b", re.IGNORECASE)

_QUOTED_BOUNDARY_RE = re.compile(
    r"^\s*(from:|sent:|subject:|de:|envoy[ée]?:|objet:|on .+ wrote:|[-]{2,}\s*original message\s*[-]{2,}|>)",
    re.IGNORECASE,
)


def _has_test_marker(text: str) -> bool:
    if not text:
        return False
    return bool(_NON_PROD_RE.search(text) or _TEST_RE.search(text))


def _primary_segment(text: str) -> str:
    """
    Keep only the top message block before quoted history markers.
    This reduces false env picks from older replies in long chains.
    """
    if not text:
        return ""
    lines = text.splitlines()
    out = []
    for line in lines:
        if _QUOTED_BOUNDARY_RE.match(line):
            break
        out.append(line)
    return "\n".join(out).strip()


def _segments_for_detection(text: str):
    if not text:
        return []
    full = text
    top = _primary_segment(text)
    if top and top != full:
        return [top, full]
    return [full]


def _resolve_explicit_env(text: str) -> str:
    if not text:
        return ""

    force_prod_phrases = (
        "daily task hyparchive",
        "severes warnings and errors in eai/es aws/es symphony",
        "files from es to grp",
    )

    for seg in _segments_for_detection(text):
        s = seg.lower()
        if any(p in s for p in force_prod_phrases):
            return "PROD"

        # Explicit environment markers (strong signals)
        has_prod = bool(_PROD_RE.search(seg))
        has_uat = bool(_UAT_RE.search(seg))
        has_qa = bool(_QA_RE.search(seg))
        has_dev = bool(_DEV_RE.search(seg))

        if has_prod:
            return "PROD"
        # If both QA and UAT are present in same content, prefer UAT.
        if has_uat and has_qa:
            return "UAT"
        if has_uat:
            return "UAT"
        if has_qa:
            return "QA"
        if has_dev:
            return "DEV"
    return ""


def _resolve_env_for_test_override(text: str) -> str:
    """
    When subject is generic TEST/non-prod, prefer consultant/body explicit markers.
    Per requested behavior, FCT in replies is treated as PROD override here.
    """
    if not text:
        return ""

    for seg in _segments_for_detection(text):
        has_prod = bool(_TEST_OVERRIDE_PROD_RE.search(seg))
        has_uat = bool(_TEST_OVERRIDE_UAT_RE.search(seg))
        has_qa = bool(_QA_RE.search(seg))
        has_dev = bool(_DEV_RE.search(seg))

        if has_prod:
            return "PROD"
        # If both QA and UAT are present in same content, prefer UAT.
        if has_uat and has_qa:
            return "UAT"
        if has_uat:
            return "UAT"
        if has_qa:
            return "QA"
        if has_dev:
            return "DEV"
    return ""


def resolve_environment(subject_text: str, body_text: str = "") -> str:
    """
    Resolve environment using subject first, then fall back to mail body/content.
    Leave empty only if no match is found in either.
    """
    subject_text = subject_text or ""
    body_text = body_text or ""

    # 1) If subject has explicit env, trust it directly.
    subject_env = _resolve_explicit_env(subject_text)
    if subject_env:
        return subject_env

    # 2) If subject is generic test/non-prod, let consultant/body explicit env override.
    if _has_test_marker(subject_text):
        body_override = _resolve_env_for_test_override(body_text)
        if body_override:
            return body_override
        return "UAT"

    # 3) Normal fallback to body explicit env.
    body_env = _resolve_explicit_env(body_text)
    if body_env:
        return body_env

    # 4) Last fallback: test/non-prod in body means UAT.
    if _has_test_marker(body_text):
        return "UAT"
    return ""
