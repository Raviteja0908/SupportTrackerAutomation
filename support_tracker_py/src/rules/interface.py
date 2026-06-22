import re
from functools import lru_cache


_RE_PREFIX = re.compile(r"(?i)^\s*re:\s*")
_RE_SPACE = re.compile(r"\s+")
_LEADING_INTERFACE_RE = re.compile(r"^\s*([a-z]{2,8}\d{2,4}[a-z]{0,3})\b", re.IGNORECASE)
_ANY_INTERFACE_RE = re.compile(r"\b([a-z]{2,8}\d{2,4}[a-z]{0,3})\b", re.IGNORECASE)
_INTERFACE_VALIDATION_RE = re.compile(r"^[a-z]{2,8}\d{2,4}[a-z]{0,3}$", re.IGNORECASE)
_BANNED_PREFIXES = ("inc", "sr")


@lru_cache(maxsize=4096)
def _canonical_text(text: str) -> str:
    value = text or ""
    value = value.replace("\n", " ").strip()
    value = _RE_PREFIX.sub("", value)
    _arrow_split = re.compile(r"--\.?>|->|→|➔|➡|=>")
    m = _arrow_split.search(value)
    if m:
        value = value[:m.start()].strip()
    value = _RE_SPACE.sub(" ", value).strip()
    return value


@lru_cache(maxsize=4096)
def _extract_interface_token(text: str) -> str:
    if not text:
        return ""
    # Prefer a strong leading interface-like token first, then fall back to
    # the first strong token found anywhere in the cleaned subject.
    for pattern in (_LEADING_INTERFACE_RE, _ANY_INTERFACE_RE):
        for match in pattern.finditer(text):
            token = match.group(1)
            lower = token.lower()
            if lower.startswith(_BANNED_PREFIXES):
                continue
            return token.upper()
    return ""


@lru_cache(maxsize=4096)
def resolve_interface_code(description: str) -> str:
    if not description:
        return ""
    base = _canonical_text(description)
    lower = base.lower()

    if "severes warnings and errors in eai/es aws/es symphony" in lower:
        return "SEVERES WARNINGS"
    if "daily task hyparchive" in lower:
        return "Daily Task Hyparchive"

    # Safe fallback for subjects with or without "-->":
    # if a strong interface token exists in the cleaned left-side text,
    # prefer that over returning the whole subject line.
    token = _extract_interface_token(base)
    if token:
        return token

    return base
