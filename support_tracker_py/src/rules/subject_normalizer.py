import re


_prefix_re = re.compile(r"^\s*(re|fw|fwd|aw|wg|sv)\s*:\s*", re.IGNORECASE)
_external_re = re.compile(r"\[external\]\s*", re.IGNORECASE)
_arrow_re = "(?:--\\.?>|→|➔|➡|=>)"
_fancy_quotes = {
    "“": "\"",
    "”": "\"",
    "‘": "'",
    "’": "'",
}

_ess_name_re = re.compile(
    r"^\s*ESS\s*-\s*[^-]{1,60}\s*-\s*",
    re.IGNORECASE,
)


def normalize_subject(subject: str) -> str:
    if not subject:
        return ""
    s = subject.strip()
    for k, v in _fancy_quotes.items():
        s = s.replace(k, v)
    s = _external_re.sub("", s)
    s = _strip_leading_symbols(s)
    while True:
        new_s = _prefix_re.sub("", s)
        if new_s == s:
            break
        s = new_s
    s = _ess_name_re.sub("", s)
    while True:
        new_s = _prefix_re.sub("", s)
        if new_s == s:
            break
        s = new_s
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _strip_leading_symbols(text: str) -> str:
    if not text:
        return ""
    s = text
    # Remove leading emoji/symbols while preserving bracketed prefixes like "[UMG ...]"
    while s and not s[0].isalnum() and s[0] not in "[(":
        s = s[1:]
    return s.strip()


def normalize_subject_for_match(subject: str) -> str:
    if not subject:
        return ""
    s = normalize_subject(subject)
    s = _strip_interface_prefix(s)
    s = _strip_trailing_date(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _strip_interface_prefix(text: str) -> str:
    if not text:
        return ""
    # Remove leading interface-like tokens before an arrow/colon/dash/pipe.
    m = re.match(rf"^\s*(.+?)\s*({_arrow_re}|:|-|\|)\s*(.+)$", text)
    if not m:
        return text
    left = m.group(1).strip()
    right = m.group(3).strip()
    if _looks_like_interface_prefix(left):
        while True:
            new_right = _prefix_re.sub("", right)
            if new_right == right:
                break
            right = new_right
        return right
    return text


def _strip_trailing_date(text: str) -> str:
    if not text:
        return ""
    # Remove trailing date-like suffixes
    patterns = [
        r"\s*--\.>\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*-->\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*->\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*[-–>]?\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*--\.>\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\s*$",
        r"\s*-->\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\s*$",
        r"\s*->\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\s*$",
        r"\s*[-–>]?\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\s*$",
    ]
    s = text
    for pat in patterns:
        s = re.sub(pat, "", s)
    return s.strip()


def extract_subject_from_description(description: str) -> str:
    if not description:
        return ""
    m = re.split(rf"\s*{_arrow_re}\s*", description, maxsplit=1)
    if len(m) >= 2:
        left, right = m[0], m[1]
        if _looks_like_interface_prefix(left):
            return right.strip()
    return description.strip()


def _looks_like_interface_prefix(text: str) -> bool:
    if not text:
        return False
    prefix = text.strip()
    if not prefix:
        return False
    tokens = re.split(r"[,\s/&;]+", prefix)
    tokens = [t for t in tokens if t]
    if not tokens:
        return False
    for token in tokens:
        if not _is_interface_token(token):
            return False
    return True


def _is_interface_token(token: str) -> bool:
    if not token:
        return False
    if not re.match(r"^[A-Za-z0-9][A-Za-z0-9_-]*$", token):
        return False
    # Require digit OR hyphen/underscore so plain words don't match
    return any(ch.isdigit() for ch in token) or "-" in token or "_" in token
