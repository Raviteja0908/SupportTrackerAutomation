import re


_prefix_re = re.compile(r"^\s*(re|fw|fwd|aw|wg|sv)\s*:\s*", re.IGNORECASE)
_external_re = re.compile(r"\[external\]\s*", re.IGNORECASE)
_fancy_quotes = {
    "“": "\"",
    "”": "\"",
    "‘": "'",
    "’": "'",
}


def normalize_subject(subject: str) -> str:
    if not subject:
        return ""
    s = subject.strip()
    for k, v in _fancy_quotes.items():
        s = s.replace(k, v)
    s = _external_re.sub("", s)
    while True:
        new_s = _prefix_re.sub("", s)
        if new_s == s:
            break
        s = new_s
    s = re.sub(r"\s+", " ", s).strip()
    return s


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
    # Remove leading interface-like tokens (e.g., CS001-->, ID082:, VMI001 -)
    # Supports multiple interface tokens separated by commas or slashes.
    return re.sub(
        r"^([a-z]{1,5}\d{2,}(?:[,\s/]+[a-z]{1,5}\d{2,})*)\s*(--?>|:|-|\|)\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )


def _strip_trailing_date(text: str) -> str:
    if not text:
        return ""
    # Remove trailing date-like suffixes
    patterns = [
        r"\s*-->\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*->\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
        r"\s*[-–>]?\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$",
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
    marker = "-->"
    if marker in description:
        left, right = description.split(marker, 1)
        if _looks_like_interface_prefix(left):
            return right.strip()
    return description.strip()


def _looks_like_interface_prefix(text: str) -> bool:
    if not text:
        return False
    prefix = text.strip()
    if not prefix:
        return False
    tokens = re.split(r"[,\s/]+", prefix)
    tokens = [t for t in tokens if t]
    if not tokens:
        return False
    for token in tokens:
        if not re.match(r"^[A-Za-z]{1,5}\d{2,}$", token):
            return False
    return True
