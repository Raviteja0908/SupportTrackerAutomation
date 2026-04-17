import argparse
import html
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None

from src.rules.subject_normalizer import normalize_subject
from src.rules.time_resolver import (
    _classify_reply_kind,
    _email_has_short_ess_ack_signal,
    _extract_canonical_message_lines,
    _extract_canonical_quoted_header_candidates,
    _is_ack_like_reply,
    _is_explicit_ack_signal,
    _is_nonfinal_followup_reply,
    _is_thanks_info_reply,
    _is_ess_sender,
)


@dataclass
class DebugEmail:
    subject: str
    sender_name: str
    sender_email: str
    sent_time: datetime | None
    body: str
    body_html: str
    body_html_raw: str
    path: str


def _parse_eml(path: Path) -> DebugEmail | None:
    try:
        with path.open("rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
    except Exception:
        return None

    plain_parts = []
    html_parts = []
    try:
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                try:
                    content = part.get_content()
                except Exception:
                    content = None
                if ctype == "text/plain" and isinstance(content, str):
                    plain_parts.append(content.strip())
                elif ctype == "text/html" and isinstance(content, str):
                    html_parts.append(content)
        else:
            content = msg.get_content()
            if isinstance(content, str):
                plain_parts.append(content.strip())
    except Exception:
        pass

    sent_time = None
    try:
        sent_time = msg["Date"].datetime if msg["Date"] else None
    except Exception:
        sent_time = None

    return DebugEmail(
        subject=str(msg.get("Subject", "") or ""),
        sender_name=str(msg.get("From", "") or ""),
        sender_email=str(msg.get("From", "") or ""),
        sent_time=sent_time,
        body="\n".join(p for p in plain_parts if p),
        body_html="\n".join(h for h in html_parts if h),
        body_html_raw="\n".join(h for h in html_parts if h),
        path=str(path),
    )


def _load_ess_team(config_dir: Path) -> list[str]:
    path = config_dir / "ess_team.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _clean_quoted_message_lines(email_obj: DebugEmail) -> list[str]:
    return _extract_canonical_message_lines(email_obj)


def _extract_bounded_quoted_header_candidates(email_obj: DebugEmail) -> list[tuple[str, datetime, str]]:
    blocks = []
    for candidate in _extract_canonical_quoted_header_candidates(email_obj, allow_relaxed=False):
        from_line = re.sub(r"(?i)^from\b\s*:?\s*", "", candidate.from_line or "").strip()
        subj_text = re.sub(r"(?i)^(subject|objet)\b\s*:?\s*", "", candidate.subject_line or "").strip()
        if candidate.sent_dt:
            blocks.append((from_line, candidate.sent_dt, subj_text))
    return blocks


def _quoted_from_line_is_ess(from_line: str, ess_team: list[str]) -> bool | None:
    addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", from_line or "", flags=re.I)
    emails_l = [em.lower() for em in addr_hits]
    ess_set = {e.strip().lower() for e in ess_team if e}
    if any(em in ess_set for em in emails_l):
        return True
    if not addr_hits and not from_line:
        return None
    return False


def _shared_reply_flags(email_obj: DebugEmail) -> dict:
    cls = _classify_reply_kind(email_obj)
    direct_resolution = bool(cls.get("direct_resolution"))
    explicit_ack = bool(cls.get("explicit_ack"))
    short_ess_ack = bool(cls.get("short_ess_ack"))
    thanks_info = bool(cls.get("thanks_info"))
    nonfinal_followup = bool(cls.get("nonfinal_followup"))
    ack_like = bool(cls.get("ack_like"))
    ignore_reply = thanks_info or nonfinal_followup
    ack_candidate = (explicit_ack or short_ess_ack or ack_like) and not ignore_reply and not direct_resolution
    real_reply = bool(cls.get("real_reply"))
    return {
        **cls,
        "ignore_reply": ignore_reply,
        "ack_candidate": ack_candidate,
        "substantive_reply": real_reply,
    }


def _analyze_unique_shape(email_obj: DebugEmail, ess_team: list[str]) -> dict:
    flags = _shared_reply_flags(email_obj)
    blocks = _extract_bounded_quoted_header_candidates(email_obj)
    first_block = blocks[0] if blocks else None
    first_is_ess = _quoted_from_line_is_ess(first_block[0], ess_team) if first_block else None
    reply_ackish = bool(
        flags["ack_candidate"]
        or flags["thanks_info"]
        or flags["nonfinal_followup"]
        or _email_has_short_ess_ack_signal(email_obj)
        or _is_ack_like_reply(email_obj)
        or _is_thanks_info_reply(email_obj)
        or _is_explicit_ack_signal(email_obj.body or "")
    )
    def _lower_non_ess_below(reference_pos: int, reference_ist: datetime):
        for pos, (from_line, sent_ist, _subj) in enumerate(blocks):
            if pos <= reference_pos:
                continue
            if sent_ist >= reference_ist:
                continue
            if _quoted_from_line_is_ess(from_line, ess_team) is False:
                return sent_ist
        return None

    lower_non_ess = None
    paired_request = None
    if first_block:
        lower_non_ess = _lower_non_ess_below(0, first_block[1])
        for pos, (from_line, sent_ist, _subj) in enumerate(blocks[1:], start=1):
            if _quoted_from_line_is_ess(from_line, ess_team) is not True:
                continue
            paired_request = _lower_non_ess_below(pos, sent_ist)
            if paired_request is not None:
                break
    predicted = "direct-reply"
    why = "fallback"
    if first_is_ess is True and reply_ackish and _is_ess_sender(email_obj, ess_team):
        if lower_non_ess is None:
            predicted = "all-three-same"
            why = "first-quoted-ess + ack-like live reply"
        else:
            predicted = "hybrid"
            why = "first-quoted-ess but lower non-ess request exists"
    elif first_is_ess is False:
        predicted = "direct-reply"
        why = "first-quoted-non-ess direct reply"

    return {
        "flags": flags,
        "blocks": blocks,
        "first_is_ess": first_is_ess,
        "lower_non_ess": lower_non_ess,
        "paired_request": paired_request,
        "predicted": predicted,
        "why": why,
    }


def main():
    parser = argparse.ArgumentParser(description="Debug unique-row quoted seed behavior.")
    parser.add_argument("--eml-dir", required=True)
    parser.add_argument("--subject", required=True)
    parser.add_argument("--config-dir", default="/app/config")
    args = parser.parse_args()

    eml_dir = Path(args.eml_dir)
    config_dir = Path(args.config_dir)
    ess_team = _load_ess_team(config_dir)
    wanted = normalize_subject(args.subject)

    matches = []
    for path in eml_dir.rglob("*.eml"):
        email_obj = _parse_eml(path)
        if not email_obj:
            continue
        if normalize_subject(email_obj.subject) == wanted:
            matches.append(email_obj)

    if not matches:
        print("No matching EMLs found.")
        return 1

    matches.sort(key=lambda e: e.sent_time or datetime.min)
    for email_obj in matches:
        analysis = _analyze_unique_shape(email_obj, ess_team)
        flags = analysis["flags"]
        print("=" * 100)
        print(f"path={email_obj.path}")
        print(f"subject={email_obj.subject}")
        print(f"sent={email_obj.sent_time}")
        print(
            "reply_flags="
            f"kind={flags['kind']} | direct={flags['direct_resolution']} | "
            f"ack_candidate={flags['ack_candidate']} | thanks={flags['thanks_info']} | "
            f"nonfinal={flags['nonfinal_followup']} | real_reply={flags['substantive_reply']}"
        )
        print(
            f"first_quoted_is_ess={analysis['first_is_ess']} | "
            f"lower_non_ess={analysis['lower_non_ess']} | "
            f"paired_request={analysis['paired_request']} | "
            f"predicted_shape={analysis['predicted']} | why={analysis['why']}"
        )
        print("quoted_blocks:")
        for idx, (from_line, sent_ist, subj) in enumerate(analysis["blocks"][:6], start=1):
            print(f"  {idx}. ess={_quoted_from_line_is_ess(from_line, ess_team)} | sent={sent_ist} | from={from_line} | subj={subj}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
