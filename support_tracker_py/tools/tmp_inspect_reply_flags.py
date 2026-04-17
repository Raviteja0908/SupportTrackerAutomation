from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from types import SimpleNamespace

from src.rules.time_resolver import (
    _classify_reply_kind,
    _is_ack_like_reply,
    _is_thanks_info_reply,
    _email_has_explicit_ack_signal,
    _email_has_short_ess_ack_signal,
)

path = Path('/app/output/eml/export_filtered/Outlook Data File/raviteja.dwarampudi@invenio-solutions.com/My Team/310.eml')
with path.open('rb') as handle:
    msg = BytesParser(policy=policy.default).parse(handle)
plain = ''
html = ''
if msg.is_multipart():
    for part in msg.walk():
        ctype = (part.get_content_type() or '').lower()
        try:
            content = part.get_content()
        except Exception:
            content = None
        if ctype == 'text/plain' and isinstance(content, str) and not plain:
            plain = content
        elif ctype == 'text/html' and isinstance(content, str) and not html:
            html = content
else:
    try:
        content = msg.get_content()
    except Exception:
        content = ''
    if isinstance(content, str):
        plain = content
sender_name = ''
sender_email = ''
addrs = getaddresses([msg.get('from', '')])
if addrs:
    sender_name, sender_email = addrs[0]
email_obj = SimpleNamespace(
    subject=str(msg.get('subject', '') or ''),
    sender_name=sender_name or '',
    sender_email=(sender_email or '').lower(),
    sent_time=parsedate_to_datetime(msg.get('date', '')),
    body=plain or '',
    body_html=html or '',
    path=path,
)
print('classify=', _classify_reply_kind(email_obj))
print('ack_like=', _is_ack_like_reply(email_obj))
print('thanks_info=', _is_thanks_info_reply(email_obj))
print('explicit_ack=', _email_has_explicit_ack_signal(email_obj))
print('short_ess_ack=', _email_has_short_ess_ack_signal(email_obj))
