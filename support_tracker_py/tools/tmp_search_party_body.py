from email import policy
from email.parser import BytesParser
from pathlib import Path
path = Path('/app/output/eml/export_filtered/Outlook Data File/raviteja.dwarampudi@invenio-solutions.com/My Team/310.eml')
with path.open('rb') as handle:
    msg = BytesParser(policy=policy.default).parse(handle)
plain=''
html=''
if msg.is_multipart():
    for part in msg.walk():
        ctype=(part.get_content_type() or '').lower()
        try:
            content=part.get_content()
        except Exception:
            content=None
        if ctype=='text/plain' and isinstance(content,str) and not plain:
            plain=content
        elif ctype=='text/html' and isinstance(content,str) and not html:
            html=content
raw = (plain or '') + '\n' + (html or '')
for needle in ['we will change','thank you for the information','thanks for the information','we will process','we will do the same']:
    idx = raw.lower().find(needle)
    print(needle, idx)
    if idx >= 0:
        print(raw[max(0, idx-120):idx+200])
        print('---')
