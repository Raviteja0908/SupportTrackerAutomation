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
else:
    try:
        content=msg.get_content()
    except Exception:
        content=''
    if isinstance(content,str):
        plain=content
print('PLAIN_START')
for i,line in enumerate((plain or '').splitlines()[:40], start=1):
    print(f'{i:02d}: {line}')
print('HTML_LEN', len(html or ''))
