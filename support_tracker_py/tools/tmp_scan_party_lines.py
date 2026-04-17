from email import policy
from email.parser import BytesParser
from pathlib import Path
import html,re
try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup=None
path = Path('/app/output/eml/export_filtered/Outlook Data File/raviteja.dwarampudi@invenio-solutions.com/My Team/310.eml')
with path.open('rb') as f:
    msg=BytesParser(policy=policy.default).parse(f)
plain=''; html_body=''
if msg.is_multipart():
    for part in msg.walk():
        ctype=(part.get_content_type() or '').lower()
        try: content=part.get_content()
        except Exception: content=None
        if ctype=='text/plain' and isinstance(content,str) and not plain: plain=content
        elif ctype=='text/html' and isinstance(content,str) and not html_body: html_body=content
raw=f"{plain}\n{html_body}"
txt=re.sub(r'(?is)<style.*?>.*?</style>',' ',raw)
txt=re.sub(r'(?is)<script.*?>.*?</script>',' ',txt)
txt=re.sub(r'(?i)<\s*br\s*/?>','\n',txt)
txt=re.sub(r'(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>','\n',txt)
txt=re.sub(r'(?is)<[^>]+>',' ',txt)
txt=html.unescape(txt)
lines=[ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
if BeautifulSoup is not None and html_body:
    soup=BeautifulSoup(html_body,'html.parser')
    bs4_lines=[ln.strip() for ln in html.unescape(soup.get_text('\n')).splitlines() if ln and ln.strip()]
    def score(ls):
        s=0
        for line in ls:
            if re.match(r'(?i)^(from|sent|to|cc|bcc|subject|objet)\b\s*:?',line): s+=3
            elif re.search(r'(?i)\bfrom\b\s*:',line): s+=2
            elif re.search(r'(?i)\b(sent|subject|objet)\b\s*:',line): s+=2
        return s
    if bs4_lines and (score(bs4_lines)>=score(lines) or not lines): lines=bs4_lines
for i,line in enumerate(lines[:220],start=1):
    low=line.lower()
    if 'thank you for the information' in low or 'we will change' in low or 'thanks for the information' in low:
        print('MATCH',i,line)
        for j in range(max(1,i-4),min(len(lines),i+6)+1):
            print(f'{j:03d}: {lines[j-1]}')
        print('---')
        break
else:
    print('NO_MATCH')
    for i,line in enumerate(lines[:120],start=1):
        print(f'{i:03d}: {line}')
