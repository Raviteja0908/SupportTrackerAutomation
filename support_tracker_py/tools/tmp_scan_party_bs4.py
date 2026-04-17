from email import policy
from email.parser import BytesParser
from pathlib import Path
import html
from bs4 import BeautifulSoup
path = Path('/app/output/eml/export_filtered/Outlook Data File/raviteja.dwarampudi@invenio-solutions.com/My Team/310.eml')
with path.open('rb') as f:
    msg=BytesParser(policy=policy.default).parse(f)
html_body=''
if msg.is_multipart():
    for part in msg.walk():
        ctype=(part.get_content_type() or '').lower()
        try: content=part.get_content()
        except Exception: content=None
        if ctype=='text/html' and isinstance(content,str) and not html_body: html_body=content
soup=BeautifulSoup(html_body,'html.parser')
lines=[ln.strip() for ln in html.unescape(soup.get_text('\n')).splitlines() if ln and ln.strip()]
for i,line in enumerate(lines[:120],start=1):
    print(f'{i:03d}: {line}')
print('---SEARCH---')
for needle in ['thank you for the information','we will change']:
    for i,line in enumerate(lines,start=1):
        if needle in line.lower():
            print('MATCH',needle,i,line)
            for j in range(max(1,i-4),min(len(lines),i+6)+1):
                print(f'{j:03d}: {lines[j-1]}')
            print('---')
            break
