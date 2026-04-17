import csv, json, re, html
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from src.rules.subject_normalizer import extract_subject_from_description, normalize_subject
from src.rules.time_resolver import _match_requester, _is_ess_sender, _is_ack_like_reply, _classify_reply_kind, _to_ist

@dataclass
class E:
    subject:str; sender_name:str; sender_email:str; sent_time:datetime|None; body:str; body_html:str; path:str

def parse_eml(path: Path):
    try:
        with path.open('rb') as h: msg=BytesParser(policy=policy.default).parse(h)
    except Exception:
        return None
    plain=''; html_body=''
    if msg.is_multipart():
        for part in msg.walk():
            c=(part.get_content_type() or '').lower()
            try: content=part.get_content()
            except Exception: content=None
            if c=='text/plain' and isinstance(content,str) and not plain: plain=content
            elif c=='text/html' and isinstance(content,str) and not html_body: html_body=content
    else:
        try: content=msg.get_content()
        except Exception: content=''
        if isinstance(content,str): plain=content
    sender_name=''; sender_email=''
    addrs=getaddresses([msg.get('from','')])
    if addrs: sender_name,sender_email=addrs[0]
    try: sent=parsedate_to_datetime(msg.get('date',''))
    except Exception: sent=None
    return E(str(msg.get('subject','') or ''), sender_name or '', (sender_email or '').lower(), sent, plain or '', html_body or '', str(path))

def clean_lines(email_obj):
    raw=f"{email_obj.body}\n{email_obj.body_html}"
    txt=re.sub(r'(?is)<style.*?>.*?</style>',' ',raw)
    txt=re.sub(r'(?is)<script.*?>.*?</script>',' ',txt)
    txt=re.sub(r'(?i)<\s*br\s*/?>','\n',txt)
    txt=re.sub(r'(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>','\n',txt)
    txt=re.sub(r'(?is)<[^>]+>',' ',txt)
    txt=html.unescape(txt)
    return [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]

def parse_sent(line):
    line=re.sub(r'(?i)^sent\b\s*:?\s*','',line).strip()
    fmts=['%A, %B %d, %Y %I:%M %p','%A, %d %B %Y %I:%M %p','%d %B %Y %I:%M %p','%d %B %Y %H:%M','%d-%m-%Y %H:%M','%d/%m/%Y %H:%M']
    for fmt in fmts:
        try: return datetime.strptime(line,fmt)
        except Exception: pass
    return None

def quoted_blocks(e):
    lines=clean_lines(e)
    out=[]
    i=0
    while i < len(lines):
        if not re.search(r'(?i)\bfrom\b\s*:', lines[i] or ''):
            i += 1; continue
        from_line=''; sent_line=''; subj=''; end=i
        for j in range(i, min(i+16, len(lines))):
            cur=(lines[j] or '').strip()
            if j>i and (re.search(r'(?i)\bfrom\b\s*:', cur or '') or re.match(r'(?i)^[-_]{3,}$',cur)):
                break
            m=re.match(r'(?i)^(from|sent|subject|objet)\b\s*:?\s*(.*)$', cur or '')
            lab=m.group(1).lower() if m else None
            val=(m.group(2) if m else '').strip()
            if lab=='from' and not from_line: from_line=val
            elif lab=='sent' and not sent_line: sent_line=cur
            elif lab in {'subject','objet'} and not subj: subj=val
            end=j
        if sent_line:
            out.append((from_line, parse_sent(sent_line), subj))
        i=max(i+1,end+1)
    return out

def same_subject(a,b):
    a=normalize_subject(a or ''); b=normalize_subject(b or '')
    return bool(a and b and (a==b or a in b or b in a))

def get(row,*names):
    lowered={str(k).strip().lower(): v for k,v in row.items()}
    for name in names:
        if name in row and row[name] not in (None,''): return str(row[name]).strip()
        value=lowered.get(name.lower())
        if value not in (None,''): return str(value).strip()
    return ''

rows=list(csv.DictReader(open('/app/output/automation_output_support_tracker_feb_2026_incident_business_09.csv', encoding='utf-8-sig')))
row=next(r for r in rows if 'ODS015-->RE: [EXTERNAL] RE: Vendor Migration- Daily call-->16-02-2026'.lower() in (r.get('Description') or '').lower())
family=normalize_subject(extract_subject_from_description(get(row,'Description')))
requester=get(row,'Requested By','Requester')
config=json.loads(Path('/app/config/ess_team.json').read_text(encoding='utf-8'))
emails=[]
for p in Path('/app/output/eml/export_filtered').rglob('*.eml'):
    rec=parse_eml(p)
    if rec and same_subject(rec.subject, family):
        emails.append(rec)
for e in sorted(emails, key=lambda x: _to_ist(x.sent_time) if x.sent_time else datetime.min):
    if not _is_ess_sender(e, config):
        continue
    if not _match_requester(e.sender_name, e.sender_email, requester):
        continue
    if _is_ack_like_reply(e):
        continue
    parent=None
    for from_line,sent_dt,subj in quoted_blocks(e):
        if subj and not same_subject(subj, family):
            continue
        parent = from_line
        break
    parent_ess = None
    if parent is not None:
        pl=parent.lower()
        parent_ess = any((ess.lower() in pl) for ess in config)
    print(_to_ist(e.sent_time).strftime('%d-%m-%Y %H:%M'), e.sender_email, _classify_reply_kind(e), 'parent_ess=', parent_ess, 'subject=', normalize_subject(e.subject))
