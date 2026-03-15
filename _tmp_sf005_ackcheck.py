import os, re
from email import policy
from email.parser import BytesParser

SUBJ = "SF005 IDoc Failed at ES PROD"

markers = (
    "could you please provide us an update regarding the below",
    "could you please provide us an update on the below",
    "please provide us an update regarding the below",
    "please provide us an update on the below",
    "thank you for the information",
    "thanks for the information",
    "thank you for the update",
    "thanks for the update",
    "thanks for the info",
    "noted with thanks",
    "duly noted",
)

def ack_like_text_fallback(txt):
    txt = (txt or "").lower()
    return any(m in txt for m in markers)

def ess_only_short_ack(raw):
    if not raw:
        return False
    txt = re.sub(r"(?is)<style.*?>.*?</style>", " ", raw)
    txt = re.sub(r"(?is)<script.*?>.*?</script>", " ", raw)
    txt = re.sub(r"(?i)<\\s*br\\s*/?>", "\\n", txt)
    txt = re.sub(r"(?i)</\\s*(p|div|tr|td|th|li|h[1-6])\\s*>", "\\n", txt)
    txt = re.sub(r"(?is)<[^>]+>", " ", txt)
    txt = re.sub(r"\\s+", " ", txt).strip()
    if not txt:
        return False
    lines = [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
    content = " ".join(lines).strip().lower()
    if len(lines) <= 2 and len(content) <= 140:
        strong = (
            "resolved","fixed","completed","success","processed","root cause",
            "issue was","closed","done"
        )
        if any(w in content for w in strong):
            return False
        if any(k in content for k in ("attachment","attached","snippet","snippet:","snippets","see attached")):
            return False
        file_words = ("file","files")
        file_actions = ("add","added","adding","attach","attached","attachment","resend","re-sent","resent","reupload","re-upload","please find")
        if "adding more files" in content or "adding one more file" in content:
            return False
        if any(w in content for w in file_words) and any(a in content for a in file_actions):
            return False
        if "cid:" in raw.lower() or "<img" in raw.lower():
            return False
        return True
    return False

def get_parts(msg):
    html = ""
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/html":
                html = part.get_content()
            elif ctype == "text/plain":
                text = part.get_content()
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            html = msg.get_content()
        elif ctype == "text/plain":
            text = msg.get_content()
    return html, text

count = 0
root_dir = r"d:\\Support_Tracker\\DockerOutput\\eml"
for root, _, files in os.walk(root_dir):
    for fn in files:
        if not fn.lower().endswith(".eml"):
            continue
        path = os.path.join(root, fn)
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        subj = msg.get("Subject","")
        if SUBJ.lower() not in subj.lower():
            continue
        html, text = get_parts(msg)
        raw = (text or "") + "\\n" + (html or "")
        print("="*80)
        print(f"FILE: {fn}")
        print(f"SUBJECT: {subj}")
        print(f"ACK_LIKE_FALLBACK: {ack_like_text_fallback(raw)}")
        print(f"ESS_ONLY_SHORT_ACK: {ess_only_short_ack(raw)}")
        print("HTML_SNIPPET:")
        print((html or "")[:1200])
        count += 1

print(f"\\nTOTAL SF005 FILES: {count}")
