from pathlib import Path
from collections import defaultdict
from tools.debug_unique_quote_seed import _parse_eml, _load_ess_team, _analyze_unique_shape
from src.rules.subject_normalizer import normalize_subject

eml_dir = Path('/app/output/eml/export_filtered')
config_dir = Path('/app/config')
ess_team = _load_ess_team(config_dir)
subjects = defaultdict(list)
for path in eml_dir.rglob('*.eml'):
    email_obj = _parse_eml(path)
    if not email_obj:
        continue
    subj = normalize_subject(email_obj.subject)
    if not subj:
        continue
    subjects[subj].append(email_obj)

candidates = []
for subj, msgs in subjects.items():
    if len(msgs) != 1:
        continue
    email_obj = msgs[0]
    analysis = _analyze_unique_shape(email_obj, ess_team)
    if analysis['predicted'] == 'all-three-same':
        candidates.append((subj, email_obj.sent_time, email_obj.path, analysis['first_is_ess'], analysis['paired_request'], analysis['why']))

for subj, sent_time, path, first_is_ess, paired_request, why in sorted(candidates)[:30]:
    print(f'subject={subj}')
    print(f'sent={sent_time}')
    print(f'first_is_ess={first_is_ess}')
    print(f'paired_request={paired_request}')
    print(f'why={why}')
    print(f'path={path}')
    print('---')
print(f'TOTAL={len(candidates)}')
