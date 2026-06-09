import pandas as pd
import re
from difflib import SequenceMatcher
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

MATCH_THRESHOLD = 65
COL_NAME = "TEST CENTER NAME"

DISTRICT_ALIASES = {
    "KANPUR NAGAR": "KANPUR",
    "KANPUR DEHAT": "KANPUR DEHAT",
}

CANONICAL = [
    (r'\bgovernment\s+inter\s+col(?:lege)?\b',       'gic'),
    (r'\bgovt\.?\s+inter\s+col(?:lege)?\b',          'gic'),
    (r'\brajkiya\s+inter\s+col(?:lege)?\b',          'gic'),
    (r'\bpm\s+shri\s+(?:govt|government|rajkiya)\b', 'gic'),
    (r'\bg\.?\s*i\.?\s*c\.?\b',                      'gic'),
    (r'\bpost\s*gradu(?:ate|ation)\s*college\b',     'pgcollege'),
    (r'\bdegree\s*college\b',                        'pgcollege'),
    (r'\bmahavidyalaya\b',                           'pgcollege'),
    (r'\bsnatakottar\s*(?:college)?\b',              'pgcollege'),
    (r'\bsnatakor\s*(?:college)?\b',                 'pgcollege'),
    (r'\bpg\s*college\b',                            'pgcollege'),
    (r'\binter(?:mediate)?\s*col(?:lege)?\b',        'ic'),
    (r'\bgovernment\b',  'govt'),
    (r'\bgovt\.?\b',     'govt'),
    (r'\brajkiya\b',     'govt'),
    (r'\bbalika\b',  'girls'),
    (r'\bkanya\b',   'girls'),
    (r'\bkumari\b',  'girls'),
    (r'\bmahila\b',  'girls'),
    (r'\bwomens?\b', 'girls'),
    (r'\bbalak\b',   'boys'),
    (r'\bvidyalayam?\b', 'school'),
    (r'\bvidyalay\b',    'school'),
    (r'\bd\s*\.?\s*a\s*\.?\s*v\s*\.?\b',  'dav'),
    (r'\bpost\s*gradu(?:ate|ation)\b',  'pg'),
    (r'\bp\.?\s*g\.?\b',               'pg'),
    (r'\b(?:senior|higher)\s*secondary\b', 'srsc'),
    (r'\bsr\.?\s*sec\.?\b',               'srsc'),
    (r'\bss\s*school\b',                  'srsc'),
    (r'\bf\.?\s*e\.?\s*t\.?\b',  'fet'),
    (r'\badarsh\b',  'ideal'),
    (r'\bsmarak\b',  'memorial'),
    (r'\bsmriti\b',  'memorial'),
    (r'\bpublic\s*school\b',  'ps'),
    (r'\bvidyapeeth\b',   'vidyapith'),
    (r'\bvidyapeetha\b',  'vidyapith'),
    (r'\bshree\b',  'shri'),
    (r'\bsri\b',    'shri'),
    (r'\bkendriya\s*school\b',   'kv'),
    (r'\bkendriya\s*vidyalya\b', 'kv'),
    (r'\bkv\b',                  'kv'),
    (r'\bhindi\s*medium\b',   ''),
    (r'\benglish\s*medium\b', ''),
    (r'\bco\s*ed\b',          ''),
    (r'\bpm\s+shri\b',        ''),
    (r'\bautonomous\b',       ''),
]

NOISE = {
    'the', 'and', 'of', 'for', 'in', 'at', 'by', 'near', 'new', 'old',
    'no', 'number', 'sh', 'smt', 'dr', 'km', 'late', 'pt', 'prof',
    'agra', 'lucknow', 'kanpur', 'varanasi', 'prayagraj', 'aligarh',
    'ayodhya', 'azamgarh', 'mathura', 'bareilly', 'moradabad', 'meerut',
    'ghaziabad', 'noida', 'gorakhpur', 'allahabad', 'jhansi', 'firozabad',
    'basti', 'saharanpur', 'mirzapur', 'banda',
}

STRUCTURAL_TOKENS = {'gic', 'ic', 'govt', 'pgcollege', 'pg', 'srsc', 'ps',
                     'kv', 'dav', 'school', 'girls', 'boys', 'shri', 'ideal',
                     'memorial', 'vidyapith', 'fet', 'women',
                     'blka', 'blkb', 'blkc', 'blkd', 'blke', 'blkf',
                     'blk1', 'blk2', 'blk3', 'blk4', 'blk5'}


def clean_text(x):
    x = str(x).lower()
    x = re.sub(r'\(?\s*block[-\s]*([a-z0-9])\s*\)?', r' blk\1 ', x)
    x = re.sub(r'(?<!\w)([a-z])\.\s*([a-z])\.\s*([a-z])\.?(?!\w)', r'\1\2\3', x)
    x = re.sub(r'(?<!\w)([a-z])\.\s*([a-z])\.?(?!\w)', r'\1\2', x)
    x = re.sub(r'[^a-z0-9 ]', ' ', x)
    for pattern, replacement in CANONICAL:
        x = re.sub(pattern, replacement, x)
    words = [w for w in x.split() if w not in NOISE and len(w) > 0]
    return " ".join(words).strip()


def similarity(a, b):
    if not a or not b:
        return 0
    s1 = int(SequenceMatcher(None, a, b).ratio() * 100)
    s2 = int(SequenceMatcher(None, " ".join(sorted(a.split())), " ".join(sorted(b.split()))).ratio() * 100)
    wa = set(a.split())
    wb = set(b.split())
    common = wa & wb
    content_common = [w for w in common if len(w) > 2 and w not in STRUCTURAL_TOKENS]
    s3 = 0
    if content_common:
        shorter = min(len(wa), len(wb))
        if shorter > 0:
            s3 = int(len(content_common) / shorter * 100)
    s4 = 0
    if len(a.split()) <= 2 or len(b.split()) <= 2:
        short = a if len(a) <= len(b) else b
        long_ = b if len(a) <= len(b) else a
        if short and short in long_:
            s4 = int(SequenceMatcher(None, short, long_).ratio() * 100)
    raw = max(s1, s2, s3, s4)
    blk_a = {w for w in a.split() if w.startswith('blk')}
    blk_b = {w for w in b.split() if w.startswith('blk')}
    if blk_a and blk_b:
        if blk_a == blk_b:
            raw = min(100, raw + 5)
        else:
            raw = min(raw, 55)
    return raw


def normalize_district(d):
    d = str(d).upper().strip()
    return DISTRICT_ALIASES.get(d, d)


def run_matching(file_a_path, file_b_path, output_path):
    df_a = pd.read_excel(file_a_path)
    df_b = pd.read_excel(file_b_path)
    df_a.columns = df_a.columns.str.strip()
    df_b.columns = df_b.columns.str.strip()

    COL_A_DIST = [c for c in df_a.columns if 'district' in c.lower()][0]
    COL_B_DIST = [c for c in df_b.columns if 'district' in c.lower()][0]

    b_state_col = [c for c in df_b.columns if 'state' in c.lower()]
    if b_state_col:
        df_b_filtered = df_b[df_b[b_state_col[0]].str.upper().str.contains('UTTAR PRADESH', na=False)].copy()
    else:
        df_b_filtered = df_b.copy()

    df_a["_name"] = df_a[COL_NAME].apply(clean_text)
    df_b_filtered["_name"] = df_b_filtered[COL_NAME].apply(clean_text)
    df_a["_dist"] = df_a[COL_A_DIST].apply(normalize_district)
    df_b_filtered["_dist"] = df_b_filtered[COL_B_DIST].apply(normalize_district)

    b_grouped = df_b_filtered.groupby("_dist")

    candidates_all = []
    for idx_a, row_a in df_a.iterrows():
        if row_a["_dist"] in b_grouped.groups:
            group = b_grouped.get_group(row_a["_dist"])
            for idx_b, row_b in group.iterrows():
                score = similarity(row_a["_name"], row_b["_name"])
                if score >= MATCH_THRESHOLD:
                    candidates_all.append((score, idx_a, idx_b))

    candidates_all.sort(key=lambda x: -x[0])
    used_a, used_b, matches = set(), set(), {}
    for score, idx_a, idx_b in candidates_all:
        if idx_a not in used_a and idx_b not in used_b:
            matches[idx_a] = (idx_b, score)
            used_a.add(idx_a)
            used_b.add(idx_b)

    matched_rows = []
    for idx_a, row_a in df_a.iterrows():
        a_cols = list(row_a.drop(["_name", "_dist"]))
        if idx_a in matches:
            idx_b, score = matches[idx_a]
            b_cols = list(df_b_filtered.loc[idx_b].drop(["_name", "_dist"]))
            matched_rows.append(a_cols + b_cols + [score, "MATCHED"])
        else:
            best_score = 0
            if row_a["_dist"] in b_grouped.groups:
                group = b_grouped.get_group(row_a["_dist"])
                for idx_b, row_b in group.iterrows():
                    s = similarity(row_a["_name"], row_b["_name"])
                    if s > best_score:
                        best_score = s
            b_cols = [None] * (len(df_b_filtered.columns) - 2)
            matched_rows.append(a_cols + b_cols + [best_score, "UNMATCHED_A"])

    for idx_b, row_b in df_b_filtered.iterrows():
        if idx_b not in used_b:
            a_cols = [None] * (len(df_a.columns) - 2)
            b_cols = list(row_b.drop(["_name", "_dist"]))
            matched_rows.append(a_cols + b_cols + [0, "UNMATCHED_B"])

    a_final_cols = list(df_a.columns.drop(["_name", "_dist"]))
    b_final_cols = list(df_b_filtered.columns.drop(["_name", "_dist"]))
    b_renamed = ["B_" + c if c in a_final_cols else c for c in b_final_cols]
    columns = a_final_cols + b_renamed + ["Match %", "Status"]
    final_df = pd.DataFrame(matched_rows, columns=columns)

    matched_count   = len(final_df[final_df["Status"] == "MATCHED"])
    unmatched_a_cnt = len(final_df[final_df["Status"] == "UNMATCHED_A"])
    unmatched_b_cnt = len(final_df[final_df["Status"] == "UNMATCHED_B"])
    match_rate      = round(matched_count / max(len(df_a), 1) * 100, 1)

    final_df.to_excel(output_path, index=False)

    wb = load_workbook(output_path)
    ws = wb.active
    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    RED    = PatternFill("solid", fgColor="FFC7CE")
    YELLOW = PatternFill("solid", fgColor="FFEB9C")
    HEADER = PatternFill("solid", fgColor="4472C4")
    WHITE_FONT = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = HEADER
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    status_col = columns.index("Status") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        status = row[status_col - 1].value
        fill = GREEN if status == "MATCHED" else (RED if status == "UNMATCHED_A" else YELLOW)
        for cell in row:
            cell.fill = fill
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = max((len(str(c.value)) if c.value else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)
    ws.freeze_panes = "A2"
    wb.save(output_path)

    return {
        "total": len(final_df),
        "matched": matched_count,
        "unmatched_a": unmatched_a_cnt,
        "unmatched_b": unmatched_b_cnt,
        "match_rate": match_rate,
    }
