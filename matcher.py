import pandas as pd
import re
from difflib import SequenceMatcher
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

MATCH_THRESHOLD = 65
COL_NAME = "TEST CENTER NAME"

# Agar tumhe pata chal jaaye ki kaunsa district A vs B mein alag likha jaata hai,
# to yahan explicit alias daal dena -- ye sabse pehle try hota hai.
DISTRICT_ALIASES = {
    "KANPUR NAGAR": "KANPUR",
    "KANPUR DEHAT": "KANPUR DEHAT",
}

# Fuzzy district fallback ke liye threshold (0-1). 0.82 kaafi safe hai,
# isse chhoti spelling/spacing difference match ho jaayegi lekin
# alag districts aapas mein nahi milenge.
DISTRICT_FUZZY_THRESHOLD = 0.82

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
    x = re.sub(r'\(?\s*block[-\s]*([a-z0-9]+)\s*\)?', r' blk\1 ', x)
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
    """
    Pehle se: sirf upper() + strip() hota tha, jisse chhoti si
    punctuation/space/typo difference (A vs B file) ke wajah se
    district string kabhi exact match nahi hoti thi -> group hi
    nahi milta -> ZERO candidates generate hote the poori file mein.

    Fix: punctuation hata do, multiple spaces collapse karo, common
    suffix words normalize karo, phir alias lookup karo.
    """
    d = str(d).upper().strip()
    d = re.sub(r'[.\-_/,()]', ' ', d)      # punctuation ko space se replace
    d = re.sub(r'\bDISTT\.?\b', '', d)     # "DISTT" jaisa filler hata do
    d = re.sub(r'\bDISTRICT\b', '', d)
    d = re.sub(r'\s+', ' ', d).strip()     # multiple/extra spaces collapse
    return DISTRICT_ALIASES.get(d, d)


class DistrictMatcher:
    """
    Exact district match na milne par closest existing district
    (B side) dhoondh leta hai, taaki chhoti si spelling/spacing
    difference se pura match zero na ho jaaye. Result cache hota
    hai taaki baar baar recompute na ho.
    """
    def __init__(self, grouped, threshold=DISTRICT_FUZZY_THRESHOLD):
        self.grouped = grouped
        self.threshold = threshold
        self.cache = {}
        self.unmatched_districts = set()

    def get_group(self, dist):
        if dist in self.grouped.groups:
            return self.grouped.get_group(dist)
        if dist in self.cache:
            cand = self.cache[dist]
            return self.grouped.get_group(cand) if cand else None
        best, best_score = None, 0.0
        for cand in self.grouped.groups.keys():
            score = SequenceMatcher(None, dist, cand).ratio()
            if score > best_score:
                best_score, best = score, cand
        if best_score >= self.threshold:
            self.cache[dist] = best
            return self.grouped.get_group(best)
        self.cache[dist] = None
        self.unmatched_districts.add(dist)
        return None


# ═══════════════════════════════════════════════════════════════
#  SINGLE B FILE MATCHING (purana, single tab output)
# ═══════════════════════════════════════════════════════════════
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
    dmatch = DistrictMatcher(b_grouped)

    candidates_all = []
    for idx_a, row_a in df_a.iterrows():
        group = dmatch.get_group(row_a["_dist"])
        if group is not None:
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
            group = dmatch.get_group(row_a["_dist"])
            if group is not None:
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
        "unmatched_districts": sorted(dmatch.unmatched_districts),
    }


# ═══════════════════════════════════════════════════════════════
#  MULTI B FILE MATCHING (naya, multi-tab output)
# ═══════════════════════════════════════════════════════════════
def run_matching_multi(file_a_path, file_b_paths_with_names, output_path):
    """
    file_a_path: List A ka path (string)
    file_b_paths_with_names: list of tuples [(tab_name, file_path), ...]
    output_path: output excel ka path
    """
    df_a = pd.read_excel(file_a_path)
    df_a.columns = df_a.columns.str.strip()

    COL_A_DIST = [c for c in df_a.columns if 'district' in c.lower()][0]
    COL_A_STATE = [c for c in df_a.columns if 'state' in c.lower()]

    df_a["_name"] = df_a[COL_NAME].apply(clean_text)
    df_a["_dist"] = df_a[COL_A_DIST].apply(normalize_district)
    if COL_A_STATE:
        df_a["_state"] = df_a[COL_A_STATE[0]].str.upper().str.strip()
    else:
        df_a["_state"] = ""

    sheet_results = {}
    per_file_stats = []
    all_unmatched_districts = set()

    for tab_name, file_b_path in file_b_paths_with_names:
        df_b = pd.read_excel(file_b_path)
        df_b.columns = df_b.columns.str.strip()

        COL_B_DIST = [c for c in df_b.columns if 'district' in c.lower()][0]
        b_state_col = [c for c in df_b.columns if 'state' in c.lower()]

        df_b["_name"] = df_b[COL_NAME].apply(clean_text)
        df_b["_dist"] = df_b[COL_B_DIST].apply(normalize_district)

        if b_state_col:
            b_state = str(df_b[b_state_col[0]].iloc[0]).upper().strip()
            df_a_f = df_a[df_a["_state"] == b_state].copy() if COL_A_STATE else df_a.copy()
        else:
            df_a_f = df_a.copy()

        if df_a_f.empty:
            a_raw_cols = [c for c in df_a.columns if not c.startswith('_')]
            b_raw_cols = [c for c in df_b.columns if not c.startswith('_')]
            b_renamed = ["B_" + c if c in a_raw_cols else c for c in b_raw_cols]
            rows = []
            for _, row_b in df_b.iterrows():
                b_cols = list(row_b[b_raw_cols])
                rows.append([None] * len(a_raw_cols) + b_cols + [0, "UNMATCHED_B"])
            cols = a_raw_cols + b_renamed + ["Match %", "Status"]
            result_df = pd.DataFrame(rows, columns=cols)
            sheet_results[tab_name] = result_df
            per_file_stats.append({
                "tab": tab_name, "matched": 0, "unmatched_a": 0,
                "unmatched_b": len(df_b), "match_rate": 0.0
            })
            continue

        b_grouped = df_b.groupby("_dist")
        dmatch = DistrictMatcher(b_grouped)

        candidates_all = []
        for idx_a, row_a in df_a_f.iterrows():
            group = dmatch.get_group(row_a["_dist"])
            if group is not None:
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

        a_raw_cols = [c for c in df_a_f.columns if not c.startswith('_')]
        b_raw_cols = [c for c in df_b.columns if not c.startswith('_')]
        b_renamed = ["B_" + c if c in a_raw_cols else c for c in b_raw_cols]

        rows = []
        for idx_a, row_a in df_a_f.iterrows():
            a_cols = list(row_a[a_raw_cols])
            if idx_a in matches:
                idx_b, score = matches[idx_a]
                b_cols = list(df_b.loc[idx_b, b_raw_cols])
                rows.append(a_cols + b_cols + [score, "MATCHED"])
            else:
                best_score = 0
                group = dmatch.get_group(row_a["_dist"])
                if group is not None:
                    for idx_b, row_b in group.iterrows():
                        s = similarity(row_a["_name"], row_b["_name"])
                        if s > best_score:
                            best_score = s
                rows.append(a_cols + [None] * len(b_raw_cols) + [best_score, "UNMATCHED_A"])

        for idx_b, row_b in df_b.iterrows():
            if idx_b not in used_b:
                rows.append([None] * len(a_raw_cols) + list(row_b[b_raw_cols]) + [0, "UNMATCHED_B"])

        cols = a_raw_cols + b_renamed + ["Match %", "Status"]
        result_df = pd.DataFrame(rows, columns=cols)
        sheet_results[tab_name] = result_df

        m  = (result_df["Status"] == "MATCHED").sum()
        ua = (result_df["Status"] == "UNMATCHED_A").sum()
        ub = (result_df["Status"] == "UNMATCHED_B").sum()
        per_file_stats.append({
            "tab": tab_name, "matched": int(m), "unmatched_a": int(ua),
            "unmatched_b": int(ub),
            "match_rate": round(m / max(m + ua, 1) * 100, 1)
        })
        all_unmatched_districts |= dmatch.unmatched_districts

    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    RED    = PatternFill("solid", fgColor="FFC7CE")
    YELLOW = PatternFill("solid", fgColor="FFEB9C")
    HEADER = PatternFill("solid", fgColor="4472C4")
    WHITE_FONT = Font(bold=True, color="FFFFFF")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df = pd.DataFrame(per_file_stats)
        summary_df = summary_df[["tab", "matched", "unmatched_a", "unmatched_b", "match_rate"]]
        summary_df.columns = ["Tab", "Matched", "Unmatched A", "Unmatched B", "Match Rate %"]
        summary_df.to_excel(writer, sheet_name="SUMMARY", index=False)

        for tab_name, df in sheet_results.items():
            df.to_excel(writer, sheet_name=tab_name[:31], index=False)

    wb = load_workbook(output_path)

    def format_ws(ws, has_status=True):
        for cell in ws[1]:
            cell.fill = HEADER
            cell.font = WHITE_FONT
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.freeze_panes = "A2"
        if has_status:
            status_col = next((i for i, c in enumerate(ws[1], 1) if c.value == "Status"), None)
            if status_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    s = row[status_col - 1].value
                    fill = GREEN if s == "MATCHED" else (RED if s == "UNMATCHED_A" else YELLOW)
                    for cell in row:
                        cell.fill = fill
        for col_idx, col_cells in enumerate(ws.columns, 1):
            max_len = max((len(str(c.value)) if c.value else 0) for c in col_cells)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    format_ws(wb["SUMMARY"], has_status=False)
    for tab_name in sheet_results:
        sheet_name = tab_name[:31]
        if sheet_name in wb.sheetnames:
            format_ws(wb[sheet_name])
    wb.save(output_path)

    total_matched = sum(s["matched"] for s in per_file_stats)
    total_unmatched_a = sum(s["unmatched_a"] for s in per_file_stats)
    total_unmatched_b = sum(s["unmatched_b"] for s in per_file_stats)

    return {
        "per_file": per_file_stats,
        "total_matched": total_matched,
        "total_unmatched_a": total_unmatched_a,
        "total_unmatched_b": total_unmatched_b,
        "unmatched_districts": sorted(all_unmatched_districts),
    }
