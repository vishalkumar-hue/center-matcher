# ═══════════════════════════════════════════════════════════════
#  YE FUNCTION matcher.py KE END MEIN ADD KARNA HAI
#  (existing run_matching() function ko mat hatao, bas neeche ye jodo)
# ═══════════════════════════════════════════════════════════════

def run_matching_multi(file_a_path, file_b_paths_with_names, output_path):
    """
    file_a_path: List A ka path (string)
    file_b_paths_with_names: list of tuples [(tab_name, file_path), ...]
    output_path: output excel ka path

    Returns: dict with per-file stats + overall summary
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

    for tab_name, file_b_path in file_b_paths_with_names:
        df_b = pd.read_excel(file_b_path)
        df_b.columns = df_b.columns.str.strip()

        COL_B_DIST = [c for c in df_b.columns if 'district' in c.lower()][0]
        b_state_col = [c for c in df_b.columns if 'state' in c.lower()]

        df_b["_name"] = df_b[COL_NAME].apply(clean_text)
        df_b["_dist"] = df_b[COL_B_DIST].apply(normalize_district)

        # B file ka state nikalo, A ko usi state pe filter karo
        if b_state_col:
            b_state = str(df_b[b_state_col[0]].iloc[0]).upper().strip()
            df_a_f = df_a[df_a["_state"] == b_state].copy() if COL_A_STATE else df_a.copy()
        else:
            df_a_f = df_a.copy()

        if df_a_f.empty:
            # State match nahi mila — sab B rows UNMATCHED_B
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
        candidates_all = []
        for idx_a, row_a in df_a_f.iterrows():
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
                if row_a["_dist"] in b_grouped.groups:
                    group = b_grouped.get_group(row_a["_dist"])
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

    # ── Excel likho: SUMMARY tab + har B file ka apna tab ─────
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
    }
