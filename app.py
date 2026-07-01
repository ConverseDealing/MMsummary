import streamlit as st
import openpyxl
import csv
import io
from datetime import date, timedelta

st.set_page_config(page_title="Money Market Summary", page_icon="📊", layout="centered")

AYL_CELL   = "\u0531\u0545\u053c"
CURRENCIES = ["AMD", "USD", "EUR", "RUB", "AYL"]

CSS = """ + repr(css) + """

st.markdown(CSS, unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────

def get_previous_week_range(today=None):
    if today is None:
        today = date.today()
    this_monday = today - timedelta(days=today.weekday())
    prev_monday = this_monday - timedelta(weeks=1)
    return prev_monday, prev_monday + timedelta(days=4)

def parse_num(s):
    try: return float(str(s).strip())
    except: return None

def parse_date_dmy(s):
    from datetime import datetime
    try: return datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except: return None

def weighted_avg(values, weights):
    tw = sum(weights)
    return None if tw == 0 else sum(v * w for v, w in zip(values, weights)) / tw

def fmt_num(n):
    return "\u2014" if not n else f"{n:,.0f}"

def fmt_rate(p, decimals=4):
    return "\u2014" if p is None else f"{p:.{decimals}f}"

def table_html(rows, headers):
    th = "".join(f"<th>{h}</th>" for h in headers)
    body = ""
    for row in rows:
        body += "<tr>" + "".join(
            f'<td class="cb-dash">{v}</td>' if v == "\u2014" else f"<td>{v}</td>"
            for v in row
        ) + "</tr>"
    return f'<table class="cb-table"><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


# ── 6.2-Repo xlsx ─────────────────────────────────────────────

def get_merged_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for mc in ws.merged_cells.ranges:
        if mc.min_col <= col <= mc.max_col and mc.min_row <= row <= mc.max_row:
            return ws.cell(row=mc.min_row, column=mc.min_col).value
    return None

def parse_repo_sheet(ws):
    COL_C, COL_F, COL_H, COL_L, COL_N = 3, 6, 8, 12, 14
    label_map = {"AMD":"AMD","USD":"USD","EUR":"EUR","RUB":"RUB", AYL_CELL:"AYL"}
    data = {cur: {"F":[],"H":[],"L":[],"N":[]} for cur in CURRENCIES}
    for r in range(1, ws.max_row + 1):
        cur = label_map.get(get_merged_value(ws, r, COL_C))
        if not cur: continue
        f = ws.cell(row=r, column=COL_F).value
        h = ws.cell(row=r, column=COL_H).value
        l = ws.cell(row=r, column=COL_L).value
        n = ws.cell(row=r, column=COL_N).value
        if isinstance(f,(int,float)) and f>0 and isinstance(h,(int,float)):
            data[cur]["F"].append(f); data[cur]["H"].append(h)
        if isinstance(l,(int,float)) and l>0 and isinstance(n,(int,float)):
            data[cur]["L"].append(l); data[cur]["N"].append(n)
    return data

def process_repo_xlsx(files):
    accum = {cur:{"F_vals":[],"H_vals":[],"L_vals":[],"N_vals":[]} for cur in CURRENCIES}
    errors = []
    for uf in files:
        try: wb = openpyxl.load_workbook(io.BytesIO(uf.read()), data_only=True)
        except Exception as e: errors.append(f"{uf.name}: {e}"); continue
        if "6.2-Repo" not in wb.sheetnames:
            errors.append(f"{uf.name}: Sheet 6.2-Repo not found."); continue
        d = parse_repo_sheet(wb["6.2-Repo"])
        for cur in CURRENCIES:
            accum[cur]["F_vals"] += d[cur]["F"]; accum[cur]["H_vals"] += d[cur]["H"]
            accum[cur]["L_vals"] += d[cur]["L"]; accum[cur]["N_vals"] += d[cur]["N"]
    results = {}
    for cur in CURRENCIES:
        fl=accum[cur]["F_vals"]; hl=accum[cur]["H_vals"]
        ll=accum[cur]["L_vals"]; nl=accum[cur]["N_vals"]
        results[cur] = {
            "total": sum(fl)+sum(ll), "total_F": sum(fl), "total_L": sum(ll),
            "wavg": weighted_avg(hl+nl, fl+ll),
        }
    return results, errors


# ── Auction CSV ───────────────────────────────────────────────

def process_auction_csv(raw, mon, fri):
    reader = csv.reader(io.StringIO(raw.decode("utf-8-sig")))
    next(reader)
    declared, satisfied, rq, re_ = [], [], [], []
    for row in reader:
        if len(row) < 17: continue
        d = parse_date_dmy(row[5])
        if d is None or not (mon <= d <= fri): continue
        q = parse_num(row[16]); e = parse_num(row[4]); n = parse_num(row[13])
        if q and q > 0 and n is not None: declared.append(q); rq.append(n)
        if e and e > 0 and n is not None: satisfied.append(e); re_.append(n)
    return {
        "declared": sum(declared), "satisfied": sum(satisfied),
        "wavg_q": weighted_avg(rq, declared), "wavg_e": weighted_avg(re_, satisfied),
    }


# ── FX CSV ────────────────────────────────────────────────────

def process_fx_csv(raw, mon, fri):
    reader = csv.reader(io.StringIO(raw.decode("utf-8-sig")))
    next(reader)
    declared, satisfied, ri, re_ = [], [], [], []
    for row in reader:
        if len(row) < 12: continue
        d = parse_date_dmy(row[5])
        if d is None or not (mon <= d <= fri): continue
        i = parse_num(row[8]); e = parse_num(row[4]); l = parse_num(row[11])
        if i and i > 0 and l is not None: declared.append(i); ri.append(l)
        if e and e > 0 and l is not None: satisfied.append(e); re_.append(l)
    return {
        "declared": sum(declared), "satisfied": sum(satisfied),
        "wavg_i": weighted_avg(ri, declared), "wavg_e": weighted_avg(re_, satisfied),
    }


# ═══════════════════════════════════════════════════════════════
# PAGE
# ═══════════════════════════════════════════════════════════════

prev_mon, prev_fri = get_previous_week_range()
week_str = f"{prev_mon.strftime('%b %d')} \u2013 {prev_fri.strftime('%b %d, %Y')}"

st.markdown(f"""
<div class="cb-header">
    <div class="cb-header-sub">Conversebank</div>
    <div class="cb-header-title">Money Market Summary</div>
</div>
<div class="week-badge">&#9632;&nbsp; Previous week &nbsp;&middot;&nbsp; {week_str}</div>
""", unsafe_allow_html=True)


# ── Section 1 ─────────────────────────────────────────────────
st.markdown('<div class="cb-section">Section 1</div>', unsafe_allow_html=True)
st.markdown('<div class="cb-section-title">Repo Operations</div>', unsafe_allow_html=True)
st.markdown('<div class="cb-section-sub">Upload one .xlsx file per day (up to 5 for the full week).</div>', unsafe_allow_html=True)

repo_files = st.file_uploader("repo_xlsx", type=["xlsx"],
    accept_multiple_files=True, label_visibility="collapsed", key="repo")

if repo_files:
    badges = "".join(f'<span class="file-badge">{f.name}</span>' for f in repo_files)
    st.markdown(badges, unsafe_allow_html=True)
    results, errors = process_repo_xlsx(repo_files)
    for e in errors: st.error(e)
    cols = ["AMD", "EUR", "USD", "RUB"]
    rows = [
        ["Banks and Financial Institutions"] + [fmt_num(results[c]["total"]) for c in cols],
        ["Weighted Avg Rate"]                + [fmt_rate(results[c]["wavg"]) for c in cols],
    ]
    st.markdown(table_html(rows, [""]+cols), unsafe_allow_html=True)
    ayl = results.get("AYL", {})
    if ayl.get("total", 0) > 0:
        st.info(f"AYL: {fmt_num(ayl['total'])}  |  Weighted Avg: {fmt_rate(ayl['wavg'])}")
else:
    st.markdown('<p class="cb-empty">No files uploaded yet.</p>', unsafe_allow_html=True)

st.markdown('<hr class="cb-divider">', unsafe_allow_html=True)


# ── Section 2 ─────────────────────────────────────────────────
st.markdown('<div class="cb-section">Section 2</div>', unsafe_allow_html=True)
st.markdown('<div class="cb-section-title">Repo / Reverse Repo / Auction</div>', unsafe_allow_html=True)
st.markdown(f'<div class="cb-section-sub">Upload the CSV \u2014 rows filtered to {week_str} automatically.</div>', unsafe_allow_html=True)

auction_file = st.file_uploader("auction_csv", type=["csv"],
    accept_multiple_files=False, label_visibility="collapsed", key="auction")

if auction_file:
    st.markdown(f'<span class="file-badge">{auction_file.name}</span>', unsafe_allow_html=True)
    res = process_auction_csv(auction_file.read(), prev_mon, prev_fri)
    if not res["declared"] and not res["satisfied"]:
        st.warning(f"No data found for {week_str}.")
    else:
        rows = [
            ["Declared Volume",  fmt_num(res["declared"]),  fmt_rate(res["wavg_q"])],
            ["Satisfied Volume", fmt_num(res["satisfied"]), fmt_rate(res["wavg_e"])],
        ]
        st.markdown(table_html(rows, ["", "Total", "Weighted Avg Rate"]), unsafe_allow_html=True)
else:
    st.markdown('<p class="cb-empty">No file uploaded yet.</p>', unsafe_allow_html=True)

st.markdown('<hr class="cb-divider">', unsafe_allow_html=True)


# ── Section 3 ─────────────────────────────────────────────────
st.markdown('<div class="cb-section">Section 3</div>', unsafe_allow_html=True)
st.markdown('<div class="cb-section-title">Foreign Exchange</div>', unsafe_allow_html=True)
st.markdown(f'<div class="cb-section-sub">Upload the CSV \u2014 rows filtered to {week_str} automatically.</div>', unsafe_allow_html=True)

fx_file = st.file_uploader("fx_csv", type=["csv"],
    accept_multiple_files=False, label_visibility="collapsed", key="fx")

if fx_file:
    st.markdown(f'<span class="file-badge">{fx_file.name}</span>', unsafe_allow_html=True)
    res = process_fx_csv(fx_file.read(), prev_mon, prev_fri)
    if not res["declared"] and not res["satisfied"]:
        st.warning(f"No data found for {week_str}.")
    else:
        rows = [
            ["Declared Volume",  fmt_num(res["declared"]),  fmt_rate(res["wavg_i"])],
            ["Satisfied Volume", fmt_num(res["satisfied"]), fmt_rate(res["wavg_e"])],
        ]
        st.markdown(table_html(rows, ["", "Total", "Weighted Avg Rate"]), unsafe_allow_html=True)
else:
    st.markdown('<p class="cb-empty">No file uploaded yet.</p>', unsafe_allow_html=True)

st.markdown('<div class="cb-footer">CONVERSEBANK &middot; MONEY MARKET &middot; INTERNAL USE ONLY</div>', unsafe_allow_html=True)
