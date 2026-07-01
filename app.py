import streamlit as st
import openpyxl
import io
from pathlib import Path

st.set_page_config(page_title="Repo Summary", page_icon="📊", layout="centered")

CURRENCIES = ['AMD', 'USD', 'EUR', 'RUB', 'AYL']
AYL_CELL   = '\u0531\u0545\u053c'

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .title { font-size: 1.5rem; font-weight: 600; color: #0f172a; margin-bottom: 0; }
    .subtitle { font-size: 0.9rem; color: #64748b; margin-top: 4px; margin-bottom: 2rem; }

    .summary-table { width: 100%; border-collapse: collapse; margin: 1.5rem 0; }
    .summary-table th {
        background: #0f172a; color: #fff;
        padding: 10px 16px; text-align: right; font-weight: 500;
        font-size: 0.85rem; letter-spacing: 0.05em;
    }
    .summary-table th:first-child { text-align: left; }
    .summary-table td {
        padding: 10px 16px; text-align: right;
        border-bottom: 1px solid #e2e8f0; font-size: 0.9rem; color: #1e293b;
    }
    .summary-table td:first-child { text-align: left; color: #475569; font-weight: 500; }
    .summary-table tr:last-child td { border-bottom: none; font-weight: 600; color: #0f172a; }
    .summary-table .dash { color: #cbd5e1; }

    .file-badge {
        display: inline-block; background: #f1f5f9; border: 1px solid #e2e8f0;
        border-radius: 6px; padding: 4px 10px; font-size: 0.8rem;
        color: #475569; margin: 3px;
    }
    .section-label {
        font-size: 0.75rem; font-weight: 600; letter-spacing: 0.08em;
        text-transform: uppercase; color: #94a3b8; margin: 1.5rem 0 0.5rem;
    }
</style>
""", unsafe_allow_html=True)


def get_merged_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for mc in ws.merged_cells.ranges:
        if mc.min_col <= col <= mc.max_col and mc.min_row <= row <= mc.max_row:
            return ws.cell(row=mc.min_row, column=mc.min_col).value
    return None


def parse_sheet(ws):
    COL_C, COL_F, COL_H, COL_L, COL_N = 3, 6, 8, 12, 14
    label_map = { 'AMD': 'AMD', 'USD': 'USD', 'EUR': 'EUR', 'RUB': 'RUB', AYL_CELL: 'AYL' }
    data = {cur: {'F': [], 'H': [], 'L': [], 'N': []} for cur in CURRENCIES}

    for row_num in range(1, ws.max_row + 1):
        raw = get_merged_value(ws, row_num, COL_C)
        currency = label_map.get(raw)
        if currency is None:
            continue
        f = ws.cell(row=row_num, column=COL_F).value
        h = ws.cell(row=row_num, column=COL_H).value
        l = ws.cell(row=row_num, column=COL_L).value
        n = ws.cell(row=row_num, column=COL_N).value
        if isinstance(f, (int, float)) and f > 0 and isinstance(h, (int, float)):
            data[currency]['F'].append(f)
            data[currency]['H'].append(h)
        if isinstance(l, (int, float)) and l > 0 and isinstance(n, (int, float)):
            data[currency]['L'].append(l)
            data[currency]['N'].append(n)
    return data


def weighted_avg(values, weights):
    total_w = sum(weights)
    if total_w == 0:
        return None
    return sum(v * w for v, w in zip(values, weights)) / total_w


def process_files(uploaded_files):
    accum = {cur: {'F_vals': [], 'H_vals': [], 'L_vals': [], 'N_vals': []} for cur in CURRENCIES}
    errors = []

    for uf in uploaded_files:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(uf.read()), data_only=True)
        except Exception as e:
            errors.append(f"{uf.name}: {e}")
            continue
        if '6.2-Repo' not in wb.sheetnames:
            errors.append(f"{uf.name}: Sheet '6.2-Repo' not found.")
            continue
        ws = wb['6.2-Repo']
        day_data = parse_sheet(ws)
        for cur in CURRENCIES:
            accum[cur]['F_vals'].extend(day_data[cur]['F'])
            accum[cur]['H_vals'].extend(day_data[cur]['H'])
            accum[cur]['L_vals'].extend(day_data[cur]['L'])
            accum[cur]['N_vals'].extend(day_data[cur]['N'])

    results = {}
    for cur in CURRENCIES:
        f_list = accum[cur]['F_vals']
        h_list = accum[cur]['H_vals']
        l_list = accum[cur]['L_vals']
        n_list = accum[cur]['N_vals']
        total_f = sum(f_list)
        total_l = sum(l_list)
        results[cur] = {
            'total':   total_f + total_l,
            'total_F': total_f,
            'total_L': total_l,
            'wavg':    weighted_avg(h_list + n_list, f_list + l_list),
        }

    return results, errors


def fmt_num(n):
    return '—' if not n else f"{n:,.0f}"

def fmt_pct(p):
    return '—' if p is None else f"{p * 100:.2f}%"


# ── UI ────────────────────────────────────────────────────────

st.markdown('<div class="title">Secondary Market — Repo Summary</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload daily files (Monday – Friday) to generate the weekly summary.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Upload files", type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if uploaded:
    st.markdown('<div class="section-label">Loaded files</div>', unsafe_allow_html=True)
    badges = "".join(f'<span class="file-badge">📄 {f.name}</span>' for f in uploaded)
    st.markdown(badges, unsafe_allow_html=True)

    results, errors = process_files(uploaded)

    if errors:
        for e in errors:
            st.error(e)

    has_data = any(results[c]['total'] > 0 for c in ['AMD', 'EUR', 'USD', 'RUB'])

    if has_data:
        st.markdown('<div class="section-label">Weekly summary</div>', unsafe_allow_html=True)

        cols = ['AMD', 'EUR', 'USD', 'RUB']
        headers = "<tr><th>Category</th>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"

        def amt_cell(v):
            s = fmt_num(v)
            cls = ' class="dash"' if s == '—' else ''
            return f"<td{cls}>{s}</td>"

        def pct_cell(v):
            s = fmt_pct(v)
            cls = ' class="dash"' if s == '—' else ''
            return f"<td{cls}>{s}</td>"

        row1 = "<tr><td>Banks and Financial Institutions</td>" + "".join(amt_cell(results[c]['total']) for c in cols) + "</tr>"
        row2 = "<tr><td>Weighted Avg Rate</td>" + "".join(pct_cell(results[c]['wavg']) for c in cols) + "</tr>"

        table_html = f'''
        <table class="summary-table">
            <thead>{headers}</thead>
            <tbody>{row1}{row2}</tbody>
        </table>
        '''
        st.markdown(table_html, unsafe_allow_html=True)

        ayl = results.get('AYL', {})
        if ayl.get('total', 0) > 0:
            st.info(f"AYL (Other currency): {fmt_num(ayl['total'])}  |  Weighted Avg Rate: {fmt_pct(ayl['wavg'])}")

        with st.expander("Breakdown by currency"):
            for cur in ['AMD', 'USD', 'EUR', 'RUB', 'AYL']:
                r = results[cur]
                if r['total'] > 0:
                    st.markdown(
                        f"**{cur}** — Total: `{fmt_num(r['total'])}` "
                        f"&nbsp;(Table 1/F: `{fmt_num(r['total_F'])}`, Table 2/L: `{fmt_num(r['total_L'])}`) "
                        f"&nbsp;Weighted Avg: `{fmt_pct(r['wavg'])}`"
                    )
    else:
        st.warning("No data found. Make sure the uploaded files contain a '6.2-Repo' sheet with data.")

else:
    st.info("Upload one or more .xlsx files above to get started.")
