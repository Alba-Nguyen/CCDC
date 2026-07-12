import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, sqlite3, shutil, re
from datetime import datetime
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENTITY = "AS"
DATA_DIR = os.path.join(BASE_DIR, "CCDC 1", ENTITY)

NUM = '#,##0'
THIN = Side(style='thin')
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE = "2F75B5"
WHITE = "FFFFFF"
PAIR_COLORS = ["D6E4F0", "FCE4D6", "E2EFDA", "E8D5F5", "FFF2CC"]

VIET = str.maketrans({
    'à':'a','á':'a','ã':'a','ạ':'a','ả':'a',
    'ă':'a','ắ':'a','ằ':'a','ẳ':'a','ẵ':'a','ặ':'a',
    'â':'a','ấ':'a','ầ':'a','ẩ':'a','ẫ':'a','ậ':'a',
    'è':'e','é':'e','ẹ':'e','ẻ':'e','ẽ':'e',
    'ê':'e','ề':'e','ế':'e','ể':'e','ễ':'e','ệ':'e',
    'đ':'d',
    'ì':'i','í':'i','ị':'i','ỉ':'i','ĩ':'i',
    'ò':'o','ó':'o','ọ':'o','ỏ':'o','õ':'o',
    'ô':'o','ố':'o','ồ':'o','ổ':'o','ỗ':'o','ộ':'o',
    'ơ':'o','ớ':'o','ờ':'o','ở':'o','ỡ':'o','ợ':'o',
    'ù':'u','ú':'u','ụ':'u','ủ':'u','ũ':'u',
    'ư':'u','ứ':'u','ừ':'u','ử':'u','ữ':'u','ự':'u',
    'ỳ':'y','ý':'y','ỵ':'y','ỷ':'y','ỹ':'y',
})

def _find(folder, keywords):
    if not os.path.isdir(folder): return None
    for name in os.listdir(folder):
        plain = name.lower().translate(VIET)
        if all(k in plain for k in keywords):
            return os.path.join(folder, name)
    return None

def _f(v):
    if v is None: return 0.0
    try: return float(v)
    except: return 0.0

def _fmt(n):
    return round(float(n or 0), 0)

STRIP_PREFIXES = ["phân bổ chi phí công cụ dụng cụ:", "hạch toán chi phí chờ phân bổ:"]

def _strip_prefix(dien_giai):
    s = str(dien_giai or '').strip()
    for prefix in STRIP_PREFIXES:
        if prefix in s.lower()[:80]:
            parts = s.split(":", 1)
            if len(parts) > 1:
                s = parts[1].strip()
                break
    return s

def _has_month_ref(s):
    return bool(re.search(r'(?:[Tt]\d{1,2}(?:\.\d{4})?|tháng\s+\d{1,2}(?:\s*[-–]\s*\d{1,2})?)', s))

def _strip_suffix(name):
    return re.sub(r'\s+[Tt]\d+\.\d{4}$', '', name).strip()

def _strip_suffix_display(name):
    s = str(name or '').strip()
    m = re.search(r'\s+[Tt](\d{1,2})(?:\.\d{4})?\s*$', s)
    if m:
        body = s[:m.start()]
        if _has_month_ref(body):
            return body.strip()
    return s

def _clean_dg_display(dien_giai):
    s = _strip_prefix(dien_giai)
    s = _strip_suffix_display(s)
    return s

def _clean_dg(dien_giai):
    s = _strip_prefix(dien_giai)
    s = _strip_suffix(s)
    return s

def _fmt_date(v):
    if hasattr(v, 'strftime'):
        return v.strftime('%d/%m/%Y')
    return str(v or '').strip()

def _extract_ccdc_name(dien_giai):
    s = _strip_prefix(dien_giai)
    if s == str(dien_giai or '').strip():
        return None
    for sep in [f" t{i}." for i in range(1, 13)]:
        idx = s.lower().find(sep)
        if idx > 0:
            s = s[:idx]
    return s.strip() or None

def _match_ccdc_item(ccdc_items, bk_name, bk_amount=0):
    if not bk_name: return None
    bn = bk_name.lower().strip()
    exact = []; prefix = []
    for item in ccdc_items:
        ten = (item['ten'] or '').lower().strip()
        if not ten: continue
        if bn == ten or ten == bn:
            exact.append(item)
        elif bn.startswith(ten) or ten.startswith(bn):
            prefix.append(item)
        else:
            bn_words = set(bn.replace('-',' ').split())
            ten_words = set(ten.replace('-',' ').split())
            common = bn_words & ten_words
            if len(common) >= 2 and len(common) >= min(len(bn_words), len(ten_words)) * 0.6:
                prefix.append(item)
    candidates = exact or prefix
    if not candidates: return None
    if bk_amount:
        for item in candidates:
            if abs((item['phan_bo'] or 0) - bk_amount) < 1000:
                return item
        return min(candidates, key=lambda x: abs((x['phan_bo'] or 0) - bk_amount))
    return candidates[0]

def main():
    import sys
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    TMP = os.path.join(BASE_DIR, "~tmp")
    if os.path.exists(TMP):
        shutil.rmtree(TMP, ignore_errors=True)
    os.makedirs(TMP, exist_ok=True)
    db_path = os.path.join(TMP, "ccdc.db")

    # ── Find 5 source files ──
    f_bcdps = _find(DATA_DIR, ["can doi", "phat sinh"])
    f_ccdc  = _find(DATA_DIR, ["tinh gia tri phan bo", "cong cu"])
    f_cp    = _find(DATA_DIR, ["tong hop", "chi phi", "cho phan bo"])
    f_2421  = _find(DATA_DIR, ["so chi tiet", "2421"])
    f_2422  = _find(DATA_DIR, ["so chi tiet", "2422"])

    if not all([f_bcdps, f_ccdc, f_cp, f_2421, f_2422]):
        print("[LOI] Khong tim du 5 file trong:", DATA_DIR)
        print(f"  BCĐPS: {f_bcdps}")
        print(f"  CCDC:  {f_ccdc}")
        print(f"  CP:    {f_cp}")
        print(f"  Sổ 2421: {f_2421}")
        print(f"  Sổ 2422: {f_2422}")
        return

    def _safe_copy(src):
        dst = os.path.join(TMP, os.path.basename(src))
        try: shutil.copy2(src, dst); return dst
        except: return src

    f_bcdps_cp = _safe_copy(f_bcdps)
    f_ccdc_cp  = _safe_copy(f_ccdc)
    f_cp_cp    = _safe_copy(f_cp)
    f_2421_cp  = _safe_copy(f_2421)
    f_2422_cp  = _safe_copy(f_2422)

    conn = sqlite3.connect(db_path)
    src_sheets = []

    def _load_ws(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        return wb, wb.active

    # ── 1. BCĐPS ──
    wb, ws = _load_ws(f_bcdps_cp)
    src_sheets.append((wb, "BCĐPS"))
    bcdps_rows = {}
    conn.execute("CREATE TABLE bcdps (tk TEXT, ten TEXT, du_no_dk REAL, du_co_dk REAL, ps_no REAL, ps_co REAL, du_no_ck REAL, du_co_ck REAL)")
    for r in range(10, ws.max_row + 1):
        tk = ws.cell(r, 1).value
        if tk is not None:
            tk_str = str(tk).strip()
            if tk_str in ('242','2421','2422'):
                bcdps_rows[tk_str] = r
            conn.execute("INSERT INTO bcdps VALUES (?,?,?,?,?,?,?,?)",
                         (tk_str, ws.cell(r, 2).value,
                          _f(ws.cell(r, 4).value), _f(ws.cell(r, 5).value),
                          _f(ws.cell(r, 6).value), _f(ws.cell(r, 7).value),
                          _f(ws.cell(r, 10).value), _f(ws.cell(r, 11).value)))

    # ── 2. Bảng tính giá trị phân bổ CCDC ──
    wb, ws = _load_ws(f_ccdc_cp)
    src_sheets.append((wb, "CCDC"))
    ccdc_rows = {}
    conn.execute("CREATE TABLE ccdc (ma TEXT, ten TEXT, con_lai_dk REAL, phan_bo REAL, con_lai_ck REAL, tk_no TEXT, tk_co TEXT, tk_account TEXT, note TEXT)")
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, 1).value is None: continue
        ma = str(ws.cell(r, 1).value or '').strip()
        if ma:
            ccdc_rows[ma] = r
        tk_no = str(ws.cell(r, 18).value or '').strip()
        tk_co = str(ws.cell(r, 19).value or '').strip()
        if tk_no in ('2421','2422'): tk_account = tk_no
        elif tk_co in ('2421','2422'): tk_account = tk_co
        elif tk_no == '242': tk_account = tk_no
        elif tk_co == '242': tk_account = tk_co
        else:
            tk_account = ''
        note = ''
        if tk_account == '' and _f(ws.cell(r, 14).value) != 0:
            note = 'Check - TK Nợ=Có, cần xác nhận'
        conn.execute("INSERT INTO ccdc VALUES (?,?,?,?,?,?,?,?,?)",
                     (ma, str(ws.cell(r, 2).value or ''),
                      _f(ws.cell(r, 10).value),  # con_lai_dk = col J
                      _f(ws.cell(r, 11).value),  # phan_bo = col K
                      _f(ws.cell(r, 14).value),  # con_lai_ck = col N
                      tk_no, tk_co, tk_account, note))

    # ── 3. CP ──
    wb, ws = _load_ws(f_cp_cp)
    src_sheets.append((wb, "CP chờ PB"))
    cp_rows = {}
    conn.execute("CREATE TABLE cp (ma TEXT, ten TEXT, con_lai REAL, pb_trong_ky REAL, tk_no TEXT, tk_co TEXT)")
    for r in range(6, ws.max_row + 1):
        if ws.cell(r, 1).value is None: continue
        ma = str(ws.cell(r, 1).value or '').strip()
        tk_co = str(ws.cell(r, 11).value or '').strip()
        if tk_co not in ('242','2421','2422'): continue
        if ma:
            cp_rows[ma] = r
        conn.execute("INSERT INTO cp VALUES (?,?,?,?,?,?)",
                     (ma, str(ws.cell(r, 2).value or ''),
                      _f(ws.cell(r, 9).value), _f(ws.cell(r, 7).value),
                      str(ws.cell(r, 10).value or '').strip(), tk_co))

    # ── 4. Sổ chi tiết TK 2421 ──
    conn.execute("CREATE TABLE so_242x (tk TEXT, ngay TEXT, so_ct TEXT, dien_giai TEXT, tk_doi_ung TEXT, ps_no REAL, ps_co REAL)")
    ledger_rows = {'2421': {}, '2422': {}}
    def _load_so(path, tk_label):
        wb, ws = _load_ws(path)
        src_sheets.append((wb, f"Sổ TK {tk_label}"))
        for r in range(9, ws.max_row + 1):  # row 8 = opening, data from row 9
            so_ct = str(ws.cell(r, 2).value or '').strip()
            if not so_ct: continue
            ledger_rows[tk_label].setdefault(so_ct, []).append(r)
            ngay = _fmt_date(ws.cell(r, 1).value)
            dg = str(ws.cell(r, 3).value or '')
            tk_du = str(ws.cell(r, 5).value or '').strip()
            ps_no = _f(ws.cell(r, 6).value)
            ps_co = _f(ws.cell(r, 8).value)
            if abs(ps_no) < 1000 and abs(ps_co) < 1000: continue
            conn.execute("INSERT INTO so_242x VALUES (?,?,?,?,?,?,?)",
                         (tk_label, ngay, so_ct, dg, tk_du, ps_no, ps_co))

    _load_so(f_2421_cp, "2421")
    _load_so(f_2422_cp, "2422")
    conn.commit()

    # ── Derive aggregate data ──
    b2421 = conn.execute("SELECT du_no_dk, du_co_dk, du_no_ck, ps_no, ps_co, ten FROM bcdps WHERE tk='2421'").fetchone() or (0,0,0,0,0,'')
    b2422 = conn.execute("SELECT du_no_dk, du_co_dk, du_no_ck, ps_no, ps_co, ten FROM bcdps WHERE tk='2422'").fetchone() or (0,0,0,0,0,'')

    ccdc_sum = conn.execute("SELECT tk_account, SUM(con_lai_ck), SUM(con_lai_dk), SUM(phan_bo) FROM ccdc GROUP BY tk_account").fetchall()
    cp_sum = conn.execute("SELECT tk_co, SUM(con_lai), SUM(pb_trong_ky) FROM cp GROUP BY tk_co").fetchall()

    ccdc_map = {r[0]: {"cl_ck":r[1], "cl_dk":r[2], "pb":r[3]} for r in ccdc_sum}
    cp_map = {r[0]: {"cl":r[1], "pb":r[2]} for r in cp_sum}

    # ── CP items (full) for matching ──
    cp_items_by_tk = {'2421': [], '2422': []}
    for r in conn.execute("SELECT ma, ten, pb_trong_ky, con_lai, tk_no, tk_co FROM cp WHERE tk_co IN ('2421','2422') AND (ABS(pb_trong_ky) >= 1000 OR ABS(con_lai) >= 1000) ORDER BY tk_co, ten"):
        ma, ten, pb, cl, tk_no, tk_co = r
        cp_items_by_tk[tk_co].append({
            'ma': ma, 'ten': str(ten or '').strip(), 'pb': _fmt(pb), 'cl': _fmt(cl), 'tk_no': tk_no,
        })

    # Build CCDC items list for matching (PBCCDC)
    ccdc_items = [{"ma":r[0], "ten":r[1], "phan_bo":_f(r[2]), "tk_account":r[3]}
                  for r in conn.execute("SELECT ma, ten, phan_bo, tk_account FROM ccdc").fetchall()]

    # CCDC items grouped by tk_account for per-TK analysis
    ccdc_items_by_tk = {'2421': [], '2422': []}
    ccdc_sql = "SELECT ma, ten, phan_bo, con_lai_ck, tk_account FROM ccdc WHERE tk_account IN ('2421','2422') AND (ABS(phan_bo) >= 1000 OR ABS(con_lai_ck) >= 1000) ORDER BY tk_account, ten"
    for r in conn.execute(ccdc_sql):
        ma, ten, pb, cl_ck, acct = r
        ccdc_items_by_tk[acct].append({
            'ma': ma, 'ten': str(ten or '').strip(), 'phan_bo': _fmt(pb), 'con_lai_ck': _fmt(cl_ck),
        })

    all_ledger = list(conn.execute("SELECT tk, ngay, so_ct, dien_giai, tk_doi_ung, ps_no, ps_co FROM so_242x ORDER BY tk, ngay, so_ct"))
    entries_by_tk = {'2421': [], '2422': []}

    for r in all_ledger:
        tk, ngay, so_ct, dg, tk_du, ps_no, ps_co = r
        ps_no = _fmt(ps_no)
        ps_co = _fmt(ps_co)
        contrib = _fmt(ps_no - ps_co)  # net change to 242x BCĐPS balance
        if abs(contrib) < 1000: continue  # skip zero-net entries

        is_cp = 'PBCP' in so_ct.upper()
        is_internal = tk_du in ('242','2421','2422')

        if is_cp:
            note = 'Phân bổ CP'
            bk_amt = 0; ccdc_amt = 0; cp_amt = abs(contrib)
        elif is_internal:
            note = 'Nội bộ 242'
            bk_amt = 0; ccdc_amt = 0; cp_amt = 0
        else:
            note = ''
            bk_amt = abs(contrib); ccdc_amt = 0; cp_amt = 0
            if 'PBCCDC' in so_ct.upper():
                name = _extract_ccdc_name(dg)
                if name:
                    match = _match_ccdc_item(ccdc_items, name, bk_amt)
                    if match:
                        ccdc_amt = _fmt(match['phan_bo'])
                        note = 'CCDC'
                        if abs(ccdc_amt - bk_amt) >= 1000:
                            note += f' (SL={ccdc_amt:,.0f})'

        entries_by_tk[tk].append({
            'tk_no': tk if ps_no > 0 else tk_du,
            'tk_co': tk_du if ps_no > 0 else tk,
            'so_ct': so_ct, 'ngay': ngay, 'dg': dg,
            'bk_amt': bk_amt, 'ccdc_amt': ccdc_amt, 'cp_amt': cp_amt,
            'contrib': contrib, 'note': note,
        })

    # Aggregate movement
    agg_bk_ccdc = {}
    for tk in ('2421', '2422'):
        b = b2421 if tk == '2421' else b2422
        bcdps_mv = _fmt(b[3] - b[4])
        cc = ccdc_map.get(tk, {})
        ccdc_mv = _fmt(cc.get('cl_ck', 0) - cc.get('cl_dk', 0))  # = -pb
        agg_bk_ccdc[tk] = _fmt(bcdps_mv - ccdc_mv)

    cp_mv_total = {}
    for tk in ('2421', '2422'):
        cp_mv_total[tk] = _fmt(cp_map.get(tk, {}).get('pb', 0))

    # Opening & closing data
    close_data = {}
    for tk in ('2421', '2422'):
        bv = _fmt(b2421[2] if tk == '2421' else b2422[2])
        cv = _fmt(ccdc_map.get(tk, {}).get('cl_ck', 0))
        pv = _fmt(cp_map.get(tk, {}).get('cl', 0))
        close_data[tk] = {'bcdps': bv, 'ccdc': cv, 'cp': pv, 'lech': _fmt(bv - cv - pv)}

    cp_open_total = {}
    for tk in ('2421', '2422'):
        closing = cp_map.get(tk, {}).get('cl', 0)
        pb_total = cp_mv_total[tk]
        cp_open_total[tk] = _fmt(closing + pb_total)

    open_data = {}
    for tk in ('2421', '2422'):
        bcdps_open = _fmt((b2421[0] - b2421[1]) if tk == '2421' else (b2422[0] - b2422[1]))
        ccdc_open = _fmt(ccdc_map.get(tk, {}).get('cl_dk', 0))
        open_data[tk] = {'bcdps': bcdps_open, 'ccdc': ccdc_open, 'cp': cp_open_total[tk],
                         'lech': _fmt(bcdps_open - ccdc_open - cp_open_total[tk])}

    # Pre-fetch unassigned CCDC items (before conn closes)
    ccdc_unassigned = [dict(ma=r[0], ten=r[1], cl_dk=_fmt(r[2]), pb=_fmt(r[3]), cl_ck=_fmt(r[4]), note=r[5])
                       for r in conn.execute("SELECT ma, ten, con_lai_dk, phan_bo, con_lai_ck, note "
                                             "FROM ccdc WHERE tk_account='' AND con_lai_ck != 0 ORDER BY ten")]

    # ── Build output ──
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    ws = wb_out.create_sheet("Phiếu cần điều chỉnh")
    MC = 12

    def _t(text, ec):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ec)
        c = ws.cell(1, 1, text)
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.font = Font(name="Arial", size=14, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 28

    def _sec(row, text, ec, color="D6E4F0"):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ec)
        c = ws.cell(row, 1, text)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(name="Arial", bold=True, size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22

    def _hdr(row, values, widths=None):
        for ci, v in enumerate(values, 1):
            c = ws.cell(row, ci, v)
            c.fill = PatternFill("solid", fgColor=BLUE)
            c.font = Font(name="Arial", bold=True, size=9, color=WHITE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = GRID
        if widths:
            for ci, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(ci)].width = w

    def _rw(row, values, fmt_cols=None, bold=False, color=None, fill=None):
        for ci, v in enumerate(values, 1):
            is_formula = isinstance(v, str) and v.startswith('=')
            c = ws.cell(row, ci, v if not is_formula else v)
            c.font = Font(name="Arial", size=9, bold=bold, color=("000000" if color is None else color))
            c.alignment = Alignment(
                horizontal="right" if (fmt_cols and ci in fmt_cols) else "left",
                vertical="center", wrap_text=True)
            c.border = GRID
            if fmt_cols and ci in fmt_cols:
                c.number_format = NUM
            if fill:
                c.fill = PatternFill("solid", fgColor=fill)

    # ── Render ──
    _t(f"PHIẾU CẦN ĐIỀU CHỈNH — ĐỐI CHIẾU TÀI KHOẢN 242 ({ENTITY})", MC)
    ws.merge_cells("A2:L2")
    ws["A2"] = f"Xuất lúc: {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(name="Arial", size=9, color="7F7F7F")

    row = 4
    _sec(row, "I. TỔNG HỢP CHÊNH LỆCH DƯ NỢ CUỐI KỲ", MC)
    row += 1
    _hdr(row, ["TK", "Nguồn", "Diễn giải", "BCĐPS", "CCDC", "CP chờ pb", "Tổng detail", "Chênh lệch", "Kết luận"],
         widths=[8, 10, 40, 18, 18, 18, 18, 18, 12])
    row += 1
    fm_a = {4, 5, 6, 7, 8}
    for tk_k, desc in [("242", "Chi phí trả trước (tổng)"), ("2421", None), ("2422", None)]:
        r = row
        if tk_k == "242":
            r1, r2 = bcdps_rows.get('2421', 0), bcdps_rows.get('2422', 0)
            bv_f = f'=BCĐPS!J{r1}-BCĐPS!K{r1}+BCĐPS!J{r2}-BCĐPS!K{r2}'
            cv_f = (f'=SUMIF(CCDC!R:R,"2421",CCDC!N:N)+SUMIF(CCDC!S:S,"2421",CCDC!N:N)'
                   f'+SUMIF(CCDC!R:R,"2422",CCDC!N:N)+SUMIF(CCDC!S:S,"2422",CCDC!N:N)')
            pv_f = (f'=SUMIF(\'CP chờ PB\'!K:K,"2421",\'CP chờ PB\'!I:I)'
                   f'+SUMIF(\'CP chờ PB\'!K:K,"2422",\'CP chờ PB\'!I:I)')
        else:
            rk = bcdps_rows.get(tk_k, 0)
            bv_f = f'=BCĐPS!J{rk}-BCĐPS!K{rk}'
            cv_f = (f'=SUMIF(CCDC!R:R,"{tk_k}",CCDC!N:N)+SUMIF(CCDC!S:S,"{tk_k}",CCDC!N:N)'
                   f'-SUMIFS(CCDC!N:N,CCDC!R:R,"{tk_k}",CCDC!S:S,"{tk_k}")')
            pv_f = f'=SUMIF(\'CP chờ PB\'!K:K,"{tk_k}",\'CP chờ PB\'!I:I)'
        tot_f = f'=E{r}+F{r}'
        df_f = f'=D{r}-G{r}'
        kl_f = f'=IF(ABS(H{r})<1000,"KHỚP","LỆCH")'
        dn = desc or (b2421[5] if tk_k == "2421" else b2422[5]) or f'TK {tk_k}'
        _rw(row, [tk_k, '', dn.strip(), bv_f, cv_f, pv_f, tot_f, df_f, kl_f], fmt_cols=fm_a)
        row += 1

    # ── II. CCDC chưa phân bổ ──
    if ccdc_unassigned:
        row += 1
        _sec(row, "II. CCDC CHƯA PHÂN BỔ", MC, color="FCE4D6")
        row += 1
        _hdr(row, ["STT", "Nguồn", "Diễn giải", "BCĐPS", "CCDC", "CP chờ pb", "Tổng detail", "Chênh lệch", "Kết luận"],
             widths=[5, 10, 60, 18, 18, 18, 18, 18, 12])
        row += 1
        u_start = row
        for u in ccdc_unassigned:
            cr = ccdc_rows.get(u['ma'], 0)
            ccdc_f = f'=CCDC!N{cr}'
            tot_f = f'=E{row}'
            lech_f = f'=-E{row}'
            _rw(row, [1, '', str(u['ten'])[:60], 0, ccdc_f, 0, tot_f, lech_f, u['note']],
                fmt_cols={4,5,6,7,8}, fill='FFF2CC')
            row += 1
        u_end = row - 1
        tot_ccdc_f = f'=SUM(E{u_start}:E{u_end})'
        tot_tot_f = f'=E{row}'
        tot_lech_f = f'=-E{row}'
        _rw(row, ['', '', 'TỔNG CCDC CHƯA PHÂN BỔ', 0, tot_ccdc_f, 0, tot_tot_f, tot_lech_f, ''],
            bold=True, fmt_cols={4,5,6,7,8}, fill='FFF2CC')
        row += 1

    # ── Render per-TK analysis: match CP items ↔ PBCP entries ──
    def _match_clean(s):
        return _clean_dg(s).lower().strip()

    def _group_pbcp(pbcp_entries):
        groups = {}
        for pe in pbcp_entries:
            k = _match_clean(pe['dg'])
            if k not in groups:
                groups[k] = {'total': 0, 'socp': [], 'pe': pe}
            groups[k]['total'] += abs(pe['contrib'])
            groups[k]['socp'].append(pe['so_ct'])
        return groups

    total_displayed = 0
    # ── Phase 1: Build all matches for both TKs ──
    all_matches = {}
    for tk in ('2421', '2422'):
        cl = close_data[tk]['lech']
        cp_items = cp_items_by_tk[tk]
        ledger = entries_by_tk[tk]

        pbcp_entries = [e for e in ledger if 'PBCP' in e['so_ct'].upper()]
        pbcp_groups = _group_pbcp(pbcp_entries)

        matches = []
        used_pbcp_keys = set()
        for cp in cp_items:
            cp_k = _match_clean(cp['ten'])
            cp_amt = cp['pb'] if cp['pb'] != 0 else cp['cl']
            if cp_k in pbcp_groups and cp_k not in used_pbcp_keys:
                used_pbcp_keys.add(cp_k)
                g = pbcp_groups[cp_k]
                pbcp_total = g['total']
                if cp['pb'] != 0:
                    diff = cp['pb'] - pbcp_total
                else:
                    diff = -(cp['cl'] - pbcp_total)
                matches.append({'src': 'CP', 'ten': _clean_dg_display(cp['ten']), 'cp_amt': cp_amt,
                    'pbcp_amt': pbcp_total, 'diff': diff,
                    'note': '' if abs(diff) < 1000 else 'Lệch tiền',
                    'so_ct_ledger': ','.join(g['socp']), 'so_ct_file': cp['ma'], 'tk': tk})
            else:
                if cp['pb'] != 0:
                    diff = cp['pb']
                else:
                    diff = -cp['cl']
                matches.append({'src': 'CP', 'ten': _clean_dg_display(cp['ten']), 'cp_amt': cp_amt,
                    'pbcp_amt': 0, 'diff': diff,
                    'note': 'Thiếu PBCP', 'so_ct_ledger': '', 'so_ct_file': cp['ma'], 'tk': tk})

        for k, g in pbcp_groups.items():
            if k in used_pbcp_keys: continue
            matches.append({'src': 'PBCP', 'ten': _clean_dg_display(g['pe']['dg']), 'cp_amt': 0,
                'pbcp_amt': g['total'], 'diff': -g['total'],
                'note': 'Thừa PBCP', 'so_ct_ledger': ','.join(g['socp']), 'so_ct_file': '', 'tk': tk})

        # ── CCDC ↔ PBCCDC matching ──
        cc_items = ccdc_items_by_tk.get(tk, [])
        pbccdc_entries = [e for e in ledger if 'PBCCDC' in e['so_ct'].upper()]
        if pbccdc_entries:
            pbccdc_groups = _group_pbcp(pbccdc_entries)
            used_pbccdc_keys = set()
            for cc in cc_items:
                cc_k = _match_clean(cc['ten'])
                if cc_k in pbccdc_groups and cc_k not in used_pbccdc_keys:
                    used_pbccdc_keys.add(cc_k)
                    g = pbccdc_groups[cc_k]
                    pbccdc_total = g['total']
                    diff = cc['phan_bo'] - pbccdc_total
                    matches.append({'src': 'CCDC', 'ten': _clean_dg_display(cc['ten']), 'cp_amt': cc['phan_bo'],
                        'pbcp_amt': pbccdc_total, 'diff': diff,
                        'note': '' if abs(diff) < 1000 else 'Lệch tiền',
                        'so_ct_ledger': ','.join(g['socp']), 'so_ct_file': cc['ma'], 'tk': tk})
                else:
                    note = 'Còn dư cuối kỳ' if cc['phan_bo'] == 0 else 'Thiếu PBCCDC'
                    matches.append({'src': 'CCDC', 'ten': _clean_dg_display(cc['ten']), 'cp_amt': cc['phan_bo'],
                        'pbcp_amt': 0, 'diff': cc['phan_bo'],
                        'note': note, 'so_ct_ledger': '', 'so_ct_file': cc['ma'], 'tk': tk})
            for k, g in pbccdc_groups.items():
                if k in used_pbccdc_keys: continue
                matches.append({'src': 'PBCCDC', 'ten': _clean_dg_display(g['pe']['dg']), 'cp_amt': 0,
                    'pbcp_amt': g['total'], 'diff': -g['total'],
                    'note': 'Thừa PBCCDC', 'so_ct_ledger': ','.join(g['socp']), 'so_ct_file': '', 'tk': tk})
        else:
            for cc in cc_items:
                note = 'Còn dư cuối kỳ' if cc['phan_bo'] == 0 else 'Thiếu PBCCDC'
                matches.append({'src': 'CCDC', 'ten': _clean_dg_display(cc['ten']), 'cp_amt': cc['phan_bo'],
                    'pbcp_amt': 0, 'diff': cc['phan_bo'],
                    'note': note, 'so_ct_ledger': '', 'so_ct_file': cc['ma'], 'tk': tk})

        # ── Khác: ledger entries not matched to CP or CCDC ──
        other_entries = [e for e in ledger
                         if 'PBCP' not in e['so_ct'].upper()
                         and 'PBCCDC' not in e['so_ct'].upper()]
        internal_sum = 0
        internal_count = 0
        for oe in other_entries:
            is_internal = oe.get('note', '') == 'Nội bộ 242'
            if is_internal:
                internal_sum += oe['contrib']
                internal_count += 1
            else:
                khac_note = oe.get('note', '')
                if not khac_note:
                    other_tk = oe['tk_co'] if oe['tk_co'] else oe['tk_no']
                    khac_note = f'TK {other_tk}'
                matches.append({'src': 'Khác', 'ten': _clean_dg_display(oe['dg']), 'cp_amt': 0,
                    'pbcp_amt': 0, 'diff': oe['contrib'],
                    'note': khac_note, 'so_ct_ledger': oe['so_ct'], 'so_ct_file': '', 'tk': tk})
        if internal_count > 0:
            matches.append({'src': 'Nội bộ', 'ten': f'Chuyển giữa 2421 ↔ 2422 ({internal_count} phiếu)',
                'cp_amt': 0, 'pbcp_amt': 0, 'diff': internal_sum,
                'note': 'Nội bộ 242 (không cần sửa)', 'so_ct_ledger': '', 'so_ct_file': '',
                '_netoff': 'Nội bộ 242', 'tk': tk})

        all_matches[tk] = matches

    # ── Phase 2: Netting ──
    # Pass A: cross-TK netting by so_ct_file
    by_code = {}
    for tk in ('2421', '2422'):
        for m in all_matches[tk]:
            c = m.get('so_ct_file', '')
            if c:
                by_code.setdefault(c, []).append(m)
    for code, items in by_code.items():
        if len(items) >= 2 and abs(sum(m['diff'] for m in items)) < 1000:
            for i, m in enumerate(items):
                m['_netoff'] = "Net off"
                m['_netoff_ct'] = [(items[j]['tk'], items[j]['diff'], items[j]['ten'])
                                   for j in range(len(items)) if j != i]

    # Pass B: same-TK exact-offset netting by ten (no diff adjustment)
    for tk in ('2421', '2422'):
        mlist = all_matches[tk]
        active = [m for m in mlist if not m.get('_netoff')]
        by_ten = {}
        for i, m in enumerate(active):
            by_ten.setdefault(m['ten'], []).append(i)
        for ten, idxs in by_ten.items():
            pos = sorted([i for i in idxs if active[i]['diff'] > 0], key=lambda i: active[i]['diff'])
            neg = sorted([i for i in idxs if active[i]['diff'] < 0], key=lambda i: -active[i]['diff'])
            i = j = 0
            while i < len(pos) and j < len(neg):
                pd = active[pos[i]]['diff']
                nd = active[neg[j]]['diff']
                if abs(pd + nd) < 1000:
                    active[pos[i]]['_netoff'] = "Net off cùng TK"
                    active[neg[j]]['_netoff'] = "Net off cùng TK"
                    i += 1; j += 1
                elif pd > -nd:
                    j += 1
                else:
                    i += 1

    # Pass C: cross-TK Nội bộ netting
    noibo_items = []
    for tk in ('2421', '2422'):
        for m in all_matches[tk]:
            if m['src'] == 'Nội bộ':
                noibo_items.append(m)
    if len(noibo_items) >= 2 and abs(sum(m['diff'] for m in noibo_items)) < 1000:
        for i, m in enumerate(noibo_items):
            m['_netoff'] = "Nội bộ 242"
            m['_netoff_ct'] = [(noibo_items[j]['tk'], noibo_items[j]['diff'], noibo_items[j]['ten'])
                               for j in range(len(noibo_items)) if j != i]

    # ── Phase 3: Print combined section for TK 242 ──
    tk1, tk2 = '2421', '2422'
    matches = all_matches[tk1] + all_matches[tk2]

    # ── Console summary ──
    for tk in (tk1, tk2):
        cl = close_data[tk]['lech']
        sum_all = sum(m['diff'] for m in all_matches[tk] if m['src'] != 'ĐK')
        print(f'  {tk}: lệch={cl:>15,} = BCĐPS.cl={close_data[tk]["bcdps"]:>15,} - CCDC.cl={close_data[tk]["ccdc"]:>15,} - CP.cl={close_data[tk]["cp"]:>15,}')
        print(f'  {tk}: sum_all_non_dk={sum_all:>15,}')
    cl_total = close_data[tk1]['lech'] + close_data[tk2]['lech']
    bcdps_cl = close_data[tk1]['bcdps'] + close_data[tk2]['bcdps']
    ccdc_cl = close_data[tk1]['ccdc'] + close_data[tk2]['ccdc']
    cp_cl = close_data[tk1]['cp'] + close_data[tk2]['cp']
    print(f'  242 : lệch={cl_total:>15,} = BCĐPS.cl={bcdps_cl:>15,} - CCDC.cl={ccdc_cl:>15,} - CP.cl={cp_cl:>15,}')

    # ── Build combined formula ──
    rk1 = bcdps_rows.get(tk1, 0)
    rk2 = bcdps_rows.get(tk2, 0)
    f1 = (f'=BCĐPS!J{rk1}-BCĐPS!K{rk1}'
        f'-SUMIF(\'CP chờ PB\'!K:K,"{tk1}",\'CP chờ PB\'!I:I)'
        f'-SUMIF(CCDC!R:R,"{tk1}",CCDC!N:N)'
        f'-SUMIF(CCDC!S:S,"{tk1}",CCDC!N:N)'
        f'+SUMIFS(CCDC!N:N,CCDC!R:R,"{tk1}",CCDC!S:S,"{tk1}")')
    f2 = (f'BCĐPS!J{rk2}-BCĐPS!K{rk2}'
        f'-SUMIF(\'CP chờ PB\'!K:K,"{tk2}",\'CP chờ PB\'!I:I)'
        f'-SUMIF(CCDC!R:R,"{tk2}",CCDC!N:N)'
        f'-SUMIF(CCDC!S:S,"{tk2}",CCDC!N:N)'
        f'+SUMIFS(CCDC!N:N,CCDC!R:R,"{tk2}",CCDC!S:S,"{tk2}")')
    le_formula = '=' + f1[1:] + '+' + f2

    # ── Section header ──
    row += 1
    _sec(row, "PHÂN TÍCH TK 242", MC - 1, color="E2EFDA")
    c12 = ws.cell(row, 12, le_formula)
    c12.number_format = NUM
    c12.font = Font(name="Arial", bold=True, size=10)
    c12.alignment = Alignment(horizontal="right", vertical="center")
    row += 1
    _hdr(row, ["STT", "Nguồn", "Diễn giải", "Số CT CP", "Số CT PBCP",
               "CP Amount", "PBCP Amount", "Lệch", "Cần sửa", "Phân loại"],
         widths=[5, 6, 55, 14, 14, 16, 16, 16, 18, 14])
    row += 1
    stt = 0

    # ── 1. TRONG KỲ - CẦN SỬA ──
    need_fix = sorted([m for m in matches if m['src'] != 'ĐK'
                      and m['src'] != 'Khác'
                      and not m.get('_netoff') and abs(m['diff']) >= 1000],
                     key=lambda x: x.get('src', ''))
    if need_fix:
        _sec(row, "TRONG KỲ - CẦN SỬA", MC - 1, color="FCE4D6")
        row += 1
        cs_start = row
        for m in need_fix:
            stt += 1; total_displayed += 1
            ws.row_dimensions[row].height = 28
            _rw(row, [stt, m['src'], str(m['ten'])[:80],
                       m.get('so_ct_file', ''), m.get('so_ct_ledger', ''),
                       m['cp_amt'], m['pbcp_amt'], m['diff'],
                       m.get('note', ''), ''],
                fmt_cols={6, 7, 8}, fill="FCE4D6")
            row += 1
        cs_end = row - 1
        _rw(row, ['', '', "Tổng cần sửa (242)", '', '',
                   0, 0, f"=SUM(H{cs_start}:H{cs_end})", '', ''],
            bold=True, fmt_cols={6, 7, 8})
        row += 1

    # ── 2. TRONG KỲ - NET OFF ──
    netoff_items = sorted([m for m in matches if m.get('_netoff')],
                          key=lambda x: x.get('src', ''))
    if netoff_items:
        _sec(row, "TRONG KỲ - NET OFF", MC - 1, color="FFF2CC")
        row += 1
        no_start = row
        for m in netoff_items:
            stt += 1; total_displayed += 1
            ws.row_dimensions[row].height = 28
            _rw(row, [stt, m['src'], str(m['ten'])[:80],
                       m.get('so_ct_file', ''), m.get('so_ct_ledger', ''),
                       m['cp_amt'], m['pbcp_amt'], m['diff'],
                       '', m.get('_netoff', '')],
                fmt_cols={6, 7, 8})
            no_end = row
            row += 1
            if m.get('_netoff_ct'):
                for ct_tk, ct_diff, _ in m['_netoff_ct']:
                    ws.row_dimensions[row].height = 18
                    _rw(row, ['', '', f"  → Đối ứng TK {ct_tk}: {ct_diff:+,}",
                               '', '', 0, 0, 0, '', ''],
                        fmt_cols={6, 7, 8})
                    row += 1
        _rw(row, ['', '', "Tổng net off (242)", '', '',
                   0, 0, f"=SUM(H{no_start}:H{no_end})", '', ''],
            bold=True, fmt_cols={6, 7, 8})
        row += 1

    # ── 3. CUỐI KỲ (BCĐPS raw formula) ──
    _sec(row, "CUỐI KỲ", MC - 1, color="E2EFDA")
    row += 1
    stt += 1
    ws.row_dimensions[row].height = 22
    _rw(row, [stt, "242", "Tổng cuối kỳ", "", "", 0, 0, le_formula, "", "Cuối kỳ"],
        fmt_cols={6, 7, 8})
    row += 1

    ws.sheet_view.showGridLines = False

    # Close source workbooks
    for src_wb, _ in src_sheets:
        src_wb.close()
    del src_sheets

    # Save & merge source sheets via win32com
    temp_output = os.path.join(TMP, "_analysis.xlsx")
    wb_out.save(temp_output)
    wb_out.close()
    conn.close()

    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.Visible = False
    out = os.path.join(BASE_DIR, f"z_DoiChieu_CCDC_{ENTITY}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    try:
        wb_merged = excel.Workbooks.Open(temp_output, 0, True)  # UpdateLinks=0, ReadOnly=True
        ws_analysis = wb_merged.Worksheets("Phiếu cần điều chỉnh")
        src_list = [(f_bcdps_cp, "BCĐPS"), (f_ccdc_cp, "CCDC"), (f_cp_cp, "CP chờ PB"),
                    (f_2421_cp, "Sổ TK 2421"), (f_2422_cp, "Sổ TK 2422")]
        for src_path, sheet_name in src_list:
            if not os.path.exists(src_path): continue
            wb_src = excel.Workbooks.Open(src_path, 0, True)  # ReadOnly
            if wb_src is None: continue
            wb_src.Worksheets(1).Copy(Before=ws_analysis)
            wb_src.Close()
        names = ["BCĐPS", "CCDC", "CP chờ PB", "Sổ TK 2421", "Sổ TK 2422"]
        for i in range(wb_merged.Worksheets.Count - 1):
            wb_merged.Worksheets(i + 1).Name = names[i]
        # Auto-fit columns for source sheets
        for i in range(1, wb_merged.Worksheets.Count):
            ws_src = wb_merged.Worksheets(i)
            if ws_src.Name != "Phiếu cần điều chỉnh":
                ws_src.UsedRange.EntireColumn.AutoFit()
        # Set column widths for analysis sheet via win32com
        for ci, w in {1:5, 2:6, 3:55, 4:14, 5:14, 6:16, 7:16, 8:16, 9:18, 10:14, 12:12}.items():
            ws_analysis.Columns(ci).ColumnWidth = w
        merged_temp = os.path.join(TMP, "_merged.xlsx")
        wb_merged.SaveAs(merged_temp)
        wb_merged.Close()
    finally:
        try: excel.Quit()
        except: pass
        try: pythoncom.CoUninitialize()
        except: pass

    # Strip [N] workbook prefixes from formulas using openpyxl
    import re
    wb_fix = openpyxl.load_workbook(merged_temp)
    ws_fix = wb_fix["Phiếu cần điều chỉnh"]
    changes = 0
    for row in ws_fix.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                new_val = re.sub(r'\[\d+\]', '', cell.value)
                if new_val != cell.value:
                    cell.value = new_val
                    changes += 1
    wb_fix.save(out)
    wb_fix.close()
    print(f"  Fixed {changes} formulas (stripped [N] prefixes)")
    try: shutil.rmtree(TMP, ignore_errors=True)
    except: pass

    print(f"  Xuat xong: {out}")
    print("  Sheet: Phieu can dieu chinh")
    print("\n  === TOM TAT ===")
    for tk_k, desc in [("242", "Tổng"), ("2421", "Ngắn hạn"), ("2422", "Dài hạn")]:
        if tk_k == "242":
            bv = _fmt(b2421[2] + b2422[2])
            cv = _fmt(sum(_f(r[1]) for r in ccdc_sum))
            pv = _fmt(sum(cp_map.get(sk, {}).get('cl', 0) for sk in ('2421', '2422')))
        else:
            bv = _fmt(b2421[2] if tk_k == "2421" else b2422[2])
            cv = _fmt(ccdc_map.get(tk_k, {}).get('cl_ck', 0))
            pv = _fmt(cp_map.get(tk_k, {}).get('cl', 0))
        df = bv - (cv + pv)
        print(f"  {tk_k:5s} | {desc:8s} | BCDPS={bv:>15,.0f} | CCDC={cv:>15,.0f} | CP={pv:>15,.0f} | Lech={df:>12,.0f} | {'KHOP' if abs(df)<1000 else 'LECH'}")
    # Unassigned CCDC
    if ccdc_unassigned:
        u_cl_ck = sum(u['cl_ck'] for u in ccdc_unassigned)
        print(f"  (u)    | Unassign | BCDPS={0:>15,.0f} | CCDC={u_cl_ck:>15,.0f} | CP={0:>15,.0f} | Lech={-u_cl_ck:>12,.0f} | CHECK")
        for u in ccdc_unassigned:
            print(f"         |   {str(u['ten'])[:40]:40s} | {0:>15,.0f} | {u['cl_ck']:>15,.0f} | {0:>15,.0f} | {u['note']}")
    for tk in ('2421', '2422'):
        od = open_data[tk]['lech']
        cd = close_data[tk]['lech']
        abk = agg_bk_ccdc[tk]
        acp = cp_mv_total[tk]
        agg = _fmt(abk + acp)
        ok = "OK" if abs(od + agg - cd) < 1000 else "ERR"
        print(f"  {tk}: dau={od:>12,.0f} + BK-CCDC={abk:>12,.0f} + CP={acp:>12,.0f} = cuoi={cd:>12,.0f} {ok}")
    print(f"  {total_displayed} dong chi tiet")
    return out

if __name__ == "__main__":
    main()
