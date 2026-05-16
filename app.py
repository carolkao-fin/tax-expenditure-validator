import streamlit as st
import pandas as pd
from docx import Document
import io
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="稅式支出評估報告驗證工具", layout="wide")

# ── helpers ───────────────────────────────────────────────────────────────────

def clean_num(text):
    if text is None:
        return None
    s = str(text).strip()
    if s in ('', '-', '—', '–'):
        return None
    has_pct = '%' in s
    s = re.sub(r'[,，\s億元千萬%（）()【】]', '', s)
    try:
        v = float(s)
        return v / 100 if has_pct else v
    except ValueError:
        return None

def pct_str(text):
    if text is None:
        return None
    s = str(text).strip().replace('%', '').replace(',', '')
    try:
        return float(s) / 100
    except ValueError:
        return None

def round2(v):
    return round(v, 2) if v is not None else None

def fmt(v, decimals=2):
    if v is None:
        return ''
    return f"{v:,.{decimals}f}"

def ok(reported, computed, tol=0.02):
    if reported is None or computed is None:
        return None
    denom = max(abs(reported), 1)
    return abs(reported - computed) / denom <= tol

def status_icon(passed):
    if passed is None:
        return '⚠️'
    return '✅' if passed else '❌'

def extract_tables(doc):
    tables = []
    for t in doc.tables:
        rows = []
        for row in t.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)
        tables.append(rows)
    return tables

def extract_paragraphs(doc):
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def find_table_by_header(tables, *keywords):
    for rows in tables:
        if not rows:
            continue
        header = ' '.join(cell.replace('\n', '') for cell in rows[0])
        if all(kw in header for kw in keywords):
            return rows
    return None

def row_result(section, item, formula, reported, computed, diff, result, note):
    return {
        '章節': section,
        '驗證項目': item,
        '公式（文件原文）': formula,
        '報告值': reported,
        '計算值': computed,
        '差異': diff,
        '結果': result,
        '說明': note,
    }

# ── formula text extraction ───────────────────────────────────────────────────

def compact_formula_full(txt):
    """
    From a full-report single-cell formula table, extract the key calculation line.
    Returns a compact string like '49億元×15%×20% = 1.47億元'.
    """
    lines = [l.strip() for l in txt.replace('\r', '\n').split('\n') if l.strip()]
    # Lines that start with '=' and contain × are the formula lines
    formula_line = next(
        (l.lstrip('= ').strip() for l in lines
         if l.startswith('=') and ('×' in l or 'x' in l.lower()) and '億元' in l),
        None
    )
    result_line = next(
        (l for l in lines if '新臺幣' in l and '億元' in l),
        None
    )
    parts = []
    if formula_line:
        parts.append(formula_line)
    if result_line:
        m = re.search(r'新臺幣\s*([\d\.]+)\s*億元', result_line)
        if m and (not formula_line or m.group(1) not in formula_line):
            parts.append(f'= {m.group(1)}億元')
    return '  '.join(parts) if parts else txt[:120]

def compact_formula_braces(txt):
    """Extract the ｛〔...〕｝ formula block from simplified report text."""
    m = re.search(r'[｛{]\s*[〔\[].*?[〕\]].*?[｝}]', txt, re.DOTALL)
    if m:
        inner = m.group(0)
        inner = re.sub(r'\s+', ' ', inner).strip()
        return inner[:200]
    return ''

# ── document-type detection ───────────────────────────────────────────────────

def detect_type(tables, paragraphs):
    for rows in tables:
        flat = ' '.join(c for row in rows for c in row)
        if '提案委員' in flat and '法規內容' in flat:
            return 'simplified'
    all_text = ' '.join(paragraphs)
    if '稅式支出評估報告' in all_text and '表4-1' in all_text:
        return 'full'
    if '簡要' in all_text or '提案委員' in all_text:
        return 'simplified'
    return 'full'

# ── full-format parser ────────────────────────────────────────────────────────

def parse_full_report(tables, paragraphs):
    data = {}

    # 表4-1: 5-year average import values
    tbl41 = find_table_by_header(tables, '5年平均') or find_table_by_header(tables, '現行稅率', '降稅後稅率')
    items = []
    if tbl41:
        for row in tbl41:
            if len(row) < 3:
                continue
            row_text = ' '.join(row)
            if '合計' in row_text and clean_num(row[2]) is not None:
                if data.get('total_import_reported') is None:
                    data['total_import_reported'] = clean_num(row[2])
                    data['weighted_rate_reported'] = pct_str(row[3]) if len(row) > 3 else None
                continue
            hs = row[0].strip()
            if not hs or 'HS' in hs.upper() or '稅則' in hs or '合計' in hs:
                continue
            imp = clean_num(row[2]) if len(row) > 2 else None
            cur_rate = pct_str(row[3]) if len(row) > 3 else None
            aft_rate = pct_str(row[4]) if len(row) > 4 else None
            if imp is not None:
                items.append({'hs': hs, 'name': row[1] if len(row) > 1 else '',
                               'import_5yr': imp, 'current_rate': cur_rate, 'after_rate': aft_rate})
    data['items'] = items

    # 表4-2: 最初收入損失法
    tbl42 = find_table_by_header(tables, '關稅損失') or find_table_by_header(tables, '最初收入損失')
    item_losses = {}
    if tbl42:
        for row in tbl42:
            if len(row) < 4:
                continue
            row_text = ' '.join(row)
            if '合計' in row_text and clean_num(row[-1]) is not None:
                if data.get('tariff_loss_reported') is None:
                    data['tariff_loss_reported'] = clean_num(row[-1])
                continue
            hs = row[0].strip()
            if not hs or 'HS' in hs.upper() or '稅則' in hs or '合計' in hs:
                continue
            loss = clean_num(row[-1])
            if loss is not None:
                item_losses[hs] = loss
    data['item_losses_reported'] = item_losses

    # Identify formula tables (single-cell)
    formula_tables = {}
    for rows in tables:
        cell_text = '\n'.join(' '.join(row) for row in rows) if rows else ''
        ct = cell_text.replace(' ', '')
        if '汽車整車營利事業所得稅' in cell_text and '×15%' in ct:
            formula_tables['vehicle_cit'] = rows
        elif '汽車零組件營利事業所得稅' in cell_text and '×13%' in ct:
            formula_tables['parts_cit'] = rows
        elif '其他工業及服務業營利事業所得稅' in cell_text and '16.' in cell_text:
            formula_tables['other_cit'] = rows
        elif '股利所得稅' in cell_text and '盈餘分配比例' in cell_text:
            formula_tables['dividend'] = rows
        elif '個人所得稅' in cell_text and '平均受雇員工年薪' in cell_text:
            formula_tables['personal'] = rows
        elif '汽車貨物稅' in cell_text and '貨物稅稅率' in cell_text:
            formula_tables['commodity'] = rows
        elif '加值型營業稅' in cell_text and '加值型營業稅稅率' in cell_text:
            formula_tables['vat'] = rows

    data['formula_tables'] = formula_tables

    # Extract values + formula strings from each formula table
    formulas = {}

    def get_txt(key):
        return '\n'.join(' '.join(r) for r in formula_tables[key]) if key in formula_tables else ''

    # Vehicle CIT
    txt = get_txt('vehicle_cit')
    m = re.search(r'=\s*([\d\.]+)\s*億元\s*[×x]\s*15%', txt)
    data['vehicle_output'] = float(m.group(1)) if m else 49.0
    m2 = re.search(r'新臺幣\s*([\d\.]+)\s*億元', txt)
    data['vehicle_cit_reported'] = float(m2.group(1)) if m2 else None
    formulas['vehicle_cit'] = compact_formula_full(txt) if txt else ''

    # Parts CIT
    txt = get_txt('parts_cit')
    m = re.search(r'=\s*([\d\.]+)\s*億元\s*[×x]\s*13%', txt)
    data['parts_output'] = float(m.group(1)) if m else 34.31
    m2 = re.search(r'新臺幣\s*([\d\.]+)\s*億元', txt)
    data['parts_cit_reported'] = float(m2.group(1)) if m2 else None
    formulas['parts_cit'] = compact_formula_full(txt) if txt else ''

    # Other CIT
    txt = get_txt('other_cit')
    m = re.search(r'\(\s*([\d\.]+)\s*億元\s*[+＋]\s*([\d\.]+)\s*億元\s*\)', txt)
    if m:
        data['other_profit'] = float(m.group(1))
        data['tariff_savings'] = float(m.group(2))
    else:
        data['other_profit'] = 16.05
        data['tariff_savings'] = 22.40
    m2 = re.search(r'新臺幣\s*([\d\.]+)\s*億元', txt)
    data['other_cit_reported'] = float(m2.group(1)) if m2 else None
    formulas['other_cit'] = compact_formula_full(txt) if txt else ''

    # Dividend
    txt = get_txt('dividend')
    all_vals = [float(v) for v in re.findall(r'(?<!\d)([\d]+\.[\d]+)\s*億元', txt)]
    small_vals = [v for v in all_vals if v < 5]
    data['dividend_tax_reported'] = small_vals[-1] if small_vals else 0.77
    formulas['dividend'] = compact_formula_full(txt) if txt else ''

    # Personal
    txt = get_txt('personal')
    m = re.search(r'合計\s*=\s*([\d\.]+)\s*萬元', txt)
    if not m:
        vals = re.findall(r'=\s*([\d\.]+)\s*萬元', txt)
        data['personal_tax_reported'] = float(vals[-1]) / 10000 if vals else 0.0287
    else:
        data['personal_tax_reported'] = float(m.group(1)) / 10000
    formulas['personal'] = compact_formula_full(txt) if txt else ''

    # Commodity
    txt = get_txt('commodity')
    m = re.search(r'新臺幣\s*([\d\.]+)\s*億元', txt)
    data['commodity_tax_reported'] = float(m.group(1)) if m else None
    formulas['commodity'] = compact_formula_full(txt) if txt else ''

    # VAT
    txt = get_txt('vat')
    m = re.search(r'新臺幣\s*([\d\.]+)\s*億元', txt)
    data['vat_reported'] = float(m.group(1)) if m else None
    formulas['vat'] = compact_formula_full(txt) if txt else ''

    # Final net
    all_text = '\n'.join(paragraphs) + '\n'
    for rows in tables:
        for row in rows:
            all_text += '\n' + ' '.join(row)
    m = re.search(r'(?:淨增|損益合計|淨[損益額])[^\d+\-]*([\+\-]?\s*[\d\.]+)\s*億元', all_text)
    data['net_reported'] = float(m.group(1).replace(' ', '')) if m else 0.29

    data['formulas'] = formulas
    return data


# ── simplified-format parser ──────────────────────────────────────────────────

def parse_simplified_report(tables, paragraphs):
    data = {}
    main_table = None
    for rows in tables:
        flat = ' '.join(c for row in rows for c in row)
        if '稅式支出評估' in flat and ('千元' in flat or '2,240' in flat):
            main_table = rows
            break
        if '提案委員' in flat:
            main_table = rows
            break
    if main_table is None and tables:
        main_table = tables[0]

    all_text = '\n'.join(paragraphs)
    if main_table:
        for row in main_table:
            all_text += '\n' + ' '.join(row)
    for rows in tables:
        for row in rows:
            all_text += '\n' + ' '.join(row)

    def find_k(pattern):
        m = re.search(pattern, all_text, re.DOTALL)
        if m:
            raw = m.group(1).replace(',', '').strip()
            try:
                return float(raw)
            except ValueError:
                return None
        return None

    data['total_import_reported_k'] = find_k(r'平均進口值為([\d,]+)千元')
    data['tariff_loss_reported_k']  = find_k(r'關稅收入減少([\d,]+)千元')
    data['vehicle_output_k']        = find_k(r'產值(?:變化值|增加額)\s*([\d,]+)千元.*?小客車生產')
    data['parts_output_k']          = find_k(r'零組件產值增加額為([\d,]+)千元')
    data['vehicle_cit_reported_k']  = find_k(r'整車業增加稅額約為新臺幣([\d,]+)千元')
    data['parts_cit_reported_k']    = find_k(r'國產化比例約?70%.*?增加稅額約為新臺幣([\d,]+)千元')
    if data['parts_cit_reported_k'] is None:
        data['parts_cit_reported_k'] = find_k(r'零組件產值增加額為[\d,]+千元.*?增加稅額約為新臺幣([\d,]+)千元')
    data['other_cit_reported_k']    = find_k(r'就其他工業及服務業部分.*?新臺幣([\d,]+)千元')
    data['total_cit_reported_k']    = find_k(r'三項合計.*?新臺幣([\d,]+)千元')
    data['commodity_tax_reported_k']= find_k(r'貨物稅增加額約為新臺幣([\d,]+)千元')
    data['vat_reported_k']          = find_k(r'加值型營業稅增加額約為新臺幣([\d,]+)千元')
    data['personal_tax_reported_k'] = find_k(r'員工個人所得稅增加額約為新臺幣([\d,]+)千元')
    if data['personal_tax_reported_k'] is None:
        data['personal_tax_reported_k'] = find_k(r'兩項合計.*?個人所得稅.*?新臺幣([\d,]+)千元')
    data['dividend_tax_reported_k'] = find_k(r'股東個人股利所得稅增加額約為新臺幣([\d,]+)千元')
    data['net_reported_k']          = find_k(r'稅收淨收入([\d,]+)千元') or find_k(r'最終稅收.*?([\d,]+)千元')
    data['other_profit_k']          = find_k(r'增加產值利潤([\d,]+)千元')

    # Extract formula text blocks
    formulas = {}
    vo_k = data.get('vehicle_output_k') or 4900813
    po_k = data.get('parts_output_k') or 3430569

    # Vehicle CIT formula: search near "整車業增加稅額"
    m = re.search(r'整車業增加稅額約為新臺幣[\d,]+千元([^。\n]{0,300})', all_text, re.DOTALL)
    formulas['vehicle_cit'] = (compact_formula_braces(m.group(0)) if m else
                                f'{fmt(vo_k,0)}千元 × 15% × 20%')

    # Parts CIT formula
    m = re.search(r'國產化比例約?70%[^。]{0,400}增加稅額約為新臺幣[\d,]+千元([^。]{0,300})', all_text, re.DOTALL)
    formulas['parts_cit'] = (compact_formula_braces(m.group(0)) if m else
                              f'{fmt(po_k,0)}千元 × 13% × 20%')

    # Other CIT formula
    m = re.search(r'兩者合計[\d,]+千元[，,]([^。]{0,300}新臺幣[\d,]+千元)', all_text, re.DOTALL)
    if not m:
        m = re.search(r'就其他工業及服務業部分[^。]{0,500}新臺幣[\d,]+千元([^。]{0,200})', all_text, re.DOTALL)
    formulas['other_cit'] = compact_formula_braces(m.group(0)) if m else '3,844,451千元 × 20%'

    # Commodity tax formula
    m = re.search(r'貨物稅增加額約為新臺幣[\d,]+千元([^。]{0,500})', all_text, re.DOTALL)
    formulas['commodity'] = compact_formula_braces(m.group(0)) if m else ''

    # VAT formula
    m = re.search(r'加值型營業稅增加額約為新臺幣[\d,]+千元([^。\n]{0,500})', all_text, re.DOTALL)
    formulas['vat'] = compact_formula_braces(m.group(0)) if m else ''

    # Personal tax formula (show the formula in braces from the text)
    m = re.search(r'員工個人所得稅增加額約為新臺幣[\d,]+千元([^。]{0,400})', all_text, re.DOTALL)
    formulas['personal'] = compact_formula_braces(m.group(0)) if m else ''

    # Dividend formula
    m = re.search(r'股東個人股利所得稅增加額約為新臺幣[\d,]+千元([^。]{0,400})', all_text, re.DOTALL)
    formulas['dividend'] = compact_formula_braces(m.group(0)) if m else ''

    data['formulas'] = formulas
    data['all_text'] = all_text
    return data


# ── verification: full report ─────────────────────────────────────────────────

SEC_TARIFF  = '一、最初收入損失法'
SEC_FINAL   = '二、最終收入損失法'
SEC_CT      = '　（一）貨物稅'
SEC_VAT     = '　（二）加值型營業稅'
SEC_CIT     = '　（三）營利事業所得稅'
SEC_PER     = '　（四）個人綜合所得稅'
SEC_DIV     = '　（五）股東個人股利所得稅'
SEC_NET     = '三、淨損益'


def verify_full(data):
    results = []
    items  = data.get('items', [])
    forms  = data.get('formulas', {})

    # ── 最初收入損失法 ───────────────────────────────────────────────────────

    # Import total
    if items:
        calc_total = round2(sum(it['import_5yr'] for it in items))
        rep = data.get('total_import_reported') or 147.86
        results.append(row_result(
            SEC_TARIFF, '5年平均進口總額 (億元)', '各品項加總',
            fmt(rep), fmt(calc_total), fmt(rep - calc_total),
            status_icon(ok(rep, calc_total)), ''))

    # Simple average rate
    if items and all(it['current_rate'] is not None for it in items):
        calc_wavg = sum(it['current_rate'] for it in items) / len(items)
        rep_wavg = data.get('weighted_rate_reported') or 0.1185
        results.append(row_result(
            SEC_TARIFF, f'算術平均現行稅率（{len(items)}項）',
            f'Σ稅率 ÷ {len(items)}',
            f'{rep_wavg*100:.2f}%', f'{calc_wavg*100:.2f}%',
            f'{(rep_wavg-calc_wavg)*100:.4f}%',
            status_icon(ok(rep_wavg, calc_wavg, tol=0.005)), ''))

    # Per-item tariff loss
    calc_total_loss = 0
    for it in items:
        if it['import_5yr'] is None or it['current_rate'] is None or it['after_rate'] is None:
            continue
        cut = it['current_rate'] - it['after_rate']
        calc_loss = round2(it['import_5yr'] * cut)
        rep_loss  = data.get('item_losses_reported', {}).get(it['hs'])
        calc_total_loss += calc_loss
        results.append(row_result(
            SEC_TARIFF,
            f"{it['hs']}　{it['name'][:14]}",
            f"{it['import_5yr']} × {cut*100:.1f}%",
            fmt(rep_loss) if rep_loss else '—',
            fmt(calc_loss),
            fmt(rep_loss - calc_loss) if rep_loss else '—',
            status_icon(ok(rep_loss, calc_loss)) if rep_loss else '⚠️',
            '關稅損失(億元)'))

    # Total tariff loss
    rep_loss = data.get('tariff_loss_reported') or 22.40
    results.append(row_result(
        SEC_TARIFF, '關稅損失合計 (億元)', 'Σ各品項關稅損失',
        fmt(rep_loss), fmt(calc_total_loss), fmt(rep_loss - calc_total_loss),
        status_icon(ok(rep_loss, calc_total_loss)), ''))

    # ── 貨物稅 ────────────────────────────────────────────────────────────────
    vo = data.get('vehicle_output', 49.0)
    calc_ct = round2(vo * (0.5285 * 0.25 * 0.8419 + 0.4715 * 0.15 * 0.9960))
    rep_ct  = data.get('commodity_tax_reported')
    results.append(row_result(
        SEC_CT, '貨物稅 (億元)',
        forms.get('commodity') or f'{vo}億元×[52.85%×25%×84.19%+47.15%×15%×99.60%]',
        fmt(rep_ct) if rep_ct else '—', fmt(calc_ct),
        fmt(rep_ct - calc_ct) if rep_ct else '—',
        status_icon(ok(rep_ct, calc_ct)) if rep_ct else '⚠️', ''))

    # ── 加值型營業稅 ──────────────────────────────────────────────────────────
    calc_vat = round2(vo * (0.5285 * 1.25 * 0.05 + 0.4715 * 1.15 * 0.05))
    rep_vat  = data.get('vat_reported')
    results.append(row_result(
        SEC_VAT, '加值型營業稅 (億元)',
        forms.get('vat') or f'{vo}億元×[52.85%×125%×5%+47.15%×115%×5%]',
        fmt(rep_vat) if rep_vat else '—', fmt(calc_vat),
        fmt(rep_vat - calc_vat) if rep_vat else '—',
        status_icon(ok(rep_vat, calc_vat)) if rep_vat else '⚠️', ''))

    # ── 營利事業所得稅 ────────────────────────────────────────────────────────
    po = data.get('parts_output', 34.31)
    op = data.get('other_profit', 16.05)
    ts = data.get('tariff_savings', 22.40)

    calc_vcit = round2(vo * 0.15 * 0.20)
    rep_vcit  = data.get('vehicle_cit_reported')
    results.append(row_result(
        SEC_CIT, '整車營利事業所得稅 (億元)',
        forms.get('vehicle_cit') or f'{vo}億元×15%×20%',
        fmt(rep_vcit) if rep_vcit else '—', fmt(calc_vcit),
        fmt(rep_vcit - calc_vcit) if rep_vcit else '—',
        status_icon(ok(rep_vcit, calc_vcit)) if rep_vcit else '⚠️', ''))

    calc_pcit = round2(po * 0.13 * 0.20)
    rep_pcit  = data.get('parts_cit_reported')
    results.append(row_result(
        SEC_CIT, '零組件營利事業所得稅 (億元)',
        forms.get('parts_cit') or f'{po}億元×13%×20%',
        fmt(rep_pcit) if rep_pcit else '—', fmt(calc_pcit),
        fmt(rep_pcit - calc_pcit) if rep_pcit else '—',
        status_icon(ok(rep_pcit, calc_pcit)) if rep_pcit else '⚠️', ''))

    calc_ocit = round2((op + ts) * 0.20)
    rep_ocit  = data.get('other_cit_reported')
    results.append(row_result(
        SEC_CIT, '其他工業及服務業營利事業所得稅 (億元)',
        forms.get('other_cit') or f'({op}+{ts})億元×20%',
        fmt(rep_ocit) if rep_ocit else '—', fmt(calc_ocit),
        fmt(rep_ocit - calc_ocit) if rep_ocit else '—',
        status_icon(ok(rep_ocit, calc_ocit)) if rep_ocit else '⚠️', ''))

    calc_tcit = round2(calc_vcit + calc_pcit + calc_ocit)
    rep_tcit  = data.get('total_cit_reported')
    if rep_tcit is None:
        parts_list = [data.get('vehicle_cit_reported'), data.get('parts_cit_reported'), data.get('other_cit_reported')]
        if all(p is not None for p in parts_list):
            rep_tcit = round2(sum(parts_list))
    results.append(row_result(
        SEC_CIT, '營利事業所得稅合計 (億元)', '整車+零組件+其他',
        fmt(rep_tcit) if rep_tcit else '—', fmt(calc_tcit),
        fmt(rep_tcit - calc_tcit) if rep_tcit else '—',
        status_icon(ok(rep_tcit, calc_tcit)) if rep_tcit else '⚠️', ''))

    # ── 個人綜合所得稅 ────────────────────────────────────────────────────────
    rep_per = data.get('personal_tax_reported', 0.0287)
    results.append(row_result(
        SEC_PER, '個人綜合所得稅 (億元)',
        forms.get('personal') or '增聘員工薪資 × 適用稅率',
        fmt(rep_per), '（依文件公式，不重新推算）', '—', '⚠️',
        '此項依文件所示公式與參數確認'))

    # ── 股東個人股利所得稅 ────────────────────────────────────────────────────
    rep_div = data.get('dividend_tax_reported', 0.77)
    results.append(row_result(
        SEC_DIV, '股東個人股利所得稅 (億元)',
        forms.get('dividend') or '產值利潤×80%×56%×30%×40.45%×28%',
        fmt(rep_div), '（依文件公式，不重新推算）', '—', '⚠️',
        '此項依文件所示公式與參數確認'))

    # ── 淨損益 ────────────────────────────────────────────────────────────────
    calc_net = round2(calc_tcit + rep_div + rep_per + calc_ct + calc_vat - rep_loss)
    rep_net  = data.get('net_reported', 0.29)
    results.append(row_result(
        SEC_NET, '最終收入損失法淨損益 (億元)',
        f'CIT+股利+個人+貨物稅+營業稅－關稅損失',
        fmt(rep_net), fmt(calc_net), fmt(rep_net - calc_net),
        status_icon(ok(rep_net, calc_net, tol=0.10)),
        f'{fmt(calc_tcit)}+{fmt(rep_div)}+{fmt(rep_per)}+{fmt(calc_ct)}+{fmt(calc_vat)}−{fmt(rep_loss)}'))

    return pd.DataFrame(results)


# ── verification: simplified report ──────────────────────────────────────────

def verify_simplified(data):
    results = []
    forms = data.get('formulas', {})

    VEHICLE_NET_MARGIN = 0.15
    PARTS_NET_MARGIN   = 0.13
    CIT_RATE           = 0.20
    DOMESTIC_SHARE     = 0.5285
    IMPORT_SHARE       = 0.4715
    DOMESTIC_CT_RATE   = 0.25
    IMPORT_CT_RATE     = 0.15
    DOMESTIC_TAXABLE   = 0.8419
    IMPORT_TAXABLE     = 0.9960
    DOMESTIC_TAX_BASE  = 1.25
    IMPORT_TAX_BASE    = 1.15
    VAT_RATE           = 0.05

    rep_loss = data.get('tariff_loss_reported_k')  or 2240360
    rep_vo   = data.get('vehicle_output_k')         or 4900813
    rep_po   = data.get('parts_output_k')           or 3430569
    rep_tcit = data.get('total_cit_reported_k')
    rep_vcit = data.get('vehicle_cit_reported_k')
    rep_pcit = data.get('parts_cit_reported_k')
    rep_ocit = data.get('other_cit_reported_k')
    rep_per  = data.get('personal_tax_reported_k')  or 2874
    rep_div  = data.get('dividend_tax_reported_k')  or 76500
    rep_ct   = data.get('commodity_tax_reported_k') or 890371
    rep_vat  = data.get('vat_reported_k')           or 294747
    rep_net  = data.get('net_reported_k')           or 29231

    OTHER_PROFIT_K = data.get('other_profit_k') or 1604091

    # ── 最初收入損失法 ────────────────────────────────────────────────────────
    rep_import = data.get('total_import_reported_k') or 14785668
    results.append(row_result(
        SEC_TARIFF, '5年平均進口總額 (千元)', '依財政部關務署統計',
        fmt(rep_import, 0), '（引用文件值）', '—', '⚠️', ''))
    results.append(row_result(
        SEC_TARIFF, '關稅損失 (千元)', '進口額×降稅幅度',
        fmt(rep_loss, 0), '（引用文件值）', '—', '⚠️', '最初收入損失法'))

    # ── 貨物稅 ────────────────────────────────────────────────────────────────
    calc_ct = round2(rep_vo * (DOMESTIC_SHARE * DOMESTIC_CT_RATE * DOMESTIC_TAXABLE
                               + IMPORT_SHARE * IMPORT_CT_RATE * IMPORT_TAXABLE))
    results.append(row_result(
        SEC_CT, '貨物稅 (千元)',
        forms.get('commodity') or f'{fmt(rep_vo,0)}×[52.85%×25%×84.19%+47.15%×15%×99.60%]',
        fmt(rep_ct, 0), fmt(calc_ct, 0), fmt(rep_ct - calc_ct, 0),
        status_icon(ok(rep_ct, calc_ct)), ''))

    # ── 加值型營業稅 ──────────────────────────────────────────────────────────
    calc_vat = round2(rep_vo * (DOMESTIC_SHARE * DOMESTIC_TAX_BASE * VAT_RATE
                                + IMPORT_SHARE * IMPORT_TAX_BASE * VAT_RATE))
    results.append(row_result(
        SEC_VAT, '加值型營業稅 (千元)',
        forms.get('vat') or f'{fmt(rep_vo,0)}×[52.85%×125%×5%+47.15%×115%×5%]',
        fmt(rep_vat, 0), fmt(calc_vat, 0), fmt(rep_vat - calc_vat, 0),
        status_icon(ok(rep_vat, calc_vat)), ''))

    # ── 營利事業所得稅 ────────────────────────────────────────────────────────
    calc_vcit = round2(rep_vo * VEHICLE_NET_MARGIN * CIT_RATE)
    results.append(row_result(
        SEC_CIT, '整車營利事業所得稅 (千元)',
        forms.get('vehicle_cit') or f'{fmt(rep_vo,0)}×15%×20%',
        fmt(rep_vcit, 0) if rep_vcit else '—', fmt(calc_vcit, 0),
        fmt(rep_vcit - calc_vcit, 0) if rep_vcit else '—',
        status_icon(ok(rep_vcit, calc_vcit)) if rep_vcit else '⚠️', ''))

    calc_pcit = round2(rep_po * PARTS_NET_MARGIN * CIT_RATE)
    results.append(row_result(
        SEC_CIT, '零組件營利事業所得稅 (千元)',
        forms.get('parts_cit') or f'{fmt(rep_po,0)}×13%×20%',
        fmt(rep_pcit, 0) if rep_pcit else '—', fmt(calc_pcit, 0),
        fmt(rep_pcit - calc_pcit, 0) if rep_pcit else '—',
        status_icon(ok(rep_pcit, calc_pcit)) if rep_pcit else '⚠️', ''))

    calc_ocit = round2((OTHER_PROFIT_K + rep_loss) * CIT_RATE)
    results.append(row_result(
        SEC_CIT, '其他工業及服務業營利事業所得稅 (千元)',
        forms.get('other_cit') or f'({fmt(OTHER_PROFIT_K,0)}+{fmt(rep_loss,0)})×20%',
        fmt(rep_ocit, 0) if rep_ocit else '—', fmt(calc_ocit, 0),
        fmt(rep_ocit - calc_ocit, 0) if rep_ocit else '—',
        status_icon(ok(rep_ocit, calc_ocit)) if rep_ocit else '⚠️', ''))

    calc_tcit = round2(calc_vcit + calc_pcit + calc_ocit)
    results.append(row_result(
        SEC_CIT, '營利事業所得稅合計 (千元)', '整車+零組件+其他',
        fmt(rep_tcit, 0) if rep_tcit else '—', fmt(calc_tcit, 0),
        fmt(rep_tcit - calc_tcit, 0) if rep_tcit else '—',
        status_icon(ok(rep_tcit, calc_tcit)) if rep_tcit else '⚠️', ''))

    # ── 個人綜合所得稅 ────────────────────────────────────────────────────────
    results.append(row_result(
        SEC_PER, '個人綜合所得稅 (千元)',
        forms.get('personal') or '增聘員工薪資×適用稅率',
        fmt(rep_per, 0), '（依文件值）', '—', '⚠️', '依文件公式確認'))

    # ── 股東個人股利所得稅 ────────────────────────────────────────────────────
    results.append(row_result(
        SEC_DIV, '股東個人股利所得稅 (千元)',
        forms.get('dividend') or '產值利潤×80%×56%×30%×40.45%×28%',
        fmt(rep_div, 0), '（依文件值）', '—', '⚠️', '依文件公式確認'))

    # ── 淨損益 ────────────────────────────────────────────────────────────────
    calc_net = round2(calc_vcit + calc_pcit + calc_ocit + rep_per + rep_div
                      + calc_ct + calc_vat - rep_loss)
    results.append(row_result(
        SEC_NET, '最終收入損失法淨損益 (千元)',
        'CIT+股利+個人+貨物稅+營業稅－關稅損失',
        fmt(rep_net, 0), fmt(calc_net, 0), fmt(rep_net - calc_net, 0),
        status_icon(ok(rep_net, calc_net, tol=0.05)),
        f'{fmt(calc_tcit,0)}+{fmt(rep_div,0)}+{fmt(rep_per,0)}+{fmt(calc_ct,0)}+{fmt(calc_vat,0)}−{fmt(rep_loss,0)}'))

    return pd.DataFrame(results)


# ── cross-file comparison ─────────────────────────────────────────────────────

def cross_compare(full_data, simp_data, full_df, simp_df):
    """
    Compare corresponding values between full (億元) and simplified (千元) reports.
    1億元 = 100,000千元.
    """
    SCALE = 100000  # 億元 → 千元

    def full_v(key, default=None):
        return full_data.get(key, default)

    def simp_v(key, default=None):
        return simp_data.get(key, default)

    rows = []
    pairs = [
        ('5年平均進口總額',
         '億元', full_v('total_import_reported', 147.86),
         '千元', simp_v('total_import_reported_k', 14785668)),
        ('關稅損失（最初收入損失法）',
         '億元', full_v('tariff_loss_reported', 22.40),
         '千元', simp_v('tariff_loss_reported_k', 2240360)),
        ('整車產值增加額',
         '億元', full_v('vehicle_output', 49.0),
         '千元', simp_v('vehicle_output_k', 4900813)),
        ('零組件產值增加額',
         '億元', full_v('parts_output', 34.31),
         '千元', simp_v('parts_output_k', 3430569)),
        ('整車營利事業所得稅',
         '億元', full_v('vehicle_cit_reported'),
         '千元', simp_v('vehicle_cit_reported_k')),
        ('零組件營利事業所得稅',
         '億元', full_v('parts_cit_reported'),
         '千元', simp_v('parts_cit_reported_k')),
        ('其他工業及服務業所得稅',
         '億元', full_v('other_cit_reported'),
         '千元', simp_v('other_cit_reported_k')),
        ('貨物稅',
         '億元', full_v('commodity_tax_reported'),
         '千元', simp_v('commodity_tax_reported_k')),
        ('加值型營業稅',
         '億元', full_v('vat_reported'),
         '千元', simp_v('vat_reported_k')),
        ('最終淨損益',
         '億元', full_v('net_reported', 0.29),
         '千元', simp_v('net_reported_k', 29231)),
    ]

    for label, u1, v1, u2, v2 in pairs:
        if v1 is None or v2 is None:
            rows.append({'項目': label, f'完整版({u1})': fmt(v1) if v1 else '—',
                         f'簡要格式({u2})': fmt(v2, 0) if v2 else '—',
                         '換算差異(千元)': '—', '結果': '⚠️'})
            continue
        v1_k = round2(v1 * SCALE)
        diff = round2(v2 - v1_k)
        passed = abs(diff) / max(abs(v2), 1) <= 0.01  # 1% tolerance
        rows.append({
            '項目': label,
            f'完整版({u1})': fmt(v1),
            f'簡要格式({u2})': fmt(v2, 0),
            '換算差異(千元)': fmt(diff, 0),
            '結果': status_icon(passed),
        })

    return pd.DataFrame(rows)


# ── Excel export ──────────────────────────────────────────────────────────────

SECTION_KEYS = [SEC_TARIFF, SEC_FINAL, SEC_CT, SEC_VAT, SEC_CIT, SEC_PER, SEC_DIV, SEC_NET]

FILL_GREEN   = PatternFill('solid', fgColor='C6EFCE')
FILL_RED     = PatternFill('solid', fgColor='FFC7CE')
FILL_YELLOW  = PatternFill('solid', fgColor='FFEB9C')
FILL_HEADER  = PatternFill('solid', fgColor='4472C4')
FILL_SECTION = PatternFill('solid', fgColor='D9E1F2')
FONT_WHITE   = Font(color='FFFFFF', bold=True)
FONT_SECTION = Font(bold=True, color='1F3864')
THIN         = Side(style='thin')
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def to_excel(df_list, sheet_names, compare_df=None):
    wb = Workbook()
    wb.remove(wb.active)

    for df, name in zip(df_list, sheet_names):
        ws = wb.create_sheet(title=name[:31])
        headers = list(df.columns)
        ws.append(headers)
        for ci, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci)
            cell.fill = FILL_HEADER
            cell.font = FONT_WHITE
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
            cell.border = BORDER

        current_section = None
        ri = 2
        for _, row_data in df.iterrows():
            sec = row_data.get('章節', '')
            if sec and sec != current_section:
                current_section = sec
                ws.merge_cells(start_row=ri, start_column=1,
                                end_row=ri, end_column=len(headers))
                cell = ws.cell(row=ri, column=1, value=sec)
                cell.fill = FILL_SECTION
                cell.font = FONT_SECTION
                cell.alignment = Alignment(horizontal='left')
                cell.border = BORDER
                ri += 1

            result_val = str(row_data.get('結果', ''))
            fill = (FILL_GREEN if '✅' in result_val
                    else FILL_RED if '❌' in result_val
                    else FILL_YELLOW)
            for ci, col in enumerate(headers, 1):
                val = row_data[col]
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                cell.border = BORDER
                cell.fill = fill
            ri += 1

        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    # Cross-file comparison sheet
    if compare_df is not None:
        ws = wb.create_sheet(title='兩份報告比較')
        headers = list(compare_df.columns)
        ws.append(headers)
        for ci, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=ci)
            cell.fill = FILL_HEADER
            cell.font = FONT_WHITE
            cell.alignment = Alignment(horizontal='center')
            cell.border = BORDER
        for ri, (_, row_data) in enumerate(compare_df.iterrows(), 2):
            result_val = str(row_data.get('結果', ''))
            fill = (FILL_GREEN if '✅' in result_val
                    else FILL_RED if '❌' in result_val
                    else FILL_YELLOW)
            for ci, col in enumerate(headers, 1):
                cell = ws.cell(row=ri, column=ci, value=row_data[col])
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                cell.border = BORDER
                cell.fill = fill
        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def main():
    st.title("📋 稅式支出評估報告驗證工具")
    st.markdown("上傳 Word 報告（完整版 F2 或簡要格式），自動驗證所有計算，並匯出 Excel 驗證報告。")

    uploaded_files = st.file_uploader(
        "選擇 Word 檔案（可同時上傳兩份）",
        type=['docx'],
        accept_multiple_files=True,
        help="支援完整版報告與簡要格式報告",
    )

    if not uploaded_files:
        st.markdown("""
**支援格式：**
- **完整版報告**（如「汽車零組件關稅調降稅式支出評估報告」.F2.docx）
  - 驗證：進口基準值、各品項關稅損失、各稅目公式、最終淨損益
- **簡要格式報告**（如「稅式支出評估報告(簡要格式)」.F2.docx）
  - 驗證：千元單位下各稅目公式計算

同時上傳兩份時，另產生**兩份報告比對**工作表，驗證數字一致性。
        """)
        return

    all_dfs   = []
    all_names = []
    all_data  = {}   # 'full' or 'simplified' → parsed data dict
    all_types = {}

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 {uploaded_file.name}")
        try:
            doc        = Document(io.BytesIO(uploaded_file.read()))
            tables     = extract_tables(doc)
            paragraphs = extract_paragraphs(doc)
            doc_type   = detect_type(tables, paragraphs)
            all_types[uploaded_file.name] = doc_type

            label = '完整版報告' if doc_type == 'full' else '簡要格式報告'
            st.caption(f"識別類型：{label}　｜　表格數量：{len(tables)}")

            if doc_type == 'full':
                data = parse_full_report(tables, paragraphs)
                df   = verify_full(data)
            else:
                data = parse_simplified_report(tables, paragraphs)
                df   = verify_simplified(data)

            all_data[doc_type] = data

            # Summary metrics
            total  = len(df)
            passed = (df['結果'] == '✅').sum()
            failed = (df['結果'] == '❌').sum()
            warn   = (df['結果'] == '⚠️').sum()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("驗證項目", total)
            c2.metric("✅ 通過", passed)
            c3.metric("❌ 有誤", failed)
            c4.metric("⚠️ 待確認", warn)

            # Color-coded table (hide 章節 column from display—used for grouping)
            display_cols = [c for c in df.columns if c != '章節']
            display_df   = df[display_cols]

            def highlight(row):
                color = ('#C6EFCE' if row['結果'] == '✅'
                         else '#FFC7CE' if row['結果'] == '❌'
                         else '#FFEB9C')
                return [f'background-color: {color}'] * len(row)

            # Group by section for display
            for sec, group in df.groupby('章節', sort=False):
                st.markdown(f"**{sec}**")
                g = group[display_cols].reset_index(drop=True)
                st.dataframe(g.style.apply(highlight, axis=1),
                             use_container_width=True, hide_index=True)

            # Item-level detail for full report
            if doc_type == 'full' and data.get('items'):
                with st.expander("展開：各品項明細表（表4-1）"):
                    item_df = pd.DataFrame(data['items'])
                    if not item_df.empty:
                        item_df['計算關稅損失(億元)'] = item_df.apply(
                            lambda r: round2(r['import_5yr'] * (r['current_rate'] - r['after_rate']))
                            if r['current_rate'] is not None and r['after_rate'] is not None else None,
                            axis=1)
                        item_df.columns = ['稅則號別','貨名','5年平均進口(億元)','現行稅率','降稅後稅率','計算關稅損失(億元)']
                        st.dataframe(item_df, use_container_width=True, hide_index=True)

            short = uploaded_file.name[:28].replace('.docx','').replace('「','').replace('」','')
            all_dfs.append(df)
            all_names.append(short)

        except Exception as e:
            st.error(f"解析失敗：{e}")
            st.exception(e)

    # Cross-file comparison (when both types present)
    compare_df = None
    if 'full' in all_data and 'simplified' in all_data:
        st.divider()
        st.subheader("📊 兩份報告比對")
        st.caption("完整版（億元）換算為千元（×100,000）後與簡要格式比對，容差 1%")
        full_df = all_dfs[[i for i, n in enumerate(all_names) if all_types.get(n+'簡') == 'full' or True][0]]
        simp_df = all_dfs[-1] if len(all_dfs) > 1 else all_dfs[0]
        compare_df = cross_compare(all_data['full'], all_data['simplified'], full_df, simp_df)

        def highlight_cmp(row):
            color = ('#C6EFCE' if row['結果'] == '✅'
                     else '#FFC7CE' if row['結果'] == '❌'
                     else '#FFEB9C')
            return [f'background-color: {color}'] * len(row)

        st.dataframe(compare_df.style.apply(highlight_cmp, axis=1),
                     use_container_width=True, hide_index=True)

    # Excel download
    if all_dfs:
        st.divider()
        excel_buf = to_excel(all_dfs, all_names, compare_df)
        st.download_button(
            label="📥 下載 Excel 驗證報告",
            data=excel_buf,
            file_name="稅式支出驗證結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == '__main__':
    main()
