# -*- coding: utf-8 -*-
"""
SmartCustoms Pro | 报关协同与自动化处理系统 (Streamlit Web Edition)
极简高能、无冗余干扰信息，专为部署于 streamlit.io 设计。
"""

import io
import os
import re
import json
import zipfile
import urllib.request
import urllib.error
from datetime import datetime, timedelta
import calendar
from copy import copy

import numpy as np
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, Border, Side, Alignment
import xlsxwriter
import streamlit as st

# ==================== 1. 页面基础配置 (左侧栏默认折叠) ====================
st.set_page_config(
    page_title="报关协同系统",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 仅保留基础按钮样式微调，杜绝任何颜色/背景覆盖导致的文字对比度问题
st.markdown("""
<style>
div.stButton > button {
    border-radius: 8px;
    font-weight: 600;
}
[data-testid="stFileUploader"] {
    padding: 2px 0;
}
</style>
""", unsafe_allow_html=True)

# ==================== 2. 默认字段与配置 ====================
DEFAULT_CONFIG = {
    "airscript": {
        "webhook_url": "https://www.kdocs.cn/api/v3/ide/file/cdH0A450EedY/script/V2-6x7HgWruLz4P74YP1PIsVf/sync_task",
        "token": "RRZwCZarLOHD4gHKtfSVi"
    },
    "fba_merger": {
        "address_col": "物流中心编码",
        "fba_col": "FBA编号",
        "product_col": "中文品名",
        "unit_col": "单位",
        "output_address_col": "配送地址",
        "must_cols": ["物流中心编码", "中文品名", "单位"],
        "num_cols": ["箱数", "数量", "金额USD", "毛重KGS", "净重KGS", "体积", "货值", "单价"]
    },
    "customs_doc_generator": {
        "purchase_sku": "SKU",
        "purchase_price": "采购单价",
        "purchase_name": "品名",
        "purchase_unit": "单位",
        "purchase_hs": "HS编码",
        "purchase_en_name": "中英文品名",
        "purchase_elements": "申报要素(品牌,材质,用途)",
        "purchase_origin": "境内货源地",
        "delivery_sku": "SKU",
        "delivery_qty": "发货量",
        "delivery_ctns": "箱数",
        "delivery_box_qty": "单箱数量",
        "delivery_center": "物流中心编码",
        "delivery_supplier": "供应商",
        "delivery_shipment_id": "货件编号",
        "delivery_gross_weight": "外箱总重量(kg)",
        "delivery_volume": "外箱总体积(m³)"
    }
}

# ==================== 3. 侧边栏（默认折叠，纯工具无废话） ====================
with st.sidebar:
    st.subheader("⚙️ 业务配置")

    with st.expander("AirScript 采购单接口", expanded=False):
        air_webhook = st.text_input(
            "Webhook 地址",
            value=DEFAULT_CONFIG["airscript"]["webhook_url"]
        )
        air_token = st.text_input(
            "访问令牌 (Token)",
            value=DEFAULT_CONFIG["airscript"]["token"],
            type="password"
        )
        if st.button("测试接口", use_container_width=True):
            try:
                with st.spinner("测试中..."):
                    df_test = fetch_purchase_from_airscript(air_webhook, air_token)
                    st.success(f"连通正常，获取到 {len(df_test)} 条记录")
            except Exception as e:
                st.error(f"连接失败: {str(e)}")

    exchange_rate = st.number_input(
        "美元汇率 (USD/CNY)",
        min_value=0.1,
        max_value=20.0,
        value=7.20,
        step=0.01,
        format="%.2f",
        help="用于采购单单价与申报货值精确换算（无尾差）"
    )

# ==================== 4. 辅助与算法函数 ====================
def parse_airscript_response(res_data):
    if isinstance(res_data, dict):
        if "data" in res_data and isinstance(res_data["data"], dict) and "result" in res_data["data"]:
            result = res_data["data"]["result"]
        elif "result" in res_data:
            result = res_data["result"]
        else:
            result = res_data
    else:
        result = res_data

    if isinstance(result, list):
        if len(result) == 0: return pd.DataFrame()
        first_item = result[0]
        if isinstance(first_item, dict):
            if 'fields' in first_item and isinstance(first_item['fields'], dict):
                records = [item.get('fields', {}) for item in result if isinstance(item, dict)]
                return pd.DataFrame(records)
            return pd.DataFrame(result)
        elif isinstance(first_item, (list, tuple)):
            headers = [str(h).strip() for h in first_item]
            rows = result[1:]
            return pd.DataFrame(rows, columns=headers)
    elif isinstance(result, dict):
        if 'records' in result and isinstance(result['records'], list):
            return parse_airscript_response(result['records'])
        return pd.DataFrame([result])

    return pd.DataFrame()

@st.cache_data(ttl=300, show_spinner=False)
def fetch_purchase_from_airscript(webhook_url, token, timeout=35):
    if not webhook_url or not token:
        raise ValueError("缺少 AirScript Webhook 链接或令牌配置！")

    headers = {
        'AirScript-Token': token.strip(),
        'Content-Type': 'application/json; charset=utf-8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
    }
    payload = json.dumps({"Context": {"argv": {}}}).encode('utf-8')
    req = urllib.request.Request(webhook_url.strip(), data=payload, headers=headers, method='POST')

    with urllib.request.urlopen(req, timeout=timeout) as response:
        status_code = response.getcode()
        resp_body = response.read().decode('utf-8')
        if status_code != 200:
            raise Exception(f"HTTP 请求失败，状态码: {status_code}")
        
        json_data = json.loads(resp_body)
        if isinstance(json_data, dict) and json_data.get("status") == "error":
            msg = json_data.get("message") or json_data.get("msg") or str(json_data)
            raise Exception(f"AirScript 执行报错: {msg}")
        
        df = parse_airscript_response(json_data)
        if df.empty:
            raise Exception("AirScript 响应成功，但返回的数据表为空！")
        return df

FOOTER_KEYWORDS = [
    '合计', '总计', 'Total', 'SUM', 'Sum',
    '备注', 'Remark', 'Comments',
    '签字', 'Signature', 'Sign',
    '日期', 'Date',
    '申报单位', '填表人', 'Approved', 'Checked'
]

CLEAR_RULES = {
    1: {"报关单": (23, 34, 1, 30), "装箱单": (12, (2, 3, 4, 5), 1, 12),
        "发票": (15, (2, 3, 4, 5), 1, 25), "合同": (14, (2, 3, 4, 5), 1, 20)},
    2: {"报关单": (26, 34, 1, 30), "装箱单": (12, (3, 4, 5), 1, 12),
        "发票": (15, (3, 4, 5), 1, 25), "合同": (14, (3, 4, 5), 1, 20)},
    3: {"报关单": (29, 34, 1, 30), "装箱单": (12, (4, 5), 1, 12),
        "发票": (15, (4, 5), 1, 25), "合同": (14, (4, 5), 1, 20)},
    4: {"报关单": (32, 34, 1, 30), "装箱单": (12, (5,), 1, 12),
        "发票": (15, (5,), 1, 25), "合同": (14, (5,), 1, 20)},
    5: {"报关单": None, "装箱单": None, "发票": None, "合同": None}
}

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
FONT_STYLE = Font(name='微软雅黑', size=10)

def normalize_header(value):
    if value is None: return ""
    return str(value).strip().replace(" ", "").replace("\n", "").replace("\t", "").replace("\r", "")

def find_real_header_row_from_bytes(file_bytes, sheet_name="汇总"):
    try:
        df_temp = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=20, engine='openpyxl')
        for idx, row in df_temp.iterrows():
            row_vals = [normalize_header(val) for val in row if pd.notna(val)]
            if "FBA编号" in row_vals and "账号" in row_vals:
                return idx
    except Exception:
        pass
    return 0

def generate_unique_fba_code(index):
    now = datetime.now()
    year = now.strftime("%Y")
    month = now.strftime("%m")
    seq_num = str(index).zfill(4)
    return f"FBA-{year}-{month}-{seq_num}"

def get_top_left_cell(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

def check_row_has_data(ws, row, col_indices):
    for col in col_indices:
        cell = ws.cell(row=row, column=col)
        real_cell = get_top_left_cell(ws, cell)
        if real_cell.value is not None and str(real_cell.value).strip() != "":
            return True, str(real_cell.value)
    return False, ""

def find_footer_boundary(ws, start_row, col_indices):
    max_row = ws.max_row
    empty_row_count = 0
    for r in range(start_row, max_row + 10):
        has_data, content = check_row_has_data(ws, r, col_indices)
        if has_data:
            content_norm = str(content).strip()
            for kw in FOOTER_KEYWORDS:
                if kw in content_norm: return r
            empty_row_count = 0
        else:
            empty_row_count += 1
            if empty_row_count >= 5: return r - 4
    return max_row + 1

def clear_data_range(ws, start_row, end_row, col_indices):
    if end_row < start_row: return
    for r in range(start_row, end_row):
        for c in col_indices:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell): continue
            cell.value = None

def is_cell_merged(ws, row, col):
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row) and \
           (merged_range.min_col <= col <= merged_range.max_col):
            return True
    return False

def clear_row_range(ws, start_row, end_row, start_col, end_col):
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            if is_cell_merged(ws, row, col): continue
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.data_type = "n"
            if hasattr(cell, "border"): cell.border = copy(cell.border)

def clear_sequence_rows(ws, start_row, clear_seq, start_col, end_col):
    for seq in clear_seq:
        row_to_clear = start_row + (seq - 1)
        if row_to_clear > ws.max_row: continue
        for col in range(start_col, end_col + 1):
            if is_cell_merged(ws, row_to_clear, col): continue
            cell = ws.cell(row=row_to_clear, column=col)
            cell.value = None
            cell.data_type = "n"
            if hasattr(cell, "border"): cell.border = copy(cell.border)

def apply_template_cleanup_memory(wb, actual_count):
    if actual_count not in CLEAR_RULES: return
    rule_set = CLEAR_RULES[actual_count]
    bgd_rule = rule_set.get("报关单")
    if bgd_rule and "报关单" in wb.sheetnames:
        bgd_ws = wb["报关单"]
        r_start, r_end, c_start, c_end = bgd_rule
        clear_row_range(bgd_ws, r_start, r_end, c_start, c_end)

    for sheet in ("装箱单", "发票", "合同"):
        rule = rule_set.get(sheet)
        if rule and sheet in wb.sheetnames:
            r_start, seq_tuple, c_start, c_end = rule
            if seq_tuple:
                sheet_ws = wb[sheet]
                clear_sequence_rows(sheet_ws, r_start, seq_tuple, c_start, c_end)


def extract_date_from_contract(contract_str):
    if not contract_str:
        return None
    m8 = re.search(r'(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', str(contract_str))
    if m8:
        y, m, d = int(m8.group(1)), int(m8.group(2)), int(m8.group(3))
        return datetime(y, m, d)
    m6 = re.search(r'([2-9]\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', str(contract_str))
    if m6:
        y, m, d = 2000 + int(m6.group(1)), int(m6.group(2)), int(m6.group(3))
        return datetime(y, m, d)
    return None

def calc_customs_dates(base_dt):
    y = base_dt.year
    m = base_dt.month - 1
    if m == 0:
        m = 12
        y -= 1
    max_d = calendar.monthrange(y, m)[1]
    d = min(base_dt.day, max_d)
    contract_dt = datetime(y, m, d)

    invoice_dt = contract_dt + timedelta(days=7)
    return contract_dt.strftime('%Y-%m-%d'), invoice_dt.strftime('%Y-%m-%d')

def update_template_dates(ws, contract_date_str, invoice_date_str):
    if not contract_date_str and not invoice_date_str:
        return
    for r in range(1, 40):
        for c in range(1, 10):
            cell_obj = ws.cell(row=r, column=c)
            real_cell = get_top_left_cell(ws, cell_obj)
            val_norm = normalize_header(real_cell.value)
            if val_norm == '合同时间' and contract_date_str:
                target = ws.cell(row=r, column=c + 1)
                get_top_left_cell(ws, target).value = contract_date_str
            elif val_norm == '发票日期' and invoice_date_str:
                target = ws.cell(row=r, column=c + 1)
                get_top_left_cell(ws, target).value = invoice_date_str
            elif val_norm == '装运日期' and invoice_date_str:
                target = ws.cell(row=r, column=c + 1)
                get_top_left_cell(ws, target).value = invoice_date_str

def fill_and_clean_template_memory(template_bytes, fba_data_df, fba_code, account_info, manual_contract_date=None, manual_invoice_date=None, auto_calc_dates=True):
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False, read_only=False)
    try:
        wb.calculation.calcMode = "manual"
        wb.calculation.calcOnSave = False
        if "输入表格" not in wb.sheetnames:
            raise ValueError("模板中缺少【输入表格】")
        ws = wb["输入表格"]

        target_b12 = ws['B12']
        real_b12 = get_top_left_cell(ws, target_b12)
        real_b12.value = account_info

        field_mapping = {
            '合同号': '合同号', 'HS编码': 'HS编码', '中文品名': '中文品名', '中英文品名': '中英文品名',
            '申报要素(品牌,材质,用途)': '申报要素(品牌,材质,用途)', '箱数': '箱数', 'CTNS': 'CTNS',
            '数量': '数量', '单位': '单位', '单价': '单价', '金额USD': '金额USD',
            '毛重KGS': '毛重KGS', '净重KGS': '净重KGS', '体积': '体积', '境内货源地': '境内货源地',
            'FBA编号': 'FBA编号'
        }

        header_row = -1
        field_col = {}
        for r in range(1, 16):
            temp_field_col = {}
            match_count = 0
            for c in range(1, 51):
                cell_obj = ws.cell(row=r, column=c)
                real_cell = get_top_left_cell(ws, cell_obj)
                clean_val = normalize_header(real_cell.value)
                if clean_val in field_mapping:
                    temp_field_col[clean_val] = c
                    match_count += 1
            if match_count >= 5:
                header_row = r
                field_col = temp_field_col
                break

        if header_row == -1: raise ValueError("无法在模板【输入表格】中识别表头")

        data_start_row = header_row + 1
        mapped_col_indices = list(field_col.values())
        footer_boundary_row = find_footer_boundary(ws, data_start_row, mapped_col_indices)
        clear_data_range(ws, start_row=data_start_row, end_row=footer_boundary_row, col_indices=mapped_col_indices)

        df_clean = fba_data_df.fillna("")
        data_list = df_clean.values.tolist()

        col_order = df_clean.columns.tolist()
        normalized_col_order = [normalize_header(c) for c in col_order]
        field_data_idx = {}
        for tm_key, tm_name in field_mapping.items():
            try: field_data_idx[tm_key] = normalized_col_order.index(normalize_header(tm_name))
            except ValueError: pass

        total_rows = len(data_list)
        contract_no = ""

        for row_idx in range(total_rows):
            excel_row = data_start_row + row_idx
            row_data = data_list[row_idx]
            ws.row_dimensions[excel_row].height = 25

            for tm_field, dt_idx in field_data_idx.items():
                if tm_field in field_col:
                    col_idx = field_col[tm_field]
                    cell_value = row_data[dt_idx]
                    if tm_field == '合同号' and not contract_no and cell_value:
                        contract_no = str(cell_value).strip()

                    target_cell = ws.cell(row=excel_row, column=col_idx)
                    if isinstance(target_cell, MergedCell):
                        real_target = get_top_left_cell(ws, target_cell)
                        real_target.value = cell_value
                        real_target.border = THIN_BORDER
                        real_target.alignment = ALIGN_CENTER
                        real_target.font = FONT_STYLE
                    else:
                        target_cell.value = cell_value
                        target_cell.border = THIN_BORDER
                        target_cell.alignment = ALIGN_CENTER
                        target_cell.font = FONT_STYLE

        if not contract_no and '合同号' in fba_data_df.columns:
            contracts = fba_data_df['合同号'].dropna().astype(str).str.strip().tolist()
            if contracts and contracts[0]: contract_no = contracts[0]

        final_contract_date = ''
        final_invoice_date = ''
        if auto_calc_dates:
            dt = extract_date_from_contract(contract_no)
            if dt:
                final_contract_date, final_invoice_date = calc_customs_dates(dt)
            elif manual_contract_date and manual_invoice_date:
                final_contract_date = manual_contract_date
                final_invoice_date = manual_invoice_date
        else:
            if manual_contract_date and manual_invoice_date:
                final_contract_date = manual_contract_date
                final_invoice_date = manual_invoice_date

        if final_contract_date and final_invoice_date:
            update_template_dates(ws, final_contract_date, final_invoice_date)

        actual_items_count = min(max(total_rows, 1), 5)
        apply_template_cleanup_memory(wb, actual_items_count)

        safe_contract = re.sub(r'[/\\:*?"<>|]', '_', contract_no).strip() if contract_no else ""
        safe_fba = str(fba_code).replace("/", "_").replace("\\", "_").strip()
        safe_account = str(account_info).replace("/", "_").replace("\\", "_").strip()

        name_parts = []
        if safe_contract: name_parts.append(safe_contract)
        name_parts.append("报关单")
        if safe_account: name_parts.append(safe_account)
        if safe_fba: name_parts.append(safe_fba)

        final_filename = "_".join(name_parts) + ".xlsx"

        wb.calculation.calcMode = "auto"
        out_stream = io.BytesIO()
        wb.save(out_stream)
        out_stream.seek(0)
        return final_filename, out_stream.getvalue(), actual_items_count, (final_contract_date, final_invoice_date)
    finally:
        wb.close()

# ==================== 5. 主页面布局 ====================
st.title("📋 报关协同处理系统")

tab1, tab2, tab3 = st.tabs([
    "📦 1. 报关资料在线生成",
    "🔄 2. FBA 报关数据合并",
    "📑 3. 报关单套打与自动清理"
])

# ----------------- TAB 1: 报关资料在线生成 -----------------
with tab1:
    delivery_file = st.file_uploader(
        "上传发货单 Excel 文件 (.xlsx)",
        type=["xlsx"],
        key="tab1_delivery",
        help="系统将自动直连云端采购单，匹配品名、HS编码、要素并换算美元货值（无尾差）"
    )

    if delivery_file:
        d_id = f"{delivery_file.name}_{delivery_file.size}"
        if st.session_state.get('tab1_file_id') != d_id:
            st.session_state['tab1_result'] = None
            st.session_state['tab1_file_id'] = d_id
    else:
        st.session_state['tab1_result'] = None
        st.session_state['tab1_file_id'] = None

    if st.button("🚀 开始生成报关资料", type="primary", use_container_width=True, disabled=not delivery_file):
        try:
            with st.spinner("正在生成..."):
                cfg = DEFAULT_CONFIG["customs_doc_generator"]
                p_sku, p_price, p_name, p_unit = cfg["purchase_sku"], cfg["purchase_price"], cfg["purchase_name"], cfg["purchase_unit"]
                p_hs, p_en, p_elem, p_origin = cfg["purchase_hs"], cfg["purchase_en_name"], cfg["purchase_elements"], cfg["purchase_origin"]
                d_sku, d_qty, d_ctns, d_box_qty = cfg["delivery_sku"], cfg["delivery_qty"], cfg["delivery_ctns"], cfg["delivery_box_qty"]
                d_center, d_supp, d_ship_id = cfg["delivery_center"], cfg["delivery_supplier"], cfg["delivery_shipment_id"]
                d_gw, d_vol = cfg["delivery_gross_weight"], cfg["delivery_volume"]

                purchase = fetch_purchase_from_airscript(air_webhook, air_token)
                for req_col in [p_sku, p_price, p_name, p_unit, p_hs, p_en, p_elem, p_origin]:
                    if req_col not in purchase.columns:
                        purchase[req_col] = "" if req_col != p_price else 0

                purchase[p_price] = pd.to_numeric(purchase[p_price], errors='coerce').fillna(0)
                purchase = purchase.fillna({
                    p_hs: "", p_en: "", p_elem: "", p_origin: "",
                    p_name: "", p_unit: ""
                })
                purchase = purchase.dropna(subset=[p_sku])

                sku_price_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_price]))
                sku_name_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_name]))
                sku_unit_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_unit]))
                sku_hs_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_hs]))
                sku_en_name_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_en]))
                sku_declare_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_elem]))
                sku_origin_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_origin]))

                delivery = pd.read_excel(
                    delivery_file,
                    usecols=[d_sku, d_qty, d_ctns, d_box_qty, d_center, d_supp, d_ship_id, d_gw, d_vol]
                )
                delivery = delivery.fillna({
                    d_qty: 0, d_ctns: 0, d_gw: 0, d_vol: 0,
                    d_center: "", d_supp: "", d_ship_id: "", d_box_qty: 0
                })
                delivery = delivery.dropna(subset=[d_sku])

                delivery["账号"] = delivery[d_sku].astype(str).str[:2].str.upper().fillna("")
                clean_skus = delivery[d_sku].astype(str).str.strip()

                delivery["采购单价"] = clean_skus.map(sku_price_dict).fillna(0)
                delivery["品名"] = clean_skus.map(sku_name_dict).fillna("【无采购单匹配】")
                delivery["单位"] = clean_skus.map(sku_unit_dict).fillna("")
                delivery["HS编码"] = clean_skus.map(sku_hs_dict).fillna("")
                delivery["中英文品名"] = clean_skus.map(sku_en_name_dict).fillna("【无采购单匹配】")
                delivery["申报要素(品牌,材质,用途)"] = clean_skus.map(sku_declare_dict).fillna("【无采购单匹配，需手动录入】")
                delivery["境内货源地"] = clean_skus.map(sku_origin_dict).fillna("")

                delivery["箱数"] = delivery[d_ctns]
                delivery["发货量"] = delivery[d_qty].astype(int)
                delivery["外箱总重量(kg)"] = delivery[d_gw]
                delivery["外箱总体积(m³)"] = delivery[d_vol]
                delivery["物流中心编码"] = delivery[d_center]
                delivery["供应商"] = delivery[d_supp]
                delivery["货件编号"] = delivery[d_ship_id]
                delivery["SKU"] = delivery[d_sku]

                delivery["货值"] = (delivery["发货量"] * delivery["采购单价"]).round(2).fillna(0)
                delivery["金额USD"] = (delivery["货值"] / exchange_rate).round(2)
                delivery["发货量_calc"] = delivery["发货量"].replace(0, 1)
                delivery["单价(USD)"] = (delivery["金额USD"] / delivery["发货量_calc"]).round(2)
                delivery["金额USD"] = (delivery["发货量"] * delivery["单价(USD)"]).round(2)
                delivery.drop(columns=["发货量_calc"], inplace=True)

                delivery["净重(kg)"] = (delivery["外箱总重量(kg)"] * 0.95).round(2).fillna(0)
                delivery["外箱总体积(m³)"] = delivery["外箱总体积(m³)"].round(2).fillna(0)

                detail = delivery.sort_values("货件编号").reset_index(drop=True)

                summary = detail[detail["采购单价"] > 0].groupby(
                    by=["货件编号", "供应商", "品名", "账号"], as_index=False, dropna=False
                ).agg({
                    "箱数": "sum", "发货量": "sum", "外箱总重量(kg)": "sum", "净重(kg)": "sum",
                    "外箱总体积(m³)": "sum", "货值": "sum", "物流中心编码": "first", "金额USD": "sum",
                    "单位": "first", "HS编码": "first", "中英文品名": "first",
                    "申报要素(品牌,材质,用途)": "first", "境内货源地": "first"
                }).round(2)
                summary["单价(USD)"] = (summary["金额USD"] / summary["发货量"]).round(2)
                summary["金额USD"] = (summary["发货量"] * summary["单价(USD)"]).round(2)
                summary = summary.fillna({
                    "HS编码": "", "中英文品名": "", "申报要素(品牌,材质,用途)": "", "境内货源地": "",
                    "单价(USD)": 0, "金额USD": 0, "货值": 0, "单位": ""
                })
                summary["单价(USD)"] = summary["单价(USD)"].replace([np.inf, -np.inf], 0)

                out_buffer = io.BytesIO()
                workbook = xlsxwriter.Workbook(out_buffer, {'in_memory': True})

                header_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 11, "bold": True,
                    "font_color": "white", "bg_color": "#1e3a8a",
                    "align": "center", "valign": "vcenter", "border": 1
                })
                title_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 13, "bold": True,
                    "align": "center", "valign": "vcenter"
                })
                int_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 10, "border": 1,
                    "align": "center", "valign": "vcenter", "num_format": "#,##0"
                })
                float2_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 10, "border": 1,
                    "align": "center", "valign": "vcenter", "num_format": "#,##0.00"
                })
                text_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 10, "border": 1,
                    "align": "center", "valign": "vcenter"
                })
                unmatch_fmt = workbook.add_format({
                    "font_name": "微软雅黑", "font_size": 10, "border": 1,
                    "align": "center", "valign": "vcenter",
                    "font_color": "#dc2626", "bold": True
                })
                color1 = workbook.add_format({
                    "bg_color": "#ffffff", "border": 1, "align": "center", "valign": "vcenter",
                    "font_name": "微软雅黑", "font_size": 10
                })
                color2 = workbook.add_format({
                    "bg_color": "#f1f5f9", "border": 1, "align": "center", "valign": "vcenter",
                    "font_name": "微软雅黑", "font_size": 10
                })

                headers = [
                    "账号", "HS编码", "中文品名", "中英文品名", "申报要素(品牌,材质,用途)",
                    "箱数", "CTNS", "数量", "单位", "单价", "币制", "金额USD",
                    "毛重KGS", "净重KGS", "体积", "境内货源地",
                    "FBA编号", "合同号", "物流中心编码", "货值", "供应商", "SKU", "采购单价(元)"
                ]
                col_widths = [8, 12, 20, 25, 30, 8, 8, 10, 8, 10, 8, 12, 10, 10, 12, 15, 15, 12, 20, 12, 15, 15, 12]

                ws_detail = workbook.add_worksheet("明细")
                ws_detail.merge_range("A1:W1", f"报关明细数据（汇率：1美元={exchange_rate}人民币 | 无尾差）", title_fmt)
                for col, h in enumerate(headers): ws_detail.write(1, col, h, header_fmt)
                for col, w in enumerate(col_widths): ws_detail.set_column(col, col, w)

                for r_idx, (_, row) in enumerate(detail.iterrows(), 2):
                    is_un = row["采购单价"] == 0
                    bf = unmatch_fmt if is_un else text_fmt
                    nf = unmatch_fmt if is_un else int_fmt
                    ff = unmatch_fmt if is_un else float2_fmt

                    ws_detail.write(r_idx, 0, row["账号"], bf)
                    ws_detail.write(r_idx, 1, row["HS编码"], bf)
                    ws_detail.write(r_idx, 2, row["品名"], bf)
                    ws_detail.write(r_idx, 3, row["中英文品名"], bf)
                    ws_detail.write(r_idx, 4, row["申报要素(品牌,材质,用途)"], bf)
                    ws_detail.write(r_idx, 5, row["箱数"], nf)
                    ws_detail.write(r_idx, 6, "CTNS", bf)
                    ws_detail.write(r_idx, 7, row["发货量"], nf)
                    ws_detail.write(r_idx, 8, row["单位"], bf)
                    ws_detail.write(r_idx, 9, row["单价(USD)"], ff)
                    ws_detail.write(r_idx, 10, "USD", bf)
                    ws_detail.write(r_idx, 11, row["金额USD"], ff)
                    ws_detail.write(r_idx, 12, row["外箱总重量(kg)"], ff)
                    ws_detail.write(r_idx, 13, row["净重(kg)"], ff)
                    ws_detail.write(r_idx, 14, row["外箱总体积(m³)"], ff)
                    ws_detail.write(r_idx, 15, row["境内货源地"], bf)
                    ws_detail.write(r_idx, 16, row["货件编号"], bf)
                    ws_detail.write(r_idx, 17, "", bf)
                    ws_detail.write(r_idx, 18, row["物流中心编码"], bf)
                    ws_detail.write(r_idx, 19, row["货值"], ff)
                    ws_detail.write(r_idx, 20, row["供应商"], bf)
                    ws_detail.write(r_idx, 21, row["SKU"], bf)
                    ws_detail.write(r_idx, 22, row["采购单价"], ff)

                ws_summary = workbook.add_worksheet("汇总")
                ws_summary.merge_range("A1:W1", f"报关汇总数据（仅匹配项，汇率：1美元={exchange_rate}人民币 | 无尾差）", title_fmt)
                for col, h in enumerate(headers): ws_summary.write(1, col, h, header_fmt)
                for col, w in enumerate(col_widths): ws_summary.set_column(col, col, w)

                fba_uniques = summary["货件编号"].unique()
                color_map = {fba: color1 if i % 2 == 0 else color2 for i, fba in enumerate(fba_uniques)}

                for r_idx, (_, row) in enumerate(summary.iterrows(), 2):
                    rf = color_map.get(row["货件编号"], color1)
                    ws_summary.write(r_idx, 0, row["账号"], rf)
                    ws_summary.write(r_idx, 1, row["HS编码"], rf)
                    ws_summary.write(r_idx, 2, row["品名"], rf)
                    ws_summary.write(r_idx, 3, row["中英文品名"], rf)
                    ws_summary.write(r_idx, 4, row["申报要素(品牌,材质,用途)"], rf)
                    ws_summary.write(r_idx, 5, row["箱数"], rf)
                    ws_summary.write(r_idx, 6, "CTNS", rf)
                    ws_summary.write(r_idx, 7, row["发货量"], rf)
                    ws_summary.write(r_idx, 8, row["单位"], rf)
                    ws_summary.write(r_idx, 9, row["单价(USD)"], rf)
                    ws_summary.write(r_idx, 10, "USD", rf)
                    ws_summary.write(r_idx, 11, row["金额USD"], rf)
                    ws_summary.write(r_idx, 12, row["外箱总重量(kg)"], rf)
                    ws_summary.write(r_idx, 13, row["净重(kg)"], rf)
                    ws_summary.write(r_idx, 14, row["外箱总体积(m³)"], rf)
                    ws_summary.write(r_idx, 15, row["境内货源地"], rf)
                    ws_summary.write(r_idx, 16, row["货件编号"], rf)
                    ws_summary.write(r_idx, 17, "", rf)
                    ws_summary.write(r_idx, 18, row["物流中心编码"], rf)
                    ws_summary.write(r_idx, 19, row["货值"], rf)
                    ws_summary.write(r_idx, 20, row["供应商"], rf)
                    ws_summary.write(r_idx, 21, "汇总", rf)
                    ws_summary.write(r_idx, 22, "", rf)

                workbook.close()
                out_buffer.seek(0)

                unmatch_count = (detail["采购单价"] == 0).sum()
                st.session_state['tab1_result'] = {
                    'bytes': out_buffer.getvalue(),
                    'unmatch_count': unmatch_count
                }
        except Exception as e:
            st.error(f"处理失败: {str(e)}")

    if 'tab1_result' in st.session_state and st.session_state['tab1_result']:
        res1 = st.session_state['tab1_result']
        if res1['unmatch_count'] > 0:
            st.warning(f"注意：存在 {res1['unmatch_count']} 项未匹配采购单的 SKU，已在明细红字标注")

        st.download_button(
            label="📥 下载报关资料 (.xlsx)",
            data=res1['bytes'],
            file_name=f"报关资料_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# ----------------- TAB 2: FBA 报关数据合并 (单按钮极简交互) -----------------
with tab2:
    if 'merger_groups' not in st.session_state:
        st.session_state['merger_groups'] = []

    fba_file = st.file_uploader(
        "上传待合并的报关汇总 Excel 文件 (.xlsx)",
        type=["xlsx"],
        key="tab2_file",
        help="请上传包含【汇总】工作表的报关文件"
    )

    if fba_file:
        file_id = f"{fba_file.name}_{fba_file.size}"
        if st.session_state.get('tab2_file_id') != file_id:
            st.session_state['merger_groups'] = []
            st.session_state['tab2_merged_result'] = None
            st.session_state['tab2_file_id'] = file_id
            if 'tab2_addr_select' in st.session_state:
                st.session_state['tab2_addr_select'] = []
    else:
        st.session_state['merger_groups'] = []
        st.session_state['tab2_merged_result'] = None
        st.session_state['tab2_file_id'] = None

    if fba_file:
        try:
            cfg = DEFAULT_CONFIG["fba_merger"]
            addr_col = cfg["address_col"]
            fba_col = cfg["fba_col"]
            prod_col = cfg["product_col"]
            unit_col = cfg["unit_col"]
            out_addr_col = cfg["output_address_col"]
            must_cols = cfg["must_cols"]
            num_cols = cfg["num_cols"]

            file_bytes = fba_file.getvalue()
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="汇总", header=1)
            df_raw.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(df_raw.columns)]

            for c in must_cols:
                if c not in df_raw.columns:
                    st.error(f"缺少必要列: {c}")
                    st.stop()

            df_clean = df_raw.dropna(how='all').reset_index(drop=True)
            for col in num_cols:
                if col in df_clean.columns:
                    df_clean[col] = pd.to_numeric(
                        df_clean[col].astype(str).str.replace(',', '').str.strip(), errors='coerce'
                    ).fillna(0)

            for col in [addr_col, prod_col, unit_col, fba_col]:
                if col in df_clean.columns:
                    df_clean[col] = df_clean[col].astype(str).str.strip()

            df_clean[addr_col] = df_clean[addr_col].astype(str).apply(
                lambda x: re.sub(r'[\(（].*?[\\)）]', '', x).strip()
            )
            df_clean = df_clean[df_clean[addr_col].notna() & (df_clean[addr_col] != 'nan') & (df_clean[addr_col] != '')]

            address_list = []
            address_mapping = {}
            fba_display_map = {}
            seen_display = set()

            for _, row in df_clean.iterrows():
                addr = row[addr_col]
                fba_val = row.get(fba_col, '').strip() if fba_col in row else ''
                disp = f"{addr}({fba_val})" if (fba_val and fba_val != 'nan') else addr
                if disp not in seen_display:
                    seen_display.add(disp)
                    address_list.append(disp)
                    address_mapping[disp] = addr
                    if fba_val: fba_display_map[fba_val] = disp

            col_left, col_right = st.columns(2)

            with col_left:
                selected_from_dropdown = st.multiselect(
                    "选取待合并地址及 FBA",
                    options=address_list,
                    key="tab2_addr_select"
                )
                pasted_fbas = st.text_area(
                    "或：粘贴 FBA 编号 (每行一个)",
                    placeholder="FBA18XXXXXXX\nFBA18YYYYYYY",
                    height=70
                )
                if st.button("➕ 保存为合并组", use_container_width=True):
                    current_items = list(selected_from_dropdown)
                    if pasted_fbas.strip():
                        for line in pasted_fbas.splitlines():
                            fba_code = line.strip()
                            if fba_code in fba_display_map:
                                disp = fba_display_map[fba_code]
                                if disp not in current_items:
                                    current_items.append(disp)
                    if current_items:
                        st.session_state['merger_groups'].append(current_items)
                        st.session_state['tab2_merged_result'] = None
                        st.rerun()
                    else:
                        st.warning("请至少选取或输入一个 FBA 编号")

            with col_right:
                st.write(f"**已暂存分组 ({len(st.session_state['merger_groups'])} 组)**")
                if not st.session_state['merger_groups']:
                    st.caption("暂未添加分组，请在左侧选取并保存")
                else:
                    for g_idx, grp in enumerate(st.session_state['merger_groups']):
                        c_a, c_b = st.columns([5, 1])
                        items_str = "、".join(grp)
                        c_a.markdown(
                            f"**组 {g_idx+1}** <span style='color: #64748b;'>({len(grp)}个货件)</span><br>"
                            f"<span style='font-size: 13px; line-height: 1.4; word-break: break-all;'>{items_str}</span>",
                            unsafe_allow_html=True
                        )
                        if c_b.button("✕", key=f"del_grp_{g_idx}", help="删除此组"):
                            st.session_state['merger_groups'].pop(g_idx)
                            st.session_state['tab2_merged_result'] = None
                            st.rerun()

                    if st.button("清空所有组", type="secondary"):
                        st.session_state['merger_groups'] = []
                        st.session_state['tab2_merged_result'] = None
                        st.rerun()

            if st.session_state['merger_groups']:
                if st.button("🚀 开始合并数据并打包", type="primary", use_container_width=True):
                    try:
                        total_grps = len(st.session_state['merger_groups'])
                        pbar = st.progress(0, text="正在准备合并工作表...")

                        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
                        ws = wb["汇总"]

                        for g_idx, grp in enumerate(st.session_state['merger_groups'], start=1):
                            pbar.progress(
                                int((g_idx - 1) / total_grps * 100),
                                text=f"[{g_idx}/{total_grps}] 正在合并组 {g_idx}（共 {len(grp)} 个货件）..."
                            )

                            valid_addrs = [address_mapping[item] for item in grp if item in address_mapping]
                            filtered = df_clean[df_clean[addr_col].isin(valid_addrs)].copy()
                            if filtered.empty:
                                continue

                            full_fba_list = []
                            for item in grp:
                                if '(' in item and ')' in item:
                                    s_idx = item.rfind('(') + 1
                                    e_idx = item.rfind(')')
                                    if e_idx > s_idx:
                                        full_fba_list.append(item[s_idx:e_idx].strip())
                            FULL_FBA = ' '.join(full_fba_list).strip()

                            agg_dict = {
                                '箱数': 'sum', '数量': 'sum', '毛重KGS': 'sum', '净重KGS': 'sum',
                                '体积': 'sum', '货值': 'sum', '金额USD': 'sum',
                                'HS编码': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                fba_col: lambda x: FULL_FBA,
                                '中英文品名': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '申报要素(品牌,材质,用途)': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                'CTNS': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '币制': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '境内货源地': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '供应商': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '合同号': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique()),
                                '账号': lambda x: ' '.join(pd.Series(x).dropna().astype(str).unique())
                            }
                            agg_dict = {k: v for k, v in agg_dict.items() if k in filtered.columns}
                            df_group = filtered.groupby([prod_col, unit_col], as_index=False).agg(agg_dict)

                            for c in ['箱数', '数量', '毛重KGS', '净重KGS', '体积', '货值', '金额USD']:
                                if c in df_group.columns:
                                    df_group[c] = df_group[c].round(2)
                            if '单价' in df_group.columns:
                                df_group['单价'] = ''
                            df_group[out_addr_col] = ' '.join(grp)

                            for _, row in df_group.iterrows():
                                row_idx = ws.max_row + 1
                                for col_idx, col_name in enumerate(df_clean.columns, start=1):
                                    ws.cell(row=row_idx, column=col_idx, value=row.get(col_name, ""))

                        pbar.progress(100, text="合并完成！已生成下载文件。")
                        merged_buffer = io.BytesIO()
                        wb.save(merged_buffer)
                        wb.close()
                        merged_buffer.seek(0)

                        st.session_state['tab2_merged_result'] = {
                            'bytes': merged_buffer.getvalue(),
                            'filename': f"FBA合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            'group_count': total_grps
                        }
                    except Exception as e:
                        st.error(f"合并计算失败: {str(e)}")

            if 'tab2_merged_result' in st.session_state and st.session_state['tab2_merged_result']:
                res2 = st.session_state['tab2_merged_result']
                st.download_button(
                    label=f"📥 立即下载合并后 Excel（包含 {res2['group_count']} 组）",
                    data=res2['bytes'],
                    file_name=res2['filename'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"解析失败: {str(e)}")

# ----------------- TAB 3: 报关套打直接生成 (在线模板与智能日期) -----------------
with tab3:
    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    os.makedirs(template_dir, exist_ok=True)
    preset_templates = sorted([f for f in os.listdir(template_dir) if f.endswith(".xlsx") and not f.startswith("~$")])

    col_t3_a, col_t3_b = st.columns(2)
    tpl_bytes = None
    tpl_display_name = ""

    with col_t3_a:
        st.markdown("**1. 报关单模板选择**")
        if preset_templates:
            tpl_choice = st.selectbox(
                "选择在线模板 (.xlsx)",
                options=preset_templates + ["➕ 上传新模板到库..."],
                key="tab3_tpl_select"
            )
            if tpl_choice == "➕ 上传新模板到库...":
                uploaded_tpl = st.file_uploader("上传新模板 (.xlsx)", type=["xlsx"], key="tab3_upload_tpl")
                save_to_lib = st.checkbox("保存到模板库 (下次可直接在下拉列表选择)", value=True)
                if uploaded_tpl:
                    tpl_bytes = uploaded_tpl.getvalue()
                    tpl_display_name = uploaded_tpl.name
                    if save_to_lib:
                        save_dest = os.path.join(template_dir, uploaded_tpl.name)
                        if not os.path.exists(save_dest):
                            with open(save_dest, "wb") as f:
                                f.write(tpl_bytes)
                            st.success(f"已存入模板库: {uploaded_tpl.name}")
            else:
                tpl_path = os.path.join(template_dir, tpl_choice)
                with open(tpl_path, "rb") as f:
                    tpl_bytes = f.read()
                tpl_display_name = tpl_choice
                st.caption(f"已选用模板: `{tpl_choice}`")
        else:
            uploaded_tpl = st.file_uploader(
                "上传报关单模板文件 (.xlsx)",
                type=["xlsx"],
                key="tab3_upload_tpl",
                help="包含【输入表格】及各联报关单"
            )
            save_to_lib = st.checkbox("保存到模板库 (下次可直接在下拉列表选择，无需重复上传)", value=True)
            if uploaded_tpl:
                tpl_bytes = uploaded_tpl.getvalue()
                tpl_display_name = uploaded_tpl.name
                if save_to_lib:
                    save_dest = os.path.join(template_dir, uploaded_tpl.name)
                    with open(save_dest, "wb") as f:
                        f.write(tpl_bytes)
                    st.success("已成功保存至在线模板库！下次可直接下拉选择。")

    with col_t3_b:
        st.markdown("**2. 报关汇总数据源**")
        data_source_file = st.file_uploader(
            "上传报关汇总数据源 (.xlsx)",
            type=["xlsx"],
            key="tab3_data",
            help="包含【汇总】工作表"
        )

    with st.expander("📅 报关单日期规则配置 (自动推算合同时间与发票/装运日期)", expanded=True):
        c_date_mode, c_date_info = st.columns([1, 1])
        with c_date_mode:
            date_calc_mode = st.radio(
                "日期计算方式",
                ["自动按合同号推算（推荐）", "手动指定固定日期"],
                key="tab3_date_mode"
            )

        auto_calc_dates = (date_calc_mode == "自动按合同号推算（推荐）")
        manual_c_date = None
        manual_i_date = None

        if auto_calc_dates:
            with c_date_info:
                st.info(
                    "💡 **阿巴阿巴阿巴**：\n" +
                   
                )
        else:
            with c_date_info:
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    d_c = st.date_input("合同时间", value=datetime.now())
                    manual_c_date = d_c.strftime("%Y-%m-%d")
                with col_d2:
                    d_i = st.date_input("发票与装运日期", value=datetime.now() + timedelta(days=7))
                    manual_i_date = d_i.strftime("%Y-%m-%d")

    can_generate = (tpl_bytes is not None) and (data_source_file is not None)

    if st.button("🚀 套版生成并打包下载", type="primary", use_container_width=True, disabled=not can_generate):
        try:
            data_bytes = data_source_file.getvalue()

            header_idx = find_real_header_row_from_bytes(data_bytes, "汇总")
            df = pd.read_excel(io.BytesIO(data_bytes), sheet_name="汇总", header=header_idx, engine="openpyxl")
            df.columns = [normalize_header(c) for c in df.columns]

            required_fields = [
                "账号", "合同号", "HS编码", "中文品名", "中英文品名", "申报要素(品牌,材质,用途)",
                "箱数", "CTNS", "数量", "单位", "单价", "金额USD",
                "毛重KGS", "净重KGS", "体积", "境内货源地", "FBA编号"
            ]
            missing = [c for c in required_fields if c not in df.columns]
            if missing:
                st.error(f"缺少必要列: {', '.join(missing)}")
                st.stop()

            df = df[required_fields].dropna(how="all").reset_index(drop=True)
            fba_groups = df.groupby("FBA编号", dropna=False)

            zip_buffer = io.BytesIO()
            pbar = st.progress(0, text="正在处理...")

            total_groups = len(fba_groups)
            success_count = 0
            generated_file_details = []

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, (fba, grp) in enumerate(fba_groups, start=1):
                    final_fba = fba if (pd.notna(fba) and str(fba).strip()) else generate_unique_fba_code(idx)
                    account_val = ""
                    if "账号" in grp.columns and not grp["账号"].isnull().all():
                        first_val = grp["账号"].iloc[0]
                        account_val = str(first_val) if pd.notna(first_val) else ""

                    grp = grp.copy()
                    grp.loc[:, "FBA编号"] = final_fba

                    final_name, final_bytes, items_cnt, dates_info = fill_and_clean_template_memory(
                        template_bytes=tpl_bytes,
                        fba_data_df=grp,
                        fba_code=final_fba,
                        account_info=account_val,
                        manual_contract_date=manual_c_date,
                        manual_invoice_date=manual_i_date,
                        auto_calc_dates=auto_calc_dates
                    )

                    zip_file.writestr(final_name, final_bytes)
                    c_d, i_d = dates_info
                    date_tag = f"合同: {c_d} | 发票: {i_d}" if c_d else "原模板日期"
                    generated_file_details.append(f"{final_name} ({date_tag} | {items_cnt}品项)")
                    success_count += 1
                    pbar.progress(int(idx / total_groups * 100), text=f"[{idx}/{total_groups}] {final_name}")

            pbar.progress(100, text="完成！")
            zip_buffer.seek(0)
            st.session_state["tab3_zip"] = {
                "bytes": zip_buffer.getvalue(),
                "count": success_count,
                "files": generated_file_details
            }
        except Exception as e:
            st.error(f"处理失败: {str(e)}")

    if "tab3_zip" in st.session_state and st.session_state["tab3_zip"]:
        res3 = st.session_state["tab3_zip"]
        st.download_button(
            label=f"📦 下载全部成品报关单 ({res3['count']} 份 .zip)",
            data=res3["bytes"],
            file_name=f"成品报关单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        with st.expander(f"查看生成文件清单 ({res3['count']} 个)", expanded=True):
            for f in res3["files"]:
                st.text(f"✓ {f}")
