import os
from datetime import datetime
import numpy as np
import pandas as pd
import xlsxwriter
from core.config import load_columns_config, DEFAULT_CONFIG
from core.airscript import fetch_purchase_from_airscript

def generate_customs_document(delivery_file_path, output_dir, exchange_rate=7.2, df_purchase_override=None, status_logger=None):
    def log(msg):
        if status_logger:
            status_logger(msg)

    log(f"开始生成报关资料，汇率：1 USD = {exchange_rate} CNY")
    
    config = load_columns_config().get("customs_doc_generator", DEFAULT_CONFIG["customs_doc_generator"])
    p_sku = config.get("purchase_sku", "SKU")
    p_price = config.get("purchase_price", "采购单价")
    p_name = config.get("purchase_name", "品名")
    p_unit = config.get("purchase_unit", "单位")
    p_hs = config.get("purchase_hs", "HS编码")
    p_en = config.get("purchase_en_name", "中英文品名")
    p_elem = config.get("purchase_elements", "申报要素(品牌,材质,用途)")
    p_origin = config.get("purchase_origin", "境内货源地")

    d_sku = config.get("delivery_sku", "SKU")
    d_qty = config.get("delivery_qty", "发货量")
    d_ctns = config.get("delivery_ctns", "箱数")
    d_box_qty = config.get("delivery_box_qty", "单箱数量")
    d_center = config.get("delivery_center", "物流中心编码")
    d_supp = config.get("delivery_supplier", "供应商")
    d_ship_id = config.get("delivery_shipment_id", "货件编号")
    d_gw = config.get("delivery_gross_weight", "外箱总重量(kg)")
    d_vol = config.get("delivery_volume", "外箱总体积(m³)")

    if df_purchase_override is not None:
        log("使用指定的采购单数据...")
        purchase = df_purchase_override.copy()
    else:
        log("正在通过 AirScript Webhook 获取云端采购单...")
        purchase = fetch_purchase_from_airscript()
        log(f"成功获取云端采购单，共 {len(purchase)} 行记录")

    purchase_cols = purchase.columns.tolist()
    for req_col in [p_sku, p_price, p_name, p_unit, p_hs, p_en, p_elem, p_origin]:
        if req_col not in purchase_cols:
            purchase[req_col] = "" if req_col != p_price else 0

    purchase[p_price] = pd.to_numeric(purchase[p_price], errors='coerce').fillna(0)
    purchase = purchase.fillna({
        p_hs: "", p_en: "", p_elem: "", p_origin: "",
        p_name: "", p_unit: ""
    })
    purchase = purchase.dropna(subset=[p_sku])
    if purchase.empty:
        raise ValueError("采购单中无有效 SKU 数据！")

    sku_price_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_price]))
    sku_name_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_name]))
    sku_unit_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_unit]))
    sku_hs_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_hs]))
    sku_en_name_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_en]))
    sku_declare_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_elem]))
    sku_origin_dict = dict(zip(purchase[p_sku].astype(str).str.strip(), purchase[p_origin]))

    log(f"采购单有效 SKU 数量：{len(sku_price_dict)} 个")

    log("读取本地发货单数据...")
    delivery = pd.read_excel(delivery_file_path, usecols=[
        d_sku, d_qty, d_ctns, d_box_qty, d_center, d_supp, d_ship_id, d_gw, d_vol
    ])
    delivery = delivery.fillna({
        d_qty: 0, d_ctns: 0, d_gw: 0, d_vol: 0,
        d_center: "", d_supp: "", d_ship_id: "", d_box_qty: 0
    })
    delivery = delivery.dropna(subset=[d_sku])
    if delivery.empty:
        raise ValueError("发货单中无有效 SKU 数据！")
    log(f"发货单有效记录：{len(delivery)} 条")

    log("生成账号前缀（SKU前两位大写）并匹配字段...")
    delivery["账号"] = delivery[d_sku].astype(str).str[:2].str.upper().fillna("")
    delivery_sku_clean = delivery[d_sku].astype(str).str.strip()

    delivery["采购单价"] = delivery_sku_clean.map(sku_price_dict).fillna(0)
    delivery["品名"] = delivery_sku_clean.map(sku_name_dict).fillna("【无采购单匹配】")
    delivery["单位"] = delivery_sku_clean.map(sku_unit_dict).fillna("")
    delivery["HS编码"] = delivery_sku_clean.map(sku_hs_dict).fillna("")
    delivery["中英文品名"] = delivery_sku_clean.map(sku_en_name_dict).fillna("【无采购单匹配】")
    delivery["申报要素(品牌,材质,用途)"] = delivery_sku_clean.map(sku_declare_dict).fillna("【无采购单匹配，需手动录入】")
    delivery["境内货源地"] = delivery_sku_clean.map(sku_origin_dict).fillna("")

    delivery["箱数"] = delivery[d_ctns]
    delivery["发货量"] = delivery[d_qty]
    delivery["外箱总重量(kg)"] = delivery[d_gw]
    delivery["外箱总体积(m³)"] = delivery[d_vol]
    delivery["物流中心编码"] = delivery[d_center]
    delivery["供应商"] = delivery[d_supp]
    delivery["货件编号"] = delivery[d_ship_id]
    delivery["SKU"] = delivery[d_sku]

    delivery["货值"] = (delivery["发货量"] * delivery["采购单价"]).round(2).fillna(0)

    delivery["发货量"] = delivery["发货量"].astype(int)
    delivery["金额USD"] = (delivery["货值"] / exchange_rate).round(2)
    delivery["发货量"] = delivery["发货量"].replace(0, 1)
    delivery["单价(USD)"] = (delivery["金额USD"] / delivery["发货量"]).round(2)
    delivery["金额USD"] = (delivery["发货量"] * delivery["单价(USD)"]).round(2)

    delivery["净重(kg)"] = (delivery["外箱总重量(kg)"] * 0.95).round(2).fillna(0)
    delivery["外箱总体积(m³)"] = delivery["外箱总体积(m³)"].round(2).fillna(0)

    unmatch_count = int((delivery["采购单价"] == 0).sum())
    detail = delivery.copy()
    if unmatch_count > 0:
        log(f"⚠️ 发现 {unmatch_count} 条采购单无匹配的 SKU（已保留在明细中红字标注）")
    else:
        log("✅ 所有 SKU 均成功匹配采购单！")

    detail = detail.sort_values("货件编号").reset_index(drop=True)

    summary = detail[detail["采购单价"] > 0].groupby(
        by=["货件编号", "供应商", "品名", "账号"], as_index=False, dropna=False
    ).agg({
        "箱数": "sum",
        "发货量": "sum",
        "外箱总重量(kg)": "sum",
        "净重(kg)": "sum",
        "外箱总体积(m³)": "sum",
        "货值": "sum",
        "物流中心编码": "first",
        "金额USD": "sum",
        "单位": "first",
        "HS编码": "first",
        "中英文品名": "first",
        "申报要素(品牌,材质,用途)": "first",
        "境内货源地": "first"
    }).round(2)

    summary["单价(USD)"] = (summary["金额USD"] / summary["发货量"]).round(2)
    summary["金额USD"] = (summary["发货量"] * summary["单价(USD)"]).round(2)

    summary = summary.fillna({
        "HS编码": "", "中英文品名": "", "申报要素(品牌,材质,用途)": "", "境内货源地": "",
        "单价(USD)": 0, "金额USD": 0, "货值": 0, "单位": ""
    })
    summary["单价(USD)"] = summary["单价(USD)"].replace([np.inf, -np.inf], 0)
    detail["单价(USD)"] = detail["单价(USD)"].replace([np.inf, -np.inf], 0)

    valid_detail = detail[detail["采购单价"] > 0]
    total_amt_cny = float(valid_detail["货值"].sum())
    total_amt_usd = float(valid_detail["金额USD"].sum())
    total_gw = float(valid_detail["外箱总重量(kg)"].sum())
    total_nw = float(valid_detail["净重(kg)"].sum())
    total_cbm = float(valid_detail["外箱总体积(m³)"].sum())
    total_fba = int(summary["货件编号"].nunique()) if not summary.empty else 0
    total_supplier = int(summary["供应商"].nunique()) if not summary.empty else 0
    total_product = int(summary["品名"].nunique()) if not summary.empty else 0

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_filename = f"报关资料_{timestamp}.xlsx"
    out_filepath = os.path.join(output_dir, out_filename)

    log("正在写入 Excel 文件并应用样式...")
    _write_customs_excel(detail, summary, out_filepath, exchange_rate)
    log(f"🎉 报关资料生成完成: {out_filename}")

    stats = {
        "filename": out_filename,
        "detail_count": len(detail),
        "summary_count": len(summary),
        "unmatch_count": unmatch_count,
        "total_cny": round(total_amt_cny, 2),
        "total_usd": round(total_amt_usd, 2),
        "total_gw": round(total_gw, 2),
        "total_nw": round(total_nw, 2),
        "total_cbm": round(total_cbm, 2),
        "total_fba": total_fba,
        "total_supplier": total_supplier,
        "total_product": total_product
    }
    return out_filepath, stats

def _write_customs_excel(detail_df, summary_df, save_path, exchange_rate):
    workbook = xlsxwriter.Workbook(save_path)

    header_fmt = workbook.add_format({
        "font_name": "微软雅黑", "font_size": 11, "bold": True,
        "font_color": "white", "bg_color": "#1f2937",
        "align": "center", "valign": "vcenter", "border": 1
    })
    title_fmt = workbook.add_format({
        "font_name": "微软雅黑", "font_size": 13, "bold": True,
        "align": "center", "valign": "vcenter", "bg_color": "#f3f4f6"
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
        "font_color": "#dc2626", "bold": True, "bg_color": "#fee2e2"
    })

    color1 = workbook.add_format({
        "bg_color": "#ffffff", "border": 1,
        "align": "center", "valign": "vcenter",
        "font_name": "微软雅黑", "font_size": 10
    })
    color2 = workbook.add_format({
        "bg_color": "#f0f9ff", "border": 1,
        "align": "center", "valign": "vcenter",
        "font_name": "微软雅黑", "font_size": 10
    })

    headers = [
        "账号", "HS编码", "中文品名", "中英文品名", "申报要素(品牌,材质,用途)",
        "箱数", "CTNS", "数量", "单位", "单价", "币制", "金额USD",
        "毛重KGS", "净重KGS", "体积", "境内货源地",
        "FBA编号", "合同号", "物流中心编码", "货值", "供应商", "SKU", "采购单价(元)"
    ]
    col_widths = [
        8, 12, 20, 25, 30, 8, 8, 10, 8, 10, 8, 12,
        10, 10, 12, 15, 15, 12, 20, 12, 15, 15, 12
    ]

    ws_detail = workbook.add_worksheet("明细")
    ws_detail.merge_range("A1:W1", f"报关明细数据（含所有SKU，汇率：1美元 = {exchange_rate} 人民币 | 无尾差）", title_fmt)
    ws_detail.set_row(0, 30)
    ws_detail.set_row(1, 24)

    for col, header in enumerate(headers):
        ws_detail.write(1, col, header, header_fmt)

    for col, width in enumerate(col_widths):
        ws_detail.set_column(col, col, width)

    for row_idx, (_, row) in enumerate(detail_df.iterrows(), 2):
        is_unmatch = row["采购单价"] == 0
        base_fmt = unmatch_fmt if is_unmatch else text_fmt
        num_fmt = unmatch_fmt if is_unmatch else int_fmt
        float_fmt = unmatch_fmt if is_unmatch else float2_fmt

        ws_detail.set_row(row_idx, 20)
        ws_detail.write(row_idx, 0, str(row["账号"]), base_fmt)
        ws_detail.write(row_idx, 1, str(row["HS编码"]), base_fmt)
        ws_detail.write(row_idx, 2, str(row["品名"]), base_fmt)
        ws_detail.write(row_idx, 3, str(row["中英文品名"]), base_fmt)
        ws_detail.write(row_idx, 4, str(row["申报要素(品牌,材质,用途)"]), base_fmt)
        ws_detail.write(row_idx, 5, row["箱数"], num_fmt)
        ws_detail.write(row_idx, 6, "CTNS", base_fmt)
        ws_detail.write(row_idx, 7, row["发货量"], num_fmt)
        ws_detail.write(row_idx, 8, str(row["单位"]), base_fmt)
        ws_detail.write(row_idx, 9, row["单价(USD)"], float_fmt)
        ws_detail.write(row_idx, 10, "USD", base_fmt)
        ws_detail.write(row_idx, 11, row["金额USD"], float_fmt)
        ws_detail.write(row_idx, 12, row["外箱总重量(kg)"], float_fmt)
        ws_detail.write(row_idx, 13, row["净重(kg)"], float_fmt)
        ws_detail.write(row_idx, 14, row["外箱总体积(m³)"], float_fmt)
        ws_detail.write(row_idx, 15, str(row["境内货源地"]), base_fmt)
        ws_detail.write(row_idx, 16, str(row["货件编号"]), base_fmt)
        ws_detail.write(row_idx, 17, "", base_fmt)
        ws_detail.write(row_idx, 18, str(row["物流中心编码"]), base_fmt)
        ws_detail.write(row_idx, 19, row["货值"], float_fmt)
        ws_detail.write(row_idx, 20, str(row["供应商"]), base_fmt)
        ws_detail.write(row_idx, 21, str(row["SKU"]), base_fmt)
        ws_detail.write(row_idx, 22, row["采购单价"], float_fmt)

    ws_summary = workbook.add_worksheet("汇总")
    ws_summary.merge_range("A1:W1", f"报关汇总数据（仅采购单匹配记录，汇率：1美元 = {exchange_rate} 人民币 | 无尾差）", title_fmt)
    ws_summary.set_row(0, 30)
    ws_summary.set_row(1, 24)

    for col, header in enumerate(headers):
        ws_summary.write(1, col, header, header_fmt)

    for col, width in enumerate(col_widths):
        ws_summary.set_column(col, col, width)

    fba_unique = summary_df["货件编号"].unique()
    fba_color_map = {fba: color1 if i % 2 == 0 else color2 for i, fba in enumerate(fba_unique)}

    for row_idx, (_, row) in enumerate(summary_df.iterrows(), 2):
        fba_code = row["货件编号"]
        row_fmt = fba_color_map.get(fba_code, color1)

        ws_summary.set_row(row_idx, 20)
        ws_summary.write(row_idx, 0, str(row["账号"]), row_fmt)
        ws_summary.write(row_idx, 1, str(row["HS编码"]), row_fmt)
        ws_summary.write(row_idx, 2, str(row["品名"]), row_fmt)
        ws_summary.write(row_idx, 3, str(row["中英文品名"]), row_fmt)
        ws_summary.write(row_idx, 4, str(row["申报要素(品牌,材质,用途)"]), row_fmt)
        ws_summary.write(row_idx, 5, row["箱数"], row_fmt)
        ws_summary.write(row_idx, 6, "CTNS", row_fmt)
        ws_summary.write(row_idx, 7, row["发货量"], row_fmt)
        ws_summary.write(row_idx, 8, str(row["单位"]), row_fmt)
        ws_summary.write(row_idx, 9, row["单价(USD)"], row_fmt)
        ws_summary.write(row_idx, 10, "USD", row_fmt)
        ws_summary.write(row_idx, 11, row["金额USD"], row_fmt)
        ws_summary.write(row_idx, 12, row["外箱总重量(kg)"], row_fmt)
        ws_summary.write(row_idx, 13, row["净重(kg)"], row_fmt)
        ws_summary.write(row_idx, 14, row["外箱总体积(m³)"], row_fmt)
        ws_summary.write(row_idx, 15, str(row["境内货源地"]), row_fmt)
        ws_summary.write(row_idx, 16, str(row["货件编号"]), row_fmt)
        ws_summary.write(row_idx, 17, "", row_fmt)
        ws_summary.write(row_idx, 18, str(row["物流中心编码"]), row_fmt)
        ws_summary.write(row_idx, 19, row["货值"], row_fmt)
        ws_summary.write(row_idx, 20, str(row["供应商"]), row_fmt)
        ws_summary.write(row_idx, 21, "汇总", row_fmt)
        ws_summary.write(row_idx, 22, "", row_fmt)

    workbook.close()
