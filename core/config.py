import os
import json

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE_PATH = os.path.join(BASE_DIR, "columns_config.json")
GOODS_CONFIG_PATH = os.path.join(BASE_DIR, "goods_config.json")

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

def load_columns_config():
    if not os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"创建默认配置文件失败: {e}")
        return DEFAULT_CONFIG

    try:
        with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            if "airscript" not in cfg:
                cfg["airscript"] = DEFAULT_CONFIG["airscript"]
            if "fba_merger" not in cfg:
                cfg["fba_merger"] = DEFAULT_CONFIG["fba_merger"]
            if "customs_doc_generator" not in cfg:
                cfg["customs_doc_generator"] = DEFAULT_CONFIG["customs_doc_generator"]
            return cfg
    except Exception as e:
        print(f"读取配置文件失败，使用默认配置: {e}")
        return DEFAULT_CONFIG

def save_columns_config(cfg):
    with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return cfg

def load_exchange_rate():
    if os.path.exists(GOODS_CONFIG_PATH):
        try:
            with open(GOODS_CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return float(data.get("exchange_rate", 7.2))
        except Exception:
            pass
    return 7.2

def save_exchange_rate(rate: float):
    try:
        with open(GOODS_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"exchange_rate": rate}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存汇率失败: {e}")
