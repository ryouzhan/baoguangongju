import json
import urllib.request
import urllib.error
import pandas as pd
from core.config import load_columns_config, DEFAULT_CONFIG

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
        if len(result) == 0:
            return pd.DataFrame()
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

def fetch_purchase_from_airscript(webhook_url=None, token=None, timeout=35):
    config = load_columns_config()
    air_cfg = config.get("airscript", DEFAULT_CONFIG["airscript"])
    url = webhook_url or air_cfg.get("webhook_url", DEFAULT_CONFIG["airscript"]["webhook_url"])
    tok = token or air_cfg.get("token", DEFAULT_CONFIG["airscript"]["token"])

    if not url or not tok:
        raise ValueError("缺少 AirScript Webhook 链接或脚本令牌配置！")

    headers = {
        'AirScript-Token': tok.strip(),
        'Content-Type': 'application/json; charset=utf-8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
    }
    payload = json.dumps({"Context": {"argv": {}}}).encode('utf-8')
    req = urllib.request.Request(url.strip(), data=payload, headers=headers, method='POST')

    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            status_code = response.getcode()
            resp_body = response.read().decode('utf-8')
            if status_code != 200:
                raise Exception(f"HTTP 请求失败，状态码: {status_code}")
            
            json_data = json.loads(resp_body)
            if isinstance(json_data, dict):
                if json_data.get("status") == "error":
                    msg = json_data.get("message") or json_data.get("msg") or str(json_data)
                    raise Exception(f"AirScript 执行报错: {msg}")
            
            df = parse_airscript_response(json_data)
            if df.empty:
                raise Exception("AirScript 响应成功，但返回的数据表为空！")
            return df
    except urllib.error.HTTPError as he:
        raise Exception(f"网络请求失败 [HTTP {he.code}]: {he.reason}")
    except urllib.error.URLError as ue:
        raise Exception(f"网络连接超时或无法连接: {ue.reason}")
    except Exception as e:
        raise e
