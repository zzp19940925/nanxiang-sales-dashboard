# -*- coding: utf-8 -*-
"""
腾讯文档在线表格读取模块
替代 pd.read_excel()，从腾讯文档在线表格读取数据

支持两种模式：
1. 本地 CLI 模式（默认）：通过 WorkBuddy 连接器调用 tencentdocs.py
2. HTTP API 模式：设置环境变量 TDOC_ACCESS_TOKEN 后，直接调用腾讯文档 MCP API
   适用于 GitHub Actions / 云函数等无法安装 WorkBuddy 的环境
"""
import json
import subprocess
import io
import os
import ssl
import time
import urllib.request
import urllib.parse
import urllib.error
import pandas as pd

# 创建宽松的 SSL 上下文（解决 Windows 上 Python SSL 握手问题）
_ssl_ctx = ssl.create_default_context()
_ssl_ctx.set_ciphers("DEFAULT@SECLEVEL=1")

# ===== 配置 =====
FILE_ID = "MjTiKqUBgOnE"

# 本地 CLI 模式路径
TDOC_DIR = r"C:\Users\40543\.workbuddy\plugins\cache\workbuddy-builtin\tencent-docs-plugin\1.0.0\skills\tencent-docs"
PYTHON = r"C:\Users\40543\.workbuddy\binaries\python\versions\3.13.12\python.exe"

# HTTP API 模式配置
MCP_URL = "https://docs.qq.com/openapi/mcp"
TDOC_CLIENT_ID = os.environ.get("TDOC_CLIENT_ID", "58c50de50c51409d9e05014b48afdbad")
TDOC_ACCESS_TOKEN = os.environ.get("TDOC_ACCESS_TOKEN", "")
TDOC_REFRESH_TOKEN = os.environ.get("TDOC_REFRESH_TOKEN", "")

# 是否使用 HTTP API 模式
USE_HTTP_MODE = bool(TDOC_ACCESS_TOKEN)

# 子表名 -> sheet_id 映射
SHEET_MAP = {
    "设置": "BB08J2",
    "7月销售数据": "m8qH7r",
    "8月销售数据": "H9Ziys",
    "9月销售数据": "L4yocL",
    "10月销售数据": "Z7GIdY",
    "11月销售数据": "dus6ez",
    "12月销售数据": "dDVTwr",
    "订单明细": "f97h5Y",
    "使用说明": "1Bf9Rt",
}

# 请求范围（500行×30列=15000单元格，在20000上限内）
MAX_ROW = 499
MAX_COL = 29

# CSV 缓存（同一次运行内每个子表只请求一次）
_csv_cache = {}


# ===== HTTP API 模式 =====

def _refresh_access_token():
    """使用 refresh_token 刷新 access_token"""
    if not TDOC_REFRESH_TOKEN:
        return None

    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "client_id": TDOC_CLIENT_ID,
        "refresh_token": TDOC_REFRESH_TOKEN,
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://docs.qq.com/openapi/mcp/oauth/token",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )

    try:
        resp = urllib.request.urlopen(req, timeout=15, context=_ssl_ctx)
        result = json.loads(resp.read().decode("utf-8"))
        new_token = result.get("access_token")
        if new_token:
            # 输出到 GitHub Actions 环境文件
            env_file = os.environ.get("GITHUB_ENV")
            if env_file and os.path.exists(os.path.dirname(env_file)):
                with open(env_file, "a") as f:
                    f.write(f"TDOC_ACCESS_TOKEN={new_token}\n")
                    if "refresh_token" in result:
                        f.write(f"TDOC_REFRESH_TOKEN={result['refresh_token']}\n")
            return new_token
    except Exception as e:
        print(f"  ⚠ Token 刷新失败: {e}")
    return None


def _mcp_call(tool_name, arguments):
    """调用 MCP 工具"""
    token = TDOC_ACCESS_TOKEN
    if not token:
        raise RuntimeError("HTTP API 模式需要 TDOC_ACCESS_TOKEN 环境变量")

    data = json.dumps({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": tool_name,
            "arguments": arguments,
        },
    }).encode("utf-8")

    req = urllib.request.Request(
        MCP_URL,
        data=data,
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        },
        method="POST",
    )

    try:
        resp = urllib.request.urlopen(req, timeout=120, context=_ssl_ctx)
        body = resp.read().decode("utf-8")
        result = json.loads(body)
        sc = result.get("result", {}).get("structuredContent", {})
        if sc.get("error"):
            raise RuntimeError(f"MCP 工具返回错误: {sc['error']}")
        return sc
    except urllib.error.HTTPError as e:
        if e.code == 401:
            # Token 过期，尝试刷新
            new_token = _refresh_access_token()
            if new_token:
                # 用新 token 重试
                req.add_header("Authorization", "Bearer " + new_token)
                resp = urllib.request.urlopen(req, timeout=120, context=_ssl_ctx)
                body = resp.read().decode("utf-8")
                result = json.loads(body)
                sc = result.get("result", {}).get("structuredContent", {})
                return sc
        raise RuntimeError(f"HTTP {e.code}: {e.read().decode('utf-8')[:200]}")
    except urllib.error.URLError as e:
        # SSL/网络错误，重试一次
        if "SSL" in str(e) or "EOF" in str(e):
            time.sleep(2)
            resp = urllib.request.urlopen(req, timeout=120, context=_ssl_ctx)
            body = resp.read().decode("utf-8")
            result = json.loads(body)
            sc = result.get("result", {}).get("structuredContent", {})
            return sc
        raise


def _fetch_csv_http(sheet_name):
    """HTTP API 模式：通过 MCP 获取子表 CSV 数据（带重试）"""
    sheet_id = SHEET_MAP[sheet_name]
    last_err = None
    for attempt in range(3):
        try:
            sc = _mcp_call("sheet.get_cell_data", {
                "file_id": FILE_ID,
                "sheet_id": sheet_id,
                "start_row": 0,
                "start_col": 0,
                "end_row": MAX_ROW,
                "end_col": MAX_COL,
                "return_csv": True,
            })
            return sc.get("csv_data", "")
        except Exception as e:
            last_err = str(e)
            if attempt < 2:
                time.sleep(1.5 ** attempt)
    raise RuntimeError(f"HTTP API 调用失败（已重试3次），子表「{sheet_name}」: {last_err}")


# ===== 本地 CLI 模式 =====

def _fetch_csv_local(sheet_name):
    """本地 CLI 模式：通过 tencentdocs.py 获取子表 CSV 数据（带重试）"""
    sheet_id = SHEET_MAP[sheet_name]
    args = json.dumps({
        "file_id": FILE_ID,
        "sheet_id": sheet_id,
        "start_row": 0,
        "start_col": 0,
        "end_row": MAX_ROW,
        "end_col": MAX_COL,
        "return_csv": True,
    }, ensure_ascii=False)

    cmd = [PYTHON, "tencentdocs.py", "tdoc_call", "sheet-mcp", "get_cell_data", args]

    last_err = None
    for attempt in range(3):
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=TDOC_DIR, timeout=120)
        except subprocess.TimeoutExpired as e:
            last_err = f"请求超时（{e.timeout}秒）"
            time.sleep(1.5 ** attempt)
            continue

        if result.returncode != 0:
            last_err = result.stderr[:500] if result.stderr else "(无错误输出)"
            time.sleep(1.5 ** attempt)
            continue

        try:
            resp = json.loads(result.stdout)
        except json.JSONDecodeError as e:
            last_err = f"返回不是有效 JSON: {e}"
            time.sleep(1.5 ** attempt)
            continue

        if resp.get("error"):
            last_err = f"API 返回错误: {resp.get('error')}"
            time.sleep(1.5 ** attempt)
            continue

        content = resp.get("result", {}).get("structuredContent", {})
        csv_data = content.get("csv_data", "")
        return csv_data

    raise RuntimeError(
        f"腾讯文档 API 调用失败（已重试 3 次），子表「{sheet_name}」: {last_err}"
    )


# ===== 统一接口 =====

def _fetch_csv(sheet_name):
    """从腾讯文档 API 获取子表的 CSV 数据（带缓存）"""
    if sheet_name in _csv_cache:
        return _csv_cache[sheet_name]

    if sheet_name not in SHEET_MAP:
        raise ValueError(f"未知的子表名: {sheet_name}")

    if USE_HTTP_MODE:
        csv_data = _fetch_csv_http(sheet_name)
    else:
        csv_data = _fetch_csv_local(sheet_name)

    _csv_cache[sheet_name] = csv_data
    return csv_data


def read_excel(file_path, sheet_name, header=None, engine=None):
    """
    替代 pd.read_excel()，从腾讯文档在线表格读取数据。
    参数兼容 pd.read_excel()，file_path 和 engine 被忽略。
    """
    csv_data = _fetch_csv(sheet_name)
    if not csv_data or not csv_data.strip():
        return pd.DataFrame()

    df = pd.read_csv(io.StringIO(csv_data), header=header)
    return df


class ExcelFile:
    """模拟 pd.ExcelFile，提供 sheet_names 属性"""
    sheet_names = list(SHEET_MAP.keys())


def clear_cache():
    """清除缓存（下次读取时重新从 API 获取）"""
    _csv_cache.clear()
