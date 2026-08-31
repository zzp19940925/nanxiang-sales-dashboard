# -*- coding: utf-8 -*-
"""
快速录入订单（支持问界组 + 鸿蒙组）
用法：
  add_order.py 杨志,M6,2                    # 2 条订单：姓名=杨志，车型=M6，今天日期
  add_order.py 杨志,M6,2;杨志,M8,3          # 一行录多笔
  add_order.py 杨志,M6 EV,1,2026-08-30      # 指定日期
  add_order.py 卢念念,M9,2 9月15日,智界      # 指定日期+品牌（空格/分号均可）

字段：
  姓名   销售员（必填，匹配设置 sheet 名单）
  车型   基础款名（如 M6, M8, M9 EV）
  数量   单次成交数量
  日期   默认今天（中文"X月X日"或 YYYY-MM-DD 都可）
  品牌   默认"问界"，可填：问界/尚界/享界/智界
  牌照   默认"大牌"，可选：大牌/小牌/外牌/免费公牌/免费绿牌

提成（I/J 列）：不填，下次 sync.py 自动按车型提成表算
订单号（G 列）：自动生成
"""
import sys, os, subprocess, json, datetime, base64, re

# ===== 配置 =====
SHEET_ID = "f97h5Y"          # 订单明细
FILE_ID = "MjTiKqUBgOnE"
TDOC_DIR = r"C:\Users\40543\.workbuddy\plugins\cache\workbuddy-builtin\tencent-docs-plugin\1.0.0\skills\tencent-docs"
PYTHON = r"C:\Users\40543\.workbuddy\binaries\python\versions\3.13.12\python.exe"

BRANDS = {"问界", "尚界", "享界", "智界"}
PLATES = {"大牌", "小牌", "外牌", "免费公牌", "免费绿牌"}

# ===== 工具函数 =====
def call_mcp(tool_name, arguments):
    args_json = json.dumps(arguments, ensure_ascii=False)
    cmd = [PYTHON, "tencentdocs.py", "tdoc_call", "sheet-mcp", tool_name, args_json]
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=TDOC_DIR, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"MCP 调用失败: {result.stderr[:300]}")
    resp = json.loads(result.stdout)
    if resp.get("error"):
        raise RuntimeError(f"API 返回错误: {resp['error']}")
    return resp.get("result", {}).get("structuredContent", {})


def get_last_row():
    """读订单明细最后一行（Excel 行号 1-based）"""
    sc = call_mcp("get_cell_data", {
        "file_id": FILE_ID, "sheet_id": SHEET_ID,
        "start_row": 0, "start_col": 0,
        "end_row": 499, "end_col": 29,
        "return_csv": True,
    })
    csv = sc.get("csv_data", "")
    if not csv.strip():
        return 4  # 表头 3 行 + 数据从第 4 行起
    lines = csv.split("\n")
    # 找最后一行（csv 末尾可能有空行）
    last_row = 3  # 至少第 4 行起（第 1-3 行是表头）
    for i, line in enumerate(lines, start=1):
        # 检查 A-K 列是否有非空数据
        cells = line.split(",")[:11]
        if any(c.strip() for c in cells):
            last_row = i
    return last_row


def parse_order(s):
    """解析一条订单字符串，返回 dict 或 None"""
    s = s.strip()
    if not s:
        return None
    # 支持分隔符："," 或 "，" 或 ";" 或 全角空格
    s = re.sub(r"[，；]", ",", s)
    parts = re.split(r"[,]", s)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) < 3:
        raise ValueError(f"格式错（需至少 姓名,车型,数量）: {s}")

    name = parts[0]
    car = parts[1]
    qty_str = parts[2]
    try:
        qty = int(qty_str)
    except ValueError:
        raise ValueError(f"数量不是数字: {qty_str}")
    if qty < 1 or qty > 100:
        raise ValueError(f"数量 {qty} 不合理（1-100）")

    # 默认值
    today = datetime.date.today()
    date_str = f"{today.month}月{today.day}日"
    brand = "问界"
    plate = "大牌"

    # 可选附加参数
    for extra in parts[3:]:
        extra = extra.strip()
        # 日期
        if re.match(r"\d+月\d+日", extra):
            date_str = extra
            continue
        m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", extra)
        if m:
            y, mo, d = m.groups()
            date_str = f"{int(mo)}月{int(d)}日"
            continue
        # 品牌
        if extra in BRANDS:
            brand = extra
            continue
        # 牌照
        if extra in PLATES:
            plate = extra
            continue
        # 其他忽略
        print(f"  (忽略未知参数: {extra})")

    return {
        "name": name, "car": car, "qty": qty,
        "date": date_str, "brand": brand, "plate": plate,
    }


def make_order_no():
    """生成订单号: D+年月日+随机3位"""
    today = datetime.datetime.now()
    return f"D{today.strftime('%y%m%d%H%M%S')}{today.microsecond // 1000:03d}"


def build_cells(orders, start_row):
    """根据订单列表生成 values 数组（row/col/value_type/string_value）"""
    values = []
    today_prefix = f"D{datetime.datetime.now().strftime('%y%m%d')}"
    seq = 1

    for i, o in enumerate(orders):
        r = start_row + i
        # 顺序：A日期 / B销售 / C客户 / D车型 / E品牌 / F牌照 / G订单号 / H资质 / I销售提成 / J主管提成 / K状态
        # 默认值：客户"待补"、资质"不需要"、提成空(公式重算)、状态"正常"
        order_no = f"{today_prefix}{seq:04d}"
        seq += 1

        defaults = [
            ("STRING", o["date"]),                # A 日期
            ("STRING", o["name"]),                # B 销售
            ("STRING", "待补"),                    # C 客户
            ("STRING", o["car"]),                 # D 车型
            ("STRING", o["brand"]),               # E 品牌
            ("STRING", o["plate"]),               # F 牌照
            ("STRING", order_no),                 # G 订单号
            ("STRING", "不需要"),                  # H 资质
        ]
        for j, (vt, v) in enumerate(defaults):
            values.append({
                "row": r, "col": j, "value_type": vt,
                "string_value": v,
            })
        # K 状态 = 正常
        values.append({"row": r, "col": 10, "value_type": "STRING", "string_value": "正常"})
        # I/J 提成留空（公式自动重算）
    return values


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    # 拼接所有参数，支持空格分隔多条
    raw = " ".join(sys.argv[1:])
    # 用分号拆多笔
    raw_parts = re.split(r"[;；]", raw)
    orders = []
    for rp in raw_parts:
        rp = rp.strip()
        if not rp:
            continue
        # 单笔可能含多个空格分隔的参数，但笔与笔之间用 ; 分
        o = parse_order(rp)
        if o:
            orders.append(o)

    if not orders:
        print("❌ 没有解析到任何订单")
        sys.exit(1)

    # 展开数量（每笔订单的 qty 字段拆成 qty 行）
    expanded = []
    for o in orders:
        for _ in range(o["qty"]):
            expanded.append({k: v for k, v in o.items() if k != "qty"})

    print(f"📋 待录入订单 {len(expanded)} 条：")
    for o in expanded:
        print(f"  • {o['date']:<8} {o['name']:<6} {o['car']:<8} {o['brand']:<4} {o['plate']}")
    print()

    # 找最后一行
    last_row = get_last_row()
    start_row = last_row + 1
    print(f"📍 写入位置：第 {start_row} 行（最后一行是 {last_row}）")

    # 分批写入（每批最多 50 行）
    BATCH = 50
    total = len(expanded)
    for i in range(0, total, BATCH):
        batch = expanded[i: i + BATCH]
        cells = build_cells(batch, start_row + i)
        print(f"  ✍  写入 {len(batch)} 条订单（第 {start_row + i} - {start_row + i + len(batch) - 1} 行）...")
        call_mcp("set_range_value", {
            "file_id": FILE_ID, "sheet_id": SHEET_ID, "values": cells,
        })

    print(f"\n✅ 完成！共写入 {total} 条订单到「订单明细」")
    print(f"   等下次云端自动同步（18:20 / 21:40）或本地跑 sync.py 后，看板会显示")


if __name__ == "__main__":
    main()
