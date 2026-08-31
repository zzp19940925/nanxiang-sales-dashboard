# -*- coding: utf-8 -*-
"""
同步脚本：从腾讯文档在线表格读取数据，生成 data.js + orders_data.js + customer_data.js
数据源：腾讯文档「董逸亭南翔订单」(file_id: MjTiKqUBgOnE)
用法：双击「一键同步.bat」或 python sync.py
"""
import pandas as pd
import json
import os
import sys
import re
from datetime import datetime
import tdoc_reader

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
os.chdir(SCRIPT_DIR)


def read_weights():
    """从设置sheet动态读取积分权重（H5-H10）
    不再硬编码，表格修改权重后看板自动同步"""
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    w = {}
    # H5=电话, H6=线索, H7=邀约, H8=试驾, H9=小订, H10=订单
    labels = ["电话", "线索", "邀约", "试驾", "小订", "订单"]
    for i, label in enumerate(labels):
        val = df.iloc[4 + i, 7]  # 行5-10(0-indexed 4-9), H列(index 7)
        w[label] = float(val) if pd.notna(val) else 0
    return w


def read_holidays():
    """从设置sheet读取法定节假日列表（A86-A110）
    返回日期字符串集合（YYYY-MM-DD格式）"""
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    holidays = set()
    for i in range(85, 110):  # 0-indexed 85-109 = Excel行86-110
        val = df.iloc[i, 0]
        if pd.isna(val) or str(val).strip() == "":
            continue
        if isinstance(val, str):
            val = pd.to_datetime(val, errors="coerce")
        if pd.notna(val):
            holidays.add(val.strftime("%Y-%m-%d"))
    return holidays


# 积分权重（动态读取）
WEIGHTS = read_weights()
# 法定节假日列表（动态读取）
HOLIDAYS = read_holidays()


def read_sales_info():
    """从设置sheet动态读取销售人员信息（行5-20，预留空白行）
    新增销售时直接在表格填写，无需改代码"""
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    info = {}
    wj_members = []
    hm_members = []
    for i in range(4, 20):  # 0-indexed 行4-19 = Excel行5-20
        name = df.iloc[i, 0]
        if pd.isna(name) or str(name).strip() == "":
            continue
        name = str(name).strip()
        group = str(df.iloc[i, 2]).strip() if pd.notna(df.iloc[i, 2]) else ""
        manager = str(df.iloc[i, 3]).strip() if pd.notna(df.iloc[i, 3]) else ""
        status = str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else "在职"
        if status == "离职":
            continue  # 跳过离职人员
        info[name] = {"group": group, "manager": manager}
        if group == "问界组":
            wj_members.append(name)
        elif group == "鸿蒙组":
            hm_members.append(name)
    return info, wj_members, hm_members


# 启动时动态读取（不再硬编码）
SALES_INFO, WENJIE_MEMBERS, HM_MEMBERS = read_sales_info()


def read_workload():
    """读取所有月份销售数据sheet的工作量（7月-12月销售数据）
    自动识别 'X月销售数据' 命名的sheet并合并读取"""
    xl = tdoc_reader.ExcelFile()
    month_sheets = [s for s in xl.sheet_names if re.match(r'^\d+月销售数据$', s)]
    month_sheets.sort(key=lambda x: int(re.match(r'(\d+)', x).group(1)))
    records = []
    for sheet_name in month_sheets:
        df = tdoc_reader.read_excel(None, sheet_name=sheet_name, header=2)
        # 列: 月份/日期/销售/组别/主管/电话量/线索量/邀约量/试驾量/小订量/订单量/综合积分/备注
        for _, row in df.iterrows():
            date = row.get("日期")
            name = row.get("销售")
            if pd.isna(date) or pd.isna(name):
                continue
            name = str(name).strip()
            if name not in SALES_INFO:
                continue
            phone = row.get("电话量")
            if pd.isna(phone) or str(phone).strip() == "":
                continue  # 跳过没数据的行
            # 解析日期（支持标准格式和中文"X月Y日"格式）
            if isinstance(date, str):
                date_str = date
                date = pd.to_datetime(date_str, errors="coerce")
                if pd.isna(date):
                    # 尝试解析中文日期格式 "8月1日" → 2026-08-01
                    m = re.match(r'(\d+)月(\d+)日', date_str)
                    if m:
                        date = pd.Timestamp(year=2026, month=int(m.group(1)), day=int(m.group(2)))
            if pd.isna(date):
                continue
            info = SALES_INFO[name]
            # 提取工作量
            p = int(phone) if pd.notna(phone) else 0
            l = int(row.get("线索量", 0)) if pd.notna(row.get("线索量")) else 0
            iv = int(row.get("邀约量", 0)) if pd.notna(row.get("邀约量")) else 0
            ts = int(row.get("试驾量", 0)) if pd.notna(row.get("试驾量")) else 0
            sm = int(row.get("小订量", 0)) if pd.notna(row.get("小订量")) else 0
            od = int(row.get("订单量", 0)) if pd.notna(row.get("订单量")) else 0
            # 积分计算：工作日(周一至周五且非节假日)电话积分封顶3分，周末/节假日电话不计分
            is_weekday = date.weekday() < 5 and date.strftime("%Y-%m-%d") not in HOLIDAYS
            phone_score = min(p * WEIGHTS["电话"], 3) if is_weekday else 0
            score = round(phone_score + l * WEIGHTS["线索"] + iv * WEIGHTS["邀约"] +
                          ts * WEIGHTS["试驾"] + sm * WEIGHTS["小订"] + od * WEIGHTS["订单"], 1)
            records.append({
                "date": date.strftime("%Y-%m-%d"),
                "name": name,
                "group": info["group"],
                "manager": info["manager"],
                "phone": p,
                "lead": l,
                "invite": iv,
                "test": ts,
                "small": sm,
                "order": od,
                "score": score,
            })
    return records


def read_car_commission():
    """读取车型提成表（设置sheet，动态查找"车型提成表"标题行）"""
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    # 动态查找"车型提成表"标题行
    title_row = None
    for i in range(df.shape[0]):
        if df.iloc[i, 0] == "车型提成表":
            title_row = i
            break
    if title_row is None:
        return {}
    comm = {}
    # 数据从 title_row+2 开始（跳过标题行和表头行）
    for i in range(title_row + 2, title_row + 42):  # 最多读40行
        car = df.iloc[i, 0]
        if pd.isna(car) or str(car).strip() == "":
            continue
        # 遇到下一个区块标题就停止（含8月提成表、月度订单目标、客户分级标准）
        car_str = str(car).strip()
        if car_str in ["月度订单目标（问界组）", "客户分级标准"] or "8月提成表" in car_str:
            break
        car = str(car).strip()
        comm[car] = {
            "brand": str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else "",
            "sale_comm": int(df.iloc[i, 2]) if pd.notna(df.iloc[i, 2]) else 0,
            "mgr_comm": int(df.iloc[i, 3]) if pd.notna(df.iloc[i, 3]) else 0,
        }
    return comm


def normalize_car_name(car):
    """车型归一化：去掉 EV/EVR 后缀，映射到基础车型名
    例如：M9 EV → M9, M9 EVR → M9, S9T EVR → S9T
    """
    if not car:
        return car
    car = str(car).strip()
    for suffix in [" EVR", " EV"]:
        if car.endswith(suffix):
            car = car[:-len(suffix)]
            break
    return car


def read_august_commission():
    """读取8月提成表（设置sheet行54-69）
    8月新政：销售提成=大定+交付（留车已删除），主管提成单独
    销售不乘完成率系数，主管保持系数
    支持订单明细中的 EV/EVR 细分车型（归一化到基础车型查表）
    """
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    # 查找"8月提成表"标题行
    title_row = None
    for i in range(df.shape[0]):
        if df.iloc[i, 0] is not None and "8月提成表" in str(df.iloc[i, 0]):
            title_row = i
            break
    if title_row is None:
        return {}
    comm = {}
    # 数据从 title_row+2 起（跳过标题行+表头行）
    for i in range(title_row + 2, title_row + 20):  # 14款车
        car = df.iloc[i, 0]
        if pd.isna(car) or str(car).strip() == "":
            continue
        car = str(car).strip()
        # 直接读C列(大定)和D列(交付)求和，不依赖Excel公式缓存
        dading = int(df.iloc[i, 2]) if pd.notna(df.iloc[i, 2]) else 0
        jiaofu = int(df.iloc[i, 3]) if pd.notna(df.iloc[i, 3]) else 0
        sale_total = dading + jiaofu  # 销售提成=大定+交付
        mgr = int(df.iloc[i, 5]) if pd.notna(df.iloc[i, 5]) else 0
        comm[car] = {
            "brand": str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else "",
            "sale_comm": sale_total,  # 销售提成=大定+交付
            "mgr_comm": mgr,
        }
    return comm


def read_orders(car_comm, august_comm=None):
    """读取订单明细sheet
    8月订单使用8月提成表（销售=大定+交付+留车，主管单独），其他月份用原表"""
    df = tdoc_reader.read_excel(None, sheet_name="订单明细", header=2)
    august_comm = august_comm or {}
    records = []
    for _, row in df.iterrows():
        date = row.get("日期")
        name = row.get("销售")
        if pd.isna(date) or pd.isna(name):
            continue
        name = str(name).strip()
        # 不过滤销售人员，保留所有历史订单（含已离职人员如董逸亭）
        car = str(row.get("车型", "")).strip() if pd.notna(row.get("车型")) else ""
        if not car:
            continue
        # 解析日期（支持标准格式和中文"X月Y日"格式）
        if isinstance(date, str):
            date_str = date.strip()
            date = pd.to_datetime(date_str, errors="coerce")
            if pd.isna(date):
                # 尝试解析中文日期格式 "6月1日" → 2026-06-01
                m = re.match(r'(\d+)月(\d+)日', date_str)
                if m:
                    month, day = int(m.group(1)), int(m.group(2))
                    date = pd.Timestamp(year=2026, month=month, day=day)
        if pd.isna(date):
            continue
        month_str = f"{date.month}月"
        is_august = month_str == "8月"
        # 提成：优先用表格里的值（Excel缓存），否则从车型提成表匹配
        sale_comm = row.get("销售提成")
        mgr_comm = row.get("主管提成")
        # 8月订单：归一化车型名（去掉 EV/EVR 后缀）去查8月提成表
        lookup_car = normalize_car_name(car) if is_august else car
        if pd.isna(sale_comm) or sale_comm == "" or (is_august and august_comm and lookup_car in august_comm):
            # 8月订单：始终用8月提成表（因为Excel可能未刷新缓存）
            if is_august and august_comm and lookup_car in august_comm:
                cc = august_comm[lookup_car]
                sale_comm = cc["sale_comm"]
                mgr_comm = cc["mgr_comm"]
            else:
                cc = car_comm.get(car, {})
                sale_comm = cc.get("sale_comm", 0)
                mgr_comm = cc.get("mgr_comm", 0)
        # 品牌：8月用归一化后的车型查品牌，其他月份用原车型
        brand = car_comm.get(car, {}).get("brand", "")
        if is_august and lookup_car in august_comm:
            brand = august_comm[lookup_car].get("brand", brand)
        if not brand:
            brand = str(row.get("品牌", "")) if pd.notna(row.get("品牌")) else ""
        # 读取L列状态：空值当"正常"处理（容错），退订/转单不计入提成
        status = str(row.get("状态", "")).strip() if pd.notna(row.get("状态")) else "正常"
        if not status:
            status = "正常"
        records.append({
            "date": date.strftime("%Y-%m-%d"),
            "month": month_str,
            "sales": name,
            "customer": str(row.get("客户", "")) if pd.notna(row.get("客户")) else "",
            "car": car,
            "brand": brand,
            "plate": str(row.get("牌照", "")) if pd.notna(row.get("牌照")) else "",
            "order_no": str(row.get("订单号", "")) if pd.notna(row.get("订单号")) else "",
            "review": str(row.get("购车资质好评", "")) if pd.notna(row.get("购车资质好评")) else "",
            "sale_commission": int(sale_comm) if pd.notna(sale_comm) and sale_comm != "" else 0,
            "mgr_commission": int(mgr_comm) if pd.notna(mgr_comm) and mgr_comm != "" else 0,
            "status": status,
        })
    return records


def read_targets():
    """读取月度目标（按表头列名匹配，不按位置顺序）
    这样月度目标表加新销售列时，即使列顺序变化也能正确读取"""
    df = tdoc_reader.read_excel(None, sheet_name="设置", header=None)
    # 找"月度订单目标"行（标题曾为"月度订单目标（问界组）"，后加鸿蒙组改名）
    target_start = None
    for i, row in df.iterrows():
        if pd.notna(row[0]) and str(row[0]).strip().startswith("月度订单目标"):
            target_start = i
            break
    if target_start is None:
        return {}
    # 表头行 = target_start+1，建立 {销售名: 列索引} 映射
    header_row = target_start + 1
    col_map = {}  # {列名: 列索引(0-based)}
    for col_idx in range(2, df.shape[1]):  # 从C列开始
        val = df.iloc[header_row, col_idx]
        if pd.notna(val) and str(val).strip():
            col_map[str(val).strip()] = col_idx
    # 数据从 target_start+2 起
    targets = {}
    for i in range(6):  # 7-12月
        r = target_start + 2 + i
        month_label = df.iloc[r, 0]  # "7月"等
        if pd.isna(month_label):
            continue
        month_str = str(month_label).strip()
        targets[month_str] = {}
        # 按表头列名读取问界组+鸿蒙组成员的目标
        for name in WENJIE_MEMBERS + HM_MEMBERS:
            if name in col_map:
                val = df.iloc[r, col_map[name]]
                targets[month_str][name] = int(val) if pd.notna(val) else 0
            else:
                # 月度目标表没有该销售的列 → 目标0（提示用户手动添加）
                targets[month_str][name] = 0
        # 主管目标 = 本组所有销售目标之和（自动求和，不读取Excel公式缓存）
        targets[month_str]["主管"] = sum(
            targets[month_str].get(name, 0) for name in WENJIE_MEMBERS
        )
        targets[month_str]["鸿蒙主管"] = sum(
            targets[month_str].get(name, 0) for name in HM_MEMBERS
        )
    return targets


def get_coeff(rate):
    """完成率→系数"""
    if rate < 0.6: return 0.5
    elif rate < 0.8: return 0.7
    elif rate < 1.0: return 0.9
    elif rate < 1.2: return 1.0
    else: return 1.2


def calc_group_commission(month_orders, month_targets, members, mgr_name, mgr_target_key,
                          is_august, mgr_count_rule):
    """计算单个组的提成统计
    mgr_count_rule: 主管达标率的订单计数规则
      - "问界": 只统计品牌=问界的订单（周志鹏）
      - "非问界": 只统计品牌≠问界的订单，即鸿蒙订单（熊峰）
    提成基数始终是本组全部订单的主管提成合计
    """
    sales_stats = []
    mgr_sale_comm = 0
    mgr_target = month_targets.get(mgr_target_key, 0)
    for name in members:
        my_orders = [o for o in month_orders if o["sales"] == name]
        order_count = len(my_orders)
        sale_comm = sum(o["sale_commission"] for o in my_orders)
        mgr_comm = sum(o["mgr_commission"] for o in my_orders)
        target = month_targets.get(name, 0)
        rate = order_count / target if target > 0 else 0
        # 8月：销售不乘完成率系数（coeff=1.0）
        coeff = 1.0 if is_august else get_coeff(rate)
        final = sale_comm * coeff
        sales_stats.append({
            "name": name, "target": target, "order_count": order_count,
            "rate": round(rate, 3), "coeff": coeff,
            "sale_commission": sale_comm, "mgr_commission": mgr_comm,
            "final_commission": round(final),
        })
        mgr_sale_comm += mgr_comm
    # 主管达标率：按规则过滤品牌
    if mgr_count_rule == "问界":
        mgr_order_count = sum(
            1 for o in month_orders
            if o["sales"] in members and str(o.get("brand", "")).strip() == "问界"
        )
    else:  # 非问界（鸿蒙订单）：问界订单不计入达标率
        mgr_order_count = sum(
            1 for o in month_orders
            if o["sales"] in members and str(o.get("brand", "")).strip() != "问界"
        )
    # 主管提成：8月保持完成率系数（与销售不同）
    mgr_rate = mgr_order_count / mgr_target if mgr_target > 0 else 0
    mgr_coeff = get_coeff(mgr_rate)  # 主管始终按系数计算
    mgr_final = mgr_sale_comm * mgr_coeff
    return {
        "sales": sales_stats,
        "manager": {
            "name": mgr_name, "target": mgr_target, "order_count": mgr_order_count,
            "rate": round(mgr_rate, 3), "coeff": mgr_coeff,
            "commission_base": mgr_sale_comm, "final_commission": round(mgr_final),
        },
    }


def calc_commission(orders, targets):
    """计算问界组+鸿蒙组提成统计
    8月新政：销售提成不再乘完成率系数（coeff=1.0），主管提成保持系数
    其他月份：销售/主管都乘完成率系数
    主管达标率：周志鹏按问界品牌订单，熊峰按非问界（鸿蒙）订单
    """
    AUGUST_NO_COEFF_MONTHS = {"8月"}  # 销售取消完成率系数的月份
    result = {}
    for month_str, month_targets in targets.items():
        # 只统计状态为"正常"的订单（空状态也当正常，退订/转单不计入提成）
        month_orders = [o for o in orders if o["month"] == month_str and o.get("status", "正常") == "正常"]
        is_august = month_str in AUGUST_NO_COEFF_MONTHS
        # 问界组（主管周志鹏：达标率只算问界品牌订单）
        wj = calc_group_commission(
            month_orders, month_targets, WENJIE_MEMBERS, "周志鹏", "主管",
            is_august, mgr_count_rule="问界",
        )
        result[month_str] = wj
        # 鸿蒙组（主管熊峰：达标率只算非问界订单，问界订单不计数）
        if HM_MEMBERS:
            hm = calc_group_commission(
                month_orders, month_targets, HM_MEMBERS, "熊峰", "鸿蒙主管",
                is_august, mgr_count_rule="非问界",
            )
            result[month_str]["hm_sales"] = hm["sales"]
            result[month_str]["hm_manager"] = hm["manager"]
    return result


def main():
    print("=" * 50)
    print("  南翔销售数据同步脚本（腾讯文档在线表格）")
    print("=" * 50)

    print("  正在连接腾讯文档...")

    # 1. 读取工作量
    workload = read_workload()
    print(f"✅ 读取工作量: {len(workload)}条")
    if workload:
        dates = sorted(set(r["date"] for r in workload))
        print(f"   日期范围: {dates[0]} ~ {dates[-1]}")
        names = sorted(set(r["name"] for r in workload))
        print(f"   人员: {', '.join(names)}")

    # 2. 读取车型提成
    car_comm = read_car_commission()
    print(f"✅ 读取车型提成: {len(car_comm)}款")

    # 2.5 读取8月提成表（新政）
    august_comm = read_august_commission()
    if august_comm:
        print(f"✅ 读取8月提成表: {len(august_comm)}款（销售=大定+交付，不含留车）")

    # 3. 读取订单
    orders = read_orders(car_comm, august_comm)
    print(f"✅ 读取订单: {len(orders)}单")
    if orders:
        months = sorted(set(o["month"] for o in orders))
        print(f"   月份: {', '.join(months)}")
        for m in months:
            mc = sum(1 for o in orders if o["month"] == m)
            print(f"     {m}: {mc}单")

    # 4. 读取目标
    targets = read_targets()
    print(f"✅ 读取月度目标: {len(targets)}个月")

    # 5. 计算提成
    commission = calc_commission(orders, targets)
    print(f"✅ 计算提成统计: {len(commission)}个月")

    # 6. 生成 data.js
    with open("data.js", "w", encoding="utf-8") as f:
        f.write("// 自动生成 - 请勿手动修改\n")
        f.write(f"// 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"var SALES_DATA = {json.dumps(workload, ensure_ascii=False, indent=2)};\n\n")
        f.write(f"var WENJIE_MEMBERS = {json.dumps(WENJIE_MEMBERS, ensure_ascii=False)};\n\n")
        f.write(f"var HM_MEMBERS = {json.dumps(HM_MEMBERS, ensure_ascii=False)};\n\n")
        f.write(f"var ALL_SALES = {json.dumps(list(SALES_INFO.keys()), ensure_ascii=False)};\n\n")
        f.write(f"var SALES_INFO = {json.dumps(SALES_INFO, ensure_ascii=False, indent=2)};\n")
    print(f"✅ 生成 data.js")

    # 7. 生成 orders_data.js
    with open("orders_data.js", "w", encoding="utf-8") as f:
        f.write("// 自动生成 - 请勿手动修改\n")
        f.write(f"// 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"var ordersData = {json.dumps(orders, ensure_ascii=False, indent=2)};\n\n")
        f.write(f"var commissionData = {json.dumps(commission, ensure_ascii=False, indent=2)};\n")
    print(f"✅ 生成 orders_data.js")

    # 8. 打印提成汇总
    print("\n" + "=" * 50)
    print("  问界组提成汇总")
    print("=" * 50)
    for month_str in sorted(commission.keys(), key=lambda x: int(x.replace("月", ""))):
        data = commission[month_str]
        print(f"\n  【{month_str}】")
        print(f"  {'销售':<6} {'目标':>4} {'订单':>4} {'完成率':>8} {'系数':>6} {'提成基数':>8} {'最终提成':>8}")
        print("  " + "-" * 50)
        for s in data["sales"]:
            print(f"  {s['name']:<6} {s['target']:>4} {s['order_count']:>4} {s['rate']:>7.1%} {s['coeff']:>6.1f} {s['sale_commission']:>8} {s['final_commission']:>8}")
        m = data["manager"]
        print("  " + "-" * 50)
        print(f"  {'主管(问界)':<6} {m['target']:>4} {m['order_count']:>4} {m['rate']:>7.1%} {m['coeff']:>6.1f} {m['commission_base']:>8} {m['final_commission']:>8}")

    if HM_MEMBERS:
        print("\n" + "=" * 50)
        print("  鸿蒙组提成汇总")
        print("=" * 50)
        for month_str in sorted(commission.keys(), key=lambda x: int(x.replace("月", ""))):
            data = commission[month_str]
            if "hm_sales" not in data:
                continue
            print(f"\n  【{month_str}】")
            print(f"  {'销售':<6} {'目标':>4} {'订单':>4} {'完成率':>8} {'系数':>6} {'提成基数':>8} {'最终提成':>8}")
            print("  " + "-" * 50)
            for s in data["hm_sales"]:
                print(f"  {s['name']:<6} {s['target']:>4} {s['order_count']:>4} {s['rate']:>7.1%} {s['coeff']:>6.1f} {s['sale_commission']:>8} {s['final_commission']:>8}")
            m = data["hm_manager"]
            print("  " + "-" * 50)
            print(f"  {'主管(鸿蒙)':<6} {m['target']:>4} {m['order_count']:>4} {m['rate']:>7.1%} {m['coeff']:>6.1f} {m['commission_base']:>8} {m['final_commission']:>8}")

    print("\n✅ 同步完成！刷新看板即可查看最新数据。")
    if not os.environ.get("CI"):
        try:
            input("\n按回车退出...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
