from __future__ import annotations

from collections import defaultdict
from copy import copy
from datetime import date, timedelta
from importlib.util import module_from_spec, spec_from_file_location
from io import BytesIO
from pathlib import Path
import json

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, PatternFill

from report_engine import blank_mask, prepare_data, safe_div, web_mask


ROOT = Path(__file__).resolve().parent
WEEKLY = ROOT / "vendor"
REGION_TEMPLATE = ROOT / "assets" / "周报SEM分地区-新.xlsx"
MONTHLY_TARGETS_FILE = ROOT / "config" / "monthly_targets.json"
REPORT_RULES_FILE = ROOT / "config" / "report_rules.json"

ENGLISH = {
    "美国", "英国", "澳大利亚", "加拿大", "印度", "其他", "越南", "沙特", "沙特阿拉伯",
    "阿联酋", "阿拉伯联合酋长国", "荷兰", "缅甸", "土耳其", "埃及", "卡塔尔", "瑞典",
    "科威特", "柬埔寨", "巴基斯坦", "瑞士", "意大利",
}
TRADITIONAL = {"中国香港", "中国澳门", "中国台湾", "香港", "澳门", "台湾", "港澳台"}
CORE_REGIONS = {
    "印度": {"印度"}, "美国": {"美国"}, "港澳台": TRADITIONAL,
    "印度尼西亚": {"印度尼西亚"}, "马来西亚": {"马来西亚"},
}
ACTIVE_RULES = {}


def _configure_rules(rules=None):
    """载入本次生成使用的规则，并刷新依赖地区集合。"""
    global ACTIVE_RULES, ENGLISH, TRADITIONAL, CORE_REGIONS
    if rules is None and REPORT_RULES_FILE.exists():
        with REPORT_RULES_FILE.open("r", encoding="utf-8") as f:
            rules = json.load(f)
    ACTIVE_RULES = rules or {}
    region_rules = ACTIVE_RULES.get("regions", {})
    if region_rules.get("english"):
        ENGLISH = set(region_rules["english"])
    if region_rules.get("traditional"):
        TRADITIONAL = set(region_rules["traditional"])
    north_america = set(region_rules.get("north_america", ["美国", "加拿大"]))
    CORE_REGIONS = {
        "印度": {"印度"}, "美国": {"美国"}, "港澳台": TRADITIONAL,
        "印度尼西亚": {"印度尼西亚"}, "马来西亚": {"马来西亚"},
    }
    return north_america


def _apply_field_mappings(reg, orders, google, bing):
    frames = {"registration": reg.copy(), "orders": orders.copy(), "google": google.copy(), "bing": bing.copy()}
    mappings = ACTIVE_RULES.get("fields", {})
    warnings = []
    for group, frame in frames.items():
        rename = {}
        for canonical, source in mappings.get(group, {}).items():
            source = str(source).strip()
            if canonical not in frame.columns and source in frame.columns:
                rename[source] = canonical
                if source != canonical:
                    warnings.append(f"{group}字段“{source}”已按规则识别为“{canonical}”。")
        frames[group] = frame.rename(columns=rename)
    return frames["registration"], frames["orders"], frames["google"], frames["bing"], warnings


def _new_user_mask(frame):
    config = ACTIVE_RULES.get("new_user", {})
    field = config.get("field", "注册转化周期")
    if field not in frame.columns:
        # 如果填写的是上传文件原字段名，反查映射后的标准字段。
        for canonical, source in ACTIVE_RULES.get("fields", {}).get("orders", {}).items():
            if source == field and canonical in frame.columns:
                field = canonical
                break
    if field not in frame.columns:
        raise ValueError(f"新用户判断字段不存在：{field}")
    value = config.get("value", 0)
    if isinstance(value, (int, float)):
        return pd.to_numeric(frame[field], errors="coerce").eq(float(value))
    return frame[field].fillna("").astype(str).str.strip().eq(str(value).strip())


def _source_values(channel):
    values = ACTIVE_RULES.get("channels", {}).get("blank_cooperation_sources", {}).get(channel, [])
    return [str(value).strip().upper() for value in values]


def _load(name: str, filename: str):
    spec = spec_from_file_location(name, WEEKLY / filename)
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


ALL_MOD = _load("v3_all_original", "v3_weekly_all_report.py")
SEM_MOD = _load("v3_sem_original", "v3_weekly_sem_report.py")
REGION_MOD = _load("v3_region_original", "v3_weekly_sem_by_region.py")


def _dates(start: date, end: date):
    return ALL_MOD.calculate_dates(start, end)


def _between(series, start, end):
    return pd.to_datetime(series, errors="coerce").dt.date.between(start, end)


def _number(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0)


def _prepare(reg, orders, google, bing):
    reg, orders, google, bing, warnings = prepare_data(reg, orders, google, bing)
    for field in ("用户注册时间", "支付时间"):
        if field in orders:
            orders[field] = pd.to_datetime(orders[field], errors="coerce")
    if "用户注册时间" not in orders:
        orders["用户注册时间"] = orders["订单生成时间"]
        warnings.append("订单表缺少用户注册时间，D7已临时使用订单生成时间；建议补充该字段。")
    for field in ("末次来源", "末次来源类型", "合作渠道", "渠道类型", "支付位置"):
        if field not in orders:
            orders[field] = ""
    for field in ("渠道类型", "广告来源"):
        if field not in reg:
            reg[field] = ""
    return reg, orders, google, bing, warnings


def _attrib(orders):
    coop = orders["合作渠道"].fillna("").astype(str).str.strip().str.upper()
    source = orders["末次来源类型"].fillna("").astype(str).str.strip().str.upper()
    default_channel = ACTIVE_RULES.get("channels", {}).get("default_channel", "SEO")
    out = pd.Series(default_channel, index=orders.index, dtype="object")
    priorities = ACTIVE_RULES.get("channels", {}).get("cooperation_priority", [
        {"prefix": "KOL", "channel": "KOL"},
        {"prefix": "AFFILIATE", "channel": "Affiliate"},
    ])
    for item in priorities:
        prefix = str(item.get("prefix", "")).strip().upper()
        if prefix:
            out[coop.str.startswith(prefix)] = item.get("channel", default_channel)
    blank = coop.eq("")
    channel_rules = ACTIVE_RULES.get("channels", {}).get("blank_cooperation_sources", {
        "SEM": ["SEM"], "KOL": ["KOL", "DIRECT", "SOCIAL_ORGANIC"],
        "Affiliate": ["AFFILIATE"], "Auto_Renewal": ["AUTO_RENEWAL"],
        "Unassigned": ["UNASSIGNED"], "Blank": [""],
    })
    channel_rules = dict(channel_rules)
    subscription_values = ACTIVE_RULES.get("subscription", {}).get("source_values")
    if subscription_values is not None:
        channel_rules["Auto_Renewal"] = subscription_values
    for channel, values in channel_rules.items():
        normalized = [str(value).strip().upper() for value in values]
        if channel == "Affiliate":
            matched = pd.Series(False, index=orders.index)
            for value in normalized:
                matched |= source.str.startswith(value)
        else:
            matched = source.isin(normalized)
        out[blank & matched] = channel
    return out


def _google_period(google, start, end):
    g = google[_between(google["天"], start, end)].copy()
    campaign = g["广告系列"].fillna("").astype(str)
    excluded = ACTIVE_RULES.get("sem_cost", {}).get("google_exclude_contains", ["KOL", "generic"])
    mask = pd.Series(False, index=g.index)
    for keyword in excluded:
        if str(keyword).strip():
            mask |= campaign.str.contains(str(keyword).strip(), case=False, regex=False)
    return g[~mask]


def _bing_period(bing, start, end):
    return bing[_between(bing["stat_date"], start, end)].copy()


def _costs(google, bing, start, end):
    g = _google_period(google, start, end)
    b = _bing_period(bing, start, end)
    g_total_rows = g["地理位置"].fillna("").astype(str).eq("总计：地理位置")
    google_total = float(g.loc[g_total_rows, "费用"].sum())
    if not g_total_rows.any():
        google_total = float(g[~g["地理位置"].fillna("").astype(str).str.startswith("总计")]["费用"].sum())
    bing_total = float(b.loc[b["ad_campaign"].fillna("").astype(str).ne("Search total"), "cost"].sum())
    return g, b, google_total, bing_total


def _mobile_cost(mobile, start, end):
    if mobile is None or mobile.empty:
        return 0.0
    cols = list(mobile.columns)
    date_col = next((c for c in cols if "日期" in str(c) or str(c).lower() in {"date", "day"}), cols[0])
    cost_col = next((c for c in cols if "消耗" in str(c) or "花费" in str(c) or str(c).lower() in {"cost", "spend"}), cols[-1])
    return float(_number(mobile.loc[_between(mobile[date_col], start, end), cost_col]).sum())


def _sem_orders(orders, start, end, date_field="订单生成时间"):
    mask = _between(orders[date_field], start, end) & orders["__web"]
    sem_sources = _source_values("SEM") or ["SEM"]
    return orders[mask & blank_mask(orders["合作渠道"]) & orders["末次来源类型"].fillna("").astype(str).str.strip().str.upper().isin(sem_sources)].copy()


def _sem_period(orders, google, bing, mobile, start, end):
    o = _sem_orders(orders, start, end)
    g, b, gc, bc = _costs(google, bing, start, end)
    mobile_cost = _mobile_cost(mobile, start, end)
    source = o["末次来源"].fillna("").astype(str).str.strip().str.upper()
    new = o[_new_user_mask(o)]
    new_source = new["末次来源"].fillna("").astype(str).str.strip().str.upper()
    d7 = _sem_orders(orders, start, end, "用户注册时间")
    d7 = d7[d7["注册转化周期"].between(0, 7)]
    d7_start, d7_end = start - timedelta(days=7), end - timedelta(days=7)
    _, _, d7_gc, d7_bc = _costs(google, bing, d7_start, d7_end)
    d7_mobile = _mobile_cost(mobile, d7_start, d7_end)
    def rev(frame, countries=None):
        if countries is not None: frame = frame[frame["国家"].isin(countries)]
        return float(frame["订单金额"].sum())
    result = {
        "total_revenue": rev(o), "all_user_non_subscription_revenue": rev(o),
        "mobile_revenue": rev(o[o["支付端口"].fillna("").astype(str).str.contains("移动")]),
        "new_user_revenue": rev(new),
        "mobile_new_user_revenue": rev(new[new["支付端口"].fillna("").astype(str).str.contains("移动")]),
        "bing_revenue": rev(new[new_source.eq("BING_ADS")]),
        "google_revenue": rev(new[new_source.eq("GOOGLE_SEM") | new_source.eq("")]),
        "bing_all_user_revenue": rev(o[source.eq("BING_ADS")]),
        "google_all_user_revenue": rev(o[source.eq("GOOGLE_SEM") | source.eq("")]),
        "d7_revenue": rev(d7),
        "d7_pc_revenue": rev(d7[d7["支付端口"].fillna("").astype(str).str.contains("pc|其他", case=False, regex=True)]),
        "d7_h5_revenue": rev(d7[d7["支付端口"].fillna("").astype(str).str.contains("移动")]),
        "google_cost": gc, "bing_cost": bc, "total_cost": gc + bc,
        "mobile_cost": mobile_cost, "pc_cost": gc + bc - mobile_cost,
        "d7_total_cost": d7_gc + d7_bc, "d7_mobile_cost": d7_mobile,
        "d7_pc_cost": d7_gc + d7_bc - d7_mobile,
    }
    region_keys = {"us": {"美国"}, "in": {"印度"}, "tw": TRADITIONAL, "id": {"印度尼西亚"}, "my": {"马来西亚"}}
    google_geo = {"us": {"美国"}, "in": {"印度"}, "tw": {"台湾", "香港", "澳门"}, "id": {"印度尼西亚"}, "my": {"马来西亚"}}
    bing_geo = {"us": {"United States"}, "in": {"India"}, "tw": {"Taiwan", "Hong Kong SAR", "Macao SAR"}, "id": {"Indonesia"}, "my": {"Malaysia"}}
    b_detail = b[b["ad_campaign"].fillna("").astype(str).ne("Search total")]
    for key, countries in region_keys.items():
        result[f"rev_{key}"] = rev(new, countries)
        result[f"all_rev_{key}"] = rev(o, countries)
        result[f"google_cost_{key}"] = float(g.loc[g["地理位置"].isin(google_geo[key]), "费用"].sum())
        result[f"bing_cost_{key}"] = float(b_detail.loc[b_detail["region"].isin(bing_geo[key]), "cost"].sum())
    return result


def _all_metrics(reg, orders, google, bing, dates, inputs):
    m = defaultdict(float)
    orders = orders.copy(); orders["收入归因"] = _attrib(orders)
    periods = {
        "本周": (dates["this_week_start"], dates["this_week_end"]),
        "上周": (dates["last_week_start"], dates["last_week_end"]),
        "MTD": (dates["month_start"], dates["month_end"]),
    }
    for label, (start, end) in periods.items():
        r = reg[_between(reg["注册时间"], start, end) & reg["__web"]]
        o = orders[_between(orders["订单生成时间"], start, end) & orders["__web"]]
        registered_o = orders[_between(orders["用户注册时间"], start, end) & orders["__web"]]
        new = o[_new_user_mask(o)]
        new_non_sub = new[new["收入归因"].ne("Auto_Renewal")]
        non_sub = o[o["收入归因"].ne("Auto_Renewal")]
        sub = o[o["收入归因"].eq("Auto_Renewal")]
        pc_o = o["支付端口"].fillna("").astype(str).str.contains("pc|其他", case=False, regex=True)
        h5_o = o["支付端口"].fillna("").astype(str).str.contains("移动")
        pc_r = r["注册端口"].fillna("").astype(str).str.contains("pc|其他", case=False, regex=True)
        h5_r = r["注册端口"].fillna("").astype(str).str.contains("移动")
        m[f"新增注册_{label}"] = len(r); m[f"付费量_{label}"] = len(o)
        m[f"注册PC端_{label}"] = int(pc_r.sum()); m[f"注册移动端_{label}"] = int(h5_r.sum())
        m[f"新用户付费量PC端_{label}"] = int(new_non_sub["支付端口"].fillna("").astype(str).str.contains("pc|其他", case=False, regex=True).sum())
        m[f"新用户付费量移动端_{label}"] = int(new_non_sub["支付端口"].fillna("").astype(str).str.contains("移动").sum())
        for prefix, frame in [("营收", new_non_sub), ("除订阅营收", non_sub), ("订阅营收", sub)]:
            m[f"{prefix}PC端_{label}"] = float(frame.loc[frame["支付端口"].fillna("").astype(str).str.contains("pc|其他", case=False, regex=True), "订单金额"].sum())
            m[f"{prefix}移动端_{label}"] = float(frame.loc[frame["支付端口"].fillna("").astype(str).str.contains("移动"), "订单金额"].sum())
        m[f"总营收_{label}"] = float(new_non_sub["订单金额"].sum())
        m[f"除订阅总营收_{label}"] = float(non_sub["订单金额"].sum())
        m[f"含订阅总营收_{label}"] = float(o["订单金额"].sum())
        d0 = registered_o[_new_user_mask(registered_o)]
        d7 = registered_o[registered_o["注册转化周期"].between(0, 7)]
        m[f"D0营收_{label}"] = float(d0["订单金额"].sum())
        m[f"D7_SEM营收_{label}"] = float(d7.loc[d7["收入归因"].eq("SEM"), "订单金额"].sum())
        for channel in ["SEM", "KOL", "Affiliate", "SEO", "Unassigned", "Blank"]:
            m[f"{channel}营收_{label}"] = float(new.loc[new["收入归因"].eq(channel), "订单金额"].sum())
            m[f"新老除订阅_{channel}营收_{label}"] = float(non_sub.loc[non_sub["收入归因"].eq(channel), "订单金额"].sum())
        m[f"归因总营收_{label}"] = float(new["订单金额"].sum())
        north_america = set(ACTIVE_RULES.get("regions", {}).get("north_america", ["美国", "加拿大"]))
        for name, countries in [("英语", ENGLISH), ("繁中", TRADITIONAL), ("北美", north_america), ("印度", {"印度"})]:
            m[f"地区营收_{name}_{label}"] = float(new_non_sub.loc[new_non_sub["国家"].isin(countries), "订单金额"].sum())
            m[f"新老除订阅地区_{name}_{label}"] = float(non_sub.loc[non_sub["国家"].isin(countries), "订单金额"].sum())
            m[f"订阅地区_{name}_{label}"] = float(sub.loc[sub["国家"].isin(countries), "订单金额"].sum())
        m[f"地区营收_其他地区_{label}"] = m[f"总营收_{label}"]-m[f"地区营收_英语_{label}"]-m[f"地区营收_繁中_{label}"]
        m[f"新老除订阅地区_其他_{label}"] = m[f"除订阅总营收_{label}"]-m[f"新老除订阅地区_英语_{label}"]-m[f"新老除订阅地区_繁中_{label}"]
        m[f"订阅地区_其他_{label}"] = float(sub["订单金额"].sum())-m[f"订阅地区_英语_{label}"]-m[f"订阅地区_繁中_{label}"]
        g, b, gc, bc = _costs(google, bing, start, end)
        m[f"SEM消耗_{label}"] = gc + bc
        # 地区消耗与原脚本一致：英语/繁中/其他。
        g_en = float(g.loc[g["地理位置"].isin(ENGLISH), "费用"].sum())
        g_tc = float(g.loc[g["地理位置"].isin({"香港", "台湾", "澳门"}), "费用"].sum())
        b_en_names = {"United States","Canada","India","Vietnam","United Arab Emirates","Myanmar","Türkiye","State of Qatar","Sweden","Kuwait","Kingdom of Cambodia","Switzerland"}
        b_tc_names = {"Hong Kong SAR","Macao SAR","Taiwan"}
        b_en = float(b.loc[b["region"].isin(b_en_names), "cost"].sum()); b_tc = float(b.loc[b["region"].isin(b_tc_names), "cost"].sum())
        m[f"SEM消耗_英语_{label}"] = g_en+b_en; m[f"SEM消耗_繁中_{label}"] = g_tc+b_tc
        m[f"SEM消耗_其他_{label}"] = gc+bc-g_en-b_en-g_tc-b_tc
        m[f"PC流量_{label}"] = inputs.get(f"pc_flow_{label}", 0); m[f"H5流量_{label}"] = inputs.get(f"h5_flow_{label}", 0)
        m[f"流量_{label}"] = m[f"PC流量_{label}"] + m[f"H5流量_{label}"]
    return m


def _region_metrics(reg, orders, google, bing, start, end, region):
    countries = CORE_REGIONS[region]
    r = reg[_between(reg["注册时间"], start, end) & reg["__web"] & reg["地区"].isin(countries)]
    r = r[r["渠道类型"].fillna("").astype(str).str.upper().eq("SEM") & ~r["广告来源"].fillna("").astype(str).str.contains("FB", case=False)]
    o = _sem_orders(orders, start, end); o = o[o["地区"].isin(countries)]
    new = o[_new_user_mask(o)]
    g, b, _, _ = _costs(google, bing, start, end)
    google_names = {"港澳台": {"香港","台湾","澳门"}}.get(region, countries)
    bing_names = {"美国":{"United States"},"印度":{"India"},"港澳台":{"Taiwan","Hong Kong SAR","Macao SAR"},"印度尼西亚":{"Indonesia"},"马来西亚":{"Malaysia"}}[region]
    cost = float(g.loc[g["地理位置"].isin(google_names), "费用"].sum()) + float(b.loc[b["ad_campaign"].fillna("").astype(str).ne("Search total") & b["region"].isin(bing_names), "cost"].sum())
    reg_count, paid, revenue, new_paid, new_rev = len(r), len(o), float(o["订单金额"].sum()), len(new), float(new["订单金额"].sum())
    base_all_new = orders[_between(orders["订单生成时间"], start, end) & orders["__web"] & orders["地区"].isin(countries)]
    all_new = base_all_new[_new_user_mask(base_all_new)]
    return {"新增注册":reg_count,"注册单价":safe_div(cost,reg_count),"客单价":safe_div(revenue,paid),"新老用户注册付费率":safe_div(paid,reg_count),"新用户注册付费率":safe_div(new_paid,reg_count),"新老用户付费量":paid,"新老用户营收":revenue,"新用户订单量":new_paid,"新用户营收":new_rev,"消耗":cost,"新老用户ROI（除订阅）":safe_div(revenue,cost),"SEM新用户ROI":safe_div(new_rev,cost),"全渠道新用户ROI":safe_div(float(all_new["订单金额"].sum()),cost)}


def _country_metrics(reg, orders, google, bing, start, end, country):
    r = reg[_between(reg["注册时间"], start, end) & reg["__web"] & reg["地区"].eq(country)]
    r = r[r["渠道类型"].fillna("").astype(str).str.upper().eq("SEM") & ~r["广告来源"].fillna("").astype(str).str.contains("FB", case=False)]
    o = _sem_orders(orders, start, end); o = o[o["地区"].eq(country)]
    new = o[_new_user_mask(o)]
    g, b, _, _ = _costs(google, bing, start, end)
    google_names, bing_names = REGION_MOD.OTHER_COUNTRY_AD_MAP.get(country, ([country], [country]))
    b_detail = b[b["ad_campaign"].fillna("").astype(str).ne("Search total")]
    cost = float(g.loc[g["地理位置"].isin(google_names), "费用"].sum()) + float(b_detail.loc[b_detail["region"].isin(bing_names), "cost"].sum())
    reg_count, paid, revenue, new_paid, new_rev = len(r), len(o), float(o["订单金额"].sum()), len(new), float(new["订单金额"].sum())
    return {"新增注册":reg_count,"注册单价":safe_div(cost,reg_count),"客单价":safe_div(revenue,paid),"新老用户注册付费率":safe_div(paid,reg_count),"新用户注册付费率":safe_div(new_paid,reg_count),"新老用户付费量":paid,"新老用户营收":revenue,"新用户订单量":new_paid,"新用户营收":new_rev,"消耗":cost,"新老用户ROI（除订阅）":safe_div(revenue,cost),"SEM新用户ROI":safe_div(new_rev,cost)}


def _wb_bytes(wb):
    bio = BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def _remove_metric_rows(ws, labels):
    """删除指定指标行，并重建H列时间进度合并区域。"""
    time_progress = ws["H4"].value
    for merged in list(ws.merged_cells.ranges):
        if merged.min_col == merged.max_col == 8 and merged.min_row <= 4 <= merged.max_row:
            ws.unmerge_cells(str(merged))
    rows = [
        row for row in range(4, ws.max_row + 1)
        if str(ws.cell(row, 1).value or "").strip() in labels
    ]
    for row in reversed(rows):
        ws.delete_rows(row, 1)
    ws.merge_cells(start_row=4, start_column=8, end_row=ws.max_row, end_column=8)
    ws["H4"] = time_progress
    ws["H4"].alignment = copy(ws["G4"].alignment)
    return ws


def _simplify_region_main_sheet(ws):
    """地区主表只保留A:E，累计付费率指标作为无边框的表外信息。"""
    for merged in list(ws.merged_cells.ranges):
        if merged.max_col >= 6:
            ws.unmerge_cells(str(merged))
    ws.delete_cols(6, 3)
    for row in range(1, ws.max_row + 1):
        label = str(ws.cell(row, 1).value or "").strip()
        if label.startswith("26年"):
            for col in range(1, 6):
                cell = ws.cell(row, col)
                cell.border = Border()
                cell.fill = PatternFill(fill_type=None)
            ws.cell(row, 1).alignment = copy(ws.cell(row, 1).alignment)
    return ws


def generate_v3_reports(reg, orders, google, bing, mobile, start, end, inputs, targets=None, rules=None):
    _configure_rules(rules)
    reg, orders, google, bing, mapping_warnings = _apply_field_mappings(reg, orders, google, bing)
    reg, orders, google, bing, warnings = _prepare(reg, orders, google, bing)
    warnings = mapping_warnings + warnings
    dates = _dates(start, end)
    month_key = end.strftime("%Y-%m")
    monthly_all = dict(ALL_MOD.MONTHLY_TARGETS.get(month_key, {}))
    monthly_sem = dict(SEM_MOD.MONTHLY_TARGETS.get(month_key, {}))
    if MONTHLY_TARGETS_FILE.exists():
        with MONTHLY_TARGETS_FILE.open("r", encoding="utf-8") as f:
            configured_targets = json.load(f).get(month_key, {})
        monthly_all.update(configured_targets.get("all_channel", {}))
        monthly_sem.update(configured_targets.get("sem", {}))
    # 网站保存后的目标配置优先级最高，保证保存后立即生成即可生效。
    if targets:
        current_targets = targets.get(month_key, {})
        monthly_all.update(current_targets.get("all_channel", {}))
        monthly_sem.update(current_targets.get("sem", {}))

    all_metrics = _all_metrics(reg, orders, google, bing, dates, inputs)
    all_wb = ALL_MOD.create_report(
        dates, all_metrics,
        inputs.get("commission_本周",0), inputs.get("commission_上周",0), inputs.get("commission_MTD",0),
        inputs.get("seo_本周",0), inputs.get("seo_上周",0), inputs.get("seo_MTD",0),
        inputs.get("kol_本周",0), inputs.get("kol_上周",0), inputs.get("kol_MTD",0),
        inputs.get("refund_本周",0), inputs.get("refund_上周",0), inputs.get("refund_MTD",0), monthly_all,
    )
    _remove_metric_rows(all_wb["全渠道周报"], {"WEB—SEM投放新用户ROI（D7）"})

    period_map = {"本周":(dates["this_week_start"],dates["this_week_end"]),"上周":(dates["last_week_start"],dates["last_week_end"]),"MTD":(dates["month_start"],dates["month_end"])}
    sem_data = {k:_sem_period(orders,google,bing,mobile,*v) for k,v in period_map.items()}
    sem_wb = SEM_MOD.create_report(dates, sem_data["本周"], sem_data["上周"], sem_data["MTD"], monthly_sem)
    _remove_metric_rows(
        sem_wb["SEM周报"],
        {"PC端", "H5端", "SEM—新用户ROI（D7）"},
    )

    if not REGION_TEMPLATE.exists():
        raise FileNotFoundError(f"找不到分地区模板：{REGION_TEMPLATE}")
    region_wb = load_workbook(REGION_TEMPLATE)
    region_wb = REGION_MOD.extend_region_template(region_wb)
    region_sets = {label:{region:_region_metrics(reg,orders,google,bing,*period,region) for region in REGION_MOD.REGIONS} for label,period in period_map.items()}
    cumulative = {}
    for region in REGION_MOD.REGIONS:
        countries = CORE_REGIONS[region]
        cr = reg[_between(reg["注册时间"], date(2026,1,1), end) & reg["__web"] & reg["地区"].isin(countries)]
        sem_cr = cr[cr["渠道类型"].fillna("").astype(str).str.upper().eq("SEM")]
        co = orders[_between(orders["支付时间"] if "支付时间" in orders else orders["订单生成时间"], date(2026,1,1), end) & orders["__web"] & orders["地区"].isin(countries)]
        sem_co = co[blank_mask(co["合作渠道"]) & co["渠道类型"].fillna("").astype(str).str.upper().eq("SEM")]
        non_sub = sem_co[~sem_co["支付位置"].fillna("").astype(str).str.contains("订阅")]
        registered_paid = sem_co[_new_user_mask(sem_co)]
        cumulative[region] = {"reg":len(sem_cr),"paid":len(non_sub),"rate":safe_div(len(non_sub),len(sem_cr)),"registered_user_paid":len(registered_paid),"registered_user_rate":safe_div(len(registered_paid),len(sem_cr))}
        if region in {"印度","美国","港澳台"}:
            all_new = co[_new_user_mask(co)]
            cumulative[region].update({"all_channel_reg":len(cr),"all_channel_new_user_orders":len(all_new),"all_channel_new_user_rate":safe_div(len(all_new),len(cr))})
    region_wb = REGION_MOD.fill_report(region_wb, dates, region_sets["本周"], region_sets["上周"], region_sets["MTD"], cumulative)
    _simplify_region_main_sheet(region_wb.active)
    # 其他国家页固定顺序，只展示业务要求的六个地区。
    countries = REGION_MOD.OTHER_COUNTRIES
    other = {
        country: _country_metrics(reg, orders, google, bing, dates["this_week_start"], dates["this_week_end"], country)
        for country in countries
    }
    region_wb = REGION_MOD.fill_other_countries_sheet(region_wb, dates, other)

    stamp = f"{start:%Y.%m.%d}-{end:%Y.%m.%d}"
    return {
        f"周报-全渠道_末次归因_{stamp}.xlsx": _wb_bytes(all_wb),
        f"周报-SEM_末次归因新用户_{stamp}.xlsx": _wb_bytes(sem_wb),
        f"周报-SEM分地区（末次归因）_{stamp}.xlsx": _wb_bytes(region_wb),
    }, warnings
