from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from io import BytesIO
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


TARGET_REGIONS = [
    "印度", "美国", "港澳台", "印度尼西亚", "巴西", "日本",
    "菲律宾", "韩国", "墨西哥", "阿联酋", "马来西亚",
]

REGION_ALIASES = {
    "中国香港": "港澳台", "香港": "港澳台", "中国澳门": "港澳台",
    "澳门": "港澳台", "中国台湾": "港澳台", "台湾": "港澳台",
    "印尼": "印度尼西亚", "马来": "马来西亚",
    "阿拉伯联合酋长国": "阿联酋", "俄罗斯联邦": "俄罗斯",
}

BING_REGION_MAP = {
    "United States": "美国", "India": "印度", "Taiwan": "港澳台",
    "Hong Kong SAR": "港澳台", "Macao SAR": "港澳台",
    "Indonesia": "印度尼西亚", "Brazil": "巴西", "Japan": "日本",
    "Philippines": "菲律宾", "Republic of the Philippines": "菲律宾",
    "South Korea": "韩国", "Republic of Korea": "韩国", "Mexico": "墨西哥",
    "United Arab Emirates": "阿联酋", "Malaysia": "马来西亚",
}

REQUIRED = {
    "注册": ["注册时间", "注册端口", "注册地区"],
    "订单": ["订单生成时间", "支付端口", "国家", "订单金额", "注册转化周期"],
    "Google": ["天", "地理位置", "广告系列", "费用"],
    "Bing": ["stat_date", "region", "ad_campaign", "cost"],
}

METRICS = [
    "新增注册", "注册单价", "客单价", "注册付费率", "新用户注册付费率",
    "付费量", "营收", "新用户订单量", "新用户营收", "消耗", "ROI", "新用户ROI",
]


@dataclass
class Period:
    label: str
    start: date
    end: date


def read_uploads(files: Iterable) -> pd.DataFrame:
    frames = []
    for file in files or []:
        name = file.name.lower()
        file.seek(0)
        if name.endswith(".csv"):
            try:
                frame = pd.read_csv(file, encoding="utf-8-sig", low_memory=False)
            except UnicodeDecodeError:
                file.seek(0)
                frame = pd.read_csv(file, encoding="gb18030", low_memory=False)
        else:
            frame = pd.read_excel(file)
        frame.columns = [str(c).strip() for c in frame.columns]
        frame["__来源文件"] = file.name
        frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def normalize_region(series: pd.Series) -> pd.Series:
    values = series.fillna("").astype(str).str.strip()
    return values.replace(REGION_ALIASES)


def web_mask(series: pd.Series) -> pd.Series:
    value = series.fillna("").astype(str).str.lower()
    return value.str.contains("pc|移动|其他", regex=True)


def blank_mask(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().eq("")


def safe_div(a: float, b: float) -> float:
    return float(a) / float(b) if b else 0.0


def prepare_data(reg, orders, google, bing):
    warnings = []
    # 兼容数据库导出和历史Excel中的字段名称差异。
    if "订单金额" not in orders.columns and "兑换美元" in orders.columns:
        orders = orders.copy()
        orders["订单金额"] = orders["兑换美元"]
        warnings.append("订单金额使用订单表的“兑换美元”字段。")
    bing_aliases = {
        "Day": "stat_date",
        "Campaign": "ad_campaign",
        "Spend": "cost",
        "日期": "stat_date",
        "国家/地区（实际位置）": "region",
        "国家/地区(实际位置)": "region",
        "广告活动": "ad_campaign",
        "花费": "cost",
        "展示": "impressions",
        "点击": "clicks",
        "CTR": "ctr",
    }
    for source, target in bing_aliases.items():
        if target not in bing.columns and source in bing.columns:
            bing = bing.rename(columns={source: target})
            warnings.append(f"Bing字段“{source}”已自动识别为“{target}”。")

    if "地理位置" not in google.columns:
        geo_column = next(
            (
                column for column in google.columns
                if "地理位置" in str(column).replace("\n", "").replace("\r", "").replace(" ", "")
            ),
            None,
        )
        if geo_column is not None:
            google = google.rename(columns={geo_column: "地理位置"})
            warnings.append(f"Google字段“{geo_column}”已自动识别为“地理位置”。")

    datasets = {"注册": reg.copy(), "订单": orders.copy(), "Google": google.copy(), "Bing": bing.copy()}
    for name, frame in datasets.items():
        missing = [c for c in REQUIRED[name] if c not in frame.columns]
        if missing:
            raise ValueError(f"{name}数据缺少字段：{', '.join(missing)}")

    for key in ("用户ID", "用户id"):
        if key in reg.columns:
            before = len(reg); reg = reg.drop_duplicates(key, keep="last")
            if before != len(reg): warnings.append(f"注册数据按{key}去除{before-len(reg)}条重复记录。")
            break
    for key in ("订单id", "订单ID"):
        if key in orders.columns:
            before = len(orders); orders = orders.drop_duplicates(key, keep="last")
            if before != len(orders): warnings.append(f"订单数据按{key}去除{before-len(orders)}条重复记录。")
            break

    reg["注册时间"] = pd.to_datetime(reg["注册时间"], errors="coerce")
    reg["地区"] = normalize_region(reg["注册地区"])
    reg["__web"] = web_mask(reg["注册端口"])

    orders["订单生成时间"] = pd.to_datetime(orders["订单生成时间"], errors="coerce")
    orders["地区"] = normalize_region(orders["国家"])
    orders["订单金额"] = pd.to_numeric(orders["订单金额"], errors="coerce").fillna(0)
    orders["注册转化周期"] = pd.to_numeric(orders["注册转化周期"], errors="coerce")
    orders["__web"] = web_mask(orders["支付端口"])
    if "合作渠道" not in orders:
        orders["合作渠道"] = ""
        warnings.append("订单表缺少合作渠道，已按全部为空处理。")
    if "末次来源类型" not in orders:
        if "渠道类型" not in orders:
            orders["渠道类型"] = ""
        orders["末次来源类型"] = orders["渠道类型"]
        warnings.append("订单表缺少末次来源类型，SEM末次归因已回退使用渠道类型。")
    if "渠道类型" not in orders:
        orders["渠道类型"] = ""

    google["天"] = pd.to_datetime(google["天"], errors="coerce")
    google["地区"] = normalize_region(google["地理位置"])
    google["费用"] = pd.to_numeric(
        google["费用"].astype(str).str.replace(",", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0)

    bing["stat_date"] = pd.to_datetime(bing["stat_date"], errors="coerce")
    bing["地区"] = bing["region"].fillna("").astype(str).str.strip().replace(BING_REGION_MAP)
    bing["cost"] = pd.to_numeric(
        bing["cost"].astype(str).str.replace(",", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0)

    return reg, orders, google, bing, warnings


def period_mask(series: pd.Series, period: Period) -> pd.Series:
    return series.dt.date.between(period.start, period.end)


def aggregate_cost(google, bing, period: Period, scope: str) -> pd.Series:
    g = google[period_mask(google["天"], period)].copy()
    campaign = g["广告系列"].fillna("").astype(str)
    g = g[~campaign.str.contains("generic", case=False, na=False)]
    if scope == "SEM":
        g = g[~campaign.loc[g.index].str.contains("KOL", case=False, na=False)]
    g = g[~g["地区"].str.startswith("总计", na=False)]
    gcost = g.groupby("地区")["费用"].sum()

    b = bing[period_mask(bing["stat_date"], period)].copy()
    b = b[b["ad_campaign"].fillna("").astype(str).ne("Search total")]
    bcost = b.groupby("地区")["cost"].sum()
    return gcost.add(bcost, fill_value=0)


def aggregate_metrics(reg, orders, google, bing, period: Period, scope: str, monday=False):
    r = reg[period_mask(reg["注册时间"], period) & reg["__web"]].copy()
    o = orders[period_mask(orders["订单生成时间"], period) & orders["__web"]].copy()

    if scope == "SEM":
        if "广告来源" not in r:
            r["广告来源"] = ""
        if "渠道类型" not in r:
            r["渠道类型"] = ""
        r = r[
            ~r["广告来源"].fillna("").astype(str).str.contains("FB", case=False)
            & r["渠道类型"].fillna("").astype(str).str.upper().eq("SEM")
        ]
        if monday:
            o = o[
                o["渠道类型"].fillna("").astype(str).str.upper().eq("SEM")
                & blank_mask(o["合作渠道"])
            ]
        else:
            o = o[
                blank_mask(o["合作渠道"])
                & o["末次来源类型"].fillna("").astype(str).str.strip().str.upper().eq("SEM")
            ]

    reg_count = r.groupby("地区").size()
    paid = o.groupby("地区").size()
    revenue = o.groupby("地区")["订单金额"].sum()
    new = o[o["注册转化周期"].eq(0)]
    new_paid = new.groupby("地区").size()
    new_revenue = new.groupby("地区")["订单金额"].sum()
    cost = aggregate_cost(google, bing, period, scope)

    regions = set(reg_count.index) | set(paid.index) | set(cost.index)
    rows = []
    for region in regions:
        registrations = int(reg_count.get(region, 0))
        order_count = int(paid.get(region, 0))
        income = float(revenue.get(region, 0))
        new_orders = int(new_paid.get(region, 0))
        new_income = float(new_revenue.get(region, 0))
        spend = float(cost.get(region, 0))
        rows.append({
            "地区": region, "新增注册": registrations,
            "注册单价": safe_div(spend, registrations), "客单价": safe_div(income, order_count),
            "注册付费率": safe_div(order_count, registrations),
            "新用户注册付费率": safe_div(new_orders, registrations), "付费量": order_count,
            "营收": income, "新用户订单量": new_orders, "新用户营收": new_income,
            "消耗": spend, "ROI": safe_div(income, spend), "新用户ROI": safe_div(new_income, spend),
        })
    return pd.DataFrame(rows).set_index("地区") if rows else pd.DataFrame(columns=METRICS)


def total_metrics(region_df: pd.DataFrame) -> dict:
    if region_df.empty:
        return {m: 0 for m in METRICS}
    reg = region_df["新增注册"].sum(); orders = region_df["付费量"].sum()
    revenue = region_df["营收"].sum(); new_orders = region_df["新用户订单量"].sum()
    new_revenue = region_df["新用户营收"].sum(); cost = region_df["消耗"].sum()
    return {
        "新增注册": reg, "注册单价": safe_div(cost, reg), "客单价": safe_div(revenue, orders),
        "注册付费率": safe_div(orders, reg), "新用户注册付费率": safe_div(new_orders, reg),
        "付费量": orders, "营收": revenue, "新用户订单量": new_orders,
        "新用户营收": new_revenue, "消耗": cost, "ROI": safe_div(revenue, cost),
        "新用户ROI": safe_div(new_revenue, cost),
    }


def build_summary(current: dict, previous: dict, scope: str) -> list[str]:
    lines = []
    for metric in ["新增注册", "营收", "新用户营收", "消耗", "ROI", "新用户ROI"]:
        old, now = previous.get(metric, 0), current.get(metric, 0)
        if old:
            change = (now - old) / old
            if abs(change) >= .1:
                direction = "上升" if change > 0 else "下降"
                lines.append(f"{scope}{metric}环比{direction}{abs(change):.1%}。")
    return lines or [f"{scope}核心指标环比变化均未超过10%。"]


def style_sheet(ws):
    navy, blue, light = "1F4E78", "4472C4", "D9EAF7"
    thin = Side(style="thin", color="D9E1F2")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B2"
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy); cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(bottom=thin)
        if row[0].row > 1:
            row[0].fill = PatternFill("solid", fgColor=light)
    ws.column_dimensions["A"].width = 24
    for col in range(2, ws.max_column + 1): ws.column_dimensions[get_column_letter(col)].width = 16


def write_overview(ws, title, values):
    ws.append([title, "本周", "上周", "环比", "MTD"])
    rate_metrics = {"注册付费率", "新用户注册付费率", "ROI", "新用户ROI"}
    money_metrics = {"注册单价", "客单价", "营收", "新用户营收", "消耗"}
    for metric in METRICS:
        cur, prev, mtd = values["本周"][metric], values["上周"][metric], values["MTD"][metric]
        change = safe_div(cur - prev, prev) if prev else 0
        ws.append([metric, cur, prev, change, mtd])
        row = ws.max_row
        fmt = "0.00%" if metric in rate_metrics else "$#,##0.00" if metric in money_metrics else "#,##0"
        for col in (2, 3, 5): ws.cell(row, col).number_format = fmt
        ws.cell(row, 4).number_format = "0.00%"
    style_sheet(ws)


def generate_report(reg, orders, google, bing, start: date, end: date):
    reg, orders, google, bing, warnings = prepare_data(reg, orders, google, bing)
    periods = {
        "本周": Period("本周", start, end),
        "上周": Period("上周", start - timedelta(days=7), end - timedelta(days=7)),
        "MTD": Period("MTD", start.replace(day=1), end),
    }
    calculated = {}
    for scope in ("全渠道", "SEM"):
        calculated[scope] = {}
        for label, period in periods.items():
            region = aggregate_metrics(reg, orders, google, bing, period, scope)
            calculated[scope][label] = {"region": region, "total": total_metrics(region)}

    wb = Workbook(); wb.remove(wb.active)
    write_overview(wb.create_sheet("全渠道总览"), "全渠道指标", {k:v["total"] for k,v in calculated["全渠道"].items()})
    write_overview(wb.create_sheet("SEM总览"), "SEM末次归因指标", {k:v["total"] for k,v in calculated["SEM"].items()})

    ws = wb.create_sheet("SEM分地区")
    ws.append(["地区", "指标", "本周", "上周", "环比", "MTD"])
    for region in TARGET_REGIONS:
        for metric in METRICS:
            vals = []
            for label in ("本周", "上周", "MTD"):
                df = calculated["SEM"][label]["region"]
                vals.append(float(df.at[region, metric]) if region in df.index else 0)
            change = safe_div(vals[0]-vals[1], vals[1]) if vals[1] else 0
            ws.append([region, metric, vals[0], vals[1], change, vals[2]])
            row=ws.max_row; fmt="0.00%" if metric in {"注册付费率","新用户注册付费率","ROI","新用户ROI"} else "$#,##0.00" if metric in {"注册单价","客单价","营收","新用户营收","消耗"} else "#,##0"
            for c in (3,4,6): ws.cell(row,c).number_format=fmt
            ws.cell(row,5).number_format="0.00%"
    style_sheet(ws)

    ws = wb.create_sheet("周一监控")
    ws.append(["口径", "地区", *METRICS])
    period = periods["本周"]
    for scope in ("SEM", "全渠道"):
        df = aggregate_metrics(reg, orders, google, bing, period, scope, monday=True)
        for region in TARGET_REGIONS:
            ws.append([scope, region, *[float(df.at[region,m]) if region in df.index else 0 for m in METRICS]])
    style_sheet(ws)

    ws = wb.create_sheet("数据检查")
    ws.append(["数据", "行数", "最早日期", "最晚日期", "说明"])
    for name, frame, field in [("注册",reg,"注册时间"),("订单",orders,"订单生成时间"),("Google",google,"天"),("Bing",bing,"stat_date")]:
        ws.append([name,len(frame),frame[field].min(),frame[field].max(),""])
        ws.cell(ws.max_row,3).number_format=ws.cell(ws.max_row,4).number_format="yyyy-mm-dd"
    for warning in warnings: ws.append(["警告",None,None,None,warning])
    style_sheet(ws); ws.column_dimensions["E"].width=60

    ws = wb.create_sheet("自动总结")
    ws.append(["自动总结"])
    lines = build_summary(calculated["全渠道"]["本周"]["total"], calculated["全渠道"]["上周"]["total"], "全渠道")
    lines += build_summary(calculated["SEM"]["本周"]["total"], calculated["SEM"]["上周"]["total"], "SEM")
    for line in lines: ws.append([line])
    style_sheet(ws); ws.column_dimensions["A"].width=80

    output = BytesIO(); wb.save(output); output.seek(0)
    return output, warnings, lines

