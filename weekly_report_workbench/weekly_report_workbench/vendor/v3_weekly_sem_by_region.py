#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SEM周报-分地区脚本（末次归因）
把 v2_weekly_sem_report.py 中各地区数据拆开展示
格式对齐：/Users/yangyingyi/Desktop/格式.xlsx
"""

import importlib.util
import os
import sys
from copy import copy
from pathlib import Path


# 使用 /usr/bin/python3 启动且缺少依赖时，自动切换到项目使用的 Python 3.13。
DEPENDENCY_PYTHON = Path("/Library/Frameworks/Python.framework/Versions/3.13/bin/python3")


def dependency_missing(name):
    try:
        return importlib.util.find_spec(name) is None
    except ModuleNotFoundError:
        return True


missing_dependencies = [
    name for name in ("mysql.connector", "openpyxl") if dependency_missing(name)
]
if missing_dependencies:
    current_python = Path(sys.executable).resolve()
    if DEPENDENCY_PYTHON.exists() and current_python != DEPENDENCY_PYTHON.resolve():
        os.execv(str(DEPENDENCY_PYTHON), [str(DEPENDENCY_PYTHON), *sys.argv])
    raise SystemExit(
        "缺少 Python 依赖："
        + ", ".join(missing_dependencies)
        + "。请运行：python3 -m pip install mysql-connector-python openpyxl"
    )

import mysql.connector
from mysql.connector import Error
from datetime import datetime, timedelta, date
import calendar
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import configparser


# ======================== 配置读取 ========================
def load_config():
    """从 config.ini 读取配置，返回 (db_config, template_path, output_dir)"""
    config_path = Path(__file__).parent / "config.ini"
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    db_config = {
        "host": config.get("database", "host", fallback="localhost"),
        "port": config.getint("database", "port", fallback=3306),
        "user": config.get("database", "user", fallback="root"),
        "password": config.get("database", "password", fallback=""),
        "database": config.get("database", "database", fallback="aippt"),
        "charset": config.get("database", "charset", fallback="utf8mb4"),
    }

    # 分地区末次归因报表固定使用该模板，确保格式、合并单元格和行高不发生变化。
    template_path = str(Path(__file__).resolve().parent.parent / "assets" / "周报SEM分地区-新.xlsx")
    output_dir = os.path.expanduser(
        config.get("paths", "output_dir", fallback="~/Desktop/周报归档")
    )

    return db_config, template_path, output_dir

# ============== 地区定义 ==============
REGIONS = ['印度', '美国', '港澳台']

SEM_LAST_ATTRIBUTION_CONDITION = """
NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'SEM'
"""

# 注册表用 注册地区，订单表用 国家
COUNTRY_MAP = {
    '美国': ['美国'],
    '印度': ['印度'],
    '港澳台': ['中国香港', '中国澳门', '中国台湾'],
}

# Google消耗：地理位置字段值
GOOGLE_GEO_MAP = {
    '美国': ['美国'],
    '印度': ['印度'],
    '港澳台': ['台湾', '香港', '澳门'],
}

# Bing消耗：region字段值
BING_REGION_MAP = {
    '美国': ['United States'],
    '印度': ['India'],
    '港澳台': ['Taiwan', 'Hong Kong SAR', 'Macao SAR'],
}

# ============== 其他国家拆分 ==============
# 主表只保留三个核心地区；以下六个国家单独进入“其他国家-本周”。
OTHER_COUNTRY_EXCLUSIONS = {'美国', '印度', '中国香港', '中国澳门', '中国台湾'}
OTHER_COUNTRIES = ['印度尼西亚', '日本', '墨西哥', '阿联酋', '马来西亚', '巴西']

# 订单/注册表使用中文国家名，Google大多使用中文，Bing使用英文。
# 未配置的中文国家仍会展示注册/订单/营收，并尝试用相同中文名匹配Google；
# Bing消耗只有配置了英文映射后才能合并。
OTHER_COUNTRY_AD_MAP = {
    '印度尼西亚': (['印度尼西亚'], ['Indonesia']),
    '日本': (['日本'], ['Japan']),
    '墨西哥': (['墨西哥'], ['Mexico']),
    '阿联酋': (['阿拉伯联合酋长国', '阿联酋'], ['United Arab Emirates']),
    '马来西亚': (['马来西亚'], ['Malaysia']),
    '巴西': (['巴西'], ['Brazil']),
}

# 通用指标；只有印度、美国、港澳台额外增加“全渠道新用户ROI”。
METRICS = [
    '新增注册', '注册单价', '客单价', '新老用户注册付费率',
    '新用户注册付费率', '新老用户付费量', '新老用户营收',
    '新用户订单量', '新用户营收', '消耗',
    '新老用户ROI（除订阅）', 'SEM新用户ROI'
]
ALL_CHANNEL_ROI_REGIONS = {'印度', '美国', '港澳台'}

# 行号对应 /Users/yangyingyi/Desktop/格式.xlsx 模板
# 每地区仍占17行；三个核心地区会多使用一行展示全渠道新用户ROI。
REGION_ROW_OFFSETS = {
    '印度': 2,
    '美国': 21,
    '港澳台': 40,
}

# 模板中的地区标题名（与 REGIONS 键名可能不同）
REGION_HEADER_TITLES = {
    '印度': 'SEM-印度',
    '美国': 'SEM-美国',
    '港澳台': 'SEM-港澳台',
}

# 每个指标对应的 Excel 数字格式（强制覆盖模板旧格式）
METRIC_FORMATS = {
    '新增注册': '#,##0',
    '注册单价': '$#,##0.00',
    '客单价': '$#,##0.00',
    '新老用户注册付费率': '0.00%',
    '新用户注册付费率': '0.00%',
    '新老用户付费量': '#,##0',
    '新老用户营收': '$#,##0.00',
    '新用户订单量': '#,##0',
    '新用户营收': '$#,##0.00',
    '消耗': '$#,##0.00',
    '新老用户ROI（除订阅）': '0.00%',
    'SEM新用户ROI': '0.00%',
    '全渠道新用户ROI': '0.00%',
}


def get_db_connection(db_config):
    try:
        return mysql.connector.connect(**db_config)
    except Error as e:
        print(f"数据库连接错误: {e}")
        raise


def calculate_dates(this_week_start=None, this_week_end=None):
    """计算所需日期：本周、上周、累计
    可手动传入 本周开始和结束日期，不传则自动按上周四-本周三计算
    """
    if this_week_start and this_week_end:
        this_week_start = datetime.strptime(this_week_start, '%Y-%m-%d').date() if isinstance(this_week_start, str) else this_week_start
        this_week_end = datetime.strptime(this_week_end, '%Y-%m-%d').date() if isinstance(this_week_end, str) else this_week_end
    else:
        # 自动计算：上周四 ~ 本周三
        today = datetime.now().date()
        today_weekday = today.weekday()
        if today_weekday <= 2:
            this_wednesday = today + timedelta(days=(2 - today_weekday))
            last_thursday = this_wednesday - timedelta(days=6)
        else:
            this_wednesday = today - timedelta(days=(today_weekday - 2))
            last_thursday = this_wednesday - timedelta(days=6)
        this_week_start = last_thursday
        this_week_end = this_wednesday

    # 上周：再往前7天
    last_week_start = this_week_start - timedelta(days=7)
    last_week_end = this_week_end - timedelta(days=7)

    yesterday = this_week_end
    month_start = yesterday.replace(day=1)
    _, days_in_month = calendar.monthrange(yesterday.year, yesterday.month)

    return {
        'this_week_start': this_week_start,
        'this_week_end': this_week_end,
        'last_week_start': last_week_start,
        'last_week_end': last_week_end,
        'month_start': month_start,
        'month_end': yesterday,
        'days_in_month': days_in_month,
        'days_passed': yesterday.day,
    }


def execute_query(connection, query, params=None):
    cursor = connection.cursor()
    try:
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        return cursor.fetchone()
    finally:
        cursor.close()


def execute_all(connection, query, params=None):
    cursor = connection.cursor()
    try:
        cursor.execute(query, params or ())
        return cursor.fetchall()
    finally:
        cursor.close()


def build_placeholders(values):
    return ','.join(['%s'] * len(values))


def query_other_country_metrics(connection, start_date, end_date):
    """一次性查询三个核心地区以外所有国家的本周指标。"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

    registration_rows = execute_all(connection, """
        SELECT 注册地区, COUNT(*)
        FROM registrationof4
        WHERE DATE(注册时间) BETWEEN %s AND %s
          AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
          AND 广告来源 NOT LIKE '%FB%'
          AND 渠道类型 = 'SEM'
        GROUP BY 注册地区
    """, (s, e))
    registrations = {
        country: int(value or 0)
        for country, value in registration_rows
        if country and country not in OTHER_COUNTRY_EXCLUSIONS
    }

    order_rows = execute_all(connection, f"""
        SELECT
            国家,
            COUNT(*) AS all_paid_count,
            COALESCE(SUM(订单金额), 0) AS all_revenue,
            SUM(CASE WHEN 注册转化周期 = 0 THEN 1 ELSE 0 END) AS new_user_orders,
            COALESCE(SUM(CASE WHEN 注册转化周期 = 0 THEN 订单金额 ELSE 0 END), 0) AS new_user_revenue
        FROM ordersof4
        WHERE DATE(订单生成时间) BETWEEN %s AND %s
          AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
          AND {SEM_LAST_ATTRIBUTION_CONDITION}
        GROUP BY 国家
    """, (s, e))
    order_data = {
        country: {
            'paid': int(paid or 0),
            'revenue': float(revenue or 0),
            'new_paid': int(new_paid or 0),
            'new_revenue': float(new_revenue or 0),
        }
        for country, paid, revenue, new_paid, new_revenue in order_rows
        if country and country not in OTHER_COUNTRY_EXCLUSIONS
    }

    google_rows = execute_all(connection, """
        SELECT 地理位置, COALESCE(SUM(费用), 0)
        FROM google__ym_ads
        WHERE DATE(`天`) BETWEEN %s AND %s
          AND `广告系列` NOT LIKE '%KOL%'
          AND `广告系列` NOT LIKE '%generic%'
          AND 地理位置 NOT LIKE '总计%%'
        GROUP BY 地理位置
    """, (s, e))
    google_raw = {geo: float(value or 0) for geo, value in google_rows if geo}

    bing_rows = execute_all(connection, """
        SELECT region, COALESCE(SUM(cost), 0)
        FROM bing_region_cost
        WHERE DATE(`stat_date`) BETWEEN %s AND %s
          AND ad_campaign != 'Search total'
        GROUP BY region
    """, (s, e))
    bing_raw = {region: float(value or 0) for region, value in bing_rows if region}

    # 国家全集以业务表为主，并补充映射中本周确有消耗的国家。
    countries = set(registrations) | set(order_data)
    for country, (google_geos, bing_regions) in OTHER_COUNTRY_AD_MAP.items():
        if (
            any(google_raw.get(geo, 0) for geo in google_geos)
            or any(bing_raw.get(region, 0) for region in bing_regions)
        ):
            countries.add(country)
    countries -= OTHER_COUNTRY_EXCLUSIONS

    def safe_div(a, b):
        return a / b if b else 0

    result = {}
    for country in countries:
        registration_count = registrations.get(country, 0)
        orders = order_data.get(country, {})
        paid_count = orders.get('paid', 0)
        revenue = orders.get('revenue', 0)
        new_paid_count = orders.get('new_paid', 0)
        new_revenue = orders.get('new_revenue', 0)

        google_geos, bing_regions = OTHER_COUNTRY_AD_MAP.get(
            country, ([country], [])
        )
        google_cost = sum(google_raw.get(geo, 0) for geo in google_geos)
        bing_cost = sum(bing_raw.get(region, 0) for region in bing_regions)
        total_cost = google_cost + bing_cost

        result[country] = {
            '新增注册': registration_count,
            '注册单价': safe_div(total_cost, registration_count),
            '客单价': safe_div(revenue, paid_count),
            '新老用户注册付费率': safe_div(paid_count, registration_count),
            '新用户注册付费率': safe_div(new_paid_count, registration_count),
            '新老用户付费量': paid_count,
            '新老用户营收': revenue,
            '新用户订单量': new_paid_count,
            '新用户营收': new_revenue,
            '消耗': total_cost,
            '新老用户ROI（除订阅）': safe_div(revenue, total_cost),
            'SEM新用户ROI': safe_div(new_revenue, total_cost),
        }

    # 默认按营收降序，营收相同时按注册数降序，便于优先查看主要市场。
    return dict(sorted(
        result.items(),
        key=lambda item: (
            -item[1]['新老用户营收'],
            -item[1]['新增注册'],
            item[0],
        ),
    ))


def query_registrations(connection, start_date, end_date, countries):
    """查询指定国家的新增注册数"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)
    q = f"""
    SELECT COUNT(*) FROM registrationof4
    WHERE DATE(注册时间) BETWEEN %s AND %s
      AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
      AND 广告来源 NOT LIKE '%FB%'
      AND 渠道类型 = 'SEM'
      AND 注册地区 IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(countries))
    return int(result[0] or 0)


def query_paid_and_revenue(connection, start_date, end_date, countries):
    """查询指定国家的SEM末次归因新老用户除订阅付费量和营收。"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)
    q = f"""
    SELECT COUNT(*), COALESCE(SUM(订单金额), 0) FROM ordersof4
    WHERE DATE(订单生成时间) BETWEEN %s AND %s
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND {SEM_LAST_ATTRIBUTION_CONDITION}
      AND 国家 IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(countries))
    return int(result[0] or 0), float(result[1] or 0)


def query_google_cost(connection, start_date, end_date, geos):
    """查询Google消耗（按地理位置）"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(geos)
    q = f"""
    SELECT COALESCE(SUM(费用), 0) FROM google__ym_ads
    WHERE DATE(`天`) BETWEEN %s AND %s
      AND `广告系列` NOT LIKE '%KOL%'
      AND `广告系列` NOT LIKE '%generic%'
      AND 地理位置 IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(geos))
    return float(result[0] or 0)


def query_bing_cost(connection, start_date, end_date, regions):
    """查询Bing消耗（按region字段）"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(regions)
    q = f"""
    SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
    WHERE DATE(`stat_date`) BETWEEN %s AND %s
      AND ad_campaign != 'Search total'
      AND region IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(regions))
    return float(result[0] or 0)


def query_new_user_orders(connection, start_date, end_date, countries):
    """查询SEM末次归因新用户的订单量和营收（注册转化周期=0）。"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)
    q = f"""
    SELECT COUNT(订单id), COALESCE(SUM(订单金额), 0) FROM ordersof4
    WHERE DATE(订单生成时间) BETWEEN %s AND %s
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND {SEM_LAST_ATTRIBUTION_CONDITION}
      AND 注册转化周期 = 0
      AND 国家 IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(countries))
    return int(result[0] or 0), float(result[1] or 0)


def query_all_channel_new_user_revenue(connection, start_date, end_date, countries):
    """查询全渠道D0新用户营收，不限制合作渠道或末次归因来源。"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)
    q = f"""
    SELECT COALESCE(SUM(订单金额), 0) FROM ordersof4
    WHERE DATE(订单生成时间) BETWEEN %s AND %s
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND 注册转化周期 = 0
      AND 国家 IN ({ph})
    """
    result = execute_query(connection, q, (s, e) + tuple(countries))
    return float(result[0] or 0)


def query_cumulative_2026(connection, end_date, countries):
    """查询2026年累计注册数、累计非订阅订单量和26年注册用户订单量。"""
    e = end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)

    # 累计注册：按注册表渠道类型=SEM统计。
    q_reg = f"""
    SELECT COUNT(*) FROM registrationof4
    WHERE DATE(注册时间) >= '2026-01-01' AND DATE(注册时间) <= %s
      AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
      AND 渠道类型 = 'SEM'
      AND 注册地区 IN ({ph})
    """
    reg_result = execute_query(connection, q_reg, (e,) + tuple(countries))
    cumulative_reg = int(reg_result[0] or 0)

    # 累计付费订单：按支付时间统计，并限定用户注册时间也在2026年。
    # 26年累计统一使用历史渠道口径，不限注册转化周期（新老用户）；
    # “除订阅”的业务口径为支付位置不包含“订阅”，空白支付位置仍保留。
    q_paid = f"""
    SELECT COUNT(*) FROM ordersof4
    WHERE DATE(支付时间) >= '2026-01-01' AND DATE(支付时间) <= %s
      AND DATE(用户注册时间) >= '2026-01-01' AND DATE(用户注册时间) <= %s
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND 渠道类型 = 'SEM'
      AND 合作渠道 IS NULL
      AND COALESCE(支付位置, '') NOT LIKE '%订阅%'
      AND 国家 IN ({ph})
    """
    paid_result = execute_query(connection, q_paid, (e, e) + tuple(countries))
    cumulative_paid = int(paid_result[0] or 0)

    # 26年注册用户中注册转化周期=0的订单量；按支付时间截止到报表结束日。
    # 使用历史渠道类型（老归因），并且包含订阅订单。
    q_registered_user_paid = f"""
    SELECT COUNT(*) FROM ordersof4
    WHERE DATE(用户注册时间) >= '2026-01-01' AND DATE(用户注册时间) <= %s
      AND DATE(支付时间) >= '2026-01-01' AND DATE(支付时间) <= %s
      AND 注册转化周期 = 0
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND 渠道类型 = 'SEM'
      AND 合作渠道 IS NULL
      AND 国家 IN ({ph})
    """
    registered_user_paid_result = execute_query(
        connection, q_registered_user_paid, (e, e) + tuple(countries)
    )
    registered_user_paid = int(registered_user_paid_result[0] or 0)

    return cumulative_reg, cumulative_paid, registered_user_paid


def query_all_channel_new_user_2026(connection, end_date, countries):
    """全渠道年度D0订单量/年度新增注册，不限制订单用户注册时间。"""
    e = end_date.strftime('%Y-%m-%d')
    ph = build_placeholders(countries)
    q_reg = f"""
    SELECT COUNT(*) FROM registrationof4
    WHERE DATE(注册时间) BETWEEN '2026-01-01' AND %s
      AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
      AND 注册地区 IN ({ph})
    """
    q_paid = f"""
    SELECT COUNT(*) FROM ordersof4
    WHERE DATE(支付时间) BETWEEN '2026-01-01' AND %s
      AND 注册转化周期 = 0
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND 国家 IN ({ph})
    """
    params = (e,) + tuple(countries)
    registrations = int(execute_query(connection, q_reg, params)[0] or 0)
    orders = int(execute_query(connection, q_paid, params)[0] or 0)
    return registrations, orders


def get_region_metrics(connection, start_date, end_date, region_name):
    """获取某地区某个时间段的全部指标"""
    countries = COUNTRY_MAP[region_name]
    google_geos = GOOGLE_GEO_MAP[region_name]
    bing_regions = BING_REGION_MAP[region_name]

    registrations = query_registrations(connection, start_date, end_date, countries)
    paid_count, revenue = query_paid_and_revenue(connection, start_date, end_date, countries)
    new_user_orders, new_user_revenue = query_new_user_orders(
        connection, start_date, end_date, countries
    )
    all_channel_new_user_revenue = 0
    if region_name in ALL_CHANNEL_ROI_REGIONS:
        all_channel_new_user_revenue = query_all_channel_new_user_revenue(
            connection, start_date, end_date, countries
        )
    google_cost = query_google_cost(connection, start_date, end_date, google_geos)
    bing_cost = query_bing_cost(connection, start_date, end_date, bing_regions)
    total_cost = google_cost + bing_cost

    def safe_div(a, b):
        return a / b if b and b != 0 else 0

    metrics = {
        '新增注册': registrations,
        '注册单价': safe_div(total_cost, registrations),
        '客单价': safe_div(revenue, paid_count),
        '新老用户注册付费率': safe_div(paid_count, registrations),
        '新用户注册付费率': safe_div(new_user_orders, registrations),
        '新老用户付费量': paid_count,
        '新老用户营收': revenue,
        '新用户订单量': new_user_orders,
        '新用户营收': new_user_revenue,
        '消耗': total_cost,
        '新老用户ROI（除订阅）': safe_div(revenue, total_cost),
        'SEM新用户ROI': safe_div(new_user_revenue, total_cost),
    }
    if region_name in ALL_CHANNEL_ROI_REGIONS:
        metrics['全渠道新用户ROI'] = safe_div(
            all_channel_new_user_revenue, total_cost
        )
    return metrics


def fill_report(wb, dates, this_week_data, last_week_data, mtd_data, cumulative_data):
    """按 格式.xlsx 模板的行号写入数据。
    模板已有完整的 A 列标签 + 表头 + 合并单元格，只需填入 B~E 列数值。
    """
    ws = wb.active

    for region_name in REGIONS:
        base_row = REGION_ROW_OFFSETS[region_name]
        tw = this_week_data.get(region_name, {})
        lw = last_week_data.get(region_name, {})
        mt = mtd_data.get(region_name, {})

        # 标题行：只更新 A 列（确保和模板一致），B~E 不覆盖（模板已有"本周/上周/环比/MTD"）
        header_row = base_row - 1
        title = REGION_HEADER_TITLES.get(region_name, f"SEM-{region_name}")
        ws.cell(row=header_row, column=1, value=title)

        region_metrics = METRICS + (
            ['全渠道新用户ROI'] if region_name in ALL_CHANNEL_ROI_REGIONS else []
        )
        for i, metric in enumerate(region_metrics):
            row = base_row + i
            fmt = METRIC_FORMATS.get(metric, 'General')

            # 模板原来是首次归因旧指标，必须同步覆盖 A 列名称。
            ws.cell(row=row, column=1, value=metric)

            tw_val = tw.get(metric, 0) or 0
            lw_val = lw.get(metric, 0) or 0
            mt_val = mt.get(metric, 0) or 0

            # 环比
            mom_val = (tw_val - lw_val) / lw_val if lw_val and lw_val != 0 else 0

            # 本周 (B) / 上周 (C) / MTD (E) —— 强制设置格式
            for col_idx, val in [(2, tw_val), (3, lw_val), (5, mt_val)]:
                cell = ws.cell(row=row, column=col_idx, value=val)
                cell.number_format = fmt

            # D列(环比)：百分比格式
            mom_cell = ws.cell(row=row, column=4, value=mom_val)
            mom_cell.number_format = '0.00%'

        # ===== 26年累计注册付费率行 =====
        cum_row = base_row + len(region_metrics)
        cum = cumulative_data.get(region_name, {'reg': 0, 'paid': 0, 'rate': 0})
        ws.cell(row=cum_row, column=1, value='26年新老用户累计注册付费率（除订阅）')

        # B列：付费率（百分比）
        rate_cell = ws.cell(row=cum_row, column=2, value=cum['rate'])
        rate_cell.number_format = '0.00%'
        # E列：累计注册数 + 付费订单数
        ws.cell(row=cum_row, column=5, value=f"注册{cum['reg']}, 付费{cum['paid']}")

        # ===== 26年注册用户付费率行 =====
        registered_user_row = cum_row + 1
        ws.row_dimensions[registered_user_row].height = ws.row_dimensions[cum_row].height
        for col_idx in range(1, 6):
            source_cell = ws.cell(row=cum_row, column=col_idx)
            target_cell = ws.cell(row=registered_user_row, column=col_idx)
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)

        ws.cell(row=registered_user_row, column=1, value='26年新注册用户付费率')
        registered_user_rate_cell = ws.cell(
            row=registered_user_row,
            column=2,
            value=cum.get('registered_user_rate', 0),
        )
        registered_user_rate_cell.number_format = '0.00%'
        ws.cell(
            row=registered_user_row,
            column=5,
            value=f"注册{cum['reg']}, 付费{cum.get('registered_user_paid', 0)}",
        )

        # 使用区块内第一行空白，不移动其他地区或改变模板合并关系。
        if region_name in ('印度', '美国', '港澳台'):
            all_channel_row = registered_user_row + 1
            ws.row_dimensions[all_channel_row].height = ws.row_dimensions[cum_row].height
            for col_idx in range(1, 6):
                source_cell = ws.cell(row=cum_row, column=col_idx)
                target_cell = ws.cell(row=all_channel_row, column=col_idx)
                target_cell._style = copy(source_cell._style)
                target_cell.value = None
            ws.cell(all_channel_row, 1, '26年全渠道新用户注册付费率')
            ws.cell(all_channel_row, 2, cum.get('all_channel_new_user_rate', 0)).number_format = '0.00%'
            ws.cell(all_channel_row, 5,
                    f"注册{cum.get('all_channel_reg', 0)}, 付费{cum.get('all_channel_new_user_orders', 0)}")

    return wb


def extend_region_template(wb):
    """格式.xlsx 已包含三个核心地区，不再插行或复制区块。"""
    return wb


def fill_other_countries_sheet(wb, dates, country_data):
    """新增横向国家拆分表：第一行为国家，下面为12项末次归因指标。"""
    sheet_name = '其他国家-本周'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    header_fill = PatternFill('solid', fgColor='4472C4')
    header_font = Font(color='FFFFFF', bold=True)
    metric_fill = PatternFill('solid', fgColor='D9EAF7')
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')

    ws.cell(row=1, column=1, value='SEM')
    for col, country in enumerate(country_data, start=2):
        ws.cell(row=1, column=col, value=country)

    for col in range(1, len(country_data) + 2):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center

    for row, metric in enumerate(METRICS, start=2):
        metric_cell = ws.cell(row=row, column=1, value=metric)
        metric_cell.fill = metric_fill
        metric_cell.font = Font(bold=True)
        metric_cell.border = border
        metric_cell.alignment = center

        for col, country in enumerate(country_data, start=2):
            value = country_data[country].get(metric, 0) or 0
            cell = ws.cell(row=row, column=col, value=value)
            cell.number_format = METRIC_FORMATS.get(metric, 'General')
            cell.border = border
            cell.alignment = center

    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(country_data) + 1)}{len(METRICS) + 1}"
    ws.column_dimensions['A'].width = 22
    for col in range(2, len(country_data) + 2):
        ws.column_dimensions[get_column_letter(col)].width = 15
    ws.sheet_view.showGridLines = False
    ws.oddHeader.center.text = (
        f"SEM其他国家本周数据 "
        f"{dates['this_week_start']:%Y-%m-%d} ~ {dates['this_week_end']:%Y-%m-%d}"
    )
    return wb


def main():
    print("=" * 50)
    print("SEM周报分地区脚本（末次归因）")
    print("=" * 50)

    db_config, template_path, output_dir = load_config()

    # 交互式输入日期（直接回车则自动按上周四~本周三计算）
    print("\n请输入统计周期（格式 YYYY-MM-DD，直接回车则自动计算）:")
    tw_start = input("  本周开始日期: ").strip()
    tw_end = input("  本周结束日期: ").strip()
    if tw_start and tw_end:
        dates = calculate_dates(this_week_start=tw_start, this_week_end=tw_end)
    else:
        dates = calculate_dates()
    print(f"\n本周: {dates['this_week_start']} ~ {dates['this_week_end']}")
    print(f"上周: {dates['last_week_start']} ~ {dates['last_week_end']}")
    print(f"MTD:  {dates['month_start']} ~ {dates['month_end']}")

    print("\n正在连接数据库...")
    connection = get_db_connection(db_config)
    print("数据库连接成功")

    try:
        this_week_data = {}
        last_week_data = {}
        mtd_data = {}
        cumulative_data = {}

        for region_name in REGIONS:
            print(f"\n  查询 {region_name} 地区数据...")
            print(f"    本周...")
            this_week_data[region_name] = get_region_metrics(
                connection, dates['this_week_start'], dates['this_week_end'], region_name)
            print(f"    上周...")
            last_week_data[region_name] = get_region_metrics(
                connection, dates['last_week_start'], dates['last_week_end'], region_name)
            print(f"    MTD...")
            mtd_data[region_name] = get_region_metrics(
                connection, dates['month_start'], dates['month_end'], region_name)
            print(f"    26年累计...")
            cum_reg, cum_paid, registered_user_paid = query_cumulative_2026(
                connection, dates['this_week_end'], COUNTRY_MAP[region_name])
            cum_rate = cum_paid / cum_reg if cum_reg else 0
            registered_user_rate = registered_user_paid / cum_reg if cum_reg else 0
            cumulative_data[region_name] = {
                'reg': cum_reg,
                'paid': cum_paid,
                'rate': cum_rate,
                'registered_user_paid': registered_user_paid,
                'registered_user_rate': registered_user_rate,
            }
            if region_name in ('印度', '美国', '港澳台'):
                all_reg, all_orders = query_all_channel_new_user_2026(
                    connection, dates['this_week_end'], COUNTRY_MAP[region_name])
                cumulative_data[region_name].update({
                    'all_channel_reg': all_reg,
                    'all_channel_new_user_orders': all_orders,
                    'all_channel_new_user_rate': all_orders / all_reg if all_reg else 0,
                })

        print("\n  查询其他国家本周拆分数据...")
        other_country_data = query_other_country_metrics(
            connection, dates['this_week_start'], dates['this_week_end']
        )
        print(f"    共拆分 {len(other_country_data)} 个国家/地区")

        print("\n正在生成报告...")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"找不到指定模板: {template_path}")
        wb = load_workbook(template_path)
        wb = extend_region_template(wb)
        wb = fill_report(wb, dates, this_week_data, last_week_data, mtd_data, cumulative_data)
        wb = fill_other_countries_sheet(wb, dates, other_country_data)

        os.makedirs(output_dir, exist_ok=True)
        tw_str = f"{dates['this_week_start'].strftime('%Y.%m.%d')}-{dates['this_week_end'].strftime('%Y.%m.%d')}"
        output_file = os.path.join(output_dir, f"周报-SEM分地区（末次归因）_{tw_str}.xlsx")
        wb.save(output_file)
        print(f"\n报告已保存: {output_file}")

        # 摘要
        print("\n" + "=" * 50)
        print("数据摘要（本周）:")
        for region_name in REGIONS:
            tw = this_week_data[region_name]
            cum = cumulative_data[region_name]
            roi_str = f"{tw['新老用户ROI（除订阅）']:.0%}" if tw['新老用户ROI（除订阅）'] else "N/A"
            cum_rate_str = f"{cum['rate']:.1%}" if cum['rate'] else "N/A"
            print(f"  {region_name}: 新增注册={tw['新增注册']}, "
                  f"新老用户营收=${tw['新老用户营收']:,.0f}, "
                  f"新用户营收=${tw['新用户营收']:,.0f}, 消耗=${tw['消耗']:,.0f}, "
                  f"ROI={roi_str}, 26年累计付费率={cum_rate_str}")
        print("=" * 50)

    finally:
        connection.close()
        print("\n数据库连接已关闭")


if __name__ == "__main__":
    main()
