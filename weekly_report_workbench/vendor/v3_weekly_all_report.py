#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全渠道周报自动生成脚本（末次归因版）
每周运行一次，自动生成上周四到本周三的全渠道周报
"""

# AUTO_PYTHON_DEPENDENCY_SWITCH
import importlib.util as _dependency_importlib_util
import os as _dependency_os
import sys as _dependency_sys
from pathlib import Path as _DependencyPath

_DEPENDENCY_PYTHON = _DependencyPath(
    "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
)

def _dependency_missing(name):
    try:
        return _dependency_importlib_util.find_spec(name) is None
    except ModuleNotFoundError:
        return True

_missing_dependencies = [
    name for name in ("mysql.connector", "openpyxl")
    if _dependency_missing(name)
]
if _missing_dependencies:
    _current_python = _DependencyPath(_dependency_sys.executable).resolve()
    if _DEPENDENCY_PYTHON.exists() and _current_python != _DEPENDENCY_PYTHON.resolve():
        _dependency_os.execv(
            str(_DEPENDENCY_PYTHON),
            [str(_DEPENDENCY_PYTHON), *_dependency_sys.argv],
        )
    raise SystemExit(
        "缺少 Python 依赖：" + ", ".join(_missing_dependencies)
        + "。请运行：python3 -m pip install mysql-connector-python openpyxl"
    )

import mysql.connector
from mysql.connector import Error
from datetime import datetime, timedelta
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import os

# ============== 数据库配置 ==============
DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'user': os.getenv('DB_USER', 'root'),
    'password': os.getenv('DB_PASSWORD', ''),
    'database': os.getenv('DB_NAME', 'aippt')
}

# ============== 输出路径配置 ==============
OUTPUT_DIR = os.getenv('OUTPUT_DIR', '/tmp')

# ============== 目标日均配置（每月初手动更新）==============
# 格式: {'YYYY-MM': {指标名: 目标值, ...}}
# 未配置的指标显示 "-"
# 数值类指标填整数，百分比类指标填小数（如 5.5 表示 5.5%）
MONTHLY_TARGETS = {
    '2026-04': {
        '当日新注册用户ROI': 92,      # 百分比
        '流量': 515789,               # 整数
        '新增注册': 128947,           # 整数
        '注册率': 25,                 # 百分比
        '客单价': 19,                 # 美元整数
        '订单量': 5158,               # 整数（付费量）
        '老用户订单占比': 35,         # 百分比
        '老用户收入占比': 33,         # 百分比
        '注册付费率': 4,              # 百分比
        'PC端': 4.5,                  # 百分比
        '移动端': 1,                  # 百分比
        '付费率（流量口径）': 1.86,   # 百分比
        '退款金额': None,             # 美元整数
        '营收总计': 98000,            # 美元整数
        'SEM': 60000,                 # 美元整数
        'KOL': 8000,                  # 美元整数
        'Affiliate': None,             # 美元整数
        'Auto_Renewal': None,          # 美元整数
        'SEO': None,                   # 美元整数
        '消耗金额': 122500,           # 美元整数
        'SEM消耗': 82500,              # 美元整数
        'KOL消耗': 40000,              # 美元整数
    },
}


def get_db_connection():
    """创建数据库连接"""
    try:
        connection = mysql.connector.connect(**DB_CONFIG)
        return connection
    except Error as e:
        print(f"数据库连接错误: {e}")
        raise


def calculate_dates(this_week_start=None, this_week_end=None):
    """
    计算所需的各种日期
    可手动传入 本周开始和结束日期，不传则自动按昨天往前7天算
    上周：再往前7天
    累计：本月1号到本周结束
    """
    if this_week_start and this_week_end:
        this_week_start = datetime.strptime(this_week_start, '%Y-%m-%d').date() if isinstance(this_week_start, str) else this_week_start
        this_week_end = datetime.strptime(this_week_end, '%Y-%m-%d').date() if isinstance(this_week_end, str) else this_week_end
    else:
        today = datetime.now().date()
        this_week_end = today - timedelta(days=1)  # 昨天
        this_week_start = this_week_end - timedelta(days=6)

    week_end = this_week_end

    # 上周：本周起止日期分别往前推 7 天。
    # 不能只根据结束日倒推固定 7 天，否则手动输入非 7 天周期时会算错开始日。
    last_week_start = this_week_start - timedelta(days=7)
    last_week_end = this_week_end - timedelta(days=7)

    # 累计范围：本月1号到本周结束
    month_start = week_end.replace(day=1)
    month_end = week_end

    # 计算当月天数和时间进度
    _, days_in_month = calendar.monthrange(week_end.year, week_end.month)
    days_passed = week_end.day

    return {
        'this_week_start': this_week_start,
        'this_week_end': this_week_end,
        'last_week_start': last_week_start,
        'last_week_end': last_week_end,
        'month_start': month_start,
        'month_end': month_end,
        'days_in_month': days_in_month,
        'days_passed': days_passed,
    }


def execute_query(connection, query, params=None):
    """执行SQL查询并返回结果"""
    cursor = connection.cursor()
    try:
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        result = cursor.fetchone()
        return result
    finally:
        cursor.close()


# ===== 地区分组定义 =====
ENGLISH_COUNTRIES = (
    "'美国','英国','澳大利亚','加拿大','印度','其他',"
    "'越南','沙特','阿联酋','荷兰','缅甸','土耳其','埃及','卡塔尔','瑞典','科威特','柬埔寨','巴基斯坦','瑞士','意大利'"
)
TRADITIONAL_CHINESE_COUNTRIES = "'中国香港','中国澳门','中国台湾'"
OTHER_COUNTRIES = (
    "'印度尼西亚','马来西亚','菲律宾','巴西','日本','秘鲁','阿根廷','墨西哥','葡萄牙','智利',"
    "'厄瓜多尔','新加坡','巴拿马','西班牙','法国','玻利维亚','危地马拉','泰国','牙买加','乌拉圭',"
    "'多米尼加','哥伦比亚','韩国','特立尼达和多巴哥','萨尔瓦多','俄罗斯','委内瑞拉','洪都拉斯','巴拉圭','尼加拉瓜'"
)


# 新用户收入归因优先级：合作渠道中的 KOL / Affiliate -> 末次来源。
# 只有合作渠道为空时才按末次来源分类。
# Direct 和 Social_Organic 归入 KOL，Referral 归入 SEO。
# Unassigned 和空白末次来源单独列示。
# 空白末次归因在报表中显示为“教育认证购买页新用户收入”。
REVENUE_ATTRIBUTION_CASE = """
CASE
    WHEN UPPER(TRIM(COALESCE(合作渠道, ''))) LIKE 'KOL%' THEN 'KOL'
    WHEN UPPER(TRIM(COALESCE(合作渠道, ''))) LIKE 'AFFILIATE%' THEN 'Affiliate'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'SEM' THEN 'SEM'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'KOL' THEN 'KOL'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'DIRECT' THEN 'KOL'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'SEO' THEN 'SEO'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'AUTO_RENEWAL' THEN 'Auto_Renewal'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) LIKE 'REFERRAL%' THEN 'SEO'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) LIKE 'AFFILIATE%' THEN 'Affiliate'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'UNASSIGNED' THEN 'Unassigned'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'SOCIAL_ORGANIC' THEN 'KOL'
    WHEN NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
         AND NULLIF(TRIM(COALESCE(末次来源类型, '')), '') IS NULL THEN 'Blank'
    ELSE 'SEO'
END
"""


def _build_d0_rev_query(cycle_cond, extra_cond="", attribution=None):
    """构建D0营收查询；可按新的8类收入归因筛选。"""
    conds = [cycle_cond]
    if extra_cond:
        conds.append(extra_cond)
    if attribution:
        conds.append(f"({REVENUE_ATTRIBUTION_CASE}) = '{attribution}'")
    where_extra = " AND " + " AND ".join(conds) if conds else ""
    return f"""
        COALESCE(SUM(CASE WHEN (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%'){where_extra}
                      THEN 订单金额 ELSE 0 END), 0)
    """


def _get_d0_rev(connection, start_str, end_str, cycle_cond, extra_cond="", attribution=None):
    """执行单次D0营收查询。"""
    query = f"""
    SELECT {_build_d0_rev_query(cycle_cond, extra_cond, attribution)}
    FROM ordersof4
    WHERE DATE(用户注册时间) BETWEEN %s AND %s
    """
    result = execute_query(connection, query, (start_str, end_str))
    return float(result[0] or 0) if result else 0


def _get_sem_cost(connection, start_str, end_str):
    """SEM消耗（Google排除KOL + Bing）"""
    query = """
    SELECT
        (SELECT COALESCE(SUM(费用), 0) FROM google__ym_ads
         WHERE DATE(`天`) BETWEEN %s AND %s AND 地理位置 = '总计：地理位置' AND `广告系列` NOT LIKE '%KOL%'
           AND `广告系列` NOT LIKE '%generic%') +
        (SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
         WHERE DATE(`stat_date`) BETWEEN %s AND %s) AS sem_cost
    """
    result = execute_query(connection, query, (start_str, end_str, start_str, end_str))
    return float(result[0] or 0) if result else 0


def _get_sem_cost_by_region(connection, start_str, end_str):
    """SEM消耗按地区拆分：英语 / 繁中 / 其他（Google排除KOL + Bing）"""
    # Google 分地区消耗（排除KOL）
    google_query = """
    SELECT
        COALESCE(SUM(CASE WHEN 地理位置 IN (
            '美国','英国','澳大利亚','加拿大','印度','越南',
            '沙特阿拉伯','阿拉伯联合酋长国','荷兰','缅甸','土耳其','埃及',
            '卡塔尔','瑞典','科威特','柬埔寨','巴基斯坦','瑞士','意大利'
        ) THEN 费用 ELSE 0 END), 0) AS 英语,
        COALESCE(SUM(CASE WHEN 地理位置 IN ('香港','台湾','澳门') THEN 费用 ELSE 0 END), 0) AS 繁中,
        COALESCE(SUM(CASE WHEN 地理位置 = '总计：地理位置' THEN 费用 ELSE 0 END), 0) AS 总计
    FROM google__ym_ads
    WHERE DATE(`天`) BETWEEN %s AND %s AND `广告系列` NOT LIKE '%KOL%'
      AND `广告系列` NOT LIKE '%generic%'
    """
    result = execute_query(connection, google_query, (start_str, end_str))
    g_en = float(result[0] or 0) if result else 0
    g_tc = float(result[1] or 0) if result else 0
    g_total = float(result[2] or 0) if result else 0
    g_other = g_total - g_en - g_tc

    # Bing 分地区消耗
    bing_query = """
    SELECT
        COALESCE(SUM(CASE WHEN region IN (
            'United States', 'Canada', 'India', 'Vietnam',
            'United Arab Emirates', 'Myanmar', 'Türkiye',
            'State of Qatar', 'Sweden', 'Kuwait',
            'Kingdom of Cambodia', 'Switzerland'
        ) THEN cost ELSE 0 END), 0) AS 英语,
        COALESCE(SUM(CASE WHEN region IN ('Hong Kong SAR', 'Macao SAR', 'Taiwan') THEN cost ELSE 0 END), 0) AS 繁中,
        COALESCE(SUM(cost), 0) AS 总计
    FROM bing_region_cost
    WHERE DATE(`stat_date`) BETWEEN %s AND %s
    """
    result = execute_query(connection, bing_query, (start_str, end_str))
    b_en = float(result[0] or 0) if result else 0
    b_tc = float(result[1] or 0) if result else 0
    b_total = float(result[2] or 0) if result else 0
    b_other = b_total - b_en - b_tc

    return {
        '英语': g_en + b_en,
        '繁中': g_tc + b_tc,
        '其他': g_other + b_other,
    }


def get_all_metrics(connection, dates):
    """
    执行全套SQL获取所有指标数据
    """
    tw_start = dates['this_week_start'].strftime('%Y-%m-%d')
    tw_end = dates['this_week_end'].strftime('%Y-%m-%d')
    lw_start = dates['last_week_start'].strftime('%Y-%m-%d')
    lw_end = dates['last_week_end'].strftime('%Y-%m-%d')
    mtd_start = dates['month_start'].strftime('%Y-%m-%d')
    mtd_end = dates['month_end'].strftime('%Y-%m-%d')

    metrics = {}
    periods = [('本周', tw_start, tw_end), ('上周', lw_start, lw_end), ('MTD', mtd_start, mtd_end)]

    # ===== 1. D0 营收 =====
    for label, s, e in periods:
        # D0 综合（注册转化周期 = 0，即24小时内）
        metrics[f'D0营收_{label}'] = _get_d0_rev(connection, s, e, '注册转化周期 = 0')
        # D0 SEM 使用新的末次归因口径
        metrics[f'D0_SEM营收_{label}'] = _get_d0_rev(
            connection, s, e, '注册转化周期 = 0', attribution='SEM'
        )
        # D7 SEM：统计当期注册用户在 0～7 天内的 SEM 末次归因收入。
        metrics[f'D7_SEM营收_{label}'] = _get_d0_rev(
            connection, s, e, '注册转化周期 BETWEEN 0 AND 7', attribution='SEM'
        )
        # D0 英语地区
        metrics[f'D0_英语营收_{label}'] = _get_d0_rev(
            connection, s, e, '注册转化周期 = 0',
            extra_cond=f'国家 IN ({ENGLISH_COUNTRIES})'
        )
        # D0 繁中地区
        metrics[f'D0_繁中营收_{label}'] = _get_d0_rev(
            connection, s, e, '注册转化周期 = 0',
            extra_cond=f'国家 IN ({TRADITIONAL_CHINESE_COUNTRIES})'
        )

    # 其他地区营收 = 总营收 - 英语 - 繁中（不硬编码国家列表）
    for label in ['本周', '上周', 'MTD']:
        metrics[f'D0_其他营收_{label}'] = metrics[f'D0营收_{label}'] - metrics[f'D0_英语营收_{label}'] - metrics[f'D0_繁中营收_{label}']

    # ===== 3. 新增注册 =====
    for label, s, e in periods:
        result = execute_query(connection, """
            SELECT COUNT(*) FROM registrationof4
            WHERE DATE(注册时间) BETWEEN %s AND %s
              AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
        """, (s, e))
        metrics[f'新增注册_{label}'] = result[0] or 0 if result else 0

    # ===== 4. 付费量 =====
    for label, s, e in periods:
        result = execute_query(connection, """
            SELECT COUNT(*) FROM ordersof4
            WHERE DATE(订单生成时间) BETWEEN %s AND %s
              AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
        """, (s, e))
        metrics[f'付费量_{label}'] = result[0] or 0 if result else 0

    # ===== 5. 注册分端 =====
    for label, s, e in periods:
        result = execute_query(connection, """
            SELECT
              COUNT(CASE WHEN `注册端口` LIKE '%pc%' OR `注册端口` LIKE '%其他%' THEN 注册时间 END),
              COUNT(CASE WHEN `注册端口` LIKE '%移动%' THEN 注册时间 END)
            FROM registrationof4
            WHERE DATE(注册时间) BETWEEN %s AND %s
              AND (注册端口 LIKE '%pc%' OR 注册端口 LIKE '%移动%' OR 注册端口 LIKE '%其他%')
        """, (s, e))
        metrics[f'注册PC端_{label}'] = result[0] or 0 if result else 0
        metrics[f'注册移动端_{label}'] = result[1] or 0 if result else 0

    # ===== 6. 付费量分端 =====
    for label, s, e in periods:
        result = execute_query(connection, """
            SELECT
              COUNT(CASE WHEN 支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%' THEN 订单金额 END),
              COUNT(CASE WHEN 支付端口 LIKE '%移动%' THEN 订单金额 END)
            FROM ordersof4
            WHERE DATE(订单生成时间) BETWEEN %s AND %s
              AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
        """, (s, e))
        metrics[f'付费量PC端_{label}'] = result[0] or 0 if result else 0
        metrics[f'付费量移动端_{label}'] = result[1] or 0 if result else 0

    # ===== 7. 新用户末次归因营收/订单量分端 =====
    # 与上方子渠道新用户收入保持一致：订单日期口径、注册转化周期=0、排除 Auto_Renewal。
    for label, s, e in periods:
        result = execute_query(connection, f"""
            SELECT
              COALESCE(SUM(CASE WHEN 支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%' THEN 订单金额 ELSE 0 END), 0),
              COALESCE(SUM(CASE WHEN 支付端口 LIKE '%移动%' THEN 订单金额 ELSE 0 END), 0),
              COALESCE(SUM(订单金额), 0),
              COUNT(CASE WHEN 支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%' THEN 1 END),
              COUNT(CASE WHEN 支付端口 LIKE '%移动%' THEN 1 END)
            FROM (
                SELECT 订单金额, 支付端口, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND 注册转化周期 = 0
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
            WHERE 收入归因 <> 'Auto_Renewal'
        """, (s, e))
        metrics[f'营收PC端_{label}'] = float(result[0] or 0) if result else 0
        metrics[f'营收移动端_{label}'] = float(result[1] or 0) if result else 0
        metrics[f'总营收_{label}'] = float(result[2] or 0) if result else 0
        metrics[f'新用户付费量PC端_{label}'] = result[3] or 0 if result else 0
        metrics[f'新用户付费量移动端_{label}'] = result[4] or 0 if result else 0

        # WEB—新老综合ROI（除订阅）分子：按末次归因排除 Auto_Renewal。
        non_subscription_result = execute_query(connection, f"""
            SELECT
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(订单金额), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal'
                                   AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%')
                                  THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' AND 支付端口 LIKE '%移动%'
                                  THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal'
                                   AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%')
                                  THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' AND 支付端口 LIKE '%移动%'
                                  THEN 订单金额 ELSE 0 END), 0)
            FROM (
                SELECT 订单金额, 支付端口, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
        """, (s, e))
        metrics[f'除订阅总营收_{label}'] = float(non_subscription_result[0] or 0) if non_subscription_result else 0
        metrics[f'订阅总营收_{label}'] = float(non_subscription_result[1] or 0) if non_subscription_result else 0
        metrics[f'含订阅总营收_{label}'] = float(non_subscription_result[2] or 0) if non_subscription_result else 0
        metrics[f'除订阅营收PC端_{label}'] = float(non_subscription_result[3] or 0) if non_subscription_result else 0
        metrics[f'除订阅营收移动端_{label}'] = float(non_subscription_result[4] or 0) if non_subscription_result else 0
        metrics[f'订阅营收PC端_{label}'] = float(non_subscription_result[5] or 0) if non_subscription_result else 0
        metrics[f'订阅营收移动端_{label}'] = float(non_subscription_result[6] or 0) if non_subscription_result else 0

    # ===== 8. 新用户营收分来源（注册转化周期=0）=====
    for label, s, e in periods:
        result = execute_query(connection, f"""
            SELECT
                COALESCE(SUM(CASE WHEN 收入归因 = 'SEM' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'KOL' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Affiliate' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'SEO' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Referral' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Unassigned' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Direct' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Social_Organic' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Blank' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(订单金额), 0)
            FROM (
                SELECT 订单金额, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND 注册转化周期 = 0
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
        """, (s, e))
        metrics[f'SEM营收_{label}'] = float(result[0] or 0) if result else 0
        metrics[f'KOL营收_{label}'] = float(result[1] or 0) if result else 0
        metrics[f'Affiliate营收_{label}'] = float(result[2] or 0) if result else 0
        metrics[f'Auto_Renewal营收_{label}'] = float(result[3] or 0) if result else 0
        metrics[f'SEO营收_{label}'] = float(result[4] or 0) if result else 0
        metrics[f'Referral营收_{label}'] = float(result[5] or 0) if result else 0
        metrics[f'Unassigned营收_{label}'] = float(result[6] or 0) if result else 0
        metrics[f'Direct营收_{label}'] = float(result[7] or 0) if result else 0
        metrics[f'Social_Organic营收_{label}'] = float(result[8] or 0) if result else 0
        metrics[f'Blank营收_{label}'] = float(result[9] or 0) if result else 0
        metrics[f'归因总营收_{label}'] = float(result[10] or 0) if result else 0

        # 新老用户末次归因渠道收入（不限注册周期，排除 Auto_Renewal）。
        all_user_result = execute_query(connection, f"""
            SELECT
                COALESCE(SUM(CASE WHEN 收入归因 = 'SEM' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'KOL' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Affiliate' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'SEO' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Unassigned' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Blank' THEN 订单金额 ELSE 0 END), 0)
            FROM (
                SELECT 订单金额, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
            WHERE 收入归因 <> 'Auto_Renewal'
        """, (s, e))
        for index, channel in enumerate(['SEM', 'KOL', 'Affiliate', 'SEO', 'Unassigned', 'Blank']):
            metrics[f'新老除订阅_{channel}营收_{label}'] = float(all_user_result[index] or 0) if all_user_result else 0

    # ===== 9. 新用户末次归因营收（按英语/繁中/其他分组）=====
    # 与上方新用户总收入口径一致：注册转化周期=0，按末次归因，不计 Auto_Renewal。
    for label, s, e in periods:
        result = execute_query(connection, f"""
            SELECT
                COALESCE(SUM(CASE WHEN 国家 IN ({ENGLISH_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 国家 IN ({TRADITIONAL_CHINESE_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(订单金额), 0),
                COALESCE(SUM(CASE WHEN 国家 IN ('美国','加拿大') THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 国家 = '印度' THEN 订单金额 ELSE 0 END), 0)
            FROM (
                SELECT 订单金额, 国家, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND 注册转化周期 = 0
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
            WHERE 收入归因 <> 'Auto_Renewal'
        """, (s, e))
        metrics[f'地区营收_英语_{label}'] = float(result[0] or 0) if result else 0
        metrics[f'地区营收_繁中_{label}'] = float(result[1] or 0) if result else 0
        attributed_region_total = float(result[2] or 0) if result else 0
        metrics[f'地区营收_北美_{label}'] = float(result[3] or 0) if result else 0
        metrics[f'地区营收_印度_{label}'] = float(result[4] or 0) if result else 0
        # 其他地区 = 新用户末次归因总收入 - 英语 - 繁中。
        metrics[f'地区营收_其他地区_{label}'] = (
            attributed_region_total
            - metrics[f'地区营收_英语_{label}']
            - metrics[f'地区营收_繁中_{label}']
        )

        # 新老用户地区收入：同时返回除订阅和 Auto_Renewal，含订阅用两者相加。
        all_region_result = execute_query(connection, f"""
            SELECT
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' AND 国家 IN ({ENGLISH_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' AND 国家 IN ({TRADITIONAL_CHINESE_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' AND 国家 IN ('美国','加拿大') THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 <> 'Auto_Renewal' AND 国家 = '印度' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' AND 国家 IN ({ENGLISH_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' AND 国家 IN ({TRADITIONAL_CHINESE_COUNTRIES}) THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' AND 国家 IN ('美国','加拿大') THEN 订单金额 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN 收入归因 = 'Auto_Renewal' AND 国家 = '印度' THEN 订单金额 ELSE 0 END), 0)
            FROM (
                SELECT 订单金额, 国家, {REVENUE_ATTRIBUTION_CASE} AS 收入归因
                FROM ordersof4
                WHERE DATE(订单生成时间) BETWEEN %s AND %s
                  AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
            ) AS attributed_orders
        """, (s, e))
        for prefix, offset in [('新老除订阅', 0), ('订阅', 5)]:
            english = float(all_region_result[offset] or 0) if all_region_result else 0
            traditional = float(all_region_result[offset + 1] or 0) if all_region_result else 0
            total = float(all_region_result[offset + 2] or 0) if all_region_result else 0
            metrics[f'{prefix}地区_英语_{label}'] = english
            metrics[f'{prefix}地区_繁中_{label}'] = traditional
            metrics[f'{prefix}地区_北美_{label}'] = float(all_region_result[offset + 3] or 0) if all_region_result else 0
            metrics[f'{prefix}地区_印度_{label}'] = float(all_region_result[offset + 4] or 0) if all_region_result else 0
            metrics[f'{prefix}地区_其他_{label}'] = total - english - traditional

    # ===== 10. SEM 消耗 =====
    for label, s, e in periods:
        metrics[f'SEM消耗_{label}'] = _get_sem_cost(connection, s, e)
        region_costs = _get_sem_cost_by_region(connection, s, e)
        metrics[f'SEM消耗_英语_{label}'] = region_costs['英语']
        metrics[f'SEM消耗_繁中_{label}'] = region_costs['繁中']
        metrics[f'SEM消耗_其他_{label}'] = region_costs['其他']

    # ===== 11. 流量相关 =====
    for label, s, e in periods:
        result = execute_query(connection, """
            SELECT COALESCE(SUM(点击次数), 0) FROM google__ym_ads
            WHERE DATE(`天`) BETWEEN %s AND %s AND 地理位置 = '总计：搜索'
        """, (s, e))
        metrics[f'Google搜索点击_{label}'] = float(result[0] or 0) if result else 0

        result = execute_query(connection, """
            SELECT COALESCE(SUM(clicks), 0) FROM bing_region_cost
            WHERE DATE(`stat_date`) BETWEEN %s AND %s AND ad_campaign = 'Search total'
        """, (s, e))
        metrics[f'Bing全部点击_{label}'] = float(result[0] or 0) if result else 0

    return metrics


def calculate_flow_metrics(metrics):
    """
    根据规则计算流量指标
    流量 = (Google流量 + Bing流量 + 固定值) × 0.7（周维度）
    累计流量不乘0.7
    """
    # 本周流量（乘0.7）
    metrics['流量_本周'] = (metrics['Google搜索点击_本周'] + metrics['Bing全部点击_本周'] + 11186 * 7) * 0.7

    # 上周流量（乘0.7）
    metrics['流量_上周'] = (metrics['Google搜索点击_上周'] + metrics['Bing全部点击_上周'] + 11186 * 7) * 0.7

    # 累计流量（不乘0.7）
    metrics['流量_MTD'] = metrics['Google搜索点击_MTD'] + metrics['Bing全部点击_MTD'] + 228420

    return metrics


def create_report(dates, metrics,
                  commission_cost_this_week, commission_cost_last_week, commission_cost_mtd,
                  seo_cost_this_week, seo_cost_last_week, seo_cost_mtd,
                  kol_cost_this_week, kol_cost_last_week, kol_cost_mtd,
                  pc_refund_this_week, pc_refund_last_week, pc_refund_mtd,
                  monthly_targets):

    wb = Workbook()
    ws = wb.active
    ws.title = "全渠道周报"

    # 定义样式
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')

    # 标题行
    ws.merge_cells('A1:H1')
    ws['A1'] = "全渠道周报"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = center_align

    # 日期范围行
    date_range = f"（{dates['this_week_start'].strftime('%Y.%m.%d')}-{dates['this_week_end'].strftime('%Y.%m.%d')}）"
    ws.merge_cells('A2:H2')
    ws['A2'] = date_range
    ws['A2'].font = Font(size=12)
    ws['A2'].alignment = center_align

    # 表头
    headers = ['指标', '本周', '上周', '环比', 'MTD', '目标', '完成率', '时间进度']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align

    # 计算环比
    def calc_mom(val_this, val_last):
        if val_last and val_last != 0:
            return (val_this - val_last) / val_last
        return None

    # 时间进度
    time_progress = dates['days_passed'] / dates['days_in_month']

    # 获取目标
    def get_target(name):
        return monthly_targets.get(name) if monthly_targets else None

    # 计算完成率
    def calc_completion(mtd_val, target_val, is_pct=False):
        if target_val is None or target_val == 0:
            return None
        if mtd_val is None:
            return None
        if is_pct:
            return (mtd_val * 100) / target_val
        else:
            return mtd_val / target_val

    # 格式化函数
    def format_int(val):
        return int(round(val)) if val is not None else "-"

    def format_pct(val):
        return f"{int(round(val * 100))}%" if val is not None else "-"

    def format_pct_2decimal(val):
        return f"{round(val * 100, 2)}%" if val is not None else "-"

    def format_dollar(val):
        return f"${int(round(val))}" if val is not None else "-"

    def format_mom(val):
        return f"{int(round(val * 100))}%" if val is not None else "-"

    def format_target(val, fmt_type):
        if val is None:
            return "-"
        if fmt_type == 'pct':
            return f"{val}%"
        elif fmt_type == 'dollar':
            return f"${int(val)}"
        else:
            return int(val)

    def format_completion(rate):
        if rate is None:
            return "-"
        return f"{int(round(rate * 100))}%"

    # ===== 计算衍生指标 =====

    # 总成本 = SEM成本 + 分佣成本 + SEO外链成本 + KOL成本
    sem_tw = metrics['SEM消耗_本周']
    sem_lw = metrics['SEM消耗_上周']
    sem_mtd = metrics['SEM消耗_MTD']

    total_cost_tw = sem_tw + commission_cost_this_week + seo_cost_this_week + kol_cost_this_week
    total_cost_lw = sem_lw + commission_cost_last_week + seo_cost_last_week + kol_cost_last_week
    total_cost_mtd = sem_mtd + commission_cost_mtd + seo_cost_mtd + kol_cost_mtd

    # 辅助：安全除法
    def sdiv(a, b):
        return a / b if b and b > 0 else None

    # 总营收
    total_rev_tw = metrics['总营收_本周']
    total_rev_lw = metrics['总营收_上周']
    total_rev_mtd = metrics['总营收_MTD']

    # 新用户总收入 = SEM + KOL + Affiliate + SEO + Unassigned + Blank。
    # 新用户不统计 Auto_Renewal；Direct/Social_Organic 已归入 KOL，Referral 已归入 SEO。
    revenue_keys = [
        'SEM营收', 'KOL营收', 'Affiliate营收',
        'SEO营收', 'Unassigned营收', 'Blank营收'
    ]
    rev_total_tw = sum(metrics[f'{key}_本周'] for key in revenue_keys)
    rev_total_lw = sum(metrics[f'{key}_上周'] for key in revenue_keys)
    rev_total_mtd = sum(metrics[f'{key}_MTD'] for key in revenue_keys)

    # 自然来源新用户收入 = KOL + SEO + Unassigned + Affiliate + Blank。
    natural_revenue_keys = [
        'SEO营收', 'KOL营收', 'Unassigned营收',
        'Affiliate营收', 'Blank营收'
    ]
    natural_rev_tw = sum(metrics[f'{key}_本周'] for key in natural_revenue_keys)
    natural_rev_lw = sum(metrics[f'{key}_上周'] for key in natural_revenue_keys)
    natural_rev_mtd = sum(metrics[f'{key}_MTD'] for key in natural_revenue_keys)

    all_natural_keys = ['SEO', 'KOL', 'Unassigned', 'Affiliate', 'Blank']
    all_natural_tw = sum(metrics[f'新老除订阅_{key}营收_本周'] for key in all_natural_keys)
    all_natural_lw = sum(metrics[f'新老除订阅_{key}营收_上周'] for key in all_natural_keys)
    all_natural_mtd = sum(metrics[f'新老除订阅_{key}营收_MTD'] for key in all_natural_keys)

    # 地区收入总和
    region_total_tw = (metrics['地区营收_英语_本周'] + metrics['地区营收_繁中_本周'] + metrics['地区营收_其他地区_本周'])
    region_total_lw = (metrics['地区营收_英语_上周'] + metrics['地区营收_繁中_上周'] + metrics['地区营收_其他地区_上周'])
    region_total_mtd = (metrics['地区营收_英语_MTD'] + metrics['地区营收_繁中_MTD'] + metrics['地区营收_其他地区_MTD'])

    # 时间进度
    time_progress = dates['days_passed'] / dates['days_in_month']

    # ===== 构建数据行 =====
    rows_data = []

    # ---------- 辅助行添加函数 ----------
    def add_row(name, val_tw, val_lw, val_mtd, fmt, target_key=None, is_pct_target=False):
        mom = calc_mom(val_tw, val_lw)
        # 网站目标使用表格中显示的完整指标名。先按当前行名匹配，
        # 再兼容旧脚本中的“SEM / 营收总计”等历史键名。
        target = get_target(target_key) if target_key else get_target(name)
        if target is None and target_key:
            target = get_target(name)
        completion = calc_completion(val_mtd, target, is_pct=is_pct_target) if target else None
        rows_data.append([
            name,
            fmt(val_tw), fmt(val_lw), format_mom(mom),
            fmt(val_mtd),
            format_target(target, 'pct' if is_pct_target else ('dollar' if fmt == format_dollar else 'int')) if target else '-',
            format_completion(completion) if completion is not None else '-',
            '-'
        ])

    def add_section_header(name):
        """添加分区标题行"""
        rows_data.append([name, '-', '-', '-', '-', '-', '-', '-'])

    # ==================== 核心 ROI ====================
    add_section_header('核心ROI（末次归因）')
    add_row('WEB—新用户ROI',
            sdiv(metrics['D0营收_本周'], total_cost_tw),
            sdiv(metrics['D0营收_上周'], total_cost_lw),
            sdiv(metrics['D0营收_MTD'], total_cost_mtd),
            format_pct, is_pct_target=True)

    add_row('WEB—新老综合ROI（除订阅）',
            sdiv(metrics['除订阅总营收_本周'], total_cost_tw),
            sdiv(metrics['除订阅总营收_上周'], total_cost_lw),
            sdiv(metrics['除订阅总营收_MTD'], total_cost_mtd), format_pct,
            is_pct_target=True)
    add_row('WEB—新老综合ROI（含订阅）',
            sdiv(metrics['含订阅总营收_本周'], total_cost_tw),
            sdiv(metrics['含订阅总营收_上周'], total_cost_lw),
            sdiv(metrics['含订阅总营收_MTD'], total_cost_mtd), format_pct,
            is_pct_target=True)

    # WEB—SEM投放新用户ROI（D7）
    add_row('WEB—SEM投放新用户ROI（D7）',
            sdiv(metrics['D7_SEM营收_本周'], sem_tw),
            sdiv(metrics['D7_SEM营收_上周'], sem_lw),
            sdiv(metrics['D7_SEM营收_MTD'], sem_mtd),
            format_pct, is_pct_target=True)

    # 地区ROI：末次归因新用户地区收入 ÷ 同期地区SEM消耗。
    add_row('WEB—英语地区新用户ROI',
            sdiv(metrics['地区营收_英语_本周'], metrics['SEM消耗_英语_本周']),
            sdiv(metrics['地区营收_英语_上周'], metrics['SEM消耗_英语_上周']),
            sdiv(metrics['地区营收_英语_MTD'], metrics['SEM消耗_英语_MTD']),
            format_pct, is_pct_target=True)

    add_row('WEB—繁中地区新用户ROI',
            sdiv(metrics['地区营收_繁中_本周'], metrics['SEM消耗_繁中_本周']),
            sdiv(metrics['地区营收_繁中_上周'], metrics['SEM消耗_繁中_上周']),
            sdiv(metrics['地区营收_繁中_MTD'], metrics['SEM消耗_繁中_MTD']),
            format_pct, is_pct_target=True)

    add_row('WEB—其他地区新用户ROI',
            sdiv(metrics['地区营收_其他地区_本周'], metrics['SEM消耗_其他_本周']),
            sdiv(metrics['地区营收_其他地区_上周'], metrics['SEM消耗_其他_上周']),
            sdiv(metrics['地区营收_其他地区_MTD'], metrics['SEM消耗_其他_MTD']),
            format_pct, is_pct_target=True)

    # ==================== 收入总览 ====================
    add_section_header('收入总览')
    add_row('新用户总收入', rev_total_tw, rev_total_lw, rev_total_mtd, format_dollar, '营收总计')
    add_row('新老用户总收入（除订阅）', metrics['除订阅总营收_本周'], metrics['除订阅总营收_上周'], metrics['除订阅总营收_MTD'], format_dollar)
    add_row('新老用户总收入（含订阅）', metrics['含订阅总营收_本周'], metrics['含订阅总营收_上周'], metrics['含订阅总营收_MTD'], format_dollar)

    # ==================== 新用户收入明细 ====================
    add_section_header('新用户收入明细（末次归因）')
    add_row('SEM新用户收入', metrics['SEM营收_本周'], metrics['SEM营收_上周'], metrics['SEM营收_MTD'], format_dollar, 'SEM')
    add_row('自然来源新用户收入', natural_rev_tw, natural_rev_lw, natural_rev_mtd, format_dollar)
    add_row('SEO新用户收入', metrics['SEO营收_本周'], metrics['SEO营收_上周'], metrics['SEO营收_MTD'], format_dollar, 'SEO')
    add_row('KOL新用户收入', metrics['KOL营收_本周'], metrics['KOL营收_上周'], metrics['KOL营收_MTD'], format_dollar, 'KOL')
    add_row('affiliate新用户收入', metrics['Affiliate营收_本周'], metrics['Affiliate营收_上周'], metrics['Affiliate营收_MTD'], format_dollar, 'Affiliate')
    add_row('教育认证新用户收入', metrics['Blank营收_本周'], metrics['Blank营收_上周'], metrics['Blank营收_MTD'], format_dollar)
    add_row('Unassigned新用户收入', metrics['Unassigned营收_本周'], metrics['Unassigned营收_上周'], metrics['Unassigned营收_MTD'], format_dollar)

    # ==================== 新老用户除订阅收入明细 ====================
    add_section_header('新老用户收入明细（末次归因除订阅）')
    add_row('SEM新老用户除订阅收入', metrics['新老除订阅_SEM营收_本周'], metrics['新老除订阅_SEM营收_上周'], metrics['新老除订阅_SEM营收_MTD'], format_dollar)
    add_row('自然来源新老用户除订阅收入', all_natural_tw, all_natural_lw, all_natural_mtd, format_dollar)
    for label_name, key in [('SEO', 'SEO'), ('KOL', 'KOL'), ('affiliate', 'Affiliate'), ('教育认证', 'Blank'), ('Unassigned', 'Unassigned')]:
        add_row(f'{label_name}新老用户除订阅收入', metrics[f'新老除订阅_{key}营收_本周'], metrics[f'新老除订阅_{key}营收_上周'], metrics[f'新老除订阅_{key}营收_MTD'], format_dollar)

    # ==================== 地区收入 ====================
    add_section_header('新用户地区收入（末次归因）')
    add_row('英语组新用户收入', metrics['地区营收_英语_本周'], metrics['地区营收_英语_上周'], metrics['地区营收_英语_MTD'], format_dollar)
    add_row('北美（美国+加拿大）新用户收入', metrics['地区营收_北美_本周'], metrics['地区营收_北美_上周'], metrics['地区营收_北美_MTD'], format_dollar)
    add_row('印度新用户收入', metrics['地区营收_印度_本周'], metrics['地区营收_印度_上周'], metrics['地区营收_印度_MTD'], format_dollar)
    add_row('繁体中文新用户收入', metrics['地区营收_繁中_本周'], metrics['地区营收_繁中_上周'], metrics['地区营收_繁中_MTD'], format_dollar)
    add_row('其他新用户收入', metrics['地区营收_其他地区_本周'], metrics['地区营收_其他地区_上周'], metrics['地区营收_其他地区_MTD'], format_dollar)

    add_section_header('新老用户地区收入（末次归因除订阅）')
    for name, key in [('英语组', '英语'), ('北美（美国+加拿大）', '北美'), ('印度', '印度'), ('繁体中文', '繁中'), ('其他地区', '其他')]:
        add_row(f'{name}新老用户除订阅收入', metrics[f'新老除订阅地区_{key}_本周'], metrics[f'新老除订阅地区_{key}_上周'], metrics[f'新老除订阅地区_{key}_MTD'], format_dollar)

    add_section_header('新老用户地区收入（末次归因含订阅）')
    for name, key in [('英语组', '英语'), ('北美（美国+加拿大）', '北美'), ('印度', '印度'), ('繁体中文', '繁中'), ('其他地区', '其他')]:
        add_row(
            f'{name}新老用户含订阅收入',
            metrics[f'新老除订阅地区_{key}_本周'] + metrics[f'订阅地区_{key}_本周'],
            metrics[f'新老除订阅地区_{key}_上周'] + metrics[f'订阅地区_{key}_上周'],
            metrics[f'新老除订阅地区_{key}_MTD'] + metrics[f'订阅地区_{key}_MTD'],
            format_dollar,
        )

    # ==================== 成本 ====================
    add_section_header('成本总览')
    add_row('总成本', total_cost_tw, total_cost_lw, total_cost_mtd, format_dollar, '消耗金额')
    add_row('SEM成本', sem_tw, sem_lw, sem_mtd, format_dollar)
    add_row('affiliate成本', commission_cost_this_week, commission_cost_last_week, commission_cost_mtd, format_dollar)
    add_row('SEO—外链成本', seo_cost_this_week, seo_cost_last_week, seo_cost_mtd, format_dollar)

    # ==================== PC端 ====================
    add_section_header('PC')

    pc_flow_tw = metrics['PC流量_本周']
    pc_flow_lw = metrics['PC流量_上周']
    pc_flow_mtd = metrics['PC流量_MTD']

    add_row('流量', pc_flow_tw, pc_flow_lw, pc_flow_mtd, format_int, 'PC—流量')
    add_row('新增注册',
            metrics['注册PC端_本周'], metrics['注册PC端_上周'], metrics['注册PC端_MTD'], format_int,
            'PC—新增注册')
    add_row('注册率',
            sdiv(metrics['注册PC端_本周'], pc_flow_tw),
            sdiv(metrics['注册PC端_上周'], pc_flow_lw),
            sdiv(metrics['注册PC端_MTD'], pc_flow_mtd),
            format_pct_2decimal, 'PC—注册率', is_pct_target=True)
    add_row('新用户客单价ARPPU',
            sdiv(metrics['营收PC端_本周'], metrics['新用户付费量PC端_本周']),
            sdiv(metrics['营收PC端_上周'], metrics['新用户付费量PC端_上周']),
            sdiv(metrics['营收PC端_MTD'], metrics['新用户付费量PC端_MTD']),
            format_dollar, 'PC—新用户客单价ARPPU')
    add_row('新用户订单量',
            metrics['新用户付费量PC端_本周'], metrics['新用户付费量PC端_上周'], metrics['新用户付费量PC端_MTD'], format_int,
            'PC—新用户订单量')
    add_row('新用户营收',
            metrics['营收PC端_本周'], metrics['营收PC端_上周'], metrics['营收PC端_MTD'], format_dollar,
            'PC—新用户营收')
    add_row('新老用户营收（除订阅）', metrics['除订阅营收PC端_本周'], metrics['除订阅营收PC端_上周'], metrics['除订阅营收PC端_MTD'], format_dollar, 'PC—新老用户营收（除订阅）')
    add_row('新老用户营收（含订阅）', metrics['除订阅营收PC端_本周'] + metrics['订阅营收PC端_本周'], metrics['除订阅营收PC端_上周'] + metrics['订阅营收PC端_上周'], metrics['除订阅营收PC端_MTD'] + metrics['订阅营收PC端_MTD'], format_dollar, 'PC—新老用户营收（含订阅）')
    add_row('新用户注册付费率',
            sdiv(metrics['新用户付费量PC端_本周'], metrics['注册PC端_本周']),
            sdiv(metrics['新用户付费量PC端_上周'], metrics['注册PC端_上周']),
            sdiv(metrics['新用户付费量PC端_MTD'], metrics['注册PC端_MTD']),
            format_pct_2decimal, 'PC—新用户注册付费率', is_pct_target=True)
    add_row('付费率 (流量口径)',
            sdiv(metrics['新用户付费量PC端_本周'], pc_flow_tw),
            sdiv(metrics['新用户付费量PC端_上周'], pc_flow_lw),
            sdiv(metrics['新用户付费量PC端_MTD'], pc_flow_mtd),
            format_pct_2decimal, 'PC—付费率 (流量口径)', is_pct_target=True)
    add_row('退款金额', pc_refund_this_week, pc_refund_last_week, pc_refund_mtd, format_dollar, 'PC—退款金额')

    # ==================== H5端 ====================
    add_section_header('H5')

    h5_flow_tw = metrics['H5流量_本周']
    h5_flow_lw = metrics['H5流量_上周']
    h5_flow_mtd = metrics['H5流量_MTD']

    add_row('流量', h5_flow_tw, h5_flow_lw, h5_flow_mtd, format_int, 'H5—流量')
    add_row('新增注册',
            metrics['注册移动端_本周'], metrics['注册移动端_上周'], metrics['注册移动端_MTD'], format_int,
            'H5—新增注册')
    add_row('注册率',
            sdiv(metrics['注册移动端_本周'], h5_flow_tw),
            sdiv(metrics['注册移动端_上周'], h5_flow_lw),
            sdiv(metrics['注册移动端_MTD'], h5_flow_mtd),
            format_pct_2decimal, 'H5—注册率', is_pct_target=True)
    add_row('新用户客单价ARPPU',
            sdiv(metrics['营收移动端_本周'], metrics['新用户付费量移动端_本周']),
            sdiv(metrics['营收移动端_上周'], metrics['新用户付费量移动端_上周']),
            sdiv(metrics['营收移动端_MTD'], metrics['新用户付费量移动端_MTD']),
            format_dollar, 'H5—新用户客单价ARPPU')
    add_row('新用户订单量',
            metrics['新用户付费量移动端_本周'], metrics['新用户付费量移动端_上周'], metrics['新用户付费量移动端_MTD'], format_int,
            'H5—新用户订单量')
    add_row('新用户营收',
            metrics['营收移动端_本周'], metrics['营收移动端_上周'], metrics['营收移动端_MTD'], format_dollar,
            'H5—新用户营收')
    add_row('新老用户营收（除订阅）', metrics['除订阅营收移动端_本周'], metrics['除订阅营收移动端_上周'], metrics['除订阅营收移动端_MTD'], format_dollar, 'H5—新老用户营收（除订阅）')
    add_row('新老用户营收（含订阅）', metrics['除订阅营收移动端_本周'] + metrics['订阅营收移动端_本周'], metrics['除订阅营收移动端_上周'] + metrics['订阅营收移动端_上周'], metrics['除订阅营收移动端_MTD'] + metrics['订阅营收移动端_MTD'], format_dollar, 'H5—新老用户营收（含订阅）')
    add_row('新用户注册付费率',
            sdiv(metrics['新用户付费量移动端_本周'], metrics['注册移动端_本周']),
            sdiv(metrics['新用户付费量移动端_上周'], metrics['注册移动端_上周']),
            sdiv(metrics['新用户付费量移动端_MTD'], metrics['注册移动端_MTD']),
            format_pct_2decimal, 'H5—新用户注册付费率', is_pct_target=True)
    add_row('付费率 (流量口径)',
            sdiv(metrics['新用户付费量移动端_本周'], h5_flow_tw),
            sdiv(metrics['新用户付费量移动端_上周'], h5_flow_lw),
            sdiv(metrics['新用户付费量移动端_MTD'], h5_flow_mtd),
            format_pct_2decimal, 'H5—付费率 (流量口径)', is_pct_target=True)

    # 时间进度（PC订单量行，PC端第5行 = 总数据(1) + WEB(11) + 收入(5) + 产品线header(1) + 产品线4行(4) + 地区(4) + 成本(5) + PC端header(1) + PC行5 = 总行36，即index 35）
    # 改为H5订单量行
    # H5订单量行: 上面所有行 = 总数据(1) + WEB(11) + 收入(5) + 产品线(1+4) + 地区(4) + 成本(5) + PC(1+9) + H5(1+5) = 46 即 index 45
    # 不写死索引，后面循环中处理

    # 写入数据行
    row_num = 4
    for row_data in rows_data:
        is_section_header = row_data[1:] == ['-', '-', '-', '-', '-', '-', '-']
        for col, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col, value=value)
            cell.border = border
            cell.alignment = center_align
            if is_section_header:
                cell.fill = section_fill
                cell.font = Font(bold=True, color="1F4E78")
            elif col == 1:
                cell.font = Font(bold=True)
        row_num += 1

    # 合并时间进度列（第8列）
    end_row = row_num - 1
    ws.merge_cells(start_row=4, start_column=8, end_row=end_row, end_column=8)
    time_progress_cell = ws.cell(row=4, column=8, value=f"{int(round(time_progress * 100))}%")
    time_progress_cell.alignment = center_align
    time_progress_cell.border = border

    # 调整列宽
    column_widths = [28, 14, 14, 10, 14, 12, 10, 10]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    return wb


def _input_cost(label):
    """安全输入成本"""
    try:
        return float(input(f"请输入{label}: "))
    except ValueError:
        print("输入无效，默认使用 0")
        return 0


def main():
    """主函数"""
    print("=" * 50)
    print("全渠道周报自动生成脚本（末次归因版）")
    print("=" * 50)

    # 计算日期（手动输入或回车跳过自动计算）
    print("\n请输入统计周期（格式 YYYY-MM-DD，直接回车则自动按昨天往前7天计算）:")
    tw_start = input("  本周开始日期: ").strip()
    tw_end = input("  本周结束日期: ").strip()
    if tw_start and tw_end:
        dates = calculate_dates(this_week_start=tw_start, this_week_end=tw_end)
    else:
        dates = calculate_dates()
    print(f"\n本周范围: {dates['this_week_start']} 至 {dates['this_week_end']}")
    print(f"上周范围: {dates['last_week_start']} 至 {dates['last_week_end']}")
    print(f"累计范围: {dates['month_start']} 至 {dates['month_end']}")

    # 获取目标配置
    month_key = dates['month_end'].strftime('%Y-%m')
    monthly_targets = MONTHLY_TARGETS.get(month_key, {})
    if monthly_targets:
        print(f"本月目标已配置: {len([v for v in monthly_targets.values() if v is not None])}个指标")
    else:
        print("本月目标: 未配置")

    # 输入成本和退款金额
    print("\n请输入以下数据:")

    # 分佣成本
    commission_cost_this_week = _input_cost("本周 分佣成本")
    commission_cost_last_week = _input_cost("上周 分佣成本")
    commission_cost_mtd = _input_cost("累计 分佣成本")

    # SEO—外链成本
    seo_cost_this_week = _input_cost("本周 SEO—外链成本")
    seo_cost_last_week = _input_cost("上周 SEO—外链成本")
    seo_cost_mtd = _input_cost("累计 SEO—外链成本")

    # KOL成本
    kol_cost_this_week = _input_cost("本周 KOL成本")
    kol_cost_last_week = _input_cost("上周 KOL成本")
    kol_cost_mtd = _input_cost("累计 KOL成本")

    # PC / H5 流量分别手动输入
    pc_flow_this_week = _input_cost("本周 PC流量")
    pc_flow_last_week = _input_cost("上周 PC流量")
    pc_flow_mtd = _input_cost("累计 PC流量")

    h5_flow_this_week = _input_cost("本周 H5流量")
    h5_flow_last_week = _input_cost("上周 H5流量")
    h5_flow_mtd = _input_cost("累计 H5流量")

    # PC端退款金额
    pc_refund_this_week = _input_cost("本周 PC端退款金额")
    pc_refund_last_week = _input_cost("上周 PC端退款金额")
    pc_refund_mtd = _input_cost("累计 PC端退款金额")

    refund_this_week = pc_refund_this_week

    # 连接数据库
    print("\n正在连接数据库...")
    connection = get_db_connection()
    print("数据库连接成功")

    try:
        # 获取所有指标数据
        print("\n正在获取数据...")
        metrics = get_all_metrics(connection, dates)

        # PC / H5 流量使用手动输入值。
        metrics['PC流量_本周'] = pc_flow_this_week
        metrics['PC流量_上周'] = pc_flow_last_week
        metrics['PC流量_MTD'] = pc_flow_mtd
        metrics['H5流量_本周'] = h5_flow_this_week
        metrics['H5流量_上周'] = h5_flow_last_week
        metrics['H5流量_MTD'] = h5_flow_mtd
        metrics['流量_本周'] = pc_flow_this_week + h5_flow_this_week
        metrics['流量_上周'] = pc_flow_last_week + h5_flow_last_week
        metrics['流量_MTD'] = pc_flow_mtd + h5_flow_mtd

        print("数据获取完成")

        # 生成报告
        print("\n正在生成报告...")
        wb = create_report(dates, metrics,
                          commission_cost_this_week, commission_cost_last_week, commission_cost_mtd,
                          seo_cost_this_week, seo_cost_last_week, seo_cost_mtd,
                          kol_cost_this_week, kol_cost_last_week, kol_cost_mtd,
                          pc_refund_this_week, pc_refund_last_week, pc_refund_mtd,
                          monthly_targets)

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # 保存文件
        output_file = os.path.join(OUTPUT_DIR, f"周报-全渠道_末次归因_{dates['this_week_start'].strftime('%Y.%m.%d')}-{dates['this_week_end'].strftime('%Y.%m.%d')}.xlsx")
        wb.save(output_file)
        print(f"\n报告已保存: {output_file}")

        # 打印摘要
        print("\n" + "=" * 50)
        print("数据摘要:")
        print(f"  本周流量: {int(round(metrics['流量_本周']))}")
        print(f"  本周新增注册: {metrics['新增注册_本周']}")
        print(f"  本周付费量: {metrics['付费量_本周']}")
        print(f"  本周总营收: ${int(round(metrics['总营收_本周']))}")
        print(f"  SEM新用户收入: ${int(round(metrics['SEM营收_本周']))}")
        print(f"  KOL新用户收入: ${int(round(metrics['KOL营收_本周']))}")
        print(f"  分佣新用户收入: ${int(round(metrics['Affiliate营收_本周']))}")
        print(f"  SEO新用户收入: ${int(round(metrics['SEO营收_本周']))}")
        print(f"  Unassigned新用户收入: ${int(round(metrics['Unassigned营收_本周']))}")
        print(f"  教育认证购买页新用户收入: ${int(round(metrics['Blank营收_本周']))}")
        total_cost_tw = metrics['SEM消耗_本周'] + commission_cost_this_week + seo_cost_this_week + kol_cost_this_week
        print(f"  本周总成本: ${int(round(total_cost_tw))}")
        print(f"  本周退款金额: ${int(round(refund_this_week))}")
        target_paid = monthly_targets.get('订单量') if monthly_targets else None
        if target_paid and target_paid > 0:
            completion_rate = metrics['付费量_MTD'] / target_paid
            print(f"  累计完成率: {int(round(completion_rate * 100))}%")
        print(f"  时间进度: {int(round(dates['days_passed'] / dates['days_in_month'] * 100))}%")
        print("=" * 50)

    finally:
        connection.close()
        print("\n数据库连接已关闭")


if __name__ == "__main__":
    main()
