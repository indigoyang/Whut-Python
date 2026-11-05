#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SEM周报自动生成脚本 V2（末次归因新用户版）
- SEM 收入按合作渠道为空且末次来源类型为 SEM 统计
- 新增 D7 ROI（含 PC/H5 子项）
- 新增 5 地区营收/消耗
- 移动端消耗从 Excel 读取
- 删除 流量/CTR/注册率/客单价/注册付费率/CVR/付费量 等旧指标
"""

import mysql.connector
from mysql.connector import Error
from datetime import datetime, timedelta, date
import calendar
from openpyxl import Workbook, load_workbook
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

# ============== 移动端消耗 Excel 路径 ==============
MOBILE_COST_FILE = os.getenv('MOBILE_COST_FILE', '')

# ============== 目标配置 ==============
MONTHLY_TARGETS = {
    '2026-04': {
        'SEM—新老综合ROI（除订阅）': 75,
    }
}

# ============== Google 地区名映射 ==============
GOOGLE_REGIONS = {
    'us': '美国', 'in': '印度', 'tw': '台湾',
    'id': '印度尼西亚', 'my': '马来西亚'
}

# ============== Bing region -> 内部 key 映射 ==============
# bing_region_cost.region 值 -> 内部标识
BING_REGION_MAP = {
    'us': 'United States',
    'in': 'India',
    'tw': ('Taiwan', 'Hong Kong SAR', 'Macao SAR'),  # 港澳台
    'id': 'Indonesia',
    'my': 'Malaysia',
}

SEM_LAST_ATTRIBUTION_CONDITION = """
NULLIF(TRIM(COALESCE(合作渠道, '')), '') IS NULL
AND UPPER(TRIM(COALESCE(末次来源类型, ''))) = 'SEM'
"""

def get_db_connection():
    """创建数据库连接"""
    try:
        connection = mysql.connector.connect(**DB_CONFIG)
        return connection
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
    """执行SQL查询并返回单行结果"""
    cursor = connection.cursor()
    try:
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        return cursor.fetchone()
    finally:
        cursor.close()


def read_mobile_cost(start_date, end_date):
    """从移动端消耗Excel读取指定日期范围内的移动端消耗总和
    文件格式：A列=日期，B列=国家，C列=消耗（行式）
    """
    if not os.path.exists(MOBILE_COST_FILE):
        print(f"警告: 移动端消耗文件不存在: {MOBILE_COST_FILE}")
        return 0

    wb = load_workbook(MOBILE_COST_FILE, data_only=True)
    ws = wb.active
    total_cost = 0.0

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        d_val = row[0]
        cost_val = row[2] if len(row) > 2 else None
        if d_val is None or cost_val is None:
            continue
        d = d_val.date() if isinstance(d_val, datetime) else d_val
        if not isinstance(d, date):
            continue
        if start_date <= d <= end_date:
            try:
                total_cost += float(cost_val)
            except (ValueError, TypeError):
                pass

    wb.close()
    return total_cost


def get_period_metrics(connection, start_date, end_date):
    """获取指定时间段的指标数据"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

    # ===== 营收查询 =====
    orders_query = f"""
    SELECT
        COALESCE(SUM(订单金额), 0) AS total_rev,
        COALESCE(SUM(CASE WHEN 支付端口 LIKE '%移动%' THEN 订单金额 ELSE 0 END), 0) AS mobile_rev,
        COALESCE(SUM(CASE WHEN 注册转化周期 = 0 THEN 订单金额 ELSE 0 END), 0) AS new_user_rev,
        COALESCE(SUM(CASE WHEN 注册转化周期 = 0
                          AND 支付端口 LIKE '%移动%'
                          THEN 订单金额 ELSE 0 END), 0) AS mobile_new_user_rev,
        COALESCE(SUM(CASE WHEN 注册转化周期 = 0
                          AND UPPER(TRIM(COALESCE(末次来源, ''))) = 'BING_ADS'
                          THEN 订单金额 ELSE 0 END), 0) AS bing_new_user_rev,
        COALESCE(SUM(CASE WHEN 注册转化周期 = 0
                          AND (UPPER(TRIM(COALESCE(末次来源, ''))) = 'GOOGLE_SEM'
                               OR NULLIF(TRIM(COALESCE(末次来源, '')), '') IS NULL)
                          THEN 订单金额 ELSE 0 END), 0) AS google_new_user_rev,
        COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(末次来源, ''))) = 'BING_ADS'
                          THEN 订单金额 ELSE 0 END), 0) AS bing_all_user_rev,
        COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(末次来源, ''))) = 'GOOGLE_SEM'
                               OR NULLIF(TRIM(COALESCE(末次来源, '')), '') IS NULL
                          THEN 订单金额 ELSE 0 END), 0) AS google_all_user_rev
    FROM ordersof4
    WHERE DATE(订单生成时间) BETWEEN %s AND %s
      AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
      AND {SEM_LAST_ATTRIBUTION_CONDITION}
    """
    orders = execute_query(connection, orders_query, (s, e))

    # 地区营收
    def region_rev(country, new_user_only=True):
        cycle_condition = "AND 注册转化周期 = 0" if new_user_only else ""
        q = f"""
        SELECT COALESCE(SUM(订单金额), 0) FROM ordersof4
        WHERE DATE(订单生成时间) BETWEEN %s AND %s
          AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
          {cycle_condition}
          AND {SEM_LAST_ATTRIBUTION_CONDITION}
          AND 国家 = %s
        """
        r = execute_query(connection, q, (s, e, country))
        return float(r[0] or 0)

    rev_us = region_rev('美国')
    rev_in = region_rev('印度')
    rev_tw = region_rev('中国台湾') + region_rev('中国香港') + region_rev('中国澳门')
    rev_id = region_rev('印度尼西亚')
    rev_my = region_rev('马来西亚')

    all_rev_us = region_rev('美国', new_user_only=False)
    all_rev_in = region_rev('印度', new_user_only=False)
    all_rev_tw = (
        region_rev('中国台湾', new_user_only=False)
        + region_rev('中国香港', new_user_only=False)
        + region_rev('中国澳门', new_user_only=False)
    )
    all_rev_id = region_rev('印度尼西亚', new_user_only=False)
    all_rev_my = region_rev('马来西亚', new_user_only=False)

    # D7 营收（按用户注册时间过滤）
    def d_rev(cycle_cond, port_cond=None, d_start=None, d_end=None):
        """D7 注册时间窗口营收。"""
        base = f"""
        SELECT COALESCE(SUM(订单金额), 0) FROM ordersof4
        WHERE DATE(用户注册时间) BETWEEN %s AND %s
          AND (支付端口 LIKE '%pc%' OR 支付端口 LIKE '%移动%' OR 支付端口 LIKE '%其他%')
          AND {SEM_LAST_ATTRIBUTION_CONDITION}
          AND {{cycle}}
        """
        if port_cond:
            base += f" AND ({port_cond})"
        start = d_start if d_start else s
        end = d_end if d_end else e
        r = execute_query(connection, base.format(cycle=cycle_cond), (start, end))
        return float(r[0] or 0)

    # D7：当期注册用户在 0～7 天内的 SEM 末次归因收入。
    d7_rev = d_rev('注册转化周期 BETWEEN 0 AND 7')
    d7_pc_rev = d_rev('注册转化周期 BETWEEN 0 AND 7', "支付端口 LIKE '%pc%' OR 支付端口 LIKE '%其他%'")
    d7_h5_rev = d_rev('注册转化周期 BETWEEN 0 AND 7', "支付端口 LIKE '%移动%'")

    # ===== 消耗查询 =====
    # Google 总消耗 (排除KOL)
    google_cost_q = """
    SELECT COALESCE(SUM(费用), 0) FROM google__ym_ads
    WHERE DATE(`天`) BETWEEN %s AND %s AND 地理位置 = '总计：地理位置' AND `广告系列` NOT LIKE '%KOL%'
      AND `广告系列` NOT LIKE '%generic%'
    """
    google_cost = float(execute_query(connection, google_cost_q, (s, e))[0] or 0)

    # Google 地区消耗
    def google_region_cost(region_name):
        q = """
        SELECT COALESCE(SUM(费用), 0) FROM google__ym_ads
        WHERE DATE(`天`) BETWEEN %s AND %s AND 地理位置 = %s AND `广告系列` NOT LIKE '%KOL%'
          AND `广告系列` NOT LIKE '%generic%'
        """
        r = execute_query(connection, q, (s, e, region_name))
        return float(r[0] or 0)

    gc_us = google_region_cost('美国')
    gc_in = google_region_cost('印度')
    gc_tw = google_region_cost('台湾') + google_region_cost('香港') + google_region_cost('澳门')
    gc_id = google_region_cost('印度尼西亚')
    gc_my = google_region_cost('马来西亚')

    # Bing 总消耗 (排除 Search total)
    bing_cost_q = """
    SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
    WHERE DATE(`stat_date`) BETWEEN %s AND %s AND ad_campaign != 'Search total'
    """
    bing_cost = float(execute_query(connection, bing_cost_q, (s, e))[0] or 0)

    # Bing 分地区消耗 (从 bing_region_cost)
    def bing_region_cost(region_key):
        region_val = BING_REGION_MAP[region_key]
        if isinstance(region_val, tuple):
            # 港澳台需要三个 region 合并
            placeholders = ','.join(['%s'] * len(region_val))
            q = f"""
            SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
            WHERE DATE(`stat_date`) BETWEEN %s AND %s
              AND ad_campaign != 'Search total'
              AND region IN ({placeholders})
            """
            r = execute_query(connection, q, (s, e) + region_val)
        else:
            q = """
            SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
            WHERE DATE(`stat_date`) BETWEEN %s AND %s
              AND ad_campaign != 'Search total'
              AND region = %s
            """
            r = execute_query(connection, q, (s, e, region_val))
        return float(r[0] or 0)

    bc_us = bing_region_cost('us')
    bc_in = bing_region_cost('in')
    bc_tw = bing_region_cost('tw')
    bc_id = bing_region_cost('id')
    bc_my = bing_region_cost('my')

    # 移动端消耗
    mobile_cost = read_mobile_cost(start_date, end_date)

    total_cost = google_cost + bing_cost
    pc_cost = total_cost - mobile_cost

    # D7 ROI 分母：统计周期整体往前推 7 天后的同期 SEM 消耗。
    # 例如统计周期为 8.11-8.18，则 D7 消耗取 8.04-8.11。
    d7_cost_start = start_date - timedelta(days=7)
    d7_cost_end = end_date - timedelta(days=7)
    d7_s = d7_cost_start.strftime('%Y-%m-%d')
    d7_e = d7_cost_end.strftime('%Y-%m-%d')

    d7_google_cost = float(execute_query(connection, google_cost_q, (d7_s, d7_e))[0] or 0)
    d7_bing_cost = float(execute_query(connection, bing_cost_q, (d7_s, d7_e))[0] or 0)
    d7_mobile_cost = read_mobile_cost(d7_cost_start, d7_cost_end)
    d7_total_cost = d7_google_cost + d7_bing_cost
    d7_pc_cost = d7_total_cost - d7_mobile_cost

    total_rev = float(orders[0] or 0)

    metrics = {
        'total_revenue': total_rev,
        'all_user_non_subscription_revenue': total_rev,
        'mobile_revenue': float(orders[1] or 0),
        'new_user_revenue': float(orders[2] or 0),
        'mobile_new_user_revenue': float(orders[3] or 0),
        'bing_revenue': float(orders[4] or 0),
        'google_revenue': float(orders[5] or 0),
        'bing_all_user_revenue': float(orders[6] or 0),
        'google_all_user_revenue': float(orders[7] or 0),
        'rev_us': rev_us, 'rev_in': rev_in, 'rev_tw': rev_tw,
        'rev_id': rev_id, 'rev_my': rev_my,
        'all_rev_us': all_rev_us, 'all_rev_in': all_rev_in, 'all_rev_tw': all_rev_tw,
        'all_rev_id': all_rev_id, 'all_rev_my': all_rev_my,
        'd7_revenue': d7_rev, 'd7_pc_revenue': d7_pc_rev, 'd7_h5_revenue': d7_h5_rev,
        'd7_total_cost': d7_total_cost, 'd7_pc_cost': d7_pc_cost,
        'd7_mobile_cost': d7_mobile_cost,
        'total_cost': total_cost, 'google_cost': google_cost, 'bing_cost': bing_cost,
        'mobile_cost': mobile_cost, 'pc_cost': pc_cost,
        'google_cost_us': gc_us, 'google_cost_in': gc_in, 'google_cost_tw': gc_tw,
        'google_cost_id': gc_id, 'google_cost_my': gc_my,
        'bing_cost_us': bc_us, 'bing_cost_in': bc_in, 'bing_cost_tw': bc_tw,
        'bing_cost_id': bc_id, 'bing_cost_my': bc_my,
    }
    return metrics


def get_period_costs(connection, start_date, end_date):
    """只获取指定时间段的总消耗、PC消耗、移动端消耗（不查营收）"""
    s, e = start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

    google_cost_q = """
    SELECT COALESCE(SUM(费用), 0) FROM google__ym_ads
    WHERE DATE(`天`) BETWEEN %s AND %s AND 地理位置 = '总计：地理位置' AND `广告系列` NOT LIKE '%KOL%'
      AND `广告系列` NOT LIKE '%generic%'
    """
    google_cost = float(execute_query(connection, google_cost_q, (s, e))[0] or 0)

    bing_cost_q = """
    SELECT COALESCE(SUM(cost), 0) FROM bing_region_cost
    WHERE DATE(`stat_date`) BETWEEN %s AND %s AND ad_campaign != 'Search total'
    """
    bing_cost = float(execute_query(connection, bing_cost_q, (s, e))[0] or 0)

    mobile_cost = read_mobile_cost(start_date, end_date)
    total_cost = google_cost + bing_cost
    pc_cost = total_cost - mobile_cost

    return {'total_cost': total_cost, 'pc_cost': pc_cost, 'mobile_cost': mobile_cost,
            'google_cost': google_cost, 'bing_cost': bing_cost}


def get_mtd_metrics(connection, dates):
    """获取MTD指标数据，复用 get_period_metrics"""
    return get_period_metrics(connection, dates['month_start'], dates['month_end'])


def create_report(dates, this_week, last_week, mtd, monthly_targets):
    """创建周报 Excel 文件；D7 ROI 分母使用统计周期往前推 7 天的消耗。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "SEM周报"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    sub_font = Font(size=11)  # 子指标不加粗
    bold_font = Font(bold=True, size=11)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center')

    # 标题
    ws.merge_cells('A1:H1')
    ws['A1'] = "SEM周报"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = center_align

    date_range = f"（{dates['this_week_start'].strftime('%Y.%m.%d')}-{dates['this_week_end'].strftime('%Y.%m.%d')}）"
    ws.merge_cells('A2:H2')
    ws['A2'] = date_range
    ws['A2'].font = Font(size=12)
    ws['A2'].alignment = center_align

    headers = ['指标', '本周', '上周', '环比', 'MTD', '目标', '完成率', '时间进度']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align

    time_progress = dates['days_passed'] / dates['days_in_month']

    # ===== 计算衍生指标 =====
    def safe_div(a, b):
        return a / b if b and b != 0 else None

    def roi(rev, cost):
        return safe_div(rev, cost)

    # ----- ROI 指标 -----
    # SEM 新用户 ROI：订单日期在统计期内、注册转化周期=0、按 SEM 末次归因。
    tw_new_roi = roi(this_week['new_user_revenue'], this_week['total_cost'])
    lw_new_roi = roi(last_week['new_user_revenue'], last_week['total_cost'])
    mt_new_roi = roi(mtd['new_user_revenue'], mtd['total_cost'])

    tw_new_pc_roi = roi(this_week['new_user_revenue'] - this_week['mobile_new_user_revenue'], this_week['pc_cost'])
    lw_new_pc_roi = roi(last_week['new_user_revenue'] - last_week['mobile_new_user_revenue'], last_week['pc_cost'])
    mt_new_pc_roi = roi(mtd['new_user_revenue'] - mtd['mobile_new_user_revenue'], mtd['pc_cost'])

    tw_new_h5_roi = roi(this_week['mobile_new_user_revenue'], this_week['mobile_cost'])
    lw_new_h5_roi = roi(last_week['mobile_new_user_revenue'], last_week['mobile_cost'])
    mt_new_h5_roi = roi(mtd['mobile_new_user_revenue'], mtd['mobile_cost'])

    tw_roi = roi(this_week['total_revenue'], this_week['total_cost'])
    lw_roi = roi(last_week['total_revenue'], last_week['total_cost'])
    mt_roi = roi(mtd['total_revenue'], mtd['total_cost'])

    tw_pc_roi = roi(this_week['total_revenue'] - this_week['mobile_revenue'], this_week['pc_cost'])
    lw_pc_roi = roi(last_week['total_revenue'] - last_week['mobile_revenue'], last_week['pc_cost'])
    mt_pc_roi = roi(mtd['total_revenue'] - mtd['mobile_revenue'], mtd['pc_cost'])

    tw_h5_roi = roi(this_week['mobile_revenue'], this_week['mobile_cost'])
    lw_h5_roi = roi(last_week['mobile_revenue'], last_week['mobile_cost'])
    mt_h5_roi = roi(mtd['mobile_revenue'], mtd['mobile_cost'])

    # D7 ROI：当期注册用户 0～7 天 SEM 收入 ÷ 统计周期往前推 7 天的消耗。
    tw_d7 = roi(this_week['d7_revenue'], this_week['d7_total_cost'])
    lw_d7 = roi(last_week['d7_revenue'], last_week['d7_total_cost'])
    mt_d7 = roi(mtd['d7_revenue'], mtd['d7_total_cost'])
    tw_d7pc = roi(this_week['d7_pc_revenue'], this_week['d7_pc_cost'])
    lw_d7pc = roi(last_week['d7_pc_revenue'], last_week['d7_pc_cost'])
    mt_d7pc = roi(mtd['d7_pc_revenue'], mtd['d7_pc_cost'])
    tw_d7h5 = roi(this_week['d7_h5_revenue'], this_week['d7_mobile_cost'])
    lw_d7h5 = roi(last_week['d7_h5_revenue'], last_week['d7_mobile_cost'])
    mt_d7h5 = roi(mtd['d7_h5_revenue'], mtd['d7_mobile_cost'])

    # ----- 地区营收 -----
    def region_sum(m, key, prefix='rev'):
        revs = [m.get(f'{prefix}_{k}', 0) or 0 for k in ['us', 'in', 'tw', 'id', 'my']]
        other = (m.get(key, 0) or 0) - sum(revs)
        return revs + [other]

    tw_rev_regions = region_sum(this_week, 'new_user_revenue')
    lw_rev_regions = region_sum(last_week, 'new_user_revenue')
    mt_rev_regions = region_sum(mtd, 'new_user_revenue')
    tw_all_rev_regions = region_sum(this_week, 'all_user_non_subscription_revenue', 'all_rev')
    lw_all_rev_regions = region_sum(last_week, 'all_user_non_subscription_revenue', 'all_rev')
    mt_all_rev_regions = region_sum(mtd, 'all_user_non_subscription_revenue', 'all_rev')

    # ----- 地区消耗 (Google + Bing) -----
    def region_cost(m):
        gc_keys = ['google_cost_us', 'google_cost_in', 'google_cost_tw', 'google_cost_id', 'google_cost_my']
        bc_keys = ['bing_cost_us', 'bing_cost_in', 'bing_cost_tw', 'bing_cost_id', 'bing_cost_my']
        google_vals = [m.get(k, 0) or 0 for k in gc_keys]
        bing_vals = [m.get(k, 0) or 0 for k in bc_keys]
        total_list = [g + b for g, b in zip(google_vals, bing_vals)]
        other = (m.get('total_cost', 0) or 0) - sum(total_list)
        return total_list + [other]

    tw_cost_regions = region_cost(this_week)
    lw_cost_regions = region_cost(last_week)
    mt_cost_regions = region_cost(mtd)

    # ===== 格式化函数 =====
    def fmt_int(v):
        return int(round(v)) if v is not None else "-"

    def fmt_pct(v):
        return f"{int(round(v * 100))}%" if v is not None else "-"

    def fmt_pct2(v):
        return f"{round(v * 100, 2)}%" if v is not None else "-"

    def fmt_dollar(v):
        return f"${int(round(v))}" if v is not None else "-"

    def format_val(v, ft):
        if ft == 'int': return fmt_int(v)
        if ft == 'pct': return fmt_pct(v)
        if ft == 'pct2': return fmt_pct2(v)
        if ft == 'dollar': return fmt_dollar(v)
        return fmt_int(v)

    def mom(this_v, last_v):
        if last_v and last_v != 0 and this_v is not None:
            return (this_v - last_v) / last_v
        return None

    def fmt_mom(v):
        return f"{int(round(v * 100))}%" if v is not None else "-"

    def target_val(name):
        return monthly_targets.get(name) if monthly_targets else None

    def completion(mtd_v, tgt, ft):
        if tgt is None or tgt == 0 or mtd_v is None: return None
        if ft in ('pct', 'pct2'): return (mtd_v * 100) / tgt
        return mtd_v / tgt

    def fmt_comp(r):
        return f"{int(round(r * 100))}%" if r is not None else "-"

    # ===== 构建指标行 =====
    # (指标名, 本周值, 上周值, 累计值, 格式类型, 缩进级别)
    raw_rows = [
        # 蓝色分区标题：格式类型使用 section。
        ('核心ROI（末次归因）', None, None, None, 'section', 0),
        ('SEM—SEM新用户ROI', tw_new_roi, lw_new_roi, mt_new_roi, 'pct', 0),
        ('PC端', tw_new_pc_roi, lw_new_pc_roi, mt_new_pc_roi, 'pct', 1),
        ('H5端', tw_new_h5_roi, lw_new_h5_roi, mt_new_h5_roi, 'pct', 1),
        ('SEM—新老综合ROI（除订阅）', tw_roi, lw_roi, mt_roi, 'pct', 0),
        ('PC端', tw_pc_roi, lw_pc_roi, mt_pc_roi, 'pct', 1),
        ('H5端', tw_h5_roi, lw_h5_roi, mt_h5_roi, 'pct', 1),
        ('SEM—新用户ROI（D7）', tw_d7, lw_d7, mt_d7, 'pct2', 0),
        ('PC端', tw_d7pc, lw_d7pc, mt_d7pc, 'pct2', 1),
        ('H5端', tw_d7h5, lw_d7h5, mt_d7h5, 'pct2', 1),

        ('SEM收入总览', None, None, None, 'section', 0),
        ('SEM新用户总收入', this_week['new_user_revenue'], last_week['new_user_revenue'], mtd['new_user_revenue'], 'dollar', 0),
        ('SEM新老用户总收入（除订阅）', this_week['all_user_non_subscription_revenue'], last_week['all_user_non_subscription_revenue'], mtd['all_user_non_subscription_revenue'], 'dollar', 0),

        ('SEM新用户收入—投放平台', None, None, None, 'section', 0),
        ('Google新用户营收', this_week['google_revenue'], last_week['google_revenue'], mtd['google_revenue'], 'dollar', 1),
        ('Bing新用户营收', this_week['bing_revenue'], last_week['bing_revenue'], mtd['bing_revenue'], 'dollar', 1),

        ('SEM新老用户收入（除订阅）—投放平台', None, None, None, 'section', 0),
        ('Google新老用户（除订阅）营收', this_week['google_all_user_revenue'], last_week['google_all_user_revenue'], mtd['google_all_user_revenue'], 'dollar', 1),
        ('Bing新老用户（除订阅）营收', this_week['bing_all_user_revenue'], last_week['bing_all_user_revenue'], mtd['bing_all_user_revenue'], 'dollar', 1),

        ('SEM新用户收入—地区', None, None, None, 'section', 0),
        ('美国地区新用户营收', tw_rev_regions[0], lw_rev_regions[0], mt_rev_regions[0], 'dollar', 1),
        ('印度地区新用户营收', tw_rev_regions[1], lw_rev_regions[1], mt_rev_regions[1], 'dollar', 1),
        ('港澳台地区新用户营收', tw_rev_regions[2], lw_rev_regions[2], mt_rev_regions[2], 'dollar', 1),
        ('印度尼西亚地区新用户营收', tw_rev_regions[3], lw_rev_regions[3], mt_rev_regions[3], 'dollar', 1),
        ('马来西亚新用户营收', tw_rev_regions[4], lw_rev_regions[4], mt_rev_regions[4], 'dollar', 1),
        ('其他新用户营收', tw_rev_regions[5], lw_rev_regions[5], mt_rev_regions[5], 'dollar', 1),

        ('SEM新老用户收入（除订阅）—地区', None, None, None, 'section', 0),
        ('美国地区新老用户（除订阅）营收', tw_all_rev_regions[0], lw_all_rev_regions[0], mt_all_rev_regions[0], 'dollar', 1),
        ('印度地区新老用户（除订阅）营收', tw_all_rev_regions[1], lw_all_rev_regions[1], mt_all_rev_regions[1], 'dollar', 1),
        ('港澳台地区新老用户（除订阅）营收', tw_all_rev_regions[2], lw_all_rev_regions[2], mt_all_rev_regions[2], 'dollar', 1),
        ('印度尼西亚地区新老用户（除订阅）营收', tw_all_rev_regions[3], lw_all_rev_regions[3], mt_all_rev_regions[3], 'dollar', 1),
        ('马来西亚新老用户（除订阅）营收', tw_all_rev_regions[4], lw_all_rev_regions[4], mt_all_rev_regions[4], 'dollar', 1),
        ('其他新老用户（除订阅）营收', tw_all_rev_regions[5], lw_all_rev_regions[5], mt_all_rev_regions[5], 'dollar', 1),

        ('SEM消耗总览', None, None, None, 'section', 0),
        ('总消耗', this_week['total_cost'], last_week['total_cost'], mtd['total_cost'], 'dollar', 0),
        ('Google消耗', this_week['google_cost'], last_week['google_cost'], mtd['google_cost'], 'dollar', 1),
        ('Bing消耗', this_week['bing_cost'], last_week['bing_cost'], mtd['bing_cost'], 'dollar', 1),

        ('SEM消耗—地区', None, None, None, 'section', 0),
        ('美国地区消耗', tw_cost_regions[0], lw_cost_regions[0], mt_cost_regions[0], 'dollar', 1),
        ('印度地区消耗', tw_cost_regions[1], lw_cost_regions[1], mt_cost_regions[1], 'dollar', 1),
        ('港澳台地区消耗', tw_cost_regions[2], lw_cost_regions[2], mt_cost_regions[2], 'dollar', 1),
        ('印度尼西亚地区消耗', tw_cost_regions[3], lw_cost_regions[3], mt_cost_regions[3], 'dollar', 1),
        ('马来西亚消耗', tw_cost_regions[4], lw_cost_regions[4], mt_cost_regions[4], 'dollar', 1),
        ('其他', tw_cost_regions[5], lw_cost_regions[5], mt_cost_regions[5], 'dollar', 1),
    ]

    # 写入数据
    row_num = 4
    for name, v_tw, v_lw, v_mt, ft, indent in raw_rows:
        if ft == 'section':
            row_data = [name, '-', '-', '-', '-', '-', '-', '-']
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col, value=val)
                cell.border = border
                cell.alignment = center_align
                cell.fill = section_fill
                cell.font = Font(bold=True, color="1F4E78")
            row_num += 1
            continue

        display_name = '  ' * indent + name
        m = mom(v_tw, v_lw)
        t = target_val(name)
        c = completion(v_mt, t, ft)

        row_data = [
            display_name,
            format_val(v_tw, ft), format_val(v_lw, ft), fmt_mom(m),
            format_val(v_mt, ft),
            format_val(t, ft) if t is not None else '-',
            fmt_comp(c),
            '-'
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.border = border
            cell.alignment = center_align
            cell.font = bold_font if indent == 0 else sub_font
        row_num += 1

    # 合并时间进度列
    end_row = row_num - 1
    ws.merge_cells(start_row=4, start_column=8, end_row=end_row, end_column=8)
    tp_cell = ws.cell(row=4, column=8, value=f"{int(round(time_progress * 100))}%")
    tp_cell.alignment = center_align
    tp_cell.border = border
    tp_cell.font = sub_font

    column_widths = [36, 14, 14, 10, 14, 12, 10, 10]
    for i, w in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    return wb


def main():
    print("=" * 50)
    print("SEM周报自动生成脚本 V2")
    print("=" * 50)

    # 交互式输入日期（直接回车则自动按上周四~本周三计算）
    print("\n请输入统计周期（格式 YYYY-MM-DD，直接回车则自动计算）:")
    tw_start = input("  本周开始日期: ").strip()
    tw_end = input("  本周结束日期: ").strip()
    if tw_start and tw_end:
        dates = calculate_dates(this_week_start=tw_start, this_week_end=tw_end)
    else:
        dates = calculate_dates()
    print(f"\n本周范围: {dates['this_week_start']} 至 {dates['this_week_end']}")
    print(f"上周范围: {dates['last_week_start']} 至 {dates['last_week_end']}")
    print(f"累计范围: {dates['month_start']} 至 {dates['month_end']}")

    month_key = dates['month_end'].strftime('%Y-%m')
    monthly_targets = MONTHLY_TARGETS.get(month_key, {})
    print(f"本月目标: {'已配置' if monthly_targets else '未配置'}")

    print("\n正在连接数据库...")
    connection = get_db_connection()
    print("数据库连接成功")

    try:
        print("\n正在获取数据...")
        print("  - 本周数据...")
        this_week = get_period_metrics(connection, dates['this_week_start'], dates['this_week_end'])
        print("  - 上周数据...")
        last_week = get_period_metrics(connection, dates['last_week_start'], dates['last_week_end'])
        print("  - 累计数据...")
        mtd = get_mtd_metrics(connection, dates)

        print("数据获取完成")

        # 生成报告
        print("\n正在生成报告...")
        wb = create_report(dates, this_week, last_week, mtd, monthly_targets)

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_file = os.path.join(
            OUTPUT_DIR,
            f"周报-SEM_末次归因新用户_{dates['this_week_start'].strftime('%Y.%m.%d')}-{dates['this_week_end'].strftime('%Y.%m.%d')}.xlsx"
        )
        wb.save(output_file)
        print(f"\n报告已保存: {output_file}")

        # 摘要
        print("\n" + "=" * 50)
        print("数据摘要:")
        print(f"  本周营收: ${int(round(this_week['total_revenue']))}")
        print(f"  本周消耗: ${int(round(this_week['total_cost']))}")
        if this_week['total_cost'] > 0:
            print(f"  本周ROI: {int(round(this_week['total_revenue'] / this_week['total_cost'] * 100))}%")
        if this_week['mobile_cost'] > 0:
            print(f"  本周移动端消耗: ${int(round(this_week['mobile_cost']))}")
            print(f"  本周PC端消耗: ${int(round(this_week['pc_cost']))}")
        print(f"  时间进度: {int(round(dates['days_passed'] / dates['days_in_month'] * 100))}%")
        print("=" * 50)

    finally:
        connection.close()
        print("\n数据库连接已关闭")


if __name__ == "__main__":
    main()
