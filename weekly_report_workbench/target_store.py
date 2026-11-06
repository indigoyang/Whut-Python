from __future__ import annotations

import base64
import json
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parent
LOCAL_TARGET_FILE = ROOT / "config" / "monthly_targets.json"


ALL_CHANNEL_TARGET_METRICS = [
    "WEB—新用户ROI", "WEB—新老综合ROI（除订阅）", "WEB—新老综合ROI（含订阅）",
    "WEB—英语地区新用户ROI", "WEB—繁中地区新用户ROI", "WEB—其他地区新用户ROI",
    "新用户总收入", "新老用户总收入（除订阅）", "新老用户总收入（含订阅）",
    "SEM新用户收入", "自然来源新用户收入", "SEO新用户收入", "KOL新用户收入",
    "affiliate新用户收入", "教育认证新用户收入", "Unassigned新用户收入",
    "SEM新老用户除订阅收入", "自然来源新老用户除订阅收入",
    "SEO新老用户除订阅收入", "KOL新老用户除订阅收入",
    "affiliate新老用户除订阅收入", "教育认证新老用户除订阅收入",
    "Unassigned新老用户除订阅收入",
    "英语组新用户收入", "北美（美国+加拿大）新用户收入", "印度新用户收入",
    "繁体中文新用户收入", "其他新用户收入",
    "英语组新老用户除订阅收入", "北美（美国+加拿大）新老用户除订阅收入",
    "印度新老用户除订阅收入", "繁体中文新老用户除订阅收入",
    "其他地区新老用户除订阅收入",
    "英语组新老用户含订阅收入", "北美（美国+加拿大）新老用户含订阅收入",
    "印度新老用户含订阅收入", "繁体中文新老用户含订阅收入",
    "其他地区新老用户含订阅收入",
    "总成本", "SEM成本", "affiliate成本", "SEO—外链成本",
    "PC—流量", "PC—新增注册", "PC—注册率", "PC—新用户客单价ARPPU",
    "PC—新用户订单量", "PC—新用户营收", "PC—新老用户营收（除订阅）",
    "PC—新老用户营收（含订阅）", "PC—新用户注册付费率",
    "PC—付费率 (流量口径)", "PC—退款金额",
    "H5—流量", "H5—新增注册", "H5—注册率", "H5—新用户客单价ARPPU",
    "H5—新用户订单量", "H5—新用户营收", "H5—新老用户营收（除订阅）",
    "H5—新老用户营收（含订阅）", "H5—新用户注册付费率",
    "H5—付费率 (流量口径)",
]

SEM_TARGET_METRICS = [
    "SEM—SEM新用户ROI", "SEM—新老综合ROI（除订阅）", "SEM新用户总收入",
    "SEM新老用户总收入（除订阅）", "Google新用户营收", "Bing新用户营收",
    "Google新老用户（除订阅）营收", "Bing新老用户（除订阅）营收",
    "美国地区新用户营收", "印度地区新用户营收", "港澳台地区新用户营收",
    "印度尼西亚地区新用户营收", "马来西亚新用户营收", "其他新用户营收",
    "美国地区新老用户（除订阅）营收", "印度地区新老用户（除订阅）营收",
    "港澳台地区新老用户（除订阅）营收", "印度尼西亚地区新老用户（除订阅）营收",
    "马来西亚新老用户（除订阅）营收", "其他新老用户（除订阅）营收",
    "总消耗", "Google消耗", "Bing消耗", "美国地区消耗", "印度地区消耗",
    "港澳台地区消耗", "印度尼西亚地区消耗", "马来西亚消耗", "其他",
]

NATURAL_TARGET_METRICS = [
    "已发布达人数量", "当周曝光", "CPM", "WEB（品牌词搜索量）",
    "注册总计", "WEB自然量注册数", "APP自然量注册数",
    "WEB自然来源新用户收入", "WEB新老用户收入（除订阅）", "APP自然量收入",
    "成本总计", "达人合作费", "投流费", "新用户ROI",
]

NATURAL_TARGET_ROWS = [
    (None, "自然来源目标"),
    *[(metric, metric) for metric in NATURAL_TARGET_METRICS],
]

# 目标维护页严格按照最终Excel的显示顺序排列。
# 元组格式：(保存键, 页面显示名称)；保存键为None的是分区标题，不保存目标。
ALL_CHANNEL_TARGET_ROWS = [
    (None, "核心ROI（末次归因）"),
    ("WEB—新用户ROI", "WEB—新用户ROI"),
    ("WEB—新老综合ROI（除订阅）", "WEB—新老综合ROI（除订阅）"),
    ("WEB—新老综合ROI（含订阅）", "WEB—新老综合ROI（含订阅）"),
    ("WEB—英语地区新用户ROI", "WEB—英语地区新用户ROI"),
    ("WEB—繁中地区新用户ROI", "WEB—繁中地区新用户ROI"),
    ("WEB—其他地区新用户ROI", "WEB—其他地区新用户ROI"),
    (None, "收入总览"),
    ("新用户总收入", "新用户总收入"),
    ("新老用户总收入（除订阅）", "新老用户总收入（除订阅）"),
    ("新老用户总收入（含订阅）", "新老用户总收入（含订阅）"),
    (None, "新用户收入明细（末次归因）"),
    ("SEM新用户收入", "SEM新用户收入"),
    ("自然来源新用户收入", "自然来源新用户收入"),
    ("SEO新用户收入", "SEO新用户收入"),
    ("KOL新用户收入", "KOL新用户收入"),
    ("affiliate新用户收入", "affiliate新用户收入"),
    ("教育认证新用户收入", "教育认证新用户收入"),
    ("Unassigned新用户收入", "Unassigned新用户收入"),
    (None, "新老用户收入明细（末次归因除订阅）"),
    ("SEM新老用户除订阅收入", "SEM新老用户除订阅收入"),
    ("自然来源新老用户除订阅收入", "自然来源新老用户除订阅收入"),
    ("SEO新老用户除订阅收入", "SEO新老用户除订阅收入"),
    ("KOL新老用户除订阅收入", "KOL新老用户除订阅收入"),
    ("affiliate新老用户除订阅收入", "affiliate新老用户除订阅收入"),
    ("教育认证新老用户除订阅收入", "教育认证新老用户除订阅收入"),
    ("Unassigned新老用户除订阅收入", "Unassigned新老用户除订阅收入"),
    (None, "新用户地区收入（末次归因）"),
    ("英语组新用户收入", "英语组新用户收入"),
    ("北美（美国+加拿大）新用户收入", "北美（美国+加拿大）新用户收入"),
    ("印度新用户收入", "印度新用户收入"),
    ("繁体中文新用户收入", "繁体中文新用户收入"),
    ("其他新用户收入", "其他新用户收入"),
    (None, "新老用户地区收入（末次归因除订阅）"),
    ("英语组新老用户除订阅收入", "英语组新老用户除订阅收入"),
    ("北美（美国+加拿大）新老用户除订阅收入", "北美（美国+加拿大）新老用户除订阅收入"),
    ("印度新老用户除订阅收入", "印度新老用户除订阅收入"),
    ("繁体中文新老用户除订阅收入", "繁体中文新老用户除订阅收入"),
    ("其他地区新老用户除订阅收入", "其他地区新老用户除订阅收入"),
    (None, "新老用户地区收入（末次归因含订阅）"),
    ("英语组新老用户含订阅收入", "英语组新老用户含订阅收入"),
    ("北美（美国+加拿大）新老用户含订阅收入", "北美（美国+加拿大）新老用户含订阅收入"),
    ("印度新老用户含订阅收入", "印度新老用户含订阅收入"),
    ("繁体中文新老用户含订阅收入", "繁体中文新老用户含订阅收入"),
    ("其他地区新老用户含订阅收入", "其他地区新老用户含订阅收入"),
    (None, "成本总览"),
    ("总成本", "总成本"),
    ("SEM成本", "SEM成本"),
    ("affiliate成本", "affiliate成本"),
    ("SEO—外链成本", "SEO—外链成本"),
    (None, "PC"),
    ("PC—流量", "流量"),
    ("PC—新增注册", "新增注册"),
    ("PC—注册率", "注册率"),
    ("PC—新用户客单价ARPPU", "新用户客单价ARPPU"),
    ("PC—新用户订单量", "新用户订单量"),
    ("PC—新用户营收", "新用户营收"),
    ("PC—新老用户营收（除订阅）", "新老用户营收（除订阅）"),
    ("PC—新老用户营收（含订阅）", "新老用户营收（含订阅）"),
    ("PC—新用户注册付费率", "新用户注册付费率"),
    ("PC—付费率 (流量口径)", "付费率 (流量口径)"),
    ("PC—退款金额", "退款金额"),
    (None, "H5"),
    ("H5—流量", "流量"),
    ("H5—新增注册", "新增注册"),
    ("H5—注册率", "注册率"),
    ("H5—新用户客单价ARPPU", "新用户客单价ARPPU"),
    ("H5—新用户订单量", "新用户订单量"),
    ("H5—新用户营收", "新用户营收"),
    ("H5—新老用户营收（除订阅）", "新老用户营收（除订阅）"),
    ("H5—新老用户营收（含订阅）", "新老用户营收（含订阅）"),
    ("H5—新用户注册付费率", "新用户注册付费率"),
    ("H5—付费率 (流量口径)", "付费率 (流量口径)"),
]

SEM_TARGET_ROWS = [
    (None, "核心ROI（末次归因）"),
    ("SEM—SEM新用户ROI", "　SEM—SEM新用户ROI"),
    ("SEM—新老综合ROI（除订阅）", "　SEM—新老综合ROI（除订阅）"),
    (None, "SEM收入总览"),
    ("SEM新用户总收入", "　SEM新用户总收入"),
    ("SEM新老用户总收入（除订阅）", "　SEM新老用户总收入（除订阅）"),
    (None, "SEM新用户收入—投放平台"),
    ("Google新用户营收", "　Google新用户营收"),
    ("Bing新用户营收", "　Bing新用户营收"),
    (None, "SEM新老用户收入（除订阅）—投放平台"),
    ("Google新老用户（除订阅）营收", "　Google新老用户（除订阅）营收"),
    ("Bing新老用户（除订阅）营收", "　Bing新老用户（除订阅）营收"),
    (None, "SEM新用户收入—地区"),
    ("美国地区新用户营收", "　美国地区新用户营收"),
    ("印度地区新用户营收", "　印度地区新用户营收"),
    ("港澳台地区新用户营收", "　港澳台地区新用户营收"),
    ("印度尼西亚地区新用户营收", "　印度尼西亚地区新用户营收"),
    ("马来西亚新用户营收", "　马来西亚新用户营收"),
    ("其他新用户营收", "　其他新用户营收"),
    (None, "SEM新老用户收入（除订阅）—地区"),
    ("美国地区新老用户（除订阅）营收", "　美国地区新老用户（除订阅）营收"),
    ("印度地区新老用户（除订阅）营收", "　印度地区新老用户（除订阅）营收"),
    ("港澳台地区新老用户（除订阅）营收", "　港澳台地区新老用户（除订阅）营收"),
    ("印度尼西亚地区新老用户（除订阅）营收", "　印度尼西亚地区新老用户（除订阅）营收"),
    ("马来西亚新老用户（除订阅）营收", "　马来西亚新老用户（除订阅）营收"),
    ("其他新老用户（除订阅）营收", "　其他新老用户（除订阅）营收"),
    (None, "SEM消耗总览"),
    ("总消耗", "　总消耗"),
    ("Google消耗", "　Google消耗"),
    ("Bing消耗", "　Bing消耗"),
    (None, "SEM消耗—地区"),
    ("美国地区消耗", "　美国地区消耗"),
    ("印度地区消耗", "　印度地区消耗"),
    ("港澳台地区消耗", "　港澳台地区消耗"),
    ("印度尼西亚地区消耗", "　印度尼西亚地区消耗"),
    ("马来西亚消耗", "　马来西亚消耗"),
    ("其他", "　其他"),
]


def load_targets() -> dict:
    if not LOCAL_TARGET_FILE.exists():
        return {}
    with LOCAL_TARGET_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def github_settings(secrets) -> dict:
    github = secrets.get("github", {}) if secrets else {}
    return {
        "token": github.get("token", ""),
        "owner": github.get("owner", "indigoyang"),
        "repo": github.get("repo", "weekly-report"),
        "branch": github.get("branch", "main"),
        "path": github.get(
            "target_path", "weekly_report_workbench/config/monthly_targets.json"
        ),
        "rule_path": github.get(
            "rule_path", "weekly_report_workbench/config/report_rules.json"
        ),
    }


def save_targets(targets: dict, secrets) -> None:
    """保存到当前运行目录，并提交到 GitHub，供所有用户长期共用。"""
    LOCAL_TARGET_FILE.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(targets, ensure_ascii=False, indent=2) + "\n"
    LOCAL_TARGET_FILE.write_text(content, encoding="utf-8")

    settings = github_settings(secrets)
    if not settings["token"]:
        raise ValueError("尚未在 Streamlit Secrets 中配置 GitHub Token。")

    api = (
        f"https://api.github.com/repos/{settings['owner']}/{settings['repo']}"
        f"/contents/{settings['path']}"
    )
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {settings['token']}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "aippt-weekly-report",
    }

    sha = None
    try:
        with urlopen(Request(f"{api}?ref={settings['branch']}", headers=headers), timeout=20) as r:
            sha = json.loads(r.read().decode("utf-8")).get("sha")
    except HTTPError as exc:
        if exc.code != 404:
            raise RuntimeError(f"读取 GitHub 目标文件失败：HTTP {exc.code}") from exc

    payload = {
        "message": "更新月度周报目标",
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": settings["branch"],
    }
    if sha:
        payload["sha"] = sha
    request = Request(
        api,
        data=json.dumps(payload).encode("utf-8"),
        headers={**headers, "Content-Type": "application/json"},
        method="PUT",
    )
    try:
        with urlopen(request, timeout=30):
            pass
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"保存目标到 GitHub 失败：HTTP {exc.code} {detail}") from exc
