from __future__ import annotations

import base64
from copy import deepcopy
from datetime import datetime
import json
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from target_store import github_settings


ROOT = Path(__file__).resolve().parent
LOCAL_RULE_FILE = ROOT / "config" / "report_rules.json"


def load_rules() -> dict:
    if not LOCAL_RULE_FILE.exists():
        return {}
    with LOCAL_RULE_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_rules(rules: dict, secrets, updated_by: str = "", change_note: str = "") -> None:
    """保存规则到当前实例并提交到 GitHub，供团队后续共同使用。"""
    payload_rules = deepcopy(rules)
    payload_rules["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    payload_rules["updated_by"] = updated_by.strip()
    payload_rules["change_note"] = change_note.strip()
    content = json.dumps(payload_rules, ensure_ascii=False, indent=2) + "\n"
    LOCAL_RULE_FILE.parent.mkdir(parents=True, exist_ok=True)
    LOCAL_RULE_FILE.write_text(content, encoding="utf-8")

    settings = github_settings(secrets)
    if not settings["token"]:
        raise ValueError("尚未在 Streamlit Secrets 中配置 GitHub Token。")
    path = settings["rule_path"]
    api = f"https://api.github.com/repos/{settings['owner']}/{settings['repo']}/contents/{path}"
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
            raise RuntimeError(f"读取 GitHub 规则文件失败：HTTP {exc.code}") from exc

    body = {
        "message": f"更新周报统计规则：{change_note.strip() or '未填写说明'}",
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": settings["branch"],
    }
    if sha:
        body["sha"] = sha
    request = Request(api, data=json.dumps(body).encode("utf-8"), headers={**headers, "Content-Type": "application/json"}, method="PUT")
    try:
        with urlopen(request, timeout=30):
            pass
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"保存统计规则到 GitHub 失败：HTTP {exc.code} {detail}") from exc


def validate_rules(rules: dict) -> list[str]:
    errors = []
    for group in ("registration", "orders", "google", "bing"):
        if not rules.get("fields", {}).get(group):
            errors.append(f"缺少 {group} 字段映射")
    if not rules.get("new_user", {}).get("field"):
        errors.append("缺少新用户判断字段")
    channels = rules.get("channels", {}).get("blank_cooperation_sources", {})
    seen = {}
    for channel, values in channels.items():
        for value in values:
            key = str(value).strip().upper()
            if key and key in seen and seen[key] != channel:
                errors.append(f"末次来源 {value} 同时属于 {seen[key]} 和 {channel}")
            seen[key] = channel
    return errors
