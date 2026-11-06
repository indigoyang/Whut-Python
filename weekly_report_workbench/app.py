from datetime import date, timedelta
from copy import deepcopy
from pathlib import Path
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st
import pandas as pd

from report_engine import read_uploads
from v3_report_engine import generate_v3_reports
from target_store import (
    ALL_CHANNEL_TARGET_METRICS,
    ALL_CHANNEL_TARGET_ROWS,
    SEM_TARGET_METRICS,
    SEM_TARGET_ROWS,
    load_targets,
    save_targets,
)
from rule_store import load_rules, save_rules, validate_rules


st.set_page_config(page_title="周报—末次归因生成", page_icon="A", layout="wide", initial_sidebar_state="expanded")
admin_mode = str(st.query_params.get("admin", "")) == "1"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@400;500;600;700&display=swap');
*{box-sizing:border-box}:root{--paper:#f6f5f0;--surface:#fff;--ink:#111;--lime:#e3f85a;--muted:#6b7280;--line:#111;--soft:#ecebe5;--success:#dff8a8}
html{scroll-behavior:smooth}.stApp{background:var(--paper);color:var(--ink);font-family:'Inter','PingFang SC',sans-serif}.block-container{max-width:1480px;padding:34px 42px 80px;min-height:100vh;background:var(--paper)}
[data-testid="stHeader"]{background:transparent}[data-testid="stToolbar"]{right:1rem}[data-testid="stSidebar"]{width:80px!important;min-width:80px!important;background:var(--paper);border-right:1px solid #1112}[data-testid="stSidebar"]>div:first-child{width:80px!important;padding:24px 0}[data-testid="stSidebarCollapseButton"],[data-testid="collapsedControl"]{display:none}
.brand-mark{width:40px;height:40px;border-radius:50%;background:#111;color:var(--lime);display:grid;place-items:center;font-size:17px;font-weight:800;margin:0 auto 38px}.nav-link{width:40px;height:40px;display:grid;place-items:center;margin:14px auto;border-radius:50%;border:1px solid transparent;color:#9ca3af;text-decoration:none!important;font-size:17px;transition:.18s}.nav-link:hover{color:#111;background:#e5e7eb}.nav-link.active{color:#111;background:var(--lime);border-color:#111;box-shadow:2px 2px 0 #111;transform:translate(-1px,-1px)}.nav-space{height:300px}.nav-hint{font-size:9px;color:#9ca3af;text-align:center;line-height:1.35;margin:0 5px}
.topbar{padding:4px 0 28px;border-bottom:1px solid #1112;margin-bottom:24px}.eyebrow{font-size:11px;color:#6b7280;letter-spacing:2.2px;text-transform:uppercase;margin-bottom:9px}.topbar h1{font-family:'Noto Serif SC',serif;font-size:38px;line-height:1.2;font-weight:900;letter-spacing:-1.5px;margin:0 0 10px;color:#111}.topbar p{font-size:13px;color:#6b7280;margin:0}.steps{display:flex;align-items:center;gap:0;margin-top:28px}.step{display:flex;align-items:center;gap:8px;font-size:12px;color:#9ca3af;white-space:nowrap}.step b{width:25px;height:25px;display:grid;place-items:center;border-radius:50%;border:1px solid #9ca3af;font-size:9px}.step.done{color:#6b7280}.step.on{color:#111;font-weight:700}.step.on b{background:var(--lime);border-color:#111;color:#111}.step-line{width:48px;height:1px;background:#1113;margin:0 14px}
.upload-head{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:13px;border-bottom:1px solid #1112;margin-bottom:18px}.upload-head h2{font-family:'Noto Serif SC',serif;font-size:21px;margin:0 0 5px;color:#111;font-weight:700}.upload-head h2:before{content:'01';font-family:'Inter',sans-serif;color:#9ca3af;font-size:12px;margin-right:10px}.upload-head p{margin:0;font-size:11px;color:#6b7280}.upload-count{font-size:10px;color:#111;padding:5px 8px;background:#f3f4f6;border:1px solid #d1d5db}
[data-testid="stVerticalBlockBorderWrapper"]{border:1px solid #111!important;border-radius:0!important;padding:22px!important;background:#fff!important;box-shadow:none!important;margin-bottom:20px}[data-testid="stFileUploader"]{background:#fff;border:1px solid #111;border-radius:0;padding:13px 14px 7px;min-height:135px;transition:.18s}[data-testid="stFileUploader"]:hover{background:#e3f85a12}[data-testid="stFileUploader"] label{color:#111;font-weight:700;font-size:12px}[data-testid="stFileUploaderDropzone"]{min-height:64px;border:0;background:transparent;border-radius:0;padding:8px 0}[data-testid="stFileUploaderDropzone"] button{color:#111!important;background:var(--lime)!important;border:1px solid #111!important;border-radius:50%!important;width:35px!important;height:35px!important;min-width:35px!important;padding:0!important;font-size:0!important;box-shadow:none!important}[data-testid="stFileUploaderDropzone"] button:after{content:'＋';font-size:20px;color:#111;position:absolute;inset:0;display:grid;place-items:center}[data-testid="stFileUploaderDropzone"] small{font-size:9px;color:#9ca3af}
.file-status{margin:-6px 0 14px;padding:7px 9px;font-size:9px;letter-spacing:.4px;color:#9ca3af;background:#f7f7f5;border:1px solid #1111}.file-status.ok{color:#111;background:var(--success);border-color:#111}.panel{background:#fff;border:1px solid #111;padding:18px;margin-bottom:18px;box-shadow:none}.panel-title{display:flex;justify-content:space-between;align-items:center;font-size:13px;font-weight:700;padding-bottom:10px;border-bottom:1px solid #1112;margin-bottom:14px}.panel-title span{font-size:9px;font-weight:500;color:#9ca3af}.mini-metrics{display:flex;align-items:end;gap:12px;margin-bottom:14px}.mini-metric{border:0;background:transparent;padding:0}.mini-metric:first-child strong{font-family:'Noto Serif SC',serif;font-size:34px;font-weight:900}.mini-metric strong{display:block;font-size:15px}.mini-metric small{color:#9ca3af;font-size:9px}.progress-track{height:8px;background:#e5e7eb;overflow:hidden}.progress-track i{display:block;height:100%;background:var(--lime);border-right:1px solid #111}.task{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #1111;font-size:10px}.task:last-child{border:0}.task span{color:#9ca3af}.task .done{color:#111;background:var(--lime);border:1px solid #111;padding:1px 5px}
.timeline{border-left:1px solid #d1d5db;margin:5px 0 0 7px;padding-left:16px}.event{position:relative;margin:0 0 17px}.event:last-child{margin-bottom:0}.event:before{content:'';position:absolute;left:-21px;top:4px;width:8px;height:8px;border-radius:50%;background:#fff;border:1px solid #9ca3af}.event:first-child:before{background:var(--lime);border-color:#111}.event b{font-size:10px;display:block;margin-bottom:3px}.event span{font-size:9px;color:#9ca3af}
[data-testid="stExpander"]{border:1px solid #111!important;border-radius:0!important;background:#fff;box-shadow:none}[data-testid="stExpander"] summary{font-family:'Noto Serif SC',serif;font-weight:700}.stNumberInput input,.stDateInput input{border-radius:0!important;background:#fafafa!important}.stButton>button,.stDownloadButton>button{border-radius:0;min-height:43px;font-size:11px;font-weight:700;border:1px solid #111}.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{background:var(--lime);color:#111;border-color:#111;box-shadow:3px 3px 0 #111}.stButton>button[kind="primary"]:hover,.stDownloadButton>button[kind="primary"]:hover{transform:translate(-1px,-1px);box-shadow:4px 4px 0 #111}.stButton>button[kind="primary"]:disabled{background:#e5e7eb;border-color:#c7c7c7;color:#9ca3af;box-shadow:none}
.download-card{background:#fff;border:1px solid #111;padding:14px;margin:5px 0 8px}.download-card b{font-family:'Noto Serif SC',serif;font-size:12px}.download-card span{color:#9ca3af;font-size:9px}.section-label{font-family:'Noto Serif SC',serif;font-size:18px;font-weight:700;margin:28px 0 5px}.section-note{color:#6b7280;font-size:10px;margin-bottom:12px}.fixed-footer{height:52px;margin-top:28px;border-top:1px solid #1112;display:flex;align-items:center;justify-content:space-between}.fixed-footer span,.fixed-footer b{font-size:9px;color:#6b7280}footer{visibility:hidden}
@media(max-width:900px){[data-testid="stSidebar"]{display:none}.block-container{padding:20px 15px 55px}.topbar h1{font-size:30px}.step-line{width:13px;margin:0 6px}.steps{overflow-x:auto}[data-testid="stVerticalBlockBorderWrapper"]{padding:14px!important}}

/* 可读性修正版：收紧超宽屏、放大文字、修复上传按钮重叠 */
html{font-size:17px}
.block-container{max-width:1280px!important;padding:42px 42px 90px!important;margin:0 auto!important}
[data-testid="stSidebar"]{width:86px!important;min-width:86px!important}
[data-testid="stSidebar"]>div:first-child{width:86px!important;padding:28px 0!important}
.brand-mark{width:44px;height:44px;font-size:19px;margin-bottom:42px}
.nav-link{width:44px;height:44px;font-size:20px;margin:16px auto}
.nav-space{height:270px}.nav-hint{font-size:10px;color:#71717a;line-height:1.45}
.topbar{padding-bottom:32px;margin-bottom:28px}.eyebrow{font-size:13px;margin-bottom:12px}
.topbar h1{font-size:46px;margin-bottom:14px}.topbar p{font-size:15px}
.steps{margin-top:32px}.step{font-size:14px;gap:9px}.step b{width:29px;height:29px;font-size:11px}.step-line{width:52px;margin:0 15px}
.upload-head{padding-bottom:16px;margin-bottom:20px}.upload-head h2{font-size:25px}.upload-head h2:before{font-size:14px}.upload-head p{font-size:13px}.upload-count{font-size:12px;padding:7px 10px}
[data-testid="stVerticalBlockBorderWrapper"]{padding:24px!important;margin-bottom:22px!important}
[data-testid="stFileUploader"]{padding:16px 16px 10px!important;min-height:150px!important}
[data-testid="stFileUploader"] label{font-size:15px!important;line-height:1.45!important}
[data-testid="stFileUploaderDropzone"]{min-height:74px!important;padding:14px 0!important;display:flex!important;align-items:center!important;justify-content:flex-end!important}
[data-testid="stFileUploaderDropzoneInstructions"]{display:none!important}
[data-testid="stFileUploaderDropzone"] svg{display:none!important}
[data-testid="stFileUploaderDropzone"] small{display:none!important}
[data-testid="stFileUploaderDropzone"] button{color:transparent!important;background:var(--lime)!important;border:1px solid #111!important;border-radius:0!important;width:118px!important;min-width:118px!important;height:42px!important;padding:0!important;font-size:0!important;box-shadow:2px 2px 0 #111!important;position:relative!important}
[data-testid="stFileUploaderDropzone"] button:after{content:'上传文件'!important;font-size:14px!important;font-weight:700!important;color:#111!important;position:absolute!important;inset:0!important;display:grid!important;place-items:center!important}
.file-status{padding:9px 11px;font-size:11px;margin-bottom:16px}
.panel{padding:21px;margin-bottom:20px}.panel-title{font-size:16px;padding-bottom:12px;margin-bottom:16px}.panel-title span{font-size:11px;color:#71717a}
.mini-metrics{gap:16px;margin-bottom:16px}.mini-metric:first-child strong{font-size:40px}.mini-metric strong{font-size:18px}.mini-metric small{font-size:11px;color:#71717a}
.task{padding:10px 0;font-size:12px}.task .done{padding:2px 6px}.event b{font-size:13px}.event span{font-size:11px;color:#71717a}
.section-label{font-size:23px;margin-top:32px}.section-note{font-size:12px;margin-bottom:14px}
.download-card{padding:16px}.download-card b{font-size:15px}.download-card span{font-size:11px}
.fixed-footer span,.fixed-footer b{font-size:11px}
.stButton>button,.stDownloadButton>button{font-size:13px;min-height:45px}
@media(min-width:1600px){.block-container{max-width:1340px!important}.topbar h1{font-size:50px}}
@media(max-width:900px){html{font-size:16px}.block-container{padding:22px 16px 60px!important}.topbar h1{font-size:34px}.topbar p{font-size:13px}.step{font-size:12px}.step-line{width:14px;margin:0 6px}[data-testid="stVerticalBlockBorderWrapper"]{padding:15px!important}[data-testid="stFileUploader"]{min-height:132px!important}.panel{padding:16px}}

/* 上传卡片：整卡展示信息，右侧圆形加号触发文件选择 */
[data-testid="stFileUploader"]{position:relative!important;min-height:210px!important;padding:32px 34px!important;overflow:visible!important}
[data-testid="stFileUploader"] label{max-width:72%!important;margin:0!important;padding:0!important}
[data-testid="stFileUploader"] label p{font-size:16px!important;line-height:1.65!important;color:#6b7280!important;font-weight:500!important;margin:0!important}
[data-testid="stFileUploader"] label strong{display:block!important;font-family:'Noto Serif SC',serif!important;font-size:23px!important;line-height:1.35!important;color:#111!important;font-weight:900!important;margin-bottom:7px!important}
[data-testid="stFileUploader"]:after{content:'XLSX, XLSM, CSV · MAX 1GB';position:absolute;left:34px;bottom:30px;color:#9ca3af;font-size:13px;font-weight:600;letter-spacing:.3px}
[data-testid="stFileUploaderDropzone"]{position:absolute!important;right:30px!important;top:31px!important;width:58px!important;height:58px!important;min-height:58px!important;padding:0!important;display:block!important;background:transparent!important;overflow:visible!important}
[data-testid="stFileUploaderDropzone"] button{display:block!important;position:absolute!important;inset:0!important;width:58px!important;min-width:58px!important;height:58px!important;border-radius:50%!important;background:#fff!important;border:1.5px solid #111!important;box-shadow:none!important;padding:0!important}
[data-testid="stFileUploaderDropzone"] button:after{content:'＋'!important;font-family:'Inter',sans-serif!important;font-size:34px!important;line-height:1!important;font-weight:400!important;color:#111!important;position:absolute!important;inset:0!important;display:grid!important;place-items:center!important}
[data-testid="stFileUploader"]:hover [data-testid="stFileUploaderDropzone"] button{background:var(--lime)!important}
/* 上传后的原生文件名/删除按钮会压到右侧加号；状态由下方绿色提示条统一展示 */
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileData"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"]{display:none!important}
[data-testid="stFileUploader"] ul{display:none!important}
.file-status:not(.ok){display:none!important}.file-status.ok{margin:-8px 0 16px!important;font-size:12px!important}
@media(max-width:900px){[data-testid="stFileUploader"]{min-height:180px!important;padding:25px 22px!important}[data-testid="stFileUploader"] label strong{font-size:19px!important}[data-testid="stFileUploader"]:after{left:22px;bottom:24px;font-size:11px}[data-testid="stFileUploaderDropzone"]{right:20px!important;top:23px!important;width:50px!important;height:50px!important;min-height:50px!important}[data-testid="stFileUploaderDropzone"] button{width:50px!important;min-width:50px!important;height:50px!important}}
</style>
""", unsafe_allow_html=True)

if not admin_mode:
    st.markdown("<style>.st-key-admin_settings{display:none!important}</style>", unsafe_allow_html=True)

with st.sidebar:
    admin_link = "./" if admin_mode else "?admin=1#manual-data"
    admin_title = "退出管理模式" if admin_mode else "管理设置"
    st.markdown(f"""<div class="brand-mark" title="周报生成器">A</div><a class="nav-link active" href="#upload-data" title="上传数据">⇧</a><a class="nav-link" href="#manual-data" title="补充数据">＋</a><a class="nav-link" href="#report-result" title="生成结果">▦</a><a class="nav-link" href="{admin_link}" title="{admin_title}">⚙</a><a class="nav-link" href="#help" title="使用说明">?</a><div class="nav-space"></div><div class="nav-hint">WEEKLY<br>REPORT</div>""", unsafe_allow_html=True)

st.markdown("""<header class="topbar"><div class="eyebrow">Weekly Report Generator</div><h1>周报—末次归因生成</h1><p>上传数据，自动完成校验、归因计算并生成三份可编辑周报</p><div class="steps"><div class="step done"><b>✓</b>选择周期</div><div class="step-line"></div><div class="step on"><b>02</b>上传数据</div><div class="step-line"></div><div class="step"><b>03</b>补充信息</div><div class="step-line"></div><div class="step"><b>04</b>生成报告</div></div></header><div id="upload-data"></div>""", unsafe_allow_html=True)

today = date.today()
default_end = today - timedelta(days=1)
default_start = default_end - timedelta(days=6)
target_month_end = st.session_state.get("report_end", default_end)

main_col, side_col = st.columns([3.15, 1], gap="large")

with main_col:
    upload_container = st.container(border=True)
    with upload_container:
        upload_header = st.empty()
        u1, u2 = st.columns(2)
        with u1:
            reg_files = st.file_uploader("**注册数据**  \n用户注册与渠道来源", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if reg_files else ""}">{"✓ 已上传 "+str(len(reg_files))+" 个文件 · 等待格式检查" if reg_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
            google_files = st.file_uploader("**Google消耗**  \nGoogle Ads 投放消耗", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if google_files else ""}">{"✓ 已上传 "+str(len(google_files))+" 个文件 · 等待格式检查" if google_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
        with u2:
            order_files = st.file_uploader("**订单数据**  \n订单及收入相关数据", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if order_files else ""}">{"✓ 已上传 "+str(len(order_files))+" 个文件 · 等待格式检查" if order_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
            bing_files = st.file_uploader("**Bing消耗**  \nBing Ads 投放消耗", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if bing_files else ""}">{"✓ 已上传 "+str(len(bing_files))+" 个文件 · 等待格式检查" if bing_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
        uploaded_count = sum(bool(x) for x in [reg_files, order_files, google_files, bing_files])
        upload_header.markdown(f'<div class="upload-head"><div><h2>上传源数据</h2><p>上传4类数据，系统将自动完成格式检查和数据预处理</p></div><div class="upload-count">{uploaded_count} / 4 已完成</div></div>', unsafe_allow_html=True)

    st.markdown('<div id="manual-data"></div><div class="section-label">02 补充手动数据</div><div class="section-note">对应全渠道周报中的成本、流量和退款；没有数据时可保持为0。</div>', unsafe_allow_html=True)
    with st.expander("展开填写成本、流量和退款", expanded=False):
        input_values = {}
        periods = [("本周","本周"),("上周","上周"),("MTD","累计 / MTD")]
        groups = [("commission","affiliate分佣成本"),("seo","SEO—外链成本"),("kol","KOL成本"),("pc_flow","PC流量"),("h5_flow","H5流量"),("refund","PC端退款金额")]
        for prefix, title in groups:
            st.markdown(f"**{title}**")
            cols = st.columns(3)
            for col, (key, label) in zip(cols, periods):
                input_values[f"{prefix}_{key}"] = col.number_input(label, min_value=0.0, value=0.0, key=f"{prefix}_{key}")

    with st.expander("月度目标设置（每月填写一次）", expanded=False):
        month_key = target_month_end.strftime("%Y-%m")
        all_targets = load_targets()
        month_targets = all_targets.get(month_key, {})
        st.caption(
            f"当前维护月份：{month_key}。金额、人数直接填数值；百分比目标按百分数填写，"
            "例如75%填写75。保存后所有用户默认使用这组目标。"
        )
        target_tab_all, target_tab_sem = st.tabs(["全渠道目标", "SEM目标"])

        def target_editor(tab, rows_config, values, key):
            rows = pd.DataFrame(
                [
                    {"_目标键": metric_key or "", "指标": label,
                     "目标": values.get(metric_key) if metric_key else None}
                    for metric_key, label in rows_config
                ]
            )
            with tab:
                st.caption(
                    f"共 {len(rows_config)} 行；灰色分区标题只用于对应Excel结构，不保存目标。"
                )
                return st.data_editor(
                    rows,
                    column_config={
                        "_目标键": None,
                        "指标": st.column_config.TextColumn("指标", disabled=True),
                        "目标": st.column_config.NumberColumn("目标", min_value=0.0, format="%.2f"),
                    },
                    disabled=["指标"],
                    hide_index=True,
                    use_container_width=True,
                    height=420,
                    key=key,
                )

        edited_all = target_editor(
            target_tab_all, ALL_CHANNEL_TARGET_ROWS,
            month_targets.get("all_channel", {}), f"all_targets_{month_key}"
        )
        edited_sem = target_editor(
            target_tab_sem, SEM_TARGET_ROWS,
            month_targets.get("sem", {}), f"sem_targets_{month_key}"
        )

        if st.button("保存本月目标", type="primary", key="save_monthly_targets"):
            def clean(frame):
                result = {}
                for row in frame.to_dict("records"):
                    value = row.get("目标")
                    metric_key = str(row.get("_目标键", "")).strip()
                    if metric_key and value is not None and not pd.isna(value):
                        result[metric_key] = float(value)
                return result

            all_targets[month_key] = {
                "all_channel": clean(edited_all),
                "sem": clean(edited_sem),
            }
            try:
                save_targets(all_targets, st.secrets)
                st.success(f"{month_key}目标已保存，之后会默认使用。")
            except Exception as exc:
                st.error(str(exc))

    admin_settings = st.container(key="admin_settings")
    with admin_settings.expander("统计规则设置（口径变更时修改）", expanded=False):
        rules = load_rules()
        st.caption(
            "这里修改的是团队共用统计口径。保存后将写入 GitHub；建议先写清变更说明，"
            "并用同一周期生成新旧报告核对。"
        )

        def csv_text(values):
            return "，".join(str(v) for v in (values or []))

        def csv_values(value, keep_blank=False):
            values = [v.strip() for v in str(value).replace("，", ",").split(",")]
            return values if keep_blank else [v for v in values if v]

        field_tabs = st.tabs(["字段对应", "用户与渠道", "消耗与地区"])
        field_frames = {}
        field_names = {
            "registration": "注册数据", "orders": "订单数据",
            "google": "Google消耗", "bing": "Bing消耗",
        }
        with field_tabs[0]:
            st.info("左列是系统标准字段，不要修改；右列填写上传文件中的真实表头。")
            for group, title in field_names.items():
                st.markdown(f"**{title}**")
                mapping = rules.get("fields", {}).get(group, {})
                frame = pd.DataFrame([
                    {"系统标准字段": canonical, "上传文件字段": source}
                    for canonical, source in mapping.items()
                ])
                field_frames[group] = st.data_editor(
                    frame,
                    disabled=["系统标准字段"], hide_index=True,
                    use_container_width=True, key=f"rule_fields_{group}",
                )

        with field_tabs[1]:
            new_user_field = st.text_input(
                "新用户判断字段", rules.get("new_user", {}).get("field", "注册转化周期")
            )
            new_user_value = st.number_input(
                "新用户判断值", value=float(rules.get("new_user", {}).get("value", 0)), step=1.0
            )
            subscription_values = st.text_input(
                "续费对应的末次来源类型（多个用逗号隔开）",
                csv_text(rules.get("subscription", {}).get("source_values", ["AUTO_RENEWAL"])),
            )
            priorities = rules.get("channels", {}).get("cooperation_priority", [])
            priority_map = {item.get("channel"): item.get("prefix", "") for item in priorities}
            kol_prefix = st.text_input("KOL合作渠道前缀", priority_map.get("KOL", "KOL"))
            affiliate_prefix = st.text_input("Affiliate合作渠道前缀", priority_map.get("Affiliate", "AFFILIATE"))
            source_rules = rules.get("channels", {}).get("blank_cooperation_sources", {})
            source_inputs = {}
            for channel in ["SEM", "KOL", "Affiliate", "Unassigned"]:
                source_inputs[channel] = st.text_input(
                    f"合作渠道为空时，{channel}对应末次来源类型",
                    csv_text(source_rules.get(channel, [])), key=f"source_rule_{channel}",
                )
            default_channel = st.selectbox(
                "其他未匹配来源默认归类",
                ["SEO", "KOL", "SEM", "Affiliate", "Unassigned", "Blank"],
                index=["SEO", "KOL", "SEM", "Affiliate", "Unassigned", "Blank"].index(
                    rules.get("channels", {}).get("default_channel", "SEO")
                ),
            )

        with field_tabs[2]:
            google_excludes = st.text_input(
                "Google广告系列排除词（多个用逗号隔开）",
                csv_text(rules.get("sem_cost", {}).get("google_exclude_contains", ["KOL", "generic"])),
            )
            region_rules = rules.get("regions", {})
            english_regions = st.text_area("英语地区国家", csv_text(region_rules.get("english", [])))
            traditional_regions = st.text_area("繁中地区国家", csv_text(region_rules.get("traditional", [])))
            north_america_regions = st.text_input("北美地区国家", csv_text(region_rules.get("north_america", ["美国", "加拿大"])))

        editor_name = st.text_input("修改人", value=rules.get("updated_by", ""), key="rules_updated_by")
        change_note = st.text_input("本次修改说明（必填）", key="rules_change_note")
        if st.button("保存统计规则", type="primary", key="save_report_rules"):
            updated = deepcopy(rules)
            updated.setdefault("fields", {})
            for group, frame in field_frames.items():
                updated["fields"][group] = {
                    str(row["系统标准字段"]).strip(): str(row["上传文件字段"]).strip()
                    for row in frame.to_dict("records")
                    if str(row.get("系统标准字段", "")).strip()
                    and str(row.get("上传文件字段", "")).strip()
                }
            updated["new_user"] = {"field": new_user_field.strip(), "value": new_user_value}
            updated["subscription"] = {"source_values": csv_values(subscription_values)}
            updated["channels"] = {
                "cooperation_priority": [
                    {"prefix": kol_prefix.strip(), "channel": "KOL"},
                    {"prefix": affiliate_prefix.strip(), "channel": "Affiliate"},
                ],
                "blank_cooperation_sources": {
                    **{channel: csv_values(value) for channel, value in source_inputs.items()},
                    "Auto_Renewal": csv_values(subscription_values),
                    "Blank": [""],
                },
                "default_channel": default_channel,
            }
            updated["sem_cost"] = {"google_exclude_contains": csv_values(google_excludes)}
            updated["regions"] = {
                "english": csv_values(english_regions),
                "traditional": csv_values(traditional_regions),
                "north_america": csv_values(north_america_regions),
            }
            errors = validate_rules(updated)
            if not change_note.strip():
                errors.append("请填写本次修改说明")
            if errors:
                for error in errors:
                    st.error(error)
            else:
                try:
                    save_rules(updated, st.secrets, editor_name, change_note)
                    st.success("统计规则已保存。后续生成报告将使用新规则。")
                except Exception as exc:
                    st.error(str(exc))

    with admin_settings.expander("当前统计口径说明（团队必读）", expanded=False):
        metric_doc = Path(__file__).resolve().parent / "config" / "统计口径说明.md"
        if metric_doc.exists():
            st.markdown(metric_doc.read_text(encoding="utf-8"))
        else:
            st.warning("未找到统计口径说明文件。")

with side_col:
    st.markdown('<div class="panel"><div class="panel-title">报告信息<span>自动保存</span></div>', unsafe_allow_html=True)
    start = st.date_input("开始日期", default_start, key="report_start")
    end = st.date_input("结束日期", default_end, key="report_end")
    st.caption(f"当前周期 · {start:%Y.%m.%d} – {end:%Y.%m.%d}")
    st.markdown('</div>', unsafe_allow_html=True)

    completed = 1 + uploaded_count; readiness = int(completed / 5 * 100)
    task_rows = [("报告周期",True),("注册数据",bool(reg_files)),("订单数据",bool(order_files)),("Google消耗",bool(google_files)),("Bing消耗",bool(bing_files))]
    tasks_html = "".join(f'<div class="task"><b>{"✓" if done else "○"} {name}</b><span class="{"done" if done else ""}">{"已完成" if done else "待上传"}</span></div>' for name,done in task_rows)
    st.markdown(f'<div class="panel"><div class="panel-title">报告准备度<span>{readiness}%</span></div><div class="mini-metrics"><div class="mini-metric"><strong>{readiness}%</strong><small>当前进度</small></div><div class="mini-metric"><strong>{completed} / 5</strong><small>已完成项目</small></div></div><div class="progress-track"><i style="width:{readiness}%"></i></div><div style="height:10px"></div>{tasks_html}</div>', unsafe_allow_html=True)
    st.markdown('<div class="panel"><div class="panel-title">自动化流程<span>1–3分钟</span></div><div class="timeline"><div class="event"><b>数据文件检查</b><span>上传后自动校验必要字段</span></div><div class="event"><b>字段自动识别</b><span>兼容Google与Bing中文字段</span></div><div class="event"><b>报告生成</b><span>输出全渠道 / SEM / SEM分地区</span></div></div></div>', unsafe_allow_html=True)
    ready = all([reg_files, order_files, google_files, bing_files]) and start <= end
    generate_clicked = st.button("✦ 开始生成报告" if ready else "请先上传全部必要数据", type="primary", disabled=not ready, use_container_width=True)

if start > end:
    with side_col: st.error("开始日期不能晚于结束日期。")

if generate_clicked:
    try:
        with main_col:
            with st.status("正在读取和检查数据……", expanded=True) as status:
                reg = read_uploads(reg_files); st.write(f"注册数据：{len(reg):,} 行")
                orders = read_uploads(order_files); st.write(f"订单数据：{len(orders):,} 行")
                google = read_uploads(google_files); st.write(f"Google消耗：{len(google):,} 行")
                bing = read_uploads(bing_files); st.write(f"Bing消耗：{len(bing):,} 行")
                # 直接把当前目标配置交给生成引擎。这样刚在网站保存的目标无需等待
                # Streamlit 重新部署，也会立即写入本次生成的 Excel。
                reports, warnings = generate_v3_reports(
                    reg, orders, google, bing, None, start, end, input_values,
                    targets=load_targets(), rules=load_rules(),
                )
                st.session_state["reports"] = reports; st.session_state["warnings"] = warnings
                status.update(label="三份周报已生成", state="complete", expanded=False)
    except Exception as exc:
        with main_col: st.exception(exc)

if st.session_state.get("reports"):
    reports = st.session_state["reports"]
    with main_col:
        for warning in st.session_state.get("warnings", []): st.warning(warning)
        st.markdown('<div id="report-result"></div><div class="section-label">03 报告已就绪</div><div class="section-note">可分别下载Excel，也可以一次下载完整压缩包。</div>', unsafe_allow_html=True)
        cols = st.columns(3); metas = ["全渠道末次归因","SEM末次归因","SEM分地区末次归因"]
        for col, (filename, content), meta in zip(cols, reports.items(), metas):
            with col:
                st.markdown(f'<div class="download-card"><b>{meta}</b><br><span>Excel · {len(content)/1024:.1f} KB</span></div>', unsafe_allow_html=True)
                st.download_button("下载周报", content, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=filename)
        bundle = BytesIO()
        with ZipFile(bundle, "w", ZIP_DEFLATED) as archive:
            for filename, content in reports.items(): archive.writestr(filename, content)
        st.download_button("一键下载三份周报（ZIP）", bundle.getvalue(), f"末次归因周报_{start:%Y.%m.%d}-{end:%Y.%m.%d}.zip", "application/zip", type="primary", use_container_width=True)

st.markdown('<div id="help"></div><div class="fixed-footer"><span>文件仅用于本次浏览器会话 · 刷新页面后需重新上传</span><b>WEEKLY REPORT GENERATOR · 末次归因工作台</b></div>', unsafe_allow_html=True)
