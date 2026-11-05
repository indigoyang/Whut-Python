from datetime import date, timedelta
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from report_engine import read_uploads
from v3_report_engine import generate_v3_reports


st.set_page_config(page_title="AiPPT数据自动化", page_icon="✦", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
*{box-sizing:border-box}:root{--bg:#dfe2f2;--surface:#fff;--page:#fafafa;--primary:#7c6cf2;--primary-light:#f0eeff;--yellow:#e8f329;--success:#bdefe5;--success-text:#168477;--text:#16171b;--muted:#8a8c94;--border:#ececf0}
.stApp{background:var(--bg);color:var(--text)}.block-container{max-width:1360px;padding:24px 34px 94px;background:#fcfcfc;min-height:100vh}
[data-testid="stSidebar"]{width:72px!important;min-width:72px!important;background:#fbfbfc;border-right:1px solid var(--border)}[data-testid="stSidebar"]>div:first-child{width:72px!important;padding:18px 0}[data-testid="stSidebarCollapseButton"],[data-testid="collapsedControl"]{display:none}
.brand-mark{width:36px;height:36px;border-radius:10px;background:#15161a;color:#fff;display:grid;place-items:center;font-size:11px;font-weight:800;margin:0 auto 34px;position:relative}.brand-mark:after{content:"";width:7px;height:7px;background:var(--yellow);border-radius:50%;position:absolute;right:-2px;top:-2px}
.nav-icon{width:38px;height:38px;display:grid;place-items:center;margin:7px auto;border-radius:10px;color:#92949d;font-size:17px}.nav-icon.active{color:#fff;background:#15161a;box-shadow:0 6px 14px #15161a20}.nav-space{height:265px}.avatar{width:31px;height:31px;border-radius:50%;display:grid;place-items:center;margin:10px auto 0;background:linear-gradient(135deg,#d9d7ff,#8d82ed);color:#fff;font-size:11px;font-weight:700}
.topbar{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px}.eyebrow{font-size:11px;color:#8d8f99;letter-spacing:1.6px;margin-bottom:11px;font-weight:650}.topbar h1{font-size:30px;line-height:1.15;letter-spacing:-1px;margin:0 0 8px;color:#16171b}.topbar p{font-size:13px;color:var(--muted);margin:0}.brand-name{padding:10px 14px;border:1px solid #e7e7eb;background:#fff;border-radius:9px;color:#686a72;font-size:12px}
.steps{display:flex;align-items:center;gap:0;border-bottom:1px solid #ededf0;padding-bottom:17px;margin-bottom:22px}.step{display:flex;align-items:center;gap:8px;font-size:12px;color:#a4a5ac;white-space:nowrap}.step b{width:23px;height:23px;display:grid;place-items:center;border-radius:50%;background:#eeeef2;font-size:10px}.step.done{color:#65716f}.step.done b{background:#d6f3ed;color:#188476}.step.on{color:#15161a;font-weight:750}.step.on b{background:#16171b;color:#fff}.step-line{width:54px;height:1px;background:#e4e4e9;margin:0 13px}
.upload-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:15px}.upload-head h2{font-size:20px;margin:0 0 5px;color:#16171b}.upload-head h2:before{content:"01";display:inline-grid;place-items:center;width:27px;height:27px;margin-right:9px;border-radius:8px;background:#16171b;color:#fff;font-size:10px;vertical-align:2px}.upload-head p{margin:0;font-size:12px;color:var(--muted)}.upload-count{font-size:11px;color:#777982;padding:7px 10px;background:#fff;border:1px solid #e7e4ff;border-radius:8px}
[data-testid="stVerticalBlockBorderWrapper"]{border:2px solid #ded9ff!important;border-radius:17px!important;padding:20px!important;background:linear-gradient(135deg,#fff 0%,#f6f4ff 100%)!important;box-shadow:0 10px 28px #7c6cf218;margin-bottom:16px}[data-testid="stFileUploader"]{background:#ffffffb8;border:1px solid #ececf0;border-radius:12px;padding:11px 12px 5px;min-height:126px}[data-testid="stFileUploader"] label{color:#2c2d32;font-weight:700;font-size:12px}[data-testid="stFileUploaderDropzone"]{min-height:62px;border:1px dashed #d6d4e7;background:#fdfdff;border-radius:9px}[data-testid="stFileUploaderDropzone"] button{color:transparent!important;position:relative;min-width:88px;border-radius:8px}[data-testid="stFileUploaderDropzone"] button:after{content:"上传文件";color:#5f54cc;position:absolute;inset:0;display:grid;place-items:center}
.file-status{margin:-4px 0 13px;padding:7px 9px;border-radius:8px;font-size:10px;color:#777982;background:#f5f5f7}.file-status.ok{color:var(--success-text);background:#dff7f1}.panel{background:#fff;border:1px solid var(--border);border-radius:15px;padding:18px;margin-bottom:15px;box-shadow:0 8px 24px #4c52700a}.panel-title{display:flex;justify-content:space-between;align-items:center;font-size:14px;font-weight:750;margin-bottom:14px}.panel-title span{font-size:10px;font-weight:500;color:#a0a1a8}
.mini-metrics{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:14px}.mini-metric{border:1px solid #f0f0f2;background:#fdfdfd;border-radius:10px;padding:11px}.mini-metric strong{display:block;font-size:18px}.mini-metric small{color:#999;font-size:10px}.progress-track{height:6px;background:#e0def3;border-radius:9px;overflow:hidden}.progress-track i{display:block;height:100%;background:var(--primary);border-radius:9px}.task{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #f2f2f4;font-size:11px}.task:last-child{border:0}.task span{color:#999}.task .done{color:var(--success-text)}
.timeline{border-left:1px solid #e5e5ea;margin:5px 0 0 7px;padding-left:16px}.event{position:relative;margin:0 0 17px}.event:last-child{margin-bottom:0}.event:before{content:"";position:absolute;left:-21px;top:4px;width:8px;height:8px;border-radius:50%;background:#bcb7ff;border:3px solid #f0efff}.event b{font-size:11px;display:block;margin-bottom:3px}.event span{font-size:10px;color:#999}
[data-testid="stExpander"]{border:1px solid var(--border);border-radius:14px;background:#fff;box-shadow:0 8px 22px #4c527008}.stButton>button,.stDownloadButton>button{border-radius:9px;min-height:42px;font-size:12px;font-weight:750}.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{background:#16171b;color:#fff;border-color:#16171b}.stButton>button[kind="primary"]:disabled{background:#ddd;border-color:#ddd;color:#999}
.download-card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:13px;margin:5px 0 8px}.download-card b{font-size:12px}.download-card span{color:#999;font-size:10px}.section-label{font-size:15px;font-weight:750;margin:22px 0 5px}.section-note{color:#999;font-size:11px;margin-bottom:12px}.fixed-footer{height:54px;margin-top:26px;background:#fff;border-top:1px solid #e9e9ec;display:flex;align-items:center;justify-content:space-between;padding:0 6px}.fixed-footer span{color:#999;font-size:10px}.fixed-footer b{font-size:11px;color:#555}footer{visibility:hidden}
@media(max-width:900px){[data-testid="stSidebar"]{display:none}.block-container{padding:18px 14px 60px}.topbar{display:block}.brand-name{display:none}.step-line{width:16px;margin:0 6px}.steps{overflow-x:auto}[data-testid="stVerticalBlockBorderWrapper"]{padding:14px!important}}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("""<div class="brand-mark" title="AiPPT数据自动化">AI</div><div class="nav-icon active" title="工作台">⌂</div><div class="nav-icon" title="数据上传">⇧</div><div class="nav-icon" title="报告概览">▦</div><div class="nav-icon" title="历史报告">◫</div><div class="nav-icon" title="设置">⚙</div><div class="nav-space"></div><div class="nav-icon" title="帮助">?</div><div class="avatar" title="当前用户">Y</div>""", unsafe_allow_html=True)

st.markdown("""<header class="topbar"><div><div class="eyebrow">AIPPT DATA AUTOMATION</div><h1>周报—末次归因生成</h1><p>上传数据，自动完成校验、归因计算并生成三份可编辑周报</p></div><div class="brand-name">AiPPT数据自动化　···</div></header><div class="steps"><div class="step done"><b>✓</b>选择周期</div><div class="step-line"></div><div class="step on"><b>02</b>上传数据</div><div class="step-line"></div><div class="step"><b>03</b>补充信息</div><div class="step-line"></div><div class="step"><b>04</b>生成报告</div></div>""", unsafe_allow_html=True)

main_col, side_col = st.columns([3.15, 1], gap="large")

with main_col:
    upload_container = st.container(border=True)
    with upload_container:
        upload_header = st.empty()
        u1, u2 = st.columns(2)
        with u1:
            reg_files = st.file_uploader("注册数据 · 用户注册与渠道来源数据", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if reg_files else ""}">{"✓ 已上传 "+str(len(reg_files))+" 个文件 · 等待格式检查" if reg_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
            google_files = st.file_uploader("Google 消耗 · Google Ads 投放消耗", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if google_files else ""}">{"✓ 已上传 "+str(len(google_files))+" 个文件 · 等待格式检查" if google_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
        with u2:
            order_files = st.file_uploader("订单数据 · 订单及收入相关数据", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if order_files else ""}">{"✓ 已上传 "+str(len(order_files))+" 个文件 · 等待格式检查" if order_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
            bing_files = st.file_uploader("Bing 消耗 · Bing Ads 投放消耗", type=["xlsx","xlsm","csv"], accept_multiple_files=True)
            st.markdown(f'<div class="file-status{" ok" if bing_files else ""}">{"✓ 已上传 "+str(len(bing_files))+" 个文件 · 等待格式检查" if bing_files else "待上传 · XLSX / XLSM / CSV · 单文件最大1GB"}</div>', unsafe_allow_html=True)
        uploaded_count = sum(bool(x) for x in [reg_files, order_files, google_files, bing_files])
        upload_header.markdown(f'<div class="upload-head"><div><h2>上传源数据</h2><p>上传4类数据，系统将自动完成格式检查和数据预处理</p></div><div class="upload-count">{uploaded_count} / 4 已完成</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">02 补充手动数据</div><div class="section-note">对应全渠道周报中的成本、流量和退款；没有数据时可保持为0。</div>', unsafe_allow_html=True)
    with st.expander("展开填写成本、流量和退款", expanded=False):
        input_values = {}
        periods = [("本周","本周"),("上周","上周"),("MTD","累计 / MTD")]
        groups = [("commission","affiliate分佣成本"),("seo","SEO—外链成本"),("kol","KOL成本"),("pc_flow","PC流量"),("h5_flow","H5流量"),("refund","PC端退款金额")]
        for prefix, title in groups:
            st.markdown(f"**{title}**")
            cols = st.columns(3)
            for col, (key, label) in zip(cols, periods):
                input_values[f"{prefix}_{key}"] = col.number_input(label, min_value=0.0, value=0.0, key=f"{prefix}_{key}")

with side_col:
    st.markdown('<div class="panel"><div class="panel-title">报告信息<span>自动保存</span></div>', unsafe_allow_html=True)
    today = date.today(); default_end = today - timedelta(days=1); default_start = default_end - timedelta(days=6)
    start = st.date_input("开始日期", default_start); end = st.date_input("结束日期", default_end)
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
                reports, warnings = generate_v3_reports(reg, orders, google, bing, None, start, end, input_values)
                st.session_state["reports"] = reports; st.session_state["warnings"] = warnings
                status.update(label="三份周报已生成", state="complete", expanded=False)
    except Exception as exc:
        with main_col: st.exception(exc)

if st.session_state.get("reports"):
    reports = st.session_state["reports"]
    with main_col:
        for warning in st.session_state.get("warnings", []): st.warning(warning)
        st.markdown('<div class="section-label">03 报告已就绪</div><div class="section-note">可分别下载Excel，也可以一次下载完整压缩包。</div>', unsafe_allow_html=True)
        cols = st.columns(3); metas = ["全渠道末次归因","SEM末次归因","SEM分地区末次归因"]
        for col, (filename, content), meta in zip(cols, reports.items(), metas):
            with col:
                st.markdown(f'<div class="download-card"><b>{meta}</b><br><span>Excel · {len(content)/1024:.1f} KB</span></div>', unsafe_allow_html=True)
                st.download_button("下载周报", content, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=filename)
        bundle = BytesIO()
        with ZipFile(bundle, "w", ZIP_DEFLATED) as archive:
            for filename, content in reports.items(): archive.writestr(filename, content)
        st.download_button("一键下载三份周报（ZIP）", bundle.getvalue(), f"末次归因周报_{start:%Y.%m.%d}-{end:%Y.%m.%d}.zip", "application/zip", type="primary", use_container_width=True)

st.markdown('<div class="fixed-footer"><span>已自动保存 · 当前上传状态</span><b>AiPPT数据自动化　·　周报末次归因工作台</b></div>', unsafe_allow_html=True)
