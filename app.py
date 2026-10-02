"""
PR KPI管理ツール - Streamlit Webアプリ版
Accentureカラーテーマ
"""

import time
import random
import io
from datetime import date
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# デフォルト設定
# ============================================================
DEFAULT_COMPANIES = {
    "霞ヶ関キャピタル": "霞ヶ関キャピタル, 霞が関キャピタル, 霞ヶ関capital, Kasumigaseki Capital",
    "X NETWORK": "X NETWORK, X Network, クロスネットワーク, CROSS NETWORK",
    "東急不動産": "東急不動産, Tokyu Land, TOKYU LAND, 東急リバブル",
    "日本GLP": "日本GLP, GLP, Global Logistic Properties",
    "三井不動産": "三井不動産, Mitsui Fudosan, MITSUI FUDOSAN, 三井不動産ロジスティクスパーク",
    "プロロジス": "プロロジス, Prologis, PROLOGIS, プロロジスパーク",
    "大和ハウス": "大和ハウス, 大和ハウス工業, Daiwa House, DAIWA HOUSE, DPL",
    "野村不動産": "野村不動産, Nomura Real Estate, NOMURA REAL ESTATE, 野村不動産ロジスティクス",
    "三菱地所": "三菱地所, Mitsubishi Estate, MITSUBISHI ESTATE, 三菱地所ロジスティクス",
    "森トラスト": "森トラスト, Mori Trust, MORI TRUST",
    "東京建物": "東京建物, Tokyo Tatemono, TOKYO TATEMONO, T-LOGI",
}

DEFAULT_KEYWORDS = {
    "キーワード1": "例: 冷凍倉庫, 冷凍・冷蔵倉庫, フリーザー倉庫",
    "キーワード2": "例: 物流REIT, 物流不動産投資",
    "キーワード3": "例: 物流DX, ロジスティクスDX",
    "キーワード4": "例: ラストワンマイル, 最終配送",
    "キーワード5": "例: サプライチェーン, supply chain",
}

MEDIA_CATEGORY1 = ["全国", "ローカル・専門"]
MEDIA_CATEGORY2 = ["TV", "紙", "WEB"]
CONTACT_METHODS = ["取材", "面会", "メール"]

# 処理速度：高速設定
MAX_WORKERS = 15
TIMEOUT = 10
MAX_RETRIES = 1
DELAY_MIN = 0.3
DELAY_MAX = 1.0

PAYWALL_KEYWORDS = [
    "有料記事", "会員限定", "ログインが必要", "続きを読む",
    "購読が必要", "プレミアム会員", "有料会員",
]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4.1 Safari/605.1.15",
]

# Accentureカラー
ACN_PURPLE = "#A100FF"
ACN_BLACK  = "#000000"
ACN_GRAY   = "#F2F2F2"
ACN_DARK   = "#1A1A1A"


# ============================================================
# グローバルCSS（Accentureテーマ）
# ============================================================
def apply_accenture_theme():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Noto Sans JP', sans-serif;
    }}

    /* 背景 */
    .stApp {{
        background-color: #FFFFFF;
    }}

    /* ヘッダーバー */
    .acn-header {{
        background: {ACN_BLACK};
        padding: 16px 32px;
        display: flex;
        align-items: center;
        gap: 16px;
        margin-bottom: 8px;
    }}
    .acn-header .logo {{
        color: {ACN_PURPLE};
        font-size: 28px;
        font-weight: 700;
        letter-spacing: -1px;
    }}
    .acn-header .title {{
        color: #FFFFFF;
        font-size: 18px;
        font-weight: 400;
        letter-spacing: 0.5px;
    }}
    .acn-header .accent {{
        color: {ACN_PURPLE};
    }}

    /* タブ */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0px;
        border-bottom: 2px solid {ACN_PURPLE};
        background: {ACN_BLACK};
        padding: 0 16px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: #999999;
        background: transparent;
        border: none;
        padding: 12px 20px;
        font-size: 13px;
        font-weight: 500;
        letter-spacing: 0.3px;
    }}
    .stTabs [aria-selected="true"] {{
        color: #FFFFFF !important;
        background: {ACN_PURPLE} !important;
        border-radius: 0px;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: #FFFFFF !important;
        background: #333333 !important;
    }}

    /* プライマリボタン */
    .stButton > button[kind="primary"] {{
        background: {ACN_PURPLE};
        color: white;
        border: none;
        border-radius: 0px;
        font-weight: 600;
        font-size: 13px;
        letter-spacing: 0.5px;
        padding: 10px 24px;
        transition: all 0.2s;
    }}
    .stButton > button[kind="primary"]:hover {{
        background: #8800DD;
        transform: translateY(-1px);
    }}

    /* セカンダリボタン */
    .stButton > button {{
        border-radius: 0px;
        font-size: 13px;
        border: 1px solid #CCCCCC;
    }}

    /* メトリクスカード */
    [data-testid="metric-container"] {{
        background: {ACN_GRAY};
        border-left: 4px solid {ACN_PURPLE};
        padding: 16px;
        border-radius: 0px;
    }}
    [data-testid="metric-container"] label {{
        font-size: 12px;
        color: #666666;
        font-weight: 500;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }}
    [data-testid="metric-container"] [data-testid="stMetricValue"] {{
        font-size: 28px;
        font-weight: 700;
        color: {ACN_BLACK};
    }}

    /* expander */
    .streamlit-expanderHeader {{
        background: {ACN_GRAY};
        border-left: 3px solid {ACN_PURPLE};
        font-weight: 500;
    }}

    /* テキスト入力 */
    .stTextInput input {{
        border-radius: 0px;
        border: 1px solid #CCCCCC;
        border-bottom: 2px solid {ACN_PURPLE};
        font-size: 13px;
    }}
    .stTextInput input:focus {{
        border-color: {ACN_PURPLE};
        box-shadow: none;
    }}

    /* セレクトボックス */
    .stSelectbox select {{
        border-radius: 0px;
        border-bottom: 2px solid {ACN_PURPLE};
    }}

    /* divider */
    hr {{
        border-color: {ACN_GRAY};
        margin: 20px 0;
    }}

    /* セクションタイトル */
    .section-title {{
        font-size: 13px;
        font-weight: 700;
        color: {ACN_PURPLE};
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 8px;
        padding-bottom: 4px;
        border-bottom: 1px solid {ACN_PURPLE};
    }}

    /* 成功・警告 */
    .stSuccess {{
        border-left: 4px solid {ACN_PURPLE};
        border-radius: 0px;
    }}

    /* dataframe */
    .stDataFrame {{
        border: 1px solid #E0E0E0;
    }}

    /* ダウンロードボタン */
    .stDownloadButton > button {{
        background: {ACN_BLACK};
        color: white;
        border-radius: 0px;
        font-weight: 600;
        border: none;
        padding: 10px 24px;
    }}
    .stDownloadButton > button:hover {{
        background: #333333;
    }}

    /* プログレスバー */
    .stProgress > div > div {{
        background: {ACN_PURPLE};
    }}

    /* キャプション */
    .stCaption {{
        color: #888888;
        font-size: 12px;
    }}
    </style>
    """, unsafe_allow_html=True)

    # ヘッダー
    st.markdown(f"""
    <div class="acn-header">
        <div class="logo">></div>
        <div class="title">PR <span class="accent">KPI</span> 管理ツール &nbsp;|&nbsp; 霞ヶ関キャピタル インフラ事業本部</div>
    </div>
    """, unsafe_allow_html=True)


def section_title(text):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


# ============================================================
# 共通処理
# ============================================================
def fetch_text(url: str) -> tuple:
    for attempt in range(MAX_RETRIES):
        try:
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
            headers = {
                "User-Agent": random.choice(USER_AGENTS),
                "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Referer": "https://www.google.co.jp/",
            }
            resp = requests.get(url, headers=headers, timeout=TIMEOUT, allow_redirects=True)
            if resp.status_code in (401, 403):
                return "error", ""
            if resp.status_code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            soup = BeautifulSoup(resp.text, "html.parser")
            page_text = soup.get_text(separator=" ", strip=True)
            if any(kw in page_text for kw in PAYWALL_KEYWORDS) and len(page_text) < 300:
                return "paywall", ""
            for tag in ["article", "main", "body"]:
                container = soup.find(tag)
                if container:
                    text = container.get_text(separator=" ", strip=True)
                    if len(text) > 100:
                        return "success", text
            return "success", page_text
        except requests.exceptions.Timeout:
            if attempt == MAX_RETRIES - 1:
                return "timeout", ""
            time.sleep(2 * (attempt + 1))
        except requests.exceptions.ConnectionError:
            if attempt == MAX_RETRIES - 1:
                return "error", ""
            time.sleep(2 * (attempt + 1))
        except Exception:
            if attempt == MAX_RETRIES - 1:
                return "error", ""
    return "error", ""


def check_keywords(text: str, keywords: dict) -> dict:
    text_lower = text.lower()
    return {
        label: any(alias.strip().lower() in text_lower for alias in aliases.split(","))
        for label, aliases in keywords.items()
    }


def load_input(uploaded_file) -> pd.DataFrame:
    df_raw = pd.read_excel(uploaded_file, header=None)
    header_row = 1
    for i, row in df_raw.iterrows():
        if "No" in row.values or "記事URL" in row.values:
            header_row = i
            break
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row)
    url_col = next(
        (col for col in df.columns if "URL" in str(col).upper() or "url" in str(col).lower()),
        None
    )
    if url_col is None:
        raise ValueError("URL列が見つかりません")
    df = df.rename(columns={url_col: "記事URL"})
    df = df.dropna(subset=["記事URL"])
    df["記事URL"] = df["記事URL"].astype(str).str.strip()
    df = df[df["記事URL"].str.startswith("http")]
    return df


def run_analysis(df, keywords, progress_bar, status_text, log_area):
    results = []
    rows = df.to_dict("records")
    total = len(rows)
    done = 0
    log_lines = []

    def process_url(row_data):
        url = row_data.get("記事URL", "")
        status, text = fetch_text(url)
        kw_results = (
            check_keywords(text, keywords)
            if status == "success"
            else {k: None for k in keywords}
        )
        return {**row_data, "取得ステータス": status, **kw_results}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_url, row): row for row in rows}
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            done += 1
            progress_bar.progress(done / total)
            status_text.text(f"処理中... {done} / {total} 件")
            url_short = result.get("記事URL", "")[:60]
            st_status = result.get("取得ステータス", "")
            log_lines.append(f"`{st_status}` {url_short}")
            if len(log_lines) > 5:
                log_lines.pop(0)
            log_area.markdown("\n".join(log_lines))

    status_text.text(f"✅ 完了！ {total} 件処理しました")
    progress_bar.progress(1.0)
    return results


def render_status_metrics(results):
    status_counts = {}
    for r in results:
        s = r.get("取得ステータス", "unknown")
        status_counts[s] = status_counts.get(s, 0) + 1
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("取得成功", f"{status_counts.get('success', 0)} 件")
    c2.metric("有料記事", f"{status_counts.get('paywall', 0)} 件")
    c3.metric("エラー", f"{status_counts.get('error', 0)} 件")
    c4.metric("タイムアウト", f"{status_counts.get('timeout', 0)} 件")


def render_settings(state_key, defaults, label_placeholder, alias_placeholder):
    if state_key not in st.session_state:
        st.session_state[state_key] = [{"name": k, "aliases": v} for k, v in defaults.items()]
    to_delete = []
    for i, item in enumerate(st.session_state[state_key]):
        col1, col2, col3 = st.columns([2, 5, 0.5])
        with col1:
            st.session_state[state_key][i]["name"] = st.text_input(
                "名前", value=item["name"], key=f"{state_key}_name_{i}",
                label_visibility="collapsed", placeholder=label_placeholder)
        with col2:
            st.session_state[state_key][i]["aliases"] = st.text_input(
                "類似表現", value=item["aliases"], key=f"{state_key}_alias_{i}",
                label_visibility="collapsed", placeholder=alias_placeholder)
        with col3:
            if st.button("✕", key=f"{state_key}_del_{i}"):
                to_delete.append(i)
    for i in reversed(to_delete):
        st.session_state[state_key].pop(i)
        st.rerun()
    col_add, col_reset = st.columns(2)
    with col_add:
        if st.button("＋ 追加", key=f"{state_key}_add", use_container_width=True):
            st.session_state[state_key].append({"name": "", "aliases": ""})
            st.rerun()
    with col_reset:
        if st.button("↺ デフォルトに戻す", key=f"{state_key}_reset", use_container_width=True):
            st.session_state[state_key] = [{"name": k, "aliases": v} for k, v in defaults.items()]
            st.rerun()
    return {item["name"]: item["aliases"] for item in st.session_state[state_key] if item["name"].strip()}


def build_excel_sov(results, companies):
    kw_cols = list(companies.keys())
    df = pd.DataFrame(results)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "記事詳細"
    hf = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    hn = Font(color="FFFFFF", bold=True, size=10)
    cf = PatternFill(start_color="A100FF", end_color="A100FF", fill_type="solid")
    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + kw_cols
    ws1.append(out_cols)
    for col_idx, col_name in enumerate(out_cols, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = cf if col_name in kw_cols else hf
        cell.font = hn
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sc = {"success": "E8D5FF", "paywall": "FFF2CC", "error": "FCE4D6", "timeout": "FCE4D6"}
    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in kw_cols:
                row.append("○" if val is True else ("-" if val is None else ""))
            else:
                row.append(val if pd.notna(val) else "")
        ws1.append(row)
        cr = ws1.max_row
        status = row_data.get("取得ステータス", "")
        fill = PatternFill(start_color=sc.get(status, "FFFFFF"), end_color=sc.get(status, "FFFFFF"), fill_type="solid")
        for ci in range(1, len(out_cols) + 1):
            cell = ws1.cell(row=cr, column=ci)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center" if out_cols[ci-1] in kw_cols else "left", vertical="center")
    cw = {"No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20, "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14}
    for ci, cn in enumerate(out_cols, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = cw.get(cn, 14)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions
    ws2 = wb.create_sheet("SOV集計")
    sdf = df[df["取得ステータス"] == "success"]
    ts = len(sdf)
    ws2.append(["SOV集計レポート"])
    ws2["A1"].font = Font(bold=True, size=14, color="A100FF")
    ws2.append([])
    ws2.append(["集計対象記事数（取得成功）", ts])
    ws2.append(["有料記事・取得不可（除外）", len(df) - ts])
    ws2.append(["総記事数", len(df)])
    ws2.append([])
    ws2.append(["企業名", "記載あり記事数", "競合間SOV（%）", "総記事SOV（%）"])
    for ci in range(1, 5):
        cell = ws2.cell(row=7, column=ci)
        cell.fill = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    ws2.append(["", "", "※企業合計を母数", "※取得成功全記事を母数"])
    tm = 0
    cc = {}
    for company in kw_cols:
        if company in sdf.columns:
            count = int((sdf[company] == True).sum())
            cc[company] = count
            tm += count
    for company, count in cc.items():
        ws2.append([company, count, round(count/tm*100,1) if tm>0 else 0, round(count/ts*100,1) if ts>0 else 0])
    ws2.append(["合計", tm, 100.0, round(tm/ts*100,1) if ts>0 else 0])
    tr = ws2.max_row
    for ci in range(1, 5):
        cell = ws2.cell(row=tr, column=ci)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="E8D5FF", end_color="E8D5FF", fill_type="solid")
    for col, w in [("A",20),("B",18),("C",18),("D",18)]:
        ws2.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel_keyword(results, keywords):
    kw_cols = list(keywords.keys())
    df = pd.DataFrame(results)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "記事詳細"
    hf = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    hn = Font(color="FFFFFF", bold=True, size=10)
    kf = PatternFill(start_color="A100FF", end_color="A100FF", fill_type="solid")
    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + kw_cols
    ws1.append(out_cols)
    for ci, cn in enumerate(out_cols, 1):
        cell = ws1.cell(row=1, column=ci)
        cell.fill = kf if cn in kw_cols else hf
        cell.font = hn
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sc = {"success": "E8D5FF", "paywall": "FFF2CC", "error": "FCE4D6", "timeout": "FCE4D6"}
    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in kw_cols:
                row.append("○" if val is True else ("-" if val is None else ""))
            else:
                row.append(val if pd.notna(val) else "")
        ws1.append(row)
        cr = ws1.max_row
        status = row_data.get("取得ステータス", "")
        fill = PatternFill(start_color=sc.get(status, "FFFFFF"), end_color=sc.get(status, "FFFFFF"), fill_type="solid")
        for ci in range(1, len(out_cols)+1):
            cell = ws1.cell(row=cr, column=ci)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center" if out_cols[ci-1] in kw_cols else "left", vertical="center")
    cw = {"No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20, "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14}
    for ci, cn in enumerate(out_cols, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = cw.get(cn, 14)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions
    ws2 = wb.create_sheet("キーワード集計")
    sdf = df[df["取得ステータス"] == "success"]
    ts = len(sdf)
    ws2.append(["キーワード露出集計レポート"])
    ws2["A1"].font = Font(bold=True, size=14, color="A100FF")
    ws2.append([])
    ws2.append(["集計対象記事数（取得成功）", ts])
    ws2.append(["有料記事・取得不可（除外）", len(df)-ts])
    ws2.append(["総記事数", len(df)])
    ws2.append([])
    ws2.append(["キーワード", "露出記事数", "露出率（%）", "設定キーワード・類似表現"])
    for ci in range(1, 5):
        cell = ws2.cell(row=7, column=ci)
        cell.fill = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for kw_label, aliases in keywords.items():
        if kw_label in sdf.columns:
            count = int((sdf[kw_label] == True).sum())
            pct = round(count/ts*100,1) if ts>0 else 0
            ws2.append([kw_label, count, pct, aliases])
    for col, w in [("A",20),("B",14),("C",12),("D",50)]:
        ws2.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel_log(log_key, title):
    logs = st.session_state.get(log_key, [])
    if not logs:
        return None
    df = pd.DataFrame(logs)
    wb = Workbook()
    ws = wb.active
    ws.title = title
    hf = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    hn = Font(color="FFFFFF", bold=True)
    ws.append(list(df.columns))
    for ci in range(1, len(df.columns)+1):
        cell = ws.cell(row=1, column=ci)
        cell.fill = hf
        cell.font = hn
        cell.alignment = Alignment(horizontal="center")
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for ci, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(ci)].width = 20
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def render_contact_log_tab(tab_key, log_key):
    section_title("新規入力")
    col1, col2 = st.columns(2)
    with col1:
        input_date = st.date_input("日付", value=date.today(), key=f"{tab_key}_date")
        cat1 = st.selectbox("メディア種別①", MEDIA_CATEGORY1, key=f"{tab_key}_cat1")
    with col2:
        media_name = st.text_input("媒体名", key=f"{tab_key}_media", placeholder="例：日本経済新聞")
        cat2 = st.selectbox("メディア種別②", MEDIA_CATEGORY2, key=f"{tab_key}_cat2")
    col3, col4 = st.columns(2)
    with col3:
        method = st.selectbox("受け渡し方法", CONTACT_METHODS, key=f"{tab_key}_method")
    with col4:
        note = st.text_input("備考（用件）", key=f"{tab_key}_note", placeholder="例：KC東扇島施設について取材依頼")

    if st.button("＋ 追加する", type="primary", use_container_width=True, key=f"{tab_key}_add"):
        if not media_name.strip():
            st.warning("媒体名を入力してください")
        else:
            if log_key not in st.session_state:
                st.session_state[log_key] = []
            st.session_state[log_key].append({
                "日付": str(input_date),
                "媒体名": media_name.strip(),
                "種別①": cat1,
                "種別②": cat2,
                "受け渡し方法": method,
                "備考": note.strip(),
            })
            st.success(f"✅ 追加しました：{media_name}")
            st.rerun()

    st.divider()
    logs = st.session_state.get(log_key, [])
    if logs:
        df_log = pd.DataFrame(logs)
        df_log["月"] = pd.to_datetime(df_log["日付"]).dt.strftime("%Y年%m月")
        monthly = df_log.groupby("月").size().reset_index(name="件数").sort_values("月", ascending=False)

        section_title("集計")
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("累計件数", f"{len(logs)} 件")
        if not monthly.empty:
            latest = monthly.iloc[0]
            col_b.metric(f"{latest['月']}（直近）", f"{latest['件数']} 件")

        col_left, col_right = st.columns(2)
        with col_left:
            st.caption("月別集計")
            st.dataframe(monthly, use_container_width=True, hide_index=True)
        with col_right:
            st.caption("種別内訳")
            cat_summary = df_log.groupby(["種別①", "種別②"]).size().reset_index(name="件数")
            st.dataframe(cat_summary, use_container_width=True, hide_index=True)

        st.divider()
        section_title("詳細ログ")
        for i, log in enumerate(reversed(logs)):
            idx = len(logs) - 1 - i
            with st.expander(f"{log['日付']}　{log['媒体名']}　{log['種別①']}/{log['種別②']}　{log['受け渡し方法']}"):
                st.write(f"**備考：** {log['備考'] if log['備考'] else '（なし）'}")
                if st.button("🗑 削除", key=f"{tab_key}_del_{idx}"):
                    st.session_state[log_key].pop(idx)
                    st.rerun()

        st.divider()
        excel_data = build_excel_log(log_key, log_key)
        if excel_data:
            st.download_button(
                label="📥 Excelでダウンロード", data=excel_data,
                file_name=f"{log_key}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
    else:
        st.info("まだ記録がありません。上のフォームから追加してください。")


# ============================================================
# メイン
# ============================================================
st.set_page_config(page_title="PR KPI管理ツール", page_icon="📊", layout="wide")
apply_accenture_theme()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏢  SOV分析",
    "🔍  キーワード分析",
    "📞  メディア接触数",
    "🤝  新規媒体リレーション",
    "📈  KPI管理",
])

# ---- Tab1: SOV分析 ----
with tab1:
    st.caption("競合企業の記事露出量を比較し、Share of Voiceを集計します")
    with st.expander("⚙️ チェック対象企業の設定", expanded=False):
        companies = render_settings("companies", DEFAULT_COMPANIES, "企業名", "表記ゆれをカンマ区切りで入力")
    section_title("チェック対象企業")
    if companies:
        cols = st.columns(4)
        for i, company in enumerate(companies.keys()):
            cols[i % 4].markdown(f"- {company}")
    st.divider()
    uploaded_file_sov = st.file_uploader("クリッピングデータのExcelをアップロード", type=["xlsx","xls"], key="sov_upload")
    if uploaded_file_sov and companies:
        try:
            df_sov = load_input(uploaded_file_sov)
            st.success(f"✅ {len(df_sov)} 件のURLを検出しました")
            if st.button("▶  SOV分析を開始", type="primary", use_container_width=True, key="sov_start"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()
                results = run_analysis(df_sov, companies, progress_bar, status_text, log_area)
                st.session_state["sov_results"] = results
                st.session_state["sov_companies"] = companies

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

    # 結果表示（セッションステートから）
    if "sov_results" in st.session_state and st.session_state["sov_results"]:
        results = st.session_state["sov_results"]
        companies_used = st.session_state.get("sov_companies", companies)
        st.divider()
        section_title("SOV集計結果")
        success_results = [r for r in results if r.get("取得ステータス") == "success"]
        total_success = len(success_results)
        total_mentions = 0
        sov_data = []
        for company in companies_used:
            count = sum(1 for r in success_results if r.get(company) is True)
            total_mentions += count
            sov_data.append({"企業名": company, "記載記事数": count})
        for row in sov_data:
            count = row["記載記事数"]
            row["競合間SOV（%）"] = round(count/total_mentions*100,1) if total_mentions>0 else 0
            row["総記事SOV（%）"] = round(count/total_success*100,1) if total_success>0 else 0
        st.dataframe(pd.DataFrame(sov_data).sort_values("競合間SOV（%）", ascending=False), use_container_width=True, hide_index=True)
        section_title("処理ステータス")
        render_status_metrics(results)
        st.divider()
        st.download_button(label="📥 結果をExcelでダウンロード", data=build_excel_sov(results, companies_used), file_name="sov_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ---- Tab2: キーワード分析 ----
with tab2:
    st.caption("指定したキーワードが記事に含まれているか判定し、露出件数・露出率を集計します")
    with st.expander("⚙️ チェック対象キーワードの設定", expanded=True):
        keywords = render_settings("keywords", DEFAULT_KEYWORDS, "キーワード名", "類似表現をカンマ区切りで入力")
    section_title("チェック対象キーワード")
    if keywords:
        for label, aliases in keywords.items():
            st.markdown(f"- **{label}**：{aliases}")
    st.divider()
    uploaded_file_kw = st.file_uploader("クリッピングデータのExcelをアップロード", type=["xlsx","xls"], key="kw_upload")
    if uploaded_file_kw and keywords:
        try:
            df_kw = load_input(uploaded_file_kw)
            st.success(f"✅ {len(df_kw)} 件のURLを検出しました")
            if st.button("▶  キーワード分析を開始", type="primary", use_container_width=True, key="kw_start"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()
                results = run_analysis(df_kw, keywords, progress_bar, status_text, log_area)
                st.session_state["kw_results"] = results
                st.session_state["kw_keywords"] = keywords
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

    # 結果表示（セッションステートから）
    if "kw_results" in st.session_state and st.session_state["kw_results"]:
        results = st.session_state["kw_results"]
        keywords_used = st.session_state.get("kw_keywords", keywords)
        st.divider()
        section_title("キーワード露出集計")
        success_results = [r for r in results if r.get("取得ステータス") == "success"]
        total_success = len(success_results)
        kw_data = []
        for label in keywords_used:
            count = sum(1 for r in success_results if r.get(label) is True)
            pct = round(count/total_success*100,1) if total_success>0 else 0
            kw_data.append({"キーワード": label, "露出記事数": count, "露出率（%）": pct, "設定キーワード": keywords_used[label]})
        st.dataframe(pd.DataFrame(kw_data).sort_values("露出記事数", ascending=False), use_container_width=True, hide_index=True)
        section_title("処理ステータス")
        render_status_metrics(results)
        st.divider()
        st.download_button(label="📥 結果をExcelでダウンロード", data=build_excel_keyword(results, keywords_used), file_name="keyword_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ---- Tab3: メディア接触数 ----
with tab3:
    st.caption("メディアへのアプローチ・接触を記録します（既存リレーション含む）")
    render_contact_log_tab("contact", "contact_log")

# ---- Tab4: 新規媒体リレーション ----
with tab4:
    st.caption("今まで関わりがなかった媒体をクライアントにつないだ件数を記録します")
    render_contact_log_tab("relation", "relation_log")

# ---- Tab5: KPI管理 ----
with tab5:
    st.caption("月次KPIの記録と累計サマリーを確認できます")

    section_title("月次データ入力（キーワード露出・SOV）")
    col1, col2 = st.columns(2)
    with col1:
        kpi_month = st.selectbox("対象月", [f"{y}年{m:02d}月" for y in range(2025,2028) for m in range(1,13)], index=9, key="kpi_month")
        kw_exposure = st.number_input("キーワード露出件数", min_value=0, step=1, key="kpi_kw")
    with col2:
        competitive_sov = st.number_input("競合間SOV（%）", min_value=0.0, max_value=100.0, step=0.1, key="kpi_comp_sov")
        total_sov = st.number_input("総記事SOV（%）", min_value=0.0, max_value=100.0, step=0.1, key="kpi_total_sov")

    if st.button("＋ 月次データを追加", type="primary", use_container_width=True, key="kpi_add"):
        if "kpi_monthly" not in st.session_state:
            st.session_state["kpi_monthly"] = []
        existing = [r["月"] for r in st.session_state["kpi_monthly"]]
        if kpi_month in existing:
            st.warning(f"⚠️ {kpi_month} のデータはすでに存在します")
        else:
            st.session_state["kpi_monthly"].append({
                "月": kpi_month,
                "キーワード露出件数": kw_exposure,
                "競合間SOV（%）": competitive_sov,
                "総記事SOV（%）": total_sov,
            })
            st.success(f"✅ {kpi_month} のデータを追加しました")
            st.rerun()

    st.divider()
    contact_logs = st.session_state.get("contact_log", [])
    relation_logs = st.session_state.get("relation_log", [])
    monthly_data = st.session_state.get("kpi_monthly", [])

    section_title("KPIサマリー（累計）")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📞 メディア接触数", f"{len(contact_logs)} 件")
    c2.metric("🤝 新規媒体リレーション", f"{len(relation_logs)} 件")
    if monthly_data:
        df_m = pd.DataFrame(monthly_data).sort_values("月")
        c3.metric("🔍 キーワード露出（累計）", f"{df_m['キーワード露出件数'].sum()} 件")
        c4.metric("🏢 直近競合間SOV", f"{df_m.iloc[-1]['競合間SOV（%）']} %")
    else:
        c3.metric("🔍 キーワード露出（累計）", "- 件")
        c4.metric("🏢 直近競合間SOV", "- %")

    if monthly_data:
        st.divider()
        section_title("月次KPI推移")
        df_monthly = pd.DataFrame(monthly_data).sort_values("月", ascending=False)
        st.dataframe(df_monthly, use_container_width=True, hide_index=True)
        with st.expander("🗑 月次データを削除する"):
            del_month = st.selectbox("削除する月", [r["月"] for r in st.session_state["kpi_monthly"]])
            if st.button("削除する", key="kpi_del"):
                st.session_state["kpi_monthly"] = [r for r in st.session_state["kpi_monthly"] if r["月"] != del_month]
                st.rerun()
    else:
        st.info("まだ月次データがありません。上のフォームから追加してください。")

    st.divider()
    st.caption("⚠️ データはブラウザを閉じると消えます。定期的にExcelでダウンロードして保存してください。")

    if contact_logs or relation_logs or monthly_data:
        section_title("全データエクスポート")
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            data = build_excel_log("contact_log", "メディア接触ログ")
            if data:
                st.download_button("📞 接触ログをDL", data=data, file_name="media_contact_log.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_dl2:
            data = build_excel_log("relation_log", "新規媒体リレーションログ")
            if data:
                st.download_button("🤝 リレーションログをDL", data=data, file_name="relation_log.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_dl3:
            if monthly_data:
                buf = io.BytesIO()
                pd.DataFrame(monthly_data).to_excel(buf, index=False)
                buf.seek(0)
                st.download_button("📈 月次KPIをDL", data=buf.read(), file_name="kpi_monthly.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
