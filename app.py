"""
SOV分析ツール - Streamlit Webアプリ版（SOV分析 + キーワード露出分析）
"""

import time
import random
import io
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

MAX_WORKERS = 5       # 並列数を減らしてサーバー負荷を抑える
TIMEOUT = 20          # タイムアウトを延長
MAX_RETRIES = 3       # リトライ回数
DELAY_MIN = 1.0       # 待機時間を長めに
DELAY_MAX = 3.0

PAYWALL_KEYWORDS = [
    "有料記事", "会員限定", "ログインが必要", "続きを読む",
    "購読が必要", "プレミアム会員", "有料会員",
]

# 複数のUser-Agentをローテーション（ボット判定を回避）
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4.1 Safari/605.1.15",
]


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
                # レートリミット：少し待ってリトライ
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
            time.sleep(3 * (attempt + 1))
        except requests.exceptions.ConnectionError:
            if attempt == MAX_RETRIES - 1:
                return "error", ""
            time.sleep(3 * (attempt + 1))
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


def build_excel_sov(results: list, companies: dict) -> bytes:
    kw_cols = list(companies.keys())
    df = pd.DataFrame(results)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "記事詳細"

    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    kw_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")

    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + kw_cols
    ws1.append(out_cols)

    for col_idx, col_name in enumerate(out_cols, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = kw_fill if col_name in kw_cols else header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_colors = {"success": "E2EFDA", "paywall": "FFF2CC", "error": "FCE4D6", "timeout": "FCE4D6"}

    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in kw_cols:
                row.append("○" if val is True else ("-" if val is None else ""))
            else:
                row.append(val if pd.notna(val) else "")
        ws1.append(row)
        current_row = ws1.max_row
        status = row_data.get("取得ステータス", "")
        fill = PatternFill(
            start_color=status_colors.get(status, "FFFFFF"),
            end_color=status_colors.get(status, "FFFFFF"),
            fill_type="solid"
        )
        for col_idx in range(1, len(out_cols) + 1):
            cell = ws1.cell(row=current_row, column=col_idx)
            cell.fill = fill
            cell.alignment = Alignment(
                horizontal="center" if out_cols[col_idx-1] in kw_cols else "left",
                vertical="center"
            )

    col_widths = {"No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20,
                  "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14}
    for col_idx, col_name in enumerate(out_cols, 1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 14)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    # SOV集計シート
    ws2 = wb.create_sheet("SOV集計")
    success_df = df[df["取得ステータス"] == "success"]
    total_success = len(success_df)

    ws2.append(["SOV集計レポート"])
    ws2["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws2.append([])
    ws2.append(["集計対象記事数（取得成功）", total_success])
    ws2.append(["有料記事・取得不可（除外）", len(df) - total_success])
    ws2.append(["総記事数", len(df)])
    ws2.append([])
    ws2.append(["企業名", "記載あり記事数", "競合間SOV（%）", "総記事SOV（%）"])

    for col_idx in range(1, 5):
        cell = ws2.cell(row=7, column=col_idx)
        cell.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws2.append(["", "", "※企業合計を母数", "※取得成功全記事を母数"])
    for col_idx in [3, 4]:
        cell = ws2.cell(row=8, column=col_idx)
        cell.font = Font(size=8, color="888888", italic=True)
        cell.alignment = Alignment(horizontal="center")

    total_mentions = 0
    company_counts = {}
    for company in kw_cols:
        if company in success_df.columns:
            count = int((success_df[company] == True).sum())
            company_counts[company] = count
            total_mentions += count

    for company, count in company_counts.items():
        comp_sov = round(count / total_mentions * 100, 1) if total_mentions > 0 else 0
        total_sov = round(count / total_success * 100, 1) if total_success > 0 else 0
        ws2.append([company, count, comp_sov, total_sov])

    ws2.append(["合計", total_mentions, 100.0,
                round(total_mentions / total_success * 100, 1) if total_success > 0 else 0])
    total_row = ws2.max_row
    for col_idx in range(1, 5):
        cell = ws2.cell(row=total_row, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    for col, width in [("A", 20), ("B", 18), ("C", 18), ("D", 18)]:
        ws2.column_dimensions[col].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel_keyword(results: list, keywords: dict) -> bytes:
    kw_cols = list(keywords.keys())
    df = pd.DataFrame(results)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "記事詳細"

    header_fill = PatternFill(start_color="1F5C2E", end_color="1F5C2E", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    kw_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")

    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + kw_cols
    ws1.append(out_cols)

    for col_idx, col_name in enumerate(out_cols, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = kw_fill if col_name in kw_cols else header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_colors = {"success": "E2EFDA", "paywall": "FFF2CC", "error": "FCE4D6", "timeout": "FCE4D6"}

    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in kw_cols:
                row.append("○" if val is True else ("-" if val is None else ""))
            else:
                row.append(val if pd.notna(val) else "")
        ws1.append(row)
        current_row = ws1.max_row
        status = row_data.get("取得ステータス", "")
        fill = PatternFill(
            start_color=status_colors.get(status, "FFFFFF"),
            end_color=status_colors.get(status, "FFFFFF"),
            fill_type="solid"
        )
        for col_idx in range(1, len(out_cols) + 1):
            cell = ws1.cell(row=current_row, column=col_idx)
            cell.fill = fill
            cell.alignment = Alignment(
                horizontal="center" if out_cols[col_idx-1] in kw_cols else "left",
                vertical="center"
            )

    col_widths = {"No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20,
                  "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14}
    for col_idx, col_name in enumerate(out_cols, 1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 14)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    # キーワード集計シート
    ws2 = wb.create_sheet("キーワード集計")
    success_df = df[df["取得ステータス"] == "success"]
    total_success = len(success_df)

    ws2.append(["キーワード露出集計レポート"])
    ws2["A1"].font = Font(bold=True, size=14, color="1F5C2E")
    ws2.append([])
    ws2.append(["集計対象記事数（取得成功）", total_success])
    ws2.append(["有料記事・取得不可（除外）", len(df) - total_success])
    ws2.append(["総記事数", len(df)])
    ws2.append([])
    ws2.append(["キーワード", "露出記事数", "露出率（%）", "設定キーワード・類似表現"])

    for col_idx in range(1, 5):
        cell = ws2.cell(row=7, column=col_idx)
        cell.fill = PatternFill(start_color="1F5C2E", end_color="1F5C2E", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    for kw_label, aliases in keywords.items():
        if kw_label in success_df.columns:
            count = int((success_df[kw_label] == True).sum())
            pct = round(count / total_success * 100, 1) if total_success > 0 else 0
            ws2.append([kw_label, count, pct, aliases])

    for col, width in [("A", 20), ("B", 14), ("C", 12), ("D", 50)]:
        ws2.column_dimensions[col].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def render_settings(state_key, defaults, label_placeholder, alias_placeholder):
    """企業/キーワード設定パネルの共通UI"""
    if state_key not in st.session_state:
        st.session_state[state_key] = [
            {"name": k, "aliases": v} for k, v in defaults.items()
        ]

    to_delete = []
    for i, item in enumerate(st.session_state[state_key]):
        col1, col2, col3 = st.columns([2, 5, 0.5])
        with col1:
            st.session_state[state_key][i]["name"] = st.text_input(
                "名前", value=item["name"], key=f"{state_key}_name_{i}",
                label_visibility="collapsed", placeholder=label_placeholder
            )
        with col2:
            st.session_state[state_key][i]["aliases"] = st.text_input(
                "類似表現", value=item["aliases"], key=f"{state_key}_alias_{i}",
                label_visibility="collapsed", placeholder=alias_placeholder
            )
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
            st.session_state[state_key] = [
                {"name": k, "aliases": v} for k, v in defaults.items()
            ]
            st.rerun()

    return {
        item["name"]: item["aliases"]
        for item in st.session_state[state_key]
        if item["name"].strip()
    }


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


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="PR分析ツール", page_icon="📊", layout="wide")
st.title("📊 PR分析ツール")
st.caption("PR TIMESクリッピングデータを使ったSOV分析・キーワード露出分析ツール")
st.divider()

tab1, tab2 = st.tabs(["🏢 SOV分析", "🔍 キーワード露出分析"])


# ============================================================
# タブ1：SOV分析
# ============================================================
with tab1:
    st.subheader("🏢 SOV分析")
    st.caption("競合企業の記事露出量を比較し、Share of Voiceを集計します")

    with st.expander("⚙️ チェック対象企業の設定（クリックで開く）", expanded=False):
        st.caption("企業名と表記ゆれをカンマ区切りで入力してください")
        companies = render_settings(
            "companies", DEFAULT_COMPANIES,
            "企業名", "表記ゆれをカンマ区切りで入力"
        )

    st.subheader("チェック対象企業")
    if companies:
        cols = st.columns(4)
        for i, company in enumerate(companies.keys()):
            cols[i % 4].markdown(f"- {company}")
    else:
        st.warning("企業が設定されていません")

    st.divider()

    uploaded_file_sov = st.file_uploader(
        "クリッピングデータのExcelをアップロード",
        type=["xlsx", "xls"], key="sov_upload"
    )

    if uploaded_file_sov and companies:
        try:
            df_sov = load_input(uploaded_file_sov)
            st.success(f"✅ {len(df_sov)} 件のURLを検出しました")

            if st.button("▶ SOV分析を開始", type="primary", use_container_width=True, key="sov_start"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()
                results = run_analysis(df_sov, companies, progress_bar, status_text, log_area)

                st.divider()
                st.subheader("📈 SOV集計結果")
                success_results = [r for r in results if r.get("取得ステータス") == "success"]
                total_success = len(success_results)
                total_mentions = 0
                sov_data = []

                for company in companies:
                    count = sum(1 for r in success_results if r.get(company) is True)
                    total_mentions += count
                    sov_data.append({"企業名": company, "記載記事数": count})

                for row in sov_data:
                    count = row["記載記事数"]
                    row["競合間SOV（%）"] = round(count / total_mentions * 100, 1) if total_mentions > 0 else 0
                    row["総記事SOV（%）"] = round(count / total_success * 100, 1) if total_success > 0 else 0

                sov_df = pd.DataFrame(sov_data).sort_values("競合間SOV（%）", ascending=False)
                st.dataframe(sov_df, use_container_width=True, hide_index=True)

                st.subheader("📋 処理ステータス")
                render_status_metrics(results)

                st.divider()
                excel_data = build_excel_sov(results, companies)
                st.download_button(
                    label="📥 結果をExcelでダウンロード",
                    data=excel_data,
                    file_name="sov_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")


# ============================================================
# タブ2：キーワード露出分析
# ============================================================
with tab2:
    st.subheader("🔍 キーワード露出分析")
    st.caption("指定したキーワードが記事に含まれているか判定し、露出件数・露出率を集計します")

    with st.expander("⚙️ チェック対象キーワードの設定（クリックで開く）", expanded=True):
        st.caption("キーワードのラベルと、類似表現をカンマ区切りで入力してください")
        keywords = render_settings(
            "keywords", DEFAULT_KEYWORDS,
            "キーワード名（例：冷凍倉庫）", "類似表現をカンマ区切りで入力（例：冷凍・冷蔵倉庫, フリーザー倉庫）"
        )

    st.subheader("チェック対象キーワード")
    if keywords:
        for label, aliases in keywords.items():
            st.markdown(f"- **{label}**：{aliases}")
    else:
        st.warning("キーワードが設定されていません")

    st.divider()

    uploaded_file_kw = st.file_uploader(
        "クリッピングデータのExcelをアップロード",
        type=["xlsx", "xls"], key="kw_upload"
    )

    if uploaded_file_kw and keywords:
        try:
            df_kw = load_input(uploaded_file_kw)
            st.success(f"✅ {len(df_kw)} 件のURLを検出しました")

            if st.button("▶ キーワード分析を開始", type="primary", use_container_width=True, key="kw_start"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()
                results = run_analysis(df_kw, keywords, progress_bar, status_text, log_area)

                st.divider()
                st.subheader("📈 キーワード露出集計")
                success_results = [r for r in results if r.get("取得ステータス") == "success"]
                total_success = len(success_results)
                kw_data = []

                for label in keywords:
                    count = sum(1 for r in success_results if r.get(label) is True)
                    pct = round(count / total_success * 100, 1) if total_success > 0 else 0
                    kw_data.append({
                        "キーワード": label,
                        "露出記事数": count,
                        "露出率（%）": pct,
                        "設定キーワード": keywords[label]
                    })

                kw_df = pd.DataFrame(kw_data).sort_values("露出記事数", ascending=False)
                st.dataframe(kw_df, use_container_width=True, hide_index=True)

                st.subheader("📋 処理ステータス")
                render_status_metrics(results)

                st.divider()
                excel_data = build_excel_keyword(results, keywords)
                st.download_button(
                    label="📥 結果をExcelでダウンロード",
                    data=excel_data,
                    file_name="keyword_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
