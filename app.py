"""
SOV分析ツール - Streamlit Webアプリ版
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
# 企業設定
# ============================================================
COMPANIES = {
    "霞ヶ関キャピタル": [
        "霞ヶ関キャピタル", "霞が関キャピタル", "霞ヶ関capital",
        "Kasumigaseki Capital", "kasumigaseki capital",
    ],
    "X NETWORK": [
        "X NETWORK", "X Network", "x network",
        "クロスネットワーク", "CROSS NETWORK", "Cross Network",
    ],
    "東急不動産": [
        "東急不動産", "Tokyu Land", "TOKYU LAND", "東急リバブル"
    ],
    "日本GLP": [
        "日本GLP", "GLP", "Global Logistic Properties",
        "グローバル・ロジスティック・プロパティーズ"
    ],
    "三井不動産": [
        "三井不動産", "Mitsui Fudosan", "MITSUI FUDOSAN",
        "三井不動産ロジスティクスパーク"
    ],
    "プロロジス": [
        "プロロジス", "Prologis", "PROLOGIS", "プロロジスパーク"
    ],
    "大和ハウス": [
        "大和ハウス", "大和ハウス工業", "Daiwa House", "DAIWA HOUSE", "DPL"
    ],
    "野村不動産": [
        "野村不動産", "Nomura Real Estate", "NOMURA REAL ESTATE",
        "野村不動産ロジスティクス"
    ],
    "三菱地所": [
        "三菱地所", "Mitsubishi Estate", "MITSUBISHI ESTATE",
        "三菱地所ロジスティクス"
    ],
    "森トラスト": [
        "森トラスト", "Mori Trust", "MORI TRUST"
    ],
    "東京建物": [
        "東京建物", "Tokyo Tatemono", "TOKYO TATEMONO", "T-LOGI"
    ],
}

MAX_WORKERS = 10
TIMEOUT = 10
DELAY_MIN = 0.3
DELAY_MAX = 1.5

PAYWALL_KEYWORDS = [
    "有料記事", "会員限定", "ログインが必要", "続きを読む",
    "購読が必要", "プレミアム会員", "有料会員",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


def fetch_text(url: str) -> tuple:
    try:
        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if resp.status_code in (401, 403):
            return "error", ""
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
        return "timeout", ""
    except Exception:
        return "error", ""


def check_companies(text: str) -> dict:
    text_lower = text.lower()
    return {
        company: any(alias.lower() in text_lower for alias in aliases)
        for company, aliases in COMPANIES.items()
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


def build_excel(results: list) -> bytes:
    company_cols = list(COMPANIES.keys())
    df = pd.DataFrame(results)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "記事詳細"

    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    company_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")

    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + company_cols
    ws1.append(out_cols)

    for col_idx, col_name in enumerate(out_cols, 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = company_fill if col_name in company_cols else header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_colors = {"success": "E2EFDA", "paywall": "FFF2CC", "error": "FCE4D6", "timeout": "FCE4D6"}

    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in company_cols:
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
                horizontal="center" if out_cols[col_idx-1] in company_cols else "left",
                vertical="center"
            )

    col_widths = {"No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20,
                  "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14}
    for col_idx, col_name in enumerate(out_cols, 1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 12)
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

    ws2.append(["", "", "※5社合計を母数", "※取得成功全記事を母数"])
    for col_idx in [3, 4]:
        cell = ws2.cell(row=8, column=col_idx)
        cell.font = Font(size=8, color="888888", italic=True)
        cell.alignment = Alignment(horizontal="center")

    total_mentions = 0
    company_counts = {}
    for company in company_cols:
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


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="SOV分析ツール", page_icon="📊", layout="wide")

st.title("📊 SOV分析ツール")
st.caption("PR TIMESクリッピングデータから企業名の露出量（Share of Voice）を集計します")

st.divider()

# ファイルアップロード
uploaded_file = st.file_uploader(
    "クリッピングデータのExcelファイルをアップロードしてください",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        df = load_input(uploaded_file)
        total = len(df)
        st.success(f"✅ {total} 件のURLを検出しました")

        st.subheader("チェック対象企業")
        cols = st.columns(4)
        for i, company in enumerate(COMPANIES.keys()):
            cols[i % 4].markdown(f"- {company}")

        st.divider()

        if st.button("▶ 分析開始", type="primary", use_container_width=True):
            results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            log_area = st.empty()
            log_lines = []

            rows = df.to_dict("records")
            done = 0

            def process_url(row_data):
                url = row_data.get("記事URL", "")
                status, text = fetch_text(url)
                company_results = check_companies(text) if status == "success" else {c: None for c in COMPANIES}
                return {**row_data, "取得ステータス": status, **company_results}

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

            # SOV集計表示
            st.divider()
            st.subheader("📈 SOV集計結果")

            success_results = [r for r in results if r.get("取得ステータス") == "success"]
            total_success = len(success_results)
            total_mentions = 0
            sov_data = []

            for company in COMPANIES:
                count = sum(1 for r in success_results if r.get(company) is True)
                total_mentions += count
                sov_data.append({"企業名": company, "記載記事数": count})

            for row in sov_data:
                count = row["記載記事数"]
                row["競合間SOV（%）"] = round(count / total_mentions * 100, 1) if total_mentions > 0 else 0
                row["総記事SOV（%）"] = round(count / total_success * 100, 1) if total_success > 0 else 0

            sov_df = pd.DataFrame(sov_data).sort_values("競合間SOV（%）", ascending=False)
            st.dataframe(sov_df, use_container_width=True, hide_index=True)

            # ステータス集計
            st.subheader("📋 処理ステータス")
            status_counts = {}
            for r in results:
                s = r.get("取得ステータス", "unknown")
                status_counts[s] = status_counts.get(s, 0) + 1

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("取得成功", f"{status_counts.get('success', 0)} 件")
            c2.metric("有料記事", f"{status_counts.get('paywall', 0)} 件")
            c3.metric("エラー", f"{status_counts.get('error', 0)} 件")
            c4.metric("タイムアウト", f"{status_counts.get('timeout', 0)} 件")

            # Excelダウンロード
            st.divider()
            excel_data = build_excel(results)
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
