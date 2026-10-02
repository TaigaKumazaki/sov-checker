"""
SOV（Share of Voice）チェッカー
================================
PR TIMESクリッピングのExcel/CSVから記事URLを読み込み、
指定した企業名が本文に含まれるかをチェックしてSOV集計します。

使い方:
  python sov_checker.py input.xlsx
  python sov_checker.py input.csv

出力:
  sov_result.xlsx（記事詳細シート + SOV集計シート）
"""

import sys
import time
import random
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# ★ 設定エリア（必要に応じて変更してください）
# ============================================================

# チェック対象企業と表記ゆれ（部分一致）
COMPANIES = {
    "霞ヶ関キャピタル": [
        "霞ヶ関キャピタル", "霞が関キャピタル", "霞ヶ関capital",
        "Kasumigaseki Capital", "kasumigaseki capital", "KC"
    ],
    "X NETWORK": [
        "X NETWORK", "X Network", "x network",
        "クロスネットワーク", "CROSS NETWORK", "Cross Network",
    ],
    "東急不動産": [
        "東急不動産", "Tokyu Land", "TOKYU LAND", "東急リバブル"
    ],
    "日本GLP": [
        "日本GLP", "GLP", "Global Logistic Properties", "グローバル・ロジスティック・プロパティーズ"
    ],
    "三井不動産": [
        "三井不動産", "Mitsui Fudosan", "MITSUI FUDOSAN", "三井不動産ロジスティクスパーク"
    ],
    "プロロジス": [
        "プロロジス", "Prologis", "PROLOGIS", "プロロジスパーク"
    ],
    "大和ハウス": [
        "大和ハウス", "大和ハウス工業", "Daiwa House", "DAIWA HOUSE", "DPL"
    ],
    "野村不動産": [
        "野村不動産", "Nomura Real Estate", "NOMURA REAL ESTATE", "野村不動産ロジスティクス"
    ],
    "三菱地所": [
        "三菱地所", "Mitsubishi Estate", "MITSUBISHI ESTATE", "三菱地所ロジスティクス"
    ],
    "森トラスト": [
        "森トラスト", "Mori Trust", "MORI TRUST"
    ],
    "東京建物": [
        "東京建物", "Tokyo Tatemono", "TOKYO TATEMONO"
    ],
}

# 並列処理数（多すぎるとサーバーに負荷。10〜20推奨）
MAX_WORKERS = 15

# リクエスト間隔（秒）- ランダムに0.5〜2秒待機
DELAY_MIN = 0.5
DELAY_MAX = 2.0

# タイムアウト（秒）
TIMEOUT = 10

# リトライ回数
MAX_RETRIES = 2

# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

PAYWALL_KEYWORDS = [
    "有料記事", "会員限定", "ログインが必要", "続きを読む",
    "購読が必要", "プレミアム会員", "有料会員", "サブスクリプション",
    "paywall", "subscribe to read", "premium content"
]


def fetch_text(url: str) -> tuple[str, str]:
    """
    URLから本文テキストを取得する。
    Returns: (status, text)
      status: "success" | "paywall" | "error" | "timeout"
    """
    for attempt in range(MAX_RETRIES + 1):
        try:
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
            resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)

            # 403/401 → アクセス拒否
            if resp.status_code in (401, 403):
                return "error", f"HTTP {resp.status_code}"

            # 有料記事チェック（ステータスコードが200でも内容で判断）
            html = resp.text
            soup = BeautifulSoup(html, "html.parser")

            # メタタグ・本文からペイウォール検出
            page_text = soup.get_text(separator=" ", strip=True)
            if any(kw in page_text for kw in PAYWALL_KEYWORDS):
                # テキスト量が少ない（300文字未満）場合はペイウォール確定
                if len(page_text) < 300:
                    return "paywall", page_text

            # 本文抽出（article タグ優先 → main → body）
            for tag in ["article", "main", "body"]:
                container = soup.find(tag)
                if container:
                    text = container.get_text(separator=" ", strip=True)
                    if len(text) > 100:
                        return "success", text

            return "success", page_text

        except requests.exceptions.Timeout:
            if attempt == MAX_RETRIES:
                return "timeout", ""
        except requests.exceptions.ConnectionError:
            if attempt == MAX_RETRIES:
                return "error", "接続エラー"
        except Exception as e:
            if attempt == MAX_RETRIES:
                return "error", str(e)[:100]

    return "error", "不明なエラー"


def check_companies(text: str) -> dict[str, bool]:
    """テキスト内の各企業名の有無をチェック"""
    results = {}
    text_lower = text.lower()
    for company, aliases in COMPANIES.items():
        found = any(alias.lower() in text_lower for alias in aliases)
        results[company] = found
    return results


def process_url(row_data: dict) -> dict:
    """1行分の処理（fetch + 企業名チェック）"""
    url = row_data["記事URL"]
    status, text = fetch_text(url)

    company_results = {}
    if status == "success":
        company_results = check_companies(text)
    else:
        company_results = {c: None for c in COMPANIES}  # None = 判定不可

    return {**row_data, "取得ステータス": status, **company_results}


def load_input(path: str) -> pd.DataFrame:
    """Excel or CSV を読み込む"""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        # ヘッダー行を自動検出（"No"列を探す）
        df_raw = pd.read_excel(path, header=None)
        header_row = None
        for i, row in df_raw.iterrows():
            if "No" in row.values or "記事URL" in row.values:
                header_row = i
                break
        if header_row is None:
            header_row = 1  # デフォルト2行目
        df = pd.read_excel(path, header=header_row)
    else:
        df = pd.read_csv(path, encoding="utf-8-sig")

    # URL列を特定
    url_col = None
    for col in df.columns:
        if "URL" in str(col).upper() or "url" in str(col).lower():
            url_col = col
            break
    if url_col is None:
        raise ValueError("URL列が見つかりません。列名を確認してください。")

    df = df.rename(columns={url_col: "記事URL"})
    df = df.dropna(subset=["記事URL"])
    df["記事URL"] = df["記事URL"].astype(str).str.strip()
    df = df[df["記事URL"].str.startswith("http")]
    return df


def save_results(results: list[dict], output_path: str):
    """結果をExcelに保存（詳細シート + SOV集計シート）"""
    df = pd.DataFrame(results)
    company_cols = list(COMPANIES.keys())

    wb = Workbook()

    # ---- シート1: 記事詳細 ----
    ws1 = wb.active
    ws1.title = "記事詳細"

    # ヘッダー色
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    company_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")

    # 出力列の定義
    base_cols = ["No", "取得日", "サイト名", "媒体社名", "記事タイトル", "記事URL", "取得ステータス"]
    out_cols = [c for c in base_cols if c in df.columns] + company_cols

    headers = out_cols
    ws1.append(headers)

    # ヘッダースタイル
    for col_idx, col_name in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col_idx)
        if col_name in company_cols:
            cell.fill = company_fill
        else:
            cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # データ行
    status_colors = {
        "success": "E2EFDA",   # 緑
        "paywall": "FFF2CC",   # 黄
        "error":   "FCE4D6",   # 橙
        "timeout": "FCE4D6",   # 橙
    }

    for row_data in results:
        row = []
        for col in out_cols:
            val = row_data.get(col, "")
            if col in company_cols:
                if val is True:
                    row.append("○")
                elif val is False:
                    row.append("")
                else:
                    row.append("-")  # 取得不可
            else:
                row.append(val if pd.notna(val) else "")
        ws1.append(row)

        # 行の色付け（取得ステータスに応じて）
        current_row = ws1.max_row
        status = row_data.get("取得ステータス", "")
        fill_color = status_colors.get(status, "FFFFFF")
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        for col_idx in range(1, len(out_cols) + 1):
            cell = ws1.cell(row=current_row, column=col_idx)
            cell.fill = fill
            cell.alignment = Alignment(vertical="center")
            if out_cols[col_idx - 1] in company_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # 列幅調整
    col_widths = {
        "No": 6, "取得日": 12, "サイト名": 20, "媒体社名": 20,
        "記事タイトル": 40, "記事URL": 50, "取得ステータス": 14,
    }
    for col_idx, col_name in enumerate(out_cols, 1):
        width = col_widths.get(col_name, 12)
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    # ---- シート2: SOV集計 ----
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

    sov_header = ["企業名", "記載あり記事数", "競合間SOV（%）", "総記事SOV（%）"]
    ws2.append(sov_header)
    for col_idx in range(1, 5):
        cell = ws2.cell(row=7, column=col_idx)
        cell.fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # 補足説明行
    ws2.cell(row=7, column=3).comment = None
    note_row = 8
    ws2.insert_rows(note_row)
    ws2.cell(row=note_row, column=3).value = "※5社合計を母数"
    ws2.cell(row=note_row, column=4).value = "※取得成功全記事を母数"
    for col_idx in [3, 4]:
        cell = ws2.cell(row=note_row, column=col_idx)
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
        competitive_sov = (count / total_mentions * 100) if total_mentions > 0 else 0
        total_sov = (count / total_success * 100) if total_success > 0 else 0
        ws2.append([company, count, round(competitive_sov, 1), round(total_sov, 1)])

    # 合計行
    ws2.append(["合計", total_mentions, 100.0, round(total_mentions / total_success * 100, 1) if total_success > 0 else 0])
    total_row = ws2.max_row
    for col_idx in range(1, 5):
        cell = ws2.cell(row=total_row, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 18
    ws2.column_dimensions["D"].width = 18

    wb.save(output_path)
    print(f"\n✅ 保存完了: {output_path}")


def main():
    if len(sys.argv) < 2:
        print("使い方: python sov_checker.py input.xlsx")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = Path(input_path).stem + "_sov_result.xlsx"

    print(f"📂 読み込み中: {input_path}")
    df = load_input(input_path)
    total = len(df)
    print(f"✅ {total} 件のURLを検出しました")
    print(f"🔍 チェック企業: {', '.join(COMPANIES.keys())}")
    print(f"⚙️  並列処理数: {MAX_WORKERS} / タイムアウト: {TIMEOUT}秒\n")

    rows = df.to_dict("records")
    results = []
    done = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_url, row): row for row in rows}
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            done += 1
            status = result.get("取得ステータス", "?")
            url = result.get("記事URL", "")[:60]
            print(f"[{done}/{total}] {status:10s} {url}")

    print(f"\n📊 集計中...")
    save_results(results, output_path)

    # サマリー表示
    status_counts = {}
    for r in results:
        s = r.get("取得ステータス", "unknown")
        status_counts[s] = status_counts.get(s, 0) + 1

    print("\n── ステータス集計 ──────────────")
    for s, c in status_counts.items():
        print(f"  {s:12s}: {c} 件")

    print("\n── SOV集計 ────────────────────────")
    success_results = [r for r in results if r.get("取得ステータス") == "success"]
    total_success_count = len(success_results)
    total_mentions = 0
    company_counts = {}
    for company in COMPANIES:
        count = sum(1 for r in success_results if r.get(company) is True)
        company_counts[company] = count
        total_mentions += count

    print(f"  {'企業名':<20} {'記載記事数':>8}  {'競合間SOV':>10}  {'総記事SOV':>10}")
    print(f"  {'-'*55}")
    for company, count in company_counts.items():
        competitive_pct = count / total_mentions * 100 if total_mentions > 0 else 0
        total_pct = count / total_success_count * 100 if total_success_count > 0 else 0
        print(f"  {company:<20} {count:>8} 件  {competitive_pct:>9.1f}%  {total_pct:>9.1f}%")


if __name__ == "__main__":
    main()
