import io
import zipfile
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Rectangle, FancyBboxPatch
from matplotlib.ticker import FuncFormatter


# =========================================================
# Streamlit基本設定
# =========================================================

st.set_page_config(page_title="AFF定例資料作成", layout="wide")
st.title("🎈AFF定例資料作成")
st.caption("デイリーレポート・コストレポートから、定例資料用の画像を作成します。")
st.caption("それぞれ、必要なシートのみ新しいファイルを作成し数式は解除した状態のデータを用意してください。")

# =========================================================
# matplotlib 日本語フォント設定
# =========================================================

from pathlib import Path
from matplotlib import font_manager

FONT_PATH = Path(__file__).parent / "fonts" / "NotoSansJP-Regular.ttf"

font_manager.fontManager.addfont(str(FONT_PATH))
FONT_NAME = font_manager.FontProperties(fname=str(FONT_PATH)).get_name()

plt.rcParams["font.family"] = FONT_NAME
plt.rcParams["axes.unicode_minus"] = False

# =========================================================
# フォントサイズ
# =========================================================

FONT_HEADER = 34
FONT_BOX_TITLE = 16

FONT_GAUGE_TITLE = 22
FONT_GAUGE_LABEL = 16
FONT_GAUGE_PERCENT = 24

FONT_KPI_MAIN = 19
FONT_KPI_SUBTITLE = 17
FONT_KPI_DETAIL = 14

FONT_RIGHT_MAIN = 17
FONT_RIGHT_SUB = 12

FONT_AXIS = 10
FONT_LEGEND = 11
FONT_FOOTER = 14

# =========================================================
# 共通関数
# =========================================================

def safe_div(a, b):
    if b is None or b == 0:
        return 0
    return a / b


def fmt_num(v):
    return f"{v:,.0f}"


def fmt_yen(v):
    return f"{v:,.0f} 円"


def to_md_label(series):
    return (
        series.dt.strftime("%m/%d")
        .str.replace(r"^0", "", regex=True)
        .str.replace(r"/0", "/", regex=True)
    )


def unmerge_and_fill(ws):
    """
    結合セルを解除し、結合範囲すべてに左上セルの値を入れる。
    """
    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        value = ws.cell(min_row, min_col).value

        ws.unmerge_cells(str(merged_range))

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row, col).value = value


def render_image_block(title, image_data, file_name, generated_images):
    """
    画像プレビュー、個別保存、ZIP用リスト追加。
    """
    st.subheader(title)
    st.image(image_data, use_container_width=True)

    st.download_button(
        f"{title}を保存",
        data=image_data,
        file_name=file_name,
        mime="image/png",
    )

    generated_images.append({
        "filename": file_name,
        "data": image_data,
    })


def render_zip_download(generated_images):
    """
    生成済み画像をZIPで一括ダウンロード。
    """
    if not generated_images:
        return

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for image in generated_images:
            zip_file.writestr(image["filename"], image["data"])

    zip_buffer.seek(0)

    st.download_button(
        "全画像をまとめてダウンロード",
        data=zip_buffer.getvalue(),
        file_name="AFF定例資料画像.zip",
        mime="application/zip",
    )


# =========================================================
# デイリーレポート
# =========================================================

def read_daily_report(uploaded_file):
    """
    デイリーレポートを読み込む。

    D列: サイト名
    T列: ディテール
    U列以降: 日別数値
    3行目: 日付
    12〜15行目: ポイントサイト合計KPI用
    15〜19行目: 比較サイト合計KPI用
    32行目以降: 明細用
    """
    file_bytes = uploaded_file.read()
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    unmerge_and_fill(ws)

    raw_df = pd.DataFrame(list(ws.values))

    site_col = 3         # D列
    detail_col = 19      # T列
    date_start_col = 20  # U列

    header_row = 2       # 3行目
    data_start_row = 31  # 32行目

    date_cols = []
    for col in range(date_start_col, raw_df.shape[1]):
        v = raw_df.iloc[header_row, col]
        if isinstance(v, (datetime, date)):
            date_cols.append(col)

    if not date_cols:
        raise ValueError("デイリーレポート：U列以降から日付列を見つけられませんでした。")

    date_names = [
        pd.to_datetime(raw_df.iloc[header_row, c]).date()
        for c in date_cols
    ]

    # KPI用：ポイントサイト合計 12〜15行目
    point_total_df = raw_df.iloc[11:15, [site_col, detail_col] + date_cols].copy()
    point_total_df.columns = ["サイト名", "ディテール"] + date_names
    point_total_df["サイト名"] = "【ポイントサイト】合計"

    # KPI用：比較サイト合計 15〜19行目
    compare_total_df = raw_df.iloc[14:19, [site_col, detail_col] + date_cols].copy()
    compare_total_df.columns = ["サイト名", "ディテール"] + date_names
    compare_total_df["サイト名"] = "【比較サイト】合計"

    total_df = pd.concat([point_total_df, compare_total_df], ignore_index=True)
    total_df["サイト名"] = total_df["サイト名"].astype(str).str.strip()
    total_df["ディテール"] = total_df["ディテール"].astype(str).str.strip()

    total_long = total_df.melt(
        id_vars=["サイト名", "ディテール"],
        var_name="日付",
        value_name="値",
    )
    total_long["日付"] = pd.to_datetime(total_long["日付"])
    total_long["値"] = pd.to_numeric(total_long["値"], errors="coerce").fillna(0)

    # 明細用：32行目以降
    df = raw_df.iloc[data_start_row:, [site_col, detail_col] + date_cols].copy()
    df.columns = ["サイト名", "ディテール"] + date_names

    df = df.dropna(subset=["サイト名", "ディテール"], how="all")
    df["サイト名"] = df["サイト名"].ffill()
    df["サイト名"] = df["サイト名"].astype(str).str.strip()
    df["ディテール"] = df["ディテール"].astype(str).str.strip()

    df = df[
        (df["サイト名"] != "")
        & (df["サイト名"].str.lower() != "nan")
        & (df["ディテール"] != "")
        & (df["ディテール"].str.lower() != "nan")
    ]

    long_df = df.melt(
        id_vars=["サイト名", "ディテール"],
        var_name="日付",
        value_name="値",
    )
    long_df["日付"] = pd.to_datetime(long_df["日付"])
    long_df["値"] = pd.to_numeric(long_df["値"], errors="coerce").fillna(0)

    return long_df, total_long


# =========================================================
# コストレポート
# =========================================================

def read_cost_report(uploaded_file):
    """
    コストレポートを読み込む。

    A列: 日付
    C列: 申込_Forecast
    D列: 申込_実績
    K列: 発行_Forecast
    L列: 発行_実績
    Q列: 目標承認率
    R列: 承認率
    T列: Forecastコスト
    U列: 発行コスト
    Y列: 目標CPA
    Z列: 発行CPA
    """
    file_bytes = uploaded_file.read()
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    raw_df = pd.DataFrame(list(ws.values))

    data_start_row = 2  # 3行目からデータ

    use_cols = {
        0: "日付",
        2: "申込_Forecast",
        3: "申込_実績",
        10: "発行_Forecast",
        11: "発行_実績",
        16: "目標承認率",
        17: "承認率",
        19: "Forecastコスト",
        20: "発行コスト",
        24: "目標CPA",
        25: "発行CPA",
    }

    df = raw_df.iloc[data_start_row:, list(use_cols.keys())].copy()
    df.columns = list(use_cols.values())

    df = df[df["日付"].notna()]
    df = df[df["日付"].astype(str).str.upper() != "TOTAL"]

    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    df = df[df["日付"].notna()]

    for col in df.columns:
        if col != "日付":
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df



# =========================================================
# AF実績データ
# =========================================================

def normalize_af_code(value):
    """
    Excel由来のAFコードを文字列として正規化する。
    例：123.0 -> 123 / 前後スペース除去
    """
    if value is None or pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    return text


@st.cache_data(show_spinner=False)
def read_af_master(path):
    """
    app.py と同じ階層に置いたAFコードマスタを読み込む。

    A列：AFコード
    1行目：ヘッダー想定
    2行目以降：データ
    """
    df = pd.read_excel(
        path,
        header=None,
        usecols=[0],
        skiprows=1,
        dtype=object,
    )

    af_codes = (
        df.iloc[:, 0]
        .map(normalize_af_code)
        .loc[lambda s: s != ""]
        .drop_duplicates()
        .tolist()
    )

    if not af_codes:
        raise ValueError("AFコードマスタのA列2行目以降にAFコードが見つかりませんでした。")

    return af_codes


def read_af_result_file(uploaded_file, af_codes, value_name):
    """
    AF実績ファイルを縦持ちに変換する。

    A列：日付（20260101形式）
    1行目：AFコード
    B2以降：実績値
    """
    file_bytes = uploaded_file.read()
    raw_df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object)

    if raw_df.shape[0] < 2 or raw_df.shape[1] < 2:
        raise ValueError("AF実績ファイルの行または列が不足しています。")

    headers = raw_df.iloc[0].tolist()
    data_df = raw_df.iloc[1:].copy()

    date_text = (
        data_df.iloc[:, 0]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    data_df["日付"] = pd.to_datetime(
        date_text,
        format="%Y%m%d",
        errors="coerce",
    )

    data_df = data_df[data_df["日付"].notna()].copy()

    master_set = set(af_codes)
    matched_cols = []

    for col_idx in range(1, raw_df.shape[1]):
        af_code = normalize_af_code(headers[col_idx])
        if af_code in master_set:
            matched_cols.append((col_idx, af_code))

    if not matched_cols:
        return pd.DataFrame(columns=["日付", "AFコード", value_name])

    long_list = []

    for col_idx, af_code in matched_cols:
        tmp = data_df[["日付"]].copy()
        tmp["AFコード"] = af_code
        tmp[value_name] = pd.to_numeric(data_df.iloc[:, col_idx], errors="coerce").fillna(0)
        long_list.append(tmp)

    result_df = pd.concat(long_list, ignore_index=True)

    return result_df


def sum_af_result(af_df, start_date, end_date, value_name):
    if af_df is None or af_df.empty:
        return 0

    filtered = af_df[
        (af_df["日付"].dt.date >= start_date)
        & (af_df["日付"].dt.date <= end_date)
    ].copy()

    return filtered[value_name].sum()


def create_af_overall_values(
    af_apply_df,
    af_issue_df,
    cost_filtered,
    total_target_cv,
    start_date,
    end_date,
):
    """
    AF実績画像用の集計値を作成する。

    AF発生：AFマスタ × 申込実績
    AF承認：AFマスタ × 発行実績
    目標：Affiliate ○月度全体実績の「合計の目標」
          = デイリーレポートのポイントサイト合計目標 + 比較サイト合計目標
    コスト：既存の全体実績で使っている Forecastコスト / 発行コスト
    """
    af_apply = sum_af_result(
        af_apply_df,
        start_date,
        end_date,
        "AF発生",
    )

    af_issue = sum_af_result(
        af_issue_df,
        start_date,
        end_date,
        "AF承認",
    )

    # AF発生目標
    # 申込Forecast × 0.9
    apply_forecast = cost_filtered["申込_Forecast"].sum()
    target_cv = apply_forecast * 0.9

    # AF承認目標
    # 発行Forecast（コストレポートK列） × 0.9
    forecast_cv = cost_filtered["発行_Forecast"].sum()
    issue_target = forecast_cv * 0.9

    forecast_cost = cost_filtered["Forecastコスト"].sum()
    actual_cost = cost_filtered["発行コスト"].sum()

    af_apply_ratio = safe_div(af_apply, target_cv)
    af_issue_ratio = safe_div(af_issue, issue_target)

    apply_forecast_cpa = safe_div(forecast_cost, target_cv)
    apply_actual_cpa = safe_div(actual_cost, af_apply)
    apply_cpa_ratio = safe_div(
        apply_forecast_cpa,
        apply_actual_cpa
    )

    # AF承認CPA
    # 想定CPA：Forecastコスト ÷ 発行Forecast（コストレポートK列）
    # 実績CPA：発行コスト ÷ AF承認
    issue_forecast_cpa = safe_div(
        forecast_cost,
        forecast_cv
    )
    issue_actual_cpa = safe_div(
        actual_cost,
        af_issue
    )
    issue_cpa_ratio = safe_div(
        issue_forecast_cpa,
        issue_actual_cpa
    )

    return {
        "apply_forecast": apply_forecast,
        "forecast_cv": forecast_cv,
        "target_cv": target_cv,
        "issue_target": issue_target,
        "af_apply": af_apply,
        "af_issue": af_issue,
        "af_apply_ratio": af_apply_ratio,
        "af_issue_ratio": af_issue_ratio,
        "forecast_cost": forecast_cost,
        "actual_cost": actual_cost,
        "apply_forecast_cpa": apply_forecast_cpa,
        "apply_actual_cpa": apply_actual_cpa,
        "apply_cpa_ratio": apply_cpa_ratio,
        "issue_forecast_cpa": issue_forecast_cpa,
        "issue_actual_cpa": issue_actual_cpa,
        "issue_cpa_ratio": issue_cpa_ratio,
    }


def create_af_chart_df(cost_filtered):
    """
    AF実績下段グラフ用。
    棒：Forecast / 実績CV
    線：forecastコスト / 発行コスト
    """
    chart_df = cost_filtered[[
        "日付",
        "発行_Forecast",
        "発行_実績",
        "Forecastコスト",
        "発行コスト",
    ]].copy()

    chart_df = chart_df.rename(columns={
        "発行_Forecast": "Forecast",
        "発行_実績": "実績",
        "Forecastコスト": "forecastコスト",
        "発行コスト": "発行コスト",
    })

    chart_df = chart_df.sort_values("日付")

    return chart_df

# =========================================================
# KPI集計
# =========================================================

def get_daily_value(df, site_name, detail_name):
    return df[
        (df["サイト名"] == site_name)
        & (df["ディテール"] == detail_name)
    ]["値"].sum()


def create_kpi_values(daily_total_filtered, cost_filtered):
    # CV実績：Actual
    point_actual = get_daily_value(
        daily_total_filtered,
        "【ポイントサイト】合計",
        "Actual",
    )
    compare_actual = get_daily_value(
        daily_total_filtered,
        "【比較サイト】合計",
        "Actual",
    )

    # CV目標：Daily Target (Initiative)
    point_target = get_daily_value(
        daily_total_filtered,
        "【ポイントサイト】合計",
        "Daily Target (Initiative)",
    )
    compare_target = get_daily_value(
        daily_total_filtered,
        "【比較サイト】合計",
        "Daily Target (Initiative)",
    )

    total_cv = point_actual + compare_actual
    total_target = point_target + compare_target

    total_cv_ratio = safe_div(total_cv, total_target)
    point_ratio = safe_div(point_actual, point_target)
    compare_ratio = safe_div(compare_actual, compare_target)

    forecast_cost = cost_filtered["Forecastコスト"].sum()
    actual_cost = cost_filtered["発行コスト"].sum()
    cost_ratio = safe_div(actual_cost, forecast_cost)

    approval_rate = cost_filtered["承認率"].mean()
    target_approval_rate = cost_filtered["目標承認率"].mean()

    apply_cpa = safe_div(
        cost_filtered["発行コスト"].sum(),
        cost_filtered["申込_実績"].sum(),
    )
    target_apply_cpa = safe_div(
        cost_filtered["Forecastコスト"].sum(),
        cost_filtered["申込_Forecast"].sum(),
    )

    issue_cpa = cost_filtered["発行CPA"].mean()
    target_issue_cpa = cost_filtered["目標CPA"].mean()

    return {
        "total_cv": total_cv,
        "total_target": total_target,
        "total_cv_ratio": total_cv_ratio,
        "point_actual": point_actual,
        "point_target": point_target,
        "point_ratio": point_ratio,
        "compare_actual": compare_actual,
        "compare_target": compare_target,
        "compare_ratio": compare_ratio,
        "forecast_cost": forecast_cost,
        "actual_cost": actual_cost,
        "cost_ratio": cost_ratio,
        "approval_rate": approval_rate,
        "target_approval_rate": target_approval_rate,
        "apply_cpa": apply_cpa,
        "target_apply_cpa": target_apply_cpa,
        "issue_cpa": issue_cpa,
        "target_issue_cpa": target_issue_cpa,
    }


def create_pointsite_values(daily_total_filtered):
    """
    還元サイト（ポイントサイト）用KPIを作成。
    対象：デイリーレポートの 【ポイントサイト】合計
    Target：Daily Target (Initiative)
    Actual：Actual
    """
    target = get_daily_value(
        daily_total_filtered,
        "【ポイントサイト】合計",
        "Daily Target (Initiative)",
    )

    actual = get_daily_value(
        daily_total_filtered,
        "【ポイントサイト】合計",
        "Actual",
    )

    ratio = safe_div(actual, target)

    return {
        "target": target,
        "actual": actual,
        "ratio": ratio,
    }


def create_pointsite_daily_chart_df(daily_total_filtered):
    """
    還元サイト（ポイントサイト）の日別Target/Actualデータを作成。
    """
    target_df = daily_total_filtered[
        (daily_total_filtered["サイト名"] == "【ポイントサイト】合計")
        & (daily_total_filtered["ディテール"] == "Daily Target (Initiative)")
    ][["日付", "値"]].copy()

    target_df = target_df.rename(columns={"値": "Target"})

    actual_df = daily_total_filtered[
        (daily_total_filtered["サイト名"] == "【ポイントサイト】合計")
        & (daily_total_filtered["ディテール"] == "Actual")
    ][["日付", "値"]].copy()

    actual_df = actual_df.rename(columns={"値": "Actual"})

    chart_df = pd.merge(
        target_df,
        actual_df,
        on="日付",
        how="outer",
    )

    chart_df["Target"] = pd.to_numeric(chart_df["Target"], errors="coerce").fillna(0)
    chart_df["Actual"] = pd.to_numeric(chart_df["Actual"], errors="coerce").fillna(0)
    chart_df = chart_df.sort_values("日付")

    return chart_df


def create_comparesite_values(daily_total_filtered):
    """
    非還元サイト（比較サイト）用KPIを作成。
    対象：デイリーレポートの 【比較サイト】合計
    Target：Daily Target (Initiative)
    Actual：Actual
    """
    target = get_daily_value(
        daily_total_filtered,
        "【比較サイト】合計",
        "Daily Target (Initiative)",
    )

    actual = get_daily_value(
        daily_total_filtered,
        "【比較サイト】合計",
        "Actual",
    )

    ratio = safe_div(actual, target)

    return {
        "target": target,
        "actual": actual,
        "ratio": ratio,
    }


def create_comparesite_daily_chart_df(daily_total_filtered):
    """
    非還元サイト（比較サイト）の日別Target/Actualデータを作成。
    """
    target_df = daily_total_filtered[
        (daily_total_filtered["サイト名"] == "【比較サイト】合計")
        & (daily_total_filtered["ディテール"] == "Daily Target (Initiative)")
    ][["日付", "値"]].copy()

    target_df = target_df.rename(columns={"値": "Target"})

    actual_df = daily_total_filtered[
        (daily_total_filtered["サイト名"] == "【比較サイト】合計")
        & (daily_total_filtered["ディテール"] == "Actual")
    ][["日付", "値"]].copy()

    actual_df = actual_df.rename(columns={"値": "Actual"})

    chart_df = pd.merge(
        target_df,
        actual_df,
        on="日付",
        how="outer",
    )

    chart_df["Target"] = pd.to_numeric(chart_df["Target"], errors="coerce").fillna(0)
    chart_df["Actual"] = pd.to_numeric(chart_df["Actual"], errors="coerce").fillna(0)
    chart_df = chart_df.sort_values("日付")

    return chart_df


# =========================================================
# 画像作成
# =========================================================

def draw_gauge(ax, ratio, title):
    """
    100% = 上半円180度
    200% = 360度
    """
    ax.set_aspect("equal")
    ax.axis("off")

    base_color = "#f7caca"
    main_color = "#d50000"

    # 100%背景
    ax.add_patch(
        Wedge(
            (0, 0),
            1.0,
            0,
            180,
            width=0.28,
            facecolor=base_color,
            edgecolor="none",
        )
    )

    # 実績
    ratio_clamped = min(max(ratio, 0), 2)

    if ratio_clamped <= 1:
        # 0〜100%：上半円の中だけ塗る
        degree = ratio_clamped * 180

        ax.add_patch(
            Wedge(
                (0, 0),
                1.0,
                180 - degree,
                180,
                width=0.28,
                facecolor=main_color,
                edgecolor="none",
            )
        )
    else:
        # 100%：上半円を全部塗る
        ax.add_patch(
            Wedge(
                (0, 0),
                1.0,
                0,
                180,
                width=0.28,
                facecolor=main_color,
                edgecolor="none",
            )
        )

        # 100〜200%：下半円を追加で塗る
        extra_degree = (ratio_clamped - 1) * 180

        ax.add_patch(
            Wedge(
                (0, 0),
                1.0,
                360 - extra_degree,
                360,
                width=0.28,
                facecolor=main_color,
                edgecolor="none",
            )
        )


    # 180度線
    ax.plot([-1.25, 1.25], [0, 0], color="#333333", linewidth=1.4)

    ax.text(-1.12, 0.92, title, fontsize=FONT_GAUGE_TITLE, ha="left", va="center")
    ax.text(0, 0.38, "<目標対比>", fontsize=FONT_GAUGE_LABEL, ha="center", va="center")

    ax.text(
        0,
        0.19,
        f"{ratio * 100:.0f}%",
        fontsize=FONT_GAUGE_PERCENT,
        weight="bold",
        ha="center",
        va="center",
    )

    ax.set_xlim(-1.25, 1.25)
    ax.set_ylim(-0.18, 1.25)

# 200%時の下半円は表示範囲外にはみ出しても描画する
    for patch in ax.patches:
        patch.set_clip_on(False)

    for line in ax.lines:
        line.set_clip_on(False)

    for text in ax.texts:
        text.set_clip_on(False)

def create_header(fig, report_title, deadline_date_text):
    close_label = f"{deadline_date_text}〆"

    fig.text(
        0.045,
        0.925,
        f"Affiliate （{report_title}） {close_label}",
        fontsize=FONT_HEADER,
        weight="bold",
        ha="left",
        va="center",
        color="#c90000",
    )

    fig.lines.append(
        plt.Line2D(
            [0.035, 0.965],
            [0.875, 0.875],
            transform=fig.transFigure,
            color="#c90000",
            linewidth=3.8,
        )
    )


def create_summary_image(
    kpi,
    cost_filtered,
    target_month,
    deadline_date_text,
):
    """
    全体実績 16:9固定画像を作成。
    """
    report_title = f"{target_month}月度全体実績"
    footer_text = f"※{deadline_date_text}時点での数値"

    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    create_header(fig, report_title, deadline_date_text)

    # 上段枠
    ax_bg = fig.add_axes([0.045, 0.50, 0.91, 0.32])
    ax_bg.axis("off")
    ax_bg.add_patch(
        Rectangle(
            (0, 0),
            1,
            1,
            facecolor="white",
            edgecolor="#cccccc",
            linewidth=1.2,
        )
    )
    ax_bg.text(
        0.015,
        0.92,
        f"【{report_title}】",
        fontsize=FONT_BOX_TITLE,
        weight="bold",
        ha="left",
        va="center",
    )

    # CVゲージ
    ax_cv = fig.add_axes([0.060, 0.545, 0.225, 0.235])
    draw_gauge(ax_cv, kpi["total_cv_ratio"], "CV")

    # CVテキスト枠
    ax_cv_text = fig.add_axes([0.285, 0.535, 0.275, 0.255])
    ax_cv_text.axis("off")
    ax_cv_text.add_patch(Rectangle((0, 0), 1, 1, facecolor="#eeeeee", edgecolor="none"))

    ax_cv_text.text(
        0.04,
        0.78,
        f"全体CV： {fmt_num(kpi['total_cv'])} 件（{kpi['total_cv_ratio'] * 100:.1f}%）",
        fontsize=FONT_KPI_MAIN,
        weight="bold",
        va="center",
    )
    ax_cv_text.text(0.04, 0.50, "▼内訳", fontsize=FONT_KPI_SUBTITLE, weight="bold", va="center")
    ax_cv_text.text(
        0.04,
        0.32,
        f"還元サイトCV： {fmt_num(kpi['point_actual'])} 件（{kpi['point_ratio'] * 100:.1f}%）",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )
    ax_cv_text.text(
        0.04,
        0.18,
        f"非還元サイトCV： {fmt_num(kpi['compare_actual'])} 件（{kpi['compare_ratio'] * 100:.1f}%）",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )

    # COSTゲージ
    ax_cost = fig.add_axes([0.570, 0.545, 0.225, 0.235])
    draw_gauge(ax_cost, kpi["cost_ratio"], "COST")

    # 右テキスト枠
    ax_right = fig.add_axes([0.785, 0.525, 0.160, 0.275])
    ax_right.axis("off")
    ax_right.add_patch(Rectangle((0, 0), 1, 1, facecolor="#eeeeee", edgecolor="none"))

    ax_right.text(
        0.05,
        0.86,
        f"承認率： {kpi['approval_rate'] * 100:.1f} %",
        fontsize=FONT_RIGHT_MAIN,
        weight="bold",
        va="center",
    )
    ax_right.text(
        0.09,
        0.72,
        f"（想定： {kpi['target_approval_rate'] * 100:.1f} %）",
        fontsize=FONT_RIGHT_SUB,
        va="center",
    )

    ax_right.text(
        0.05,
        0.49,
        f"申込CPA： {fmt_yen(kpi['apply_cpa'])}",
        fontsize=FONT_RIGHT_MAIN,
        weight="bold",
        va="center",
    )
    ax_right.text(
        0.09,
        0.36,
        f"（想定： {fmt_yen(kpi['target_apply_cpa'])}）",
        fontsize=FONT_RIGHT_SUB,
        va="center",
    )

    ax_right.text(
        0.05,
        0.15,
        f"発行CPA： {fmt_yen(kpi['issue_cpa'])}",
        fontsize=FONT_RIGHT_MAIN,
        weight="bold",
        va="center",
    )
    ax_right.text(
        0.09,
        0.03,
        f"（想定： {fmt_yen(kpi['target_issue_cpa'])}）",
        fontsize=FONT_RIGHT_SUB,
        va="bottom",
    )

    # 下段グラフ
    # 棒グラフ：CV（発行_Forecast / 発行_実績）
    # 折れ線グラフ：COST（Forecastコスト / 発行コスト）
    ax = fig.add_axes([0.16, 0.15, 0.74, 0.30])
    ax2 = ax.twinx()

    chart_df = cost_filtered.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.30

    # 棒：CV
    ax.bar(
        [i - width / 2 for i in x],
        chart_df["発行_Forecast"],
        width=width,
        label="Forecast",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["発行_実績"],
        width=width,
        label="実績",
        color="#d50000",
    )

    # 線：COST
    ax2.plot(
        x,
        chart_df["Forecastコスト"],
        label="forecastコスト",
        color="#00a0df",
        linewidth=2.4,
    )
    ax2.plot(
        x,
        chart_df["発行コスト"],
        label="発行コスト",
        color="#005a8c",
        linewidth=2.4,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=FONT_AXIS)

    # 左軸：CV件数
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(axis="y", labelsize=FONT_AXIS)

    # 右軸：COST金額
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    ax2.tick_params(axis="y", labelsize=FONT_AXIS)

    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)
    for spine in ["top", "right", "left", "bottom"]:
        ax2.spines[spine].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()

    ax.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="center left",
        bbox_to_anchor=(-0.22, 0.56),
        frameon=False,
        fontsize=FONT_LEGEND,
        borderaxespad=0,
    )

    fig.text(
        0.050,
        0.075,
        footer_text,
        fontsize=FONT_FOOTER,
        ha="left",
        va="center",
        color="#333333",
    )

    img_bytes = io.BytesIO()

    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )

    plt.close(fig)
    img_bytes.seek(0)

    return img_bytes


def create_pointsite_image(
    point_values,
    point_chart_df,
    target_month,
    deadline_date_text,
):
    """
    還元サイト（ポイントサイト）用 16:9 PNG画像を作成。
    """
    report_title = f"{target_month}月度定常期間還元サイト実績"
    progress_title = f"{target_month}月度定常期間全体進捗"
    footer_text = f"※{deadline_date_text}時点での数値"

    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    create_header(fig, report_title, deadline_date_text)

    # 見出し
    fig.text(
        0.16,
        0.795,
        f"【{progress_title}】",
        fontsize=FONT_BOX_TITLE,
        weight="bold",
        ha="left",
        va="center",
        color="#111111",
    )

    # CVゲージ
    ax_cv = fig.add_axes([0.18, 0.535, 0.33, 0.28])
    draw_gauge(ax_cv, point_values["ratio"], "CV")

    # KPIテキスト枠
    ax_text = fig.add_axes([0.62, 0.55, 0.20, 0.25])
    ax_text.axis("off")
    ax_text.add_patch(Rectangle((0, 0), 1, 1, facecolor="#eeeeee", edgecolor="none"))

    ax_text.text(
        0.05,
        0.72,
        f"{target_month}月度還元サイトCV",
        fontsize=FONT_KPI_SUBTITLE,
        weight="bold",
        va="center",
    )
    ax_text.text(
        0.05,
        0.52,
        f"Target　　　： {fmt_num(point_values['target'])} 件",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )
    ax_text.text(
        0.05,
        0.37,
        f"Actual　　　： {fmt_num(point_values['actual'])} 件",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )
    ax_text.text(
        0.05,
        0.22,
        f"目標対比　 ： {point_values['ratio'] * 100:.0f} %",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )

    # 日別Target/Actual棒グラフ
    ax = fig.add_axes([0.16, 0.20, 0.74, 0.28])

    chart_df = point_chart_df.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.28

    ax.bar(
        [i - width / 2 for i in x],
        chart_df["Target"],
        width=width,
        label="Target",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["Actual"],
        width=width,
        label="Actual",
        color="#d50000",
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=FONT_AXIS)
    ax.tick_params(axis="y", labelsize=FONT_AXIS)

    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)

    ax.legend(
        loc="center left",
        bbox_to_anchor=(-0.13, 0.50),
        frameon=False,
        fontsize=FONT_LEGEND,
        borderaxespad=0,
    )

    fig.text(
        0.050,
        0.075,
        footer_text,
        fontsize=FONT_FOOTER,
        ha="left",
        va="center",
        color="#333333",
    )

    img_bytes = io.BytesIO()

    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )

    plt.close(fig)
    img_bytes.seek(0)

    return img_bytes


def create_comparesite_image(
    compare_values,
    compare_chart_df,
    target_month,
    deadline_date_text,
):
    """
    非還元サイト（比較サイト）用 16:9 PNG画像を作成。
    """
    report_title = f"{target_month}月度定常期間非還元サイト実績"
    progress_title = f"{target_month}月度定常期間全体進捗"
    footer_text = f"※{deadline_date_text}時点での数値"

    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    create_header(fig, report_title, deadline_date_text)

    # 見出し
    fig.text(
        0.16,
        0.795,
        f"【{progress_title}】",
        fontsize=FONT_BOX_TITLE,
        weight="bold",
        ha="left",
        va="center",
        color="#111111",
    )

    # CVゲージ
    ax_cv = fig.add_axes([0.18, 0.535, 0.33, 0.28])
    draw_gauge(ax_cv, compare_values["ratio"], "CV")

    # KPIテキスト枠
    ax_text = fig.add_axes([0.62, 0.55, 0.20, 0.25])
    ax_text.axis("off")
    ax_text.add_patch(Rectangle((0, 0), 1, 1, facecolor="#eeeeee", edgecolor="none"))

    ax_text.text(
        0.05,
        0.72,
        f"{target_month}月度非還元サイトCV",
        fontsize=FONT_KPI_SUBTITLE,
        weight="bold",
        va="center",
    )
    ax_text.text(
        0.05,
        0.52,
        f"Target　　　： {fmt_num(compare_values['target'])} 件",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )
    ax_text.text(
        0.05,
        0.37,
        f"Actual　　　： {fmt_num(compare_values['actual'])} 件",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )
    ax_text.text(
        0.05,
        0.22,
        f"目標対比　 ： {compare_values['ratio'] * 100:.0f} %",
        fontsize=FONT_KPI_DETAIL,
        va="center",
    )

    # 日別Target/Actual棒グラフ
    ax = fig.add_axes([0.16, 0.20, 0.74, 0.28])

    chart_df = compare_chart_df.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.28

    ax.bar(
        [i - width / 2 for i in x],
        chart_df["Target"],
        width=width,
        label="Target",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["Actual"],
        width=width,
        label="Actual",
        color="#d50000",
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=FONT_AXIS)
    ax.tick_params(axis="y", labelsize=FONT_AXIS)

    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)

    ax.legend(
        loc="center left",
        bbox_to_anchor=(-0.13, 0.50),
        frameon=False,
        fontsize=FONT_LEGEND,
        borderaxespad=0,
    )

    fig.text(
        0.050,
        0.075,
        footer_text,
        fontsize=FONT_FOOTER,
        ha="left",
        va="center",
        color="#333333",
    )

    img_bytes = io.BytesIO()

    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )

    plt.close(fig)
    img_bytes.seek(0)

    return img_bytes


def create_media_values(daily_filtered, site_name):
    """
    媒体別分析用KPIを作成。

    対象：デイリーレポート32行目以降の明細データ
    サイト名：UIで選択
    Target：Daily Target (Initiative)
    Actual：Actual
    CT：CT
    """
    target = get_daily_value(
        daily_filtered,
        site_name,
        "Daily Target (Initiative)",
    )

    actual = get_daily_value(
        daily_filtered,
        site_name,
        "Actual",
    )

    ratio = safe_div(actual, target)

    return {
        "target": target,
        "actual": actual,
        "ratio": ratio,
    }


def create_media_chart_df(daily_filtered, site_name):
    """
    媒体別分析の日別Target/Actual/CTデータを作成。
    """
    target_df = daily_filtered[
        (daily_filtered["サイト名"] == site_name)
        & (daily_filtered["ディテール"] == "Daily Target (Initiative)")
    ][["日付", "値"]].copy()
    target_df = target_df.rename(columns={"値": "Target"})

    actual_df = daily_filtered[
        (daily_filtered["サイト名"] == site_name)
        & (daily_filtered["ディテール"] == "Actual")
    ][["日付", "値"]].copy()
    actual_df = actual_df.rename(columns={"値": "Actual"})

    ct_df = daily_filtered[
        (daily_filtered["サイト名"] == site_name)
        & (daily_filtered["ディテール"] == "CT")
    ][["日付", "値"]].copy()
    ct_df = ct_df.rename(columns={"値": "CT"})

    chart_df = pd.merge(target_df, actual_df, on="日付", how="outer")
    chart_df = pd.merge(chart_df, ct_df, on="日付", how="outer")

    for col in ["Target", "Actual", "CT"]:
        chart_df[col] = pd.to_numeric(chart_df[col], errors="coerce").fillna(0)

    chart_df = chart_df.sort_values("日付")

    return chart_df


def create_media_image(
    media_values,
    media_chart_df,
    target_month,
    deadline_date_text,
    site_name,
):
    """
    媒体別分析用 16:9 PNG画像を作成。
    """
    report_title = f"媒体別分析（{site_name}）"
    footer_text = f"※{deadline_date_text}時点での数値"

    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    # ヘッダー：媒体別分析は「Affiliate」なしで見本寄せ
    fig.text(
        0.045,
        0.925,
        report_title,
        fontsize=FONT_HEADER,
        weight="bold",
        ha="left",
        va="center",
        color="#c90000",
    )

    fig.lines.append(
        plt.Line2D(
            [0.035, 0.965],
            [0.875, 0.875],
            transform=fig.transFigure,
            color="#c90000",
            linewidth=3.8,
        )
    )

    # 左見出し
    fig.text(
        0.060,
        0.795,
        "【CV数進捗】",
        fontsize=FONT_BOX_TITLE,
        weight="bold",
        ha="left",
        va="center",
        color="#111111",
    )

    # ゲージ
    ax_cv = fig.add_axes([0.070, 0.555, 0.260, 0.245])
    draw_gauge(ax_cv, media_values["ratio"], "")

    # 中央矢印
    fig.text(
        0.345,
        0.655,
        "→",
        fontsize=18,
        ha="center",
        va="center",
        color="#111111",
    )

    # 進捗テキスト
    ax_text = fig.add_axes([0.405, 0.585, 0.250, 0.160])
    ax_text.axis("off")

    ax_text.text(
        0.00,
        0.85,
        f"〈{deadline_date_text}時点での進捗〉",
        fontsize=17,
        weight="bold",
        ha="left",
        va="center",
    )
    ax_text.text(
        0.05,
        0.55,
        f"Target: {fmt_num(media_values['target'])} 件",
        fontsize=18,
        weight="bold",
        ha="left",
        va="center",
    )
    ax_text.text(
        0.05,
        0.27,
        f"Actual: {fmt_num(media_values['actual'])} 件",
        fontsize=18,
        weight="bold",
        ha="left",
        va="center",
    )

    # 下段グラフ
    # 棒：CV、線：CT
    ax = fig.add_axes([0.16, 0.18, 0.72, 0.30])
    ax2 = ax.twinx()

    chart_df = media_chart_df.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.28

    ax.bar(
        [i - width / 2 for i in x],
        chart_df["Target"],
        width=width,
        label="Target",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["Actual"],
        width=width,
        label="Actual",
        color="#d50000",
    )

    ax2.plot(
        x,
        chart_df["CT"],
        label="CT",
        color="#00a0df",
        linewidth=2.4,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=FONT_AXIS)

    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(axis="y", labelsize=FONT_AXIS)

    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax2.tick_params(axis="y", labelsize=FONT_AXIS)

    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)
    for spine in ["top", "right", "left", "bottom"]:
        ax2.spines[spine].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()

    ax.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="center left",
        bbox_to_anchor=(-0.15, 0.50),
        frameon=False,
        fontsize=FONT_LEGEND,
        borderaxespad=0,
    )

    fig.text(
        0.050,
        0.075,
        footer_text,
        fontsize=FONT_FOOTER,
        ha="left",
        va="center",
        color="#333333",
    )

    img_bytes = io.BytesIO()

    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )

    plt.close(fig)
    img_bytes.seek(0)

    return img_bytes
    

def create_media_gauge_part_image(media_values):
    """
    媒体別分析のゲージ部分だけを、元画像と同じ16:9キャンバスで作成。
    画面には表示せず、ダウンロード用に使う。
    """
    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    ax_cv = fig.add_axes([0.070, 0.555, 0.260, 0.245])
    draw_gauge(ax_cv, media_values["ratio"], "")

    img_bytes = io.BytesIO()
    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )
    plt.close(fig)
    img_bytes.seek(0)
    return img_bytes


def create_media_progress_part_image(media_values, deadline_date_text):
    """
    媒体別分析の進捗テキスト部分だけを、元画像と同じ16:9キャンバスで作成。
    画面には表示せず、ダウンロード用に使う。
    """
    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    ax_text = fig.add_axes([0.405, 0.585, 0.250, 0.160])
    ax_text.axis("off")

    ax_text.text(
        0.00,
        0.85,
        f"〈{deadline_date_text}時点での進捗〉",
        fontsize=17,
        weight="bold",
        ha="left",
        va="center",
    )
    ax_text.text(
        0.05,
        0.55,
        f"Target: {fmt_num(media_values['target'])} 件",
        fontsize=18,
        weight="bold",
        ha="left",
        va="center",
    )
    ax_text.text(
        0.05,
        0.27,
        f"Actual: {fmt_num(media_values['actual'])} 件",
        fontsize=18,
        weight="bold",
        ha="left",
        va="center",
    )

    img_bytes = io.BytesIO()
    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )
    plt.close(fig)
    img_bytes.seek(0)
    return img_bytes


def create_media_chart_part_image(media_chart_df):
    """
    媒体別分析の折れ線・棒グラフ部分だけを、元画像と同じ16:9キャンバスで作成。
    画面には表示せず、ダウンロード用に使う。
    """
    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    ax = fig.add_axes([0.16, 0.18, 0.72, 0.30])
    ax2 = ax.twinx()

    chart_df = media_chart_df.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.28

    ax.bar(
        [i - width / 2 for i in x],
        chart_df["Target"],
        width=width,
        label="Target",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["Actual"],
        width=width,
        label="Actual",
        color="#d50000",
    )
    ax2.plot(
        x,
        chart_df["CT"],
        label="CT",
        color="#00a0df",
        linewidth=2.4,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=FONT_AXIS)

    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(axis="y", labelsize=FONT_AXIS)

    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax2.tick_params(axis="y", labelsize=FONT_AXIS)

    ax.grid(axis="x", color="#dddddd", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)
    for spine in ["top", "right", "left", "bottom"]:
        ax2.spines[spine].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()

    ax.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="center left",
        bbox_to_anchor=(-0.15, 0.50),
        frameon=False,
        fontsize=FONT_LEGEND,
        borderaxespad=0,
    )

    img_bytes = io.BytesIO()
    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )
    plt.close(fig)
    img_bytes.seek(0)
    return img_bytes




def draw_af_small_gauge(ax, ratio, side_label="CV", center_label="<目標対比>", percent_text=None):
    """
    AF実績画像用の小さめゲージ。
    サンプルに寄せて、半円ゲージ＋中央パーセント＋下線で描画する。
    """
    ax.set_aspect("equal")
    ax.axis("off")

    base_color = "#f7caca"
    main_color = "#d50000"

    ax.add_patch(
        Wedge(
            (0, 0),
            1.0,
            0,
            180,
            width=0.32,
            facecolor=base_color,
            edgecolor="none",
        )
    )

    ratio_clamped = min(max(ratio, 0), 1)
    degree = ratio_clamped * 180

    ax.add_patch(
        Wedge(
            (0, 0),
            1.0,
            180 - degree,
            180,
            width=0.32,
            facecolor=main_color,
            edgecolor="none",
        )
    )

    ax.plot([-1.25, 1.25], [0, 0], color="#777777", linewidth=1.2)

    if side_label:
        ax.text(
            -1.38,
            0.92,
            side_label,
            fontsize=15,
            ha="left",
            va="center",
            color="#333333",
        )

    if percent_text is None:
        percent_text = f"{ratio * 100:.0f}%"

    ax.text(
        0,
        0.38,
        center_label,
        fontsize=11,
        weight="bold",
        ha="center",
        va="center",
        color="#333333",
    )
    ax.text(
        0,
        0.17,
        percent_text,
        fontsize=18,
        weight="bold",
        ha="center",
        va="center",
        color="#111111",
    )

    ax.set_xlim(-1.45, 1.30)
    ax.set_ylim(-0.08, 1.22)

    for patch in ax.patches:
        patch.set_clip_on(False)
    for line in ax.lines:
        line.set_clip_on(False)
    for text in ax.texts:
        text.set_clip_on(False)


def add_round_box(ax, xy, width, height, radius=0.04, facecolor="#dddddd"):
    box = FancyBboxPatch(
        xy,
        width,
        height,
        boxstyle=f"round,pad=0.018,rounding_size={radius}",
        linewidth=0,
        facecolor=facecolor,
        edgecolor="none",
    )
    ax.add_patch(box)
    return box


def create_af_overall_image(
    af_values,
    af_chart_df,
    target_month,
    deadline_date_text,
):
    """
    サンプル寄せのAF実績画像を作成する。
    """
    report_title = f"AF実績({target_month}月度全体実績)  {deadline_date_text}〆"

    fig = plt.figure(figsize=(16, 9), dpi=120)
    fig.patch.set_facecolor("white")

    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    canvas.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    # ヘッダー
    fig.text(
        0.055,
        0.925,
        report_title,
        fontsize=26,
        weight="bold",
        ha="left",
        va="center",
        color="#c90000",
    )
    fig.lines.append(
        plt.Line2D(
            [0.045, 0.965],
            [0.875, 0.875],
            transform=fig.transFigure,
            color="#c90000",
            linewidth=2.4,
        )
    )

    # ゲージ：左上 AF発生CV
    ax_g1 = fig.add_axes([0.075, 0.615, 0.230, 0.200])
    draw_af_small_gauge(
        ax_g1,
        af_values["af_apply_ratio"],
        side_label="AFCV",
        center_label="<目標対比>",
    )

    # ゲージ：左下 AF承認CV（発行は目標対比なし）
    ax_g2 = fig.add_axes([0.075, 0.405, 0.230, 0.200])
    issue_ratio = safe_div(
        af_values["af_issue"],
        af_values["issue_target"]
    )

    draw_af_small_gauge(
        ax_g2,
        issue_ratio,
        side_label="承認",
        center_label="<目標対比>",
    )

    # 実績ボックス
    ax_box = fig.add_axes([0.320, 0.395, 0.215, 0.385])
    ax_box.axis("off")
    add_round_box(ax_box, (0, 0), 1, 1, radius=0.12, facecolor="#dddddd")
    ax_box.text(0.50, 0.72, "実績", fontsize=16, weight="bold", ha="center", va="center", color="#333333")
    ax_box.text(
        0.10,
        0.50,
        f"AF発生： {fmt_num(af_values['af_apply'])}件（{af_values['af_apply_ratio'] * 100:.0f}%）",
        fontsize=15,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )
    ax_box.text(
        0.10,
        0.32,
        f"AF承認： {fmt_num(af_values['af_issue'])}件",
        fontsize=15,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )

    # 右側CPAゲージ
    cpa_ratio_for_gauge = safe_div(af_values["apply_forecast_cpa"], af_values["apply_actual_cpa"])
    ax_g3 = fig.add_axes([0.590, 0.605, 0.230, 0.200])
    draw_af_small_gauge(
        ax_g3,
        cpa_ratio_for_gauge,
        side_label="",
        center_label="<目標対比>",
    )

    issue_cpa_ratio_for_gauge = safe_div(af_values["issue_forecast_cpa"], af_values["issue_actual_cpa"])
    ax_g4 = fig.add_axes([0.590, 0.395, 0.230, 0.200])
    draw_af_small_gauge(
        ax_g4,
        issue_cpa_ratio_for_gauge,
        side_label="",
        center_label="<目標対比>",
    )

    # CPAボックス
    ax_cpa = fig.add_axes([0.825, 0.395, 0.135, 0.385])
    ax_cpa.axis("off")
    add_round_box(ax_cpa, (0, 0), 1, 1, radius=0.12, facecolor="#dddddd")
    ax_cpa.text(0.50, 0.80, "CPA", fontsize=14, weight="bold", ha="center", va="center", color="#333333")

    ax_cpa.text(
        0.08,
        0.62,
        f"AF発生： {fmt_yen(af_values['apply_actual_cpa'])}",
        fontsize=12,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )
    ax_cpa.text(
        0.14,
        0.52,
        f"（想定比：{af_values['apply_cpa_ratio'] * 100:.0f}%）",
        fontsize=11,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )

    ax_cpa.text(
        0.08,
        0.34,
        f"AF承認： {fmt_yen(af_values['issue_actual_cpa'])}",
        fontsize=12,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )
    ax_cpa.text(
        0.14,
        0.24,
        f"（想定比：{af_values['issue_cpa_ratio'] * 100:.0f}%）",
        fontsize=11,
        weight="bold",
        ha="left",
        va="center",
        color="#333333",
    )

    # 下段グラフ
    ax = fig.add_axes([0.235, 0.145, 0.650, 0.245])
    ax2 = ax.twinx()

    chart_df = af_chart_df.copy().sort_values("日付")
    chart_df["日付ラベル"] = to_md_label(chart_df["日付"])

    x = list(range(len(chart_df)))
    width = 0.30

    ax.bar(
        [i - width / 2 for i in x],
        chart_df["Forecast"],
        width=width,
        label="Forecast",
        color="#f5c6c6",
    )
    ax.bar(
        [i + width / 2 for i in x],
        chart_df["実績"],
        width=width,
        label="実績",
        color="#d50000",
    )

    ax2.plot(
        x,
        chart_df["forecastコスト"],
        label="forecastコスト",
        color="#00a0df",
        linewidth=2.2,
    )
    ax2.plot(
        x,
        chart_df["発行コスト"],
        label="発行コスト",
        color="#005a8c",
        linewidth=2.2,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(chart_df["日付ラベル"], fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(axis="y", labelsize=9)

    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    ax2.tick_params(axis="y", labelsize=9)

    ax.grid(axis="x", color="#eeeeee", linewidth=0.7)
    ax.grid(axis="y", visible=False)

    for spine in ["top", "right", "left", "bottom"]:
        ax.spines[spine].set_visible(False)
    for spine in ["top", "right", "left", "bottom"]:
        ax2.spines[spine].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()

    ax.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="center left",
        bbox_to_anchor=(-0.23, 0.55),
        frameon=False,
        fontsize=10,
        borderaxespad=0,
    )

    fig.text(
        0.050,
        0.075,
        f"※{deadline_date_text}時点での数値",
        fontsize=FONT_FOOTER,
        ha="left",
        va="center",
        color="#333333",
    )

    img_bytes = io.BytesIO()
    fig.savefig(
        img_bytes,
        format="png",
        facecolor="white",
        dpi=120,
        bbox_inches=None,
        pad_inches=0,
    )
    plt.close(fig)
    img_bytes.seek(0)

    return img_bytes


# =========================================================
# UI
# =========================================================

generated_images = []

st.sidebar.header("ファイル読込")

uploaded_af_master = st.sidebar.file_uploader(
    "AFコードマスタをアップロード",
    type=["xlsx"],
    help="BOXから最新版をダウンロードしてアップロードしてください"
)

st.sidebar.markdown(
    "[📂 AFコードマスタはこちら](https://rak.box.com/s/rtkp5rshiwqsa69pkezl13881b552oe0)"
)

daily_file = st.sidebar.file_uploader(
    "デイリーレポート ※PWと数式を解除し該当シートのみのファイルをUP",
    type=["xlsx"],
    key="daily_file",
)

cost_file = st.sidebar.file_uploader(
    "コストレポート ※PWと数式を解除し該当シートのみのファイルをUP",
    type=["xlsx"],
    key="cost_file",
)

af_apply_file = st.sidebar.file_uploader(
    "AF申込実績データ ※A列日付・1行目AFコード",
    type=["xlsx"],
    key="af_apply_file",
)

af_issue_file = st.sidebar.file_uploader(
    "AF発行実績データ ※A列日付・1行目AFコード",
    type=["xlsx"],
    key="af_issue_file",
)

daily_df = None
daily_total_df = None
cost_df = None
af_master_codes = []
af_apply_df = None
af_issue_df = None

if daily_file:
    try:
        daily_df, daily_total_df = read_daily_report(daily_file)
    except Exception as e:
        st.sidebar.error(f"デイリーレポート 読み込み失敗: {e}")

if cost_file:
    try:
        cost_df = read_cost_report(cost_file)
    except Exception as e:
        st.sidebar.error(f"コストレポート 読み込み失敗: {e}")

if uploaded_af_master is not None:
    try:
        af_master_codes = read_af_master(uploaded_af_master)
    except Exception as e:
        st.sidebar.error(f"AFコードマスタ 読み込み失敗: {e}")
else:
    st.sidebar.warning("AFコードマスタをアップロードしてください。")
    st.stop()

if af_apply_file and af_master_codes:
    try:
        af_apply_df = read_af_result_file(
            uploaded_file=af_apply_file,
            af_codes=af_master_codes,
            value_name="AF発生",
        )
    except Exception as e:
        st.sidebar.error(f"AF申込実績データ 読み込み失敗: {e}")

if af_issue_file and af_master_codes:
    try:
        af_issue_df = read_af_result_file(
            uploaded_file=af_issue_file,
            af_codes=af_master_codes,
            value_name="AF承認",
        )
    except Exception as e:
        st.sidebar.error(f"AF発行実績データ 読み込み失敗: {e}")

if daily_df is None and cost_df is None and af_apply_df is None and af_issue_df is None:
    st.info("まずファイルをアップロードしてね。")
    st.stop()


st.sidebar.header("集計条件")

date_candidates = []

if daily_df is not None:
    date_candidates.append(daily_df["日付"])

if cost_df is not None:
    date_candidates.append(cost_df["日付"])

if af_apply_df is not None:
    date_candidates.append(af_apply_df["日付"])

if af_issue_df is not None:
    date_candidates.append(af_issue_df["日付"])

all_dates = pd.concat(date_candidates)

min_date = all_dates.min().date()
max_date = all_dates.max().date()

start_date, end_date = st.sidebar.date_input(
    "集計期間",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

st.sidebar.header("画像テキスト")

default_month = start_date.month
target_month = st.sidebar.number_input(
    "対象月",
    min_value=1,
    max_value=12,
    value=default_month,
    step=1,
    help="例：5 と入力すると、画像では 5月度全体実績 と表示",
)

deadline_date_text = st.sidebar.text_input(
    "締め日",
    value=f"{end_date.month}/{end_date.day}",
    help="例：5/27 と入力すると、画像上部は 5/27〆、フッターは ※5/27時点での数値 ",
)

selected_media_site = None


# =========================================================
# 画像生成
# =========================================================

if daily_total_df is not None and cost_df is not None:
    daily_total_filtered = daily_total_df[
        (daily_total_df["日付"].dt.date >= start_date)
        & (daily_total_df["日付"].dt.date <= end_date)
    ].copy()

    cost_filtered = cost_df[
        (cost_df["日付"].dt.date >= start_date)
        & (cost_df["日付"].dt.date <= end_date)
    ].copy()

    kpi = create_kpi_values(daily_total_filtered, cost_filtered)
    point_values = create_pointsite_values(daily_total_filtered)
    point_chart_df = create_pointsite_daily_chart_df(daily_total_filtered)
    compare_values = create_comparesite_values(daily_total_filtered)
    compare_chart_df = create_comparesite_daily_chart_df(daily_total_filtered)


    # 全体実績画像
    img_bytes = create_summary_image(
        kpi=kpi,
        cost_filtered=cost_filtered,
        target_month=target_month,
        deadline_date_text=deadline_date_text,
    )

    img_data = img_bytes.getvalue()

    render_image_block(
        title="全体実績",
        image_data=img_data,
        file_name="全体実績.png",
        generated_images=generated_images,
    )

    # 還元サイト画像
    point_img_bytes = create_pointsite_image(
        point_values=point_values,
        point_chart_df=point_chart_df,
        target_month=target_month,
        deadline_date_text=deadline_date_text,
    )

    point_img_data = point_img_bytes.getvalue()

    render_image_block(
        title="還元サイト実績",
        image_data=point_img_data,
        file_name="還元サイト実績.png",
        generated_images=generated_images,
    )


    # 非還元サイト画像
    compare_img_bytes = create_comparesite_image(
        compare_values=compare_values,
        compare_chart_df=compare_chart_df,
        target_month=target_month,
        deadline_date_text=deadline_date_text,
    )

    compare_img_data = compare_img_bytes.getvalue()

    render_image_block(
        title="非還元サイト実績",
        image_data=compare_img_data,
        file_name="非還元サイト実績.png",
        generated_images=generated_images,
    )

    # =====================================================
    # AF実績
    # =====================================================

    if af_apply_df is not None or af_issue_df is not None:
        st.subheader("AF実績")

        if af_apply_df is None or af_issue_df is None:
            st.warning("AF実績画像を作るには、AF申込実績データとAF発行実績データの両方が必要です。")
        else:
            af_values = create_af_overall_values(
                af_apply_df=af_apply_df,
                af_issue_df=af_issue_df,
                cost_filtered=cost_filtered,
                total_target_cv=kpi["total_target"],
                start_date=start_date,
                end_date=end_date,
            )

            af_chart_df = create_af_chart_df(cost_filtered)

            # 確認用の数値テーブル
            af_check_df = pd.DataFrame([
                {
                    "項目": "AF発生",
                    "実績": af_values["af_apply"],
                    "目標": af_values["target_cv"],
                    "目標対比": af_values["af_apply_ratio"],
                    "CPA": af_values["apply_actual_cpa"],
                    "CPA想定比": af_values["apply_cpa_ratio"],
                },
                {
                    "項目": "AF承認",
                    "実績": af_values["af_issue"],
                    "目標": af_values["issue_target"],
                    "目標対比": af_values["af_issue_ratio"],
                    "CPA": af_values["issue_actual_cpa"],
                    "CPA想定比": af_values["issue_cpa_ratio"],
                },
            ])

            st.dataframe(af_check_df, use_container_width=True)

            af_img_bytes = create_af_overall_image(
                af_values=af_values,
                af_chart_df=af_chart_df,
                target_month=target_month,
                deadline_date_text=deadline_date_text,
            )

            af_img_data = af_img_bytes.getvalue()

            render_image_block(
                title="AF実績",
                image_data=af_img_data,
                file_name="AF実績.png",
                generated_images=generated_images,
            )

    # =====================================================
    # 媒体別分析
    # =====================================================

    st.subheader("媒体別分析")

    media_site_options = sorted(
        daily_df["サイト名"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: (s != "") & (s.str.lower() != "nan")]
        .unique()
        .tolist()
    )

    if "モッピー" in media_site_options:
        default_media_index = media_site_options.index("モッピー")
    else:
        default_media_index = 0

    selected_media_site = st.selectbox(
        "サイト名",
        options=media_site_options,
        index=default_media_index,
        key="selected_media_site",
    )

    # 媒体別分析画像
    if selected_media_site:
        daily_filtered = daily_df[
            (daily_df["日付"].dt.date >= start_date)
            & (daily_df["日付"].dt.date <= end_date)
        ].copy()

        media_values = create_media_values(
            daily_filtered=daily_filtered,
            site_name=selected_media_site,
        )

        media_chart_df = create_media_chart_df(
            daily_filtered=daily_filtered,
            site_name=selected_media_site,
        )

        media_img_bytes = create_media_image(
            media_values=media_values,
            media_chart_df=media_chart_df,
            target_month=target_month,
            deadline_date_text=deadline_date_text,
            site_name=selected_media_site,
        )

        media_img_data = media_img_bytes.getvalue()

        render_image_block(
            title=f"媒体別分析（{selected_media_site}）",
            image_data=media_img_data,
            file_name=f"媒体別分析_{selected_media_site}.png",
            generated_images=generated_images,
        )
        
        # ==========================================
        # 媒体別分析パーツ：画像は表示せず、DLボタンだけ出す
        # ==========================================

        gauge_img = create_media_gauge_part_image(
            media_values
        ).getvalue()

        progress_img = create_media_progress_part_image(
            media_values,
            deadline_date_text,
        ).getvalue()

        chart_img = create_media_chart_part_image(
            media_chart_df
        ).getvalue()

        st.markdown("##### 媒体別分析 パーツ別ダウンロード")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.download_button(
                "ゲージのみ保存",
                data=gauge_img,
                file_name=f"{selected_media_site}_ゲージ.png",
                mime="image/png",
            )

        with col2:
            st.download_button(
                "進捗テキストのみ保存",
                data=progress_img,
                file_name=f"{selected_media_site}_進捗.png",
                mime="image/png",
            )

        with col3:
            st.download_button(
                "グラフのみ保存",
                data=chart_img,
                file_name=f"{selected_media_site}_グラフ.png",
                mime="image/png",
            )

        # ZIP一括DLにはパーツ画像も含める
        generated_images.extend([
            {
                "filename": f"{selected_media_site}_ゲージ.png",
                "data": gauge_img,
            },
            {
                "filename": f"{selected_media_site}_進捗.png",
                "data": progress_img,
            },
            {
                "filename": f"{selected_media_site}_グラフ.png",
                "data": chart_img,
            },
        ])

else:
    st.warning("画像生成には、デイリーレポートとコストレポートの両方が必要です。")

st.header("一括ダウンロード")
render_zip_download(generated_images)
