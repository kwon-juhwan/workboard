# app.py
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="교환/반품 대시보드",
    layout="wide"
)

st.title("교환 / 반품 월별 지표 대시보드")

# -----------------------------
# 유틸
# -----------------------------
def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x)
    x = x.replace("\n", " ")
    x = re.sub(r"\s+", " ", x).strip()
    return x

def normalize_no_space(x):
    return normalize_text(x).replace(" ", "")

def extract_month(x):
    if pd.isna(x):
        return None
    m = re.search(r"(\d+)\s*월", str(x))
    if m:
        return int(m.group(1))
    return None

def classify_exchange_return(x):
    x = normalize_text(x)
    if "교환" in x:
        return "교환"
    if "반품" in x:
        return "반품"
    return "기타"

def safe_pct(numerator, denominator):
    if denominator == 0:
        return 0.0
    return round((numerator / denominator) * 100, 1)

def calc_delta_text(series, month_value):
    series = series.sort_index()
    if month_value not in series.index:
        return None, "0"
    current = series.loc[month_value]

    prev_months = [m for m in series.index if m < month_value]
    if not prev_months:
        return None, "전월 없음"

    prev = series.loc[max(prev_months)]
    delta = current - prev
    return delta, f"{delta:+,}"

# -----------------------------
# 데이터 로드
# -----------------------------
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # 컬럼명 정리
    cols = list(df.columns)
    rename_map = {}
    for c in cols:
        if str(c).startswith("Unnamed"):
            rename_map[c] = "채널"
    df = df.rename(columns=rename_map)

    expected_cols = ["접수일", "채널", "주문번호", "배송비", "교환/반품"]
    missing = [c for c in expected_cols if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {missing}")

    # 기본 정리
    for col in ["접수일", "채널", "주문번호", "배송비", "교환/반품"]:
        df[col] = df[col].apply(normalize_text)

    df["배송비_공백제거"] = df["배송비"].apply(normalize_no_space)
    df["월"] = df["접수일"].apply(extract_month)
    df["구분"] = df["교환/반품"].apply(classify_exchange_return)

    # 지표용 플래그
    df["미청구(N배송)"] = df["배송비_공백제거"].str.contains(r"미청구\(N배송\)", na=False)
    df["첫구매 무료반품"] = df["배송비_공백제거"].str.contains(r"첫구매무료반품", na=False)
    df["첫구매 무료교환"] = df["배송비_공백제거"].str.contains(r"첫구매무료교환", na=False)

    return df

uploaded_file = st.file_uploader("교환/반품 내역 엑셀 업로드", type=["xlsx"])

if uploaded_file is None:
    st.info("엑셀 파일을 업로드하면 대시보드가 표시됩니다.")
    st.stop()

try:
    df = load_data(uploaded_file)
except Exception as e:
    st.error(f"파일 로드 중 오류가 발생했습니다: {e}")
    st.stop()

# -----------------------------
# 사이드바 필터
# -----------------------------
st.sidebar.header("필터")

month_options = sorted([m for m in df["월"].dropna().unique().tolist()])
channel_options = sorted([c for c in df["채널"].dropna().unique().tolist() if c])

selected_months = st.sidebar.multiselect(
    "월 선택",
    options=month_options,
    default=month_options
)

selected_channels = st.sidebar.multiselect(
    "채널 선택",
    options=channel_options,
    default=channel_options
)

filtered = df.copy()
if selected_months:
    filtered = filtered[filtered["월"].isin(selected_months)]
if selected_channels:
    filtered = filtered[filtered["채널"].isin(selected_channels)]

# -----------------------------
# 월별 집계
# -----------------------------
monthly_special = (
    filtered.groupby("월")[["미청구(N배송)", "첫구매 무료반품", "첫구매 무료교환"]]
    .sum()
    .reset_index()
    .sort_values("월")
)

exchange_return_df = filtered[filtered["구분"].isin(["교환", "반품"])].copy()

monthly_type = (
    exchange_return_df.groupby(["월", "구분"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
    .sort_values("월")
)

if "교환" not in monthly_type.columns:
    monthly_type["교환"] = 0
if "반품" not in monthly_type.columns:
    monthly_type["반품"] = 0

monthly_type["총접수"] = monthly_type["교환"] + monthly_type["반품"]
monthly_type["교환비중(%)"] = monthly_type.apply(
    lambda row: safe_pct(row["교환"], row["총접수"]), axis=1
)
monthly_type["반품비중(%)"] = monthly_type.apply(
    lambda row: safe_pct(row["반품"], row["총접수"]), axis=1
)

# -----------------------------
# 상단 KPI
# -----------------------------
st.subheader("핵심 요약")

col1, col2, col3 = st.columns(3)
col1.metric("미청구(N배송)", int(filtered["미청구(N배송)"].sum()))
col2.metric("첫구매 무료반품", int(filtered["첫구매 무료반품"].sum()))
col3.metric("첫구매 무료교환", int(filtered["첫구매 무료교환"].sum()))

col4, col5, col6 = st.columns(3)
col4.metric("교환 건수", int((filtered["구분"] == "교환").sum()))
col5.metric("반품 건수", int((filtered["구분"] == "반품").sum()))
col6.metric("전체 건수", int(len(filtered)))

# -----------------------------
# 선택 월 기준 전월 대비
# -----------------------------
st.subheader("전월 대비 변화")

if len(selected_months) == 1:
    target_month = selected_months[0]

    ms_indexed = monthly_special.set_index("월")
    mt_indexed = monthly_type.set_index("월")

    d1, d1_text = calc_delta_text(ms_indexed["미청구(N배송)"], target_month) if not ms_indexed.empty else (None, "전월 없음")
    d2, d2_text = calc_delta_text(ms_indexed["첫구매 무료반품"], target_month) if not ms_indexed.empty else (None, "전월 없음")
    d3, d3_text = calc_delta_text(ms_indexed["첫구매 무료교환"], target_month) if not ms_indexed.empty else (None, "전월 없음")
    d4, d4_text = calc_delta_text(mt_indexed["교환"], target_month) if not mt_indexed.empty else (None, "전월 없음")
    d5, d5_text = calc_delta_text(mt_indexed["반품"], target_month) if not mt_indexed.empty else (None, "전월 없음")
    d6, d6_text = calc_delta_text(mt_indexed["반품비중(%)"], target_month) if not mt_indexed.empty else (None, "전월 없음")

    c1, c2, c3 = st.columns(3)
    c1.metric(f"{target_month}월 미청구(N배송)", int(ms_indexed.loc[target_month, "미청구(N배송)"]) if target_month in ms_indexed.index else 0, d1_text)
    c2.metric(f"{target_month}월 첫구매 무료반품", int(ms_indexed.loc[target_month, "첫구매 무료반품"]) if target_month in ms_indexed.index else 0, d2_text)
    c3.metric(f"{target_month}월 첫구매 무료교환", int(ms_indexed.loc[target_month, "첫구매 무료교환"]) if target_month in ms_indexed.index else 0, d3_text)

    c4, c5, c6 = st.columns(3)
    c4.metric(f"{target_month}월 교환 건수", int(mt_indexed.loc[target_month, "교환"]) if target_month in mt_indexed.index else 0, d4_text)
    c5.metric(f"{target_month}월 반품 건수", int(mt_indexed.loc[target_month, "반품"]) if target_month in mt_indexed.index else 0, d5_text)
    c6.metric(
        f"{target_month}월 반품 비중(%)",
        float(mt_indexed.loc[target_month, "반품비중(%)"]) if target_month in mt_indexed.index else 0.0,
        d6_text
    )
else:
    st.info("전월 대비는 월을 1개만 선택했을 때 표시됩니다.")

# -----------------------------
# 차트
# -----------------------------
st.subheader("월별 특수 배송비 사유 추이")
if not monthly_special.empty:
    chart_special = monthly_special.set_index("월")[["미청구(N배송)", "첫구매 무료반품", "첫구매 무료교환"]]
    st.line_chart(chart_special, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")

st.subheader("월별 교환 / 반품 건수")
if not monthly_type.empty:
    chart_count = monthly_type.set_index("월")[["교환", "반품"]]
    st.bar_chart(chart_count, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")

st.subheader("월별 교환 / 반품 비중(현재 파일 기준)")
if not monthly_type.empty:
    chart_ratio = monthly_type.set_index("월")[["교환비중(%)", "반품비중(%)"]]
    st.line_chart(chart_ratio, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")

# -----------------------------
# 표
# -----------------------------
st.subheader("월별 특수 사유 집계표")
st.dataframe(monthly_special, use_container_width=True)

st.subheader("월별 교환 / 반품 비중표")
st.dataframe(monthly_type, use_container_width=True)

# -----------------------------
# 상세 조회
# -----------------------------
st.subheader("상세 내역")

detail_type = st.selectbox(
    "상세 구분",
    ["전체", "교환", "반품", "기타", "미청구(N배송)", "첫구매 무료반품", "첫구매 무료교환"]
)

detail_df = filtered.copy()

if detail_type == "교환":
    detail_df = detail_df[detail_df["구분"] == "교환"]
elif detail_type == "반품":
    detail_df = detail_df[detail_df["구분"] == "반품"]
elif detail_type == "기타":
    detail_df = detail_df[detail_df["구분"] == "기타"]
elif detail_type == "미청구(N배송)":
    detail_df = detail_df[detail_df["미청구(N배송)"]]
elif detail_type == "첫구매 무료반품":
    detail_df = detail_df[detail_df["첫구매 무료반품"]]
elif detail_type == "첫구매 무료교환":
    detail_df = detail_df[detail_df["첫구매 무료교환"]]

show_cols = ["접수일", "월", "채널", "주문번호", "배송비", "교환/반품", "구분"]
st.dataframe(detail_df[show_cols], use_container_width=True)

# -----------------------------
# CSV 다운로드
# -----------------------------
st.subheader("다운로드")

csv_data = detail_df[show_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="현재 상세 내역 CSV 다운로드",
    data=csv_data,
    file_name="교환반품_상세내역.csv",
    mime="text/csv"
)
