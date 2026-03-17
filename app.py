# app.py
import re
from io import BytesIO

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Exchange / Return Dashboard",
    layout="wide"
)

st.title("Exchange / Return Monthly Dashboard")

# =====================================
# GitHub 설정
# =====================================
# 1) public 저장소면 RAW URL만 넣으면 됨
# 2) private 저장소면 st.secrets["GITHUB_TOKEN"] 필요
GITHUB_FILE_URL = "https://raw.githubusercontent.com/kwon-juhwan/workboard/main/data/통합문서1.xlsx"

# private 저장소 사용 여부
USE_GITHUB_TOKEN = False


# =====================================
# 유틸
# =====================================
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
        return "Exchange"
    if "반품" in x:
        return "Return"
    return "Other"

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
        return None, "No previous month"

    prev = series.loc[max(prev_months)]
    delta = current - prev

    if isinstance(delta, float):
        return delta, f"{delta:+.1f}"
    return delta, f"{delta:+,}"


# =====================================
# GitHub에서 파일 읽기
# =====================================
@st.cache_data(ttl=300)
def download_excel_from_github(url, use_token=False):
    headers = {}

    if use_token:
        token = st.secrets.get("GITHUB_TOKEN", "")
        if not token:
            raise ValueError("private 저장소 사용 중인데 st.secrets에 GITHUB_TOKEN이 없습니다.")
        headers["Authorization"] = f"token {token}"

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()

    return BytesIO(response.content)


@st.cache_data(ttl=300)
def load_data_from_github(url, use_token=False):
    file_obj = download_excel_from_github(url, use_token)
    df = pd.read_excel(file_obj)

    # 컬럼 자동 정리
    cols = list(df.columns)
    rename_map = {}

    for c in cols:
        c_str = str(c).strip()

        if c_str.startswith("Unnamed"):
            rename_map[c] = "채널"
        elif c_str in ["접수일", "채널", "주문번호", "배송비", "교환/반품"]:
            rename_map[c] = c_str

    df = df.rename(columns=rename_map)

    expected_cols = ["접수일", "채널", "주문번호", "배송비", "교환/반품"]
    missing = [c for c in expected_cols if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {missing}")

    # 기본 전처리
    for col in expected_cols:
        df[col] = df[col].apply(normalize_text)

    df["배송비_공백제거"] = df["배송비"].apply(normalize_no_space)
    df["월"] = df["접수일"].apply(extract_month)
    df["구분"] = df["교환/반품"].apply(classify_exchange_return)

    # 특수 지표
    df["미청구(N배송)"] = df["배송비_공백제거"].str.contains(r"미청구\(N배송\)", na=False)
    df["첫구매 무료반품"] = df["배송비_공백제거"].str.contains(r"첫구매무료반품", na=False)
    df["첫구매 무료교환"] = df["배송비_공백제거"].str.contains(r"첫구매무료교환", na=False)

    return df


# =====================================
# 데이터 로드
# =====================================
with st.spinner("GitHub에서 엑셀 파일 불러오는 중..."):
    try:
        df = load_data_from_github(GITHUB_FILE_URL, USE_GITHUB_TOKEN)
    except Exception as e:
        st.error(f"GitHub 파일 로드 실패: {e}")
        st.stop()

st.success("최신 GitHub 엑셀 파일을 불러왔습니다.")

with st.expander("현재 연결된 파일 정보"):
    st.write(f"**GitHub URL**: {GITHUB_FILE_URL}")
    st.write(f"**Rows**: {len(df):,}")


# =====================================
# 사이드바 필터
# =====================================
st.sidebar.header("Filter")

month_options = sorted([m for m in df["월"].dropna().unique().tolist()])
channel_options = sorted([c for c in df["채널"].dropna().unique().tolist() if c])

selected_months = st.sidebar.multiselect(
    "Month",
    options=month_options,
    default=month_options
)

selected_channels = st.sidebar.multiselect(
    "Channel",
    options=channel_options,
    default=channel_options
)

filtered = df.copy()

if selected_months:
    filtered = filtered[filtered["월"].isin(selected_months)]

if selected_channels:
    filtered = filtered[filtered["채널"].isin(selected_channels)]


# =====================================
# 월별 집계
# =====================================
monthly_special = (
    filtered.groupby("월")[["미청구(N배송)", "첫구매 무료반품", "첫구매 무료교환"]]
    .sum()
    .reset_index()
    .sort_values("월")
)

exchange_return_df = filtered[filtered["구분"].isin(["Exchange", "Return"])].copy()

monthly_type = (
    exchange_return_df.groupby(["월", "구분"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
    .sort_values("월")
)

if "Exchange" not in monthly_type.columns:
    monthly_type["Exchange"] = 0
if "Return" not in monthly_type.columns:
    monthly_type["Return"] = 0

monthly_type["Total Cases"] = monthly_type["Exchange"] + monthly_type["Return"]
monthly_type["Exchange Ratio (%)"] = monthly_type.apply(
    lambda row: safe_pct(row["Exchange"], row["Total Cases"]), axis=1
)
monthly_type["Return Ratio (%)"] = monthly_type.apply(
    lambda row: safe_pct(row["Return"], row["Total Cases"]), axis=1
)


# =====================================
# 상단 KPI
# =====================================
st.subheader("KPI Summary")

c1, c2, c3 = st.columns(3)
c1.metric("Uncharged (N Shipping)", int(filtered["미청구(N배송)"].sum()))
c2.metric("First Purchase Free Return", int(filtered["첫구매 무료반품"].sum()))
c3.metric("First Purchase Free Exchange", int(filtered["첫구매 무료교환"].sum()))

c4, c5, c6 = st.columns(3)
c4.metric("Exchange Count", int((filtered["구분"] == "Exchange").sum()))
c5.metric("Return Count", int((filtered["구분"] == "Return").sum()))
c6.metric("Total Records", int(len(filtered)))


# =====================================
# 전월 대비
# =====================================
st.subheader("Month-over-Month Change")

if len(selected_months) == 1:
    target_month = selected_months[0]

    ms_indexed = monthly_special.set_index("월") if not monthly_special.empty else pd.DataFrame()
    mt_indexed = monthly_type.set_index("월") if not monthly_type.empty else pd.DataFrame()

    d1, d1_text = calc_delta_text(ms_indexed["미청구(N배송)"], target_month) if not ms_indexed.empty else (None, "No previous month")
    d2, d2_text = calc_delta_text(ms_indexed["첫구매 무료반품"], target_month) if not ms_indexed.empty else (None, "No previous month")
    d3, d3_text = calc_delta_text(ms_indexed["첫구매 무료교환"], target_month) if not ms_indexed.empty else (None, "No previous month")
    d4, d4_text = calc_delta_text(mt_indexed["Exchange"], target_month) if not mt_indexed.empty else (None, "No previous month")
    d5, d5_text = calc_delta_text(mt_indexed["Return"], target_month) if not mt_indexed.empty else (None, "No previous month")
    d6, d6_text = calc_delta_text(mt_indexed["Return Ratio (%)"], target_month) if not mt_indexed.empty else (None, "No previous month")

    k1, k2, k3 = st.columns(3)
    k1.metric(
        f"{target_month}월 Uncharged (N Shipping)",
        int(ms_indexed.loc[target_month, "미청구(N배송)"]) if target_month in ms_indexed.index else 0,
        d1_text
    )
    k2.metric(
        f"{target_month}월 First Purchase Free Return",
        int(ms_indexed.loc[target_month, "첫구매 무료반품"]) if target_month in ms_indexed.index else 0,
        d2_text
    )
    k3.metric(
        f"{target_month}월 First Purchase Free Exchange",
        int(ms_indexed.loc[target_month, "첫구매 무료교환"]) if target_month in ms_indexed.index else 0,
        d3_text
    )

    k4, k5, k6 = st.columns(3)
    k4.metric(
        f"{target_month}월 Exchange Count",
        int(mt_indexed.loc[target_month, "Exchange"]) if target_month in mt_indexed.index else 0,
        d4_text
    )
    k5.metric(
        f"{target_month}월 Return Count",
        int(mt_indexed.loc[target_month, "Return"]) if target_month in mt_indexed.index else 0,
        d5_text
    )
    k6.metric(
        f"{target_month}월 Return Ratio (%)",
        float(mt_indexed.loc[target_month, "Return Ratio (%)"]) if target_month in mt_indexed.index else 0.0,
        d6_text
    )
else:
    st.info("전월 대비는 월을 1개만 선택했을 때 표시됩니다.")


# =====================================
# 차트
# =====================================
st.subheader("Monthly Trend of Special Shipping Reasons")
if not monthly_special.empty:
    chart_special = monthly_special.set_index("월")[["미청구(N배송)", "첫구매 무료반품", "첫구매 무료교환"]]
    st.line_chart(chart_special, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")

st.subheader("Monthly Exchange / Return Count")
if not monthly_type.empty:
    chart_count = monthly_type.set_index("월")[["Exchange", "Return"]]
    st.bar_chart(chart_count, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")

st.subheader("Monthly Exchange / Return Ratio")
if not monthly_type.empty:
    chart_ratio = monthly_type.set_index("월")[["Exchange Ratio (%)", "Return Ratio (%)"]]
    st.line_chart(chart_ratio, use_container_width=True)
else:
    st.warning("표시할 데이터가 없습니다.")


# =====================================
# 집계표
# =====================================
st.subheader("Monthly Special Reason Summary")
st.dataframe(monthly_special, use_container_width=True)

st.subheader("Monthly Exchange / Return Ratio Table")
st.dataframe(monthly_type, use_container_width=True)


# =====================================
# 상세 내역
# =====================================
st.subheader("Detail Records")

detail_type = st.selectbox(
    "Detail Type",
    [
        "All",
        "Exchange",
        "Return",
        "Other",
        "미청구(N배송)",
        "첫구매 무료반품",
        "첫구매 무료교환"
    ]
)

detail_df = filtered.copy()

if detail_type == "Exchange":
    detail_df = detail_df[detail_df["구분"] == "Exchange"]
elif detail_type == "Return":
    detail_df = detail_df[detail_df["구분"] == "Return"]
elif detail_type == "Other":
    detail_df = detail_df[detail_df["구분"] == "Other"]
elif detail_type == "미청구(N배송)":
    detail_df = detail_df[detail_df["미청구(N배송)"]]
elif detail_type == "첫구매 무료반품":
    detail_df = detail_df[detail_df["첫구매 무료반품"]]
elif detail_type == "첫구매 무료교환":
    detail_df = detail_df[detail_df["첫구매 무료교환"]]

show_cols = ["접수일", "월", "채널", "주문번호", "배송비", "교환/반품", "구분"]
st.dataframe(detail_df[show_cols], use_container_width=True)


# =====================================
# 다운로드
# =====================================
st.subheader("Download")

csv_data = detail_df[show_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="Download Current Detail CSV",
    data=csv_data,
    file_name="exchange_return_detail.csv",
    mime="text/csv"
)


# =====================================
# 새로고침 버튼
# =====================================
if st.button("GitHub 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()
