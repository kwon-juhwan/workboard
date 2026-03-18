import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="교환/반품 대시보드", layout="wide")

FILE_PATH = Path(__file__).with_name("교환반품.xlsx")
DEFAULT_SHEET = "Sheet1"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        if pd.isna(c) or str(c).startswith("Unnamed"):
            cols.append("채널")
        else:
            cols.append(str(c).strip())
    df.columns = cols
    return df


def extract_year_from_order_no(value):
    if pd.isna(value):
        return None
    s = str(value)
    m = re.match(r"(20\d{2})", s)
    return int(m.group(1)) if m else None


def parse_date_value(value, fallback_year=None):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value

    s = str(value).strip()
    if not s:
        return pd.NaT

    dt = pd.to_datetime(s, errors="coerce")
    if not pd.isna(dt):
        return dt

    m = re.match(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", s)
    if m and fallback_year:
        month = int(m.group(1))
        day = int(m.group(2))
        return pd.Timestamp(year=fallback_year, month=month, day=day)

    return pd.NaT


def classify_shipping(text: str) -> str:
    s = str(text).strip().replace("\n", " ")
    s_low = s.lower()

    if not s or s in {"nan", '"', "**"}:
        return "기타/미분류"
    if "첫" in s and ("무료반품" in s or "무료 반품" in s):
        return "첫구매 무료반품"
    if "첫" in s and ("무료교환" in s or "무료 교환" in s):
        return "첫구매 무료교환"
    if "당사" in s:
        return "당사부담"
    if "입금완료" in s or "입금 완료" in s:
        return "고객 입금완료"
    if "환불금" in s and "차감" in s:
        return "환불금 차감"
    if "차감" in s:
        return "비용 차감"
    if "보류" in s:
        return "처리보류"
    if "확인" in s_low:
        return "확인 필요"
    return "기타/미분류"


@st.cache_data(ttl=0)
def inspect_excel(file_obj):
    xls = pd.ExcelFile(file_obj)
    info = []
    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(file_obj, sheet_name=sheet_name)
        info.append(
            {
                "sheet_name": sheet_name,
                "rows": len(raw),
                "columns": list(raw.columns),
            }
        )
    return info


@st.cache_data(ttl=0)
def load_data(file_obj, selected_sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(file_obj, sheet_name=selected_sheet)
    raw = normalize_columns(raw)

    need_cols = ["접수일", "채널", "주문번호", "배송비", "교환/반품"]
    raw = raw[[c for c in raw.columns if c in need_cols]].copy()

    for col in need_cols:
        if col not in raw.columns:
            raw[col] = None

    raw["sheet_name"] = selected_sheet

    sample_years = (
        raw["주문번호"]
        .dropna()
        .astype(str)
        .head(100)
        .map(extract_year_from_order_no)
        .dropna()
    )
    fallback_year = int(sample_years.mode().iloc[0]) if not sample_years.empty else pd.Timestamp.today().year

    raw["접수일_dt"] = raw["접수일"].apply(lambda x: parse_date_value(x, fallback_year=fallback_year))
    raw["연도"] = raw["접수일_dt"].dt.year
    raw["월"] = raw["접수일_dt"].dt.month
    raw["연월"] = raw["접수일_dt"].dt.strftime("%Y-%m")
    raw["주차"] = raw["접수일_dt"].dt.isocalendar().week.astype("Int64")
    raw["요일"] = raw["접수일_dt"].dt.day_name()

    raw["채널"] = raw["채널"].fillna("미기재").astype(str).str.strip()
    raw["배송비"] = raw["배송비"].fillna("미기재").astype(str).str.strip()
    raw["교환/반품"] = raw["교환/반품"].fillna("미기재").astype(str).str.strip()
    raw["배송비_분류"] = raw["배송비"].apply(classify_shipping)

    raw = raw.dropna(subset=["접수일_dt"]).copy()
    raw = raw.sort_values("접수일_dt").reset_index(drop=True)
    return raw


def kpi_card(label, value, delta=None):
    with st.container(border=True):
        st.caption(label)
        st.subheader(f"{value:,}")
        if delta is not None:
            st.write(delta)


def main():
    st.title("교환 / 반품 통합 대시보드")
    st.caption("업로드된 Excel 파일 기준으로 교환/반품 현황을 시각화합니다.")

    with st.sidebar:
        st.header("파일")
        uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx"])
        if st.button("캐시 초기화"):
            st.cache_data.clear()
            st.success("캐시를 초기화했습니다.")

    file_source = uploaded_file if uploaded_file is not None else FILE_PATH

    if uploaded_file is None and not FILE_PATH.exists():
        st.error(f"엑셀 파일을 찾을 수 없습니다: {FILE_PATH}")
        st.stop()

    excel_info = inspect_excel(file_source)
    sheet_names = [x["sheet_name"] for x in excel_info]

    default_sheet = DEFAULT_SHEET if DEFAULT_SHEET in sheet_names else sheet_names[0]

    with st.sidebar:
        selected_sheet = st.selectbox("시트 선택", sheet_names, index=sheet_names.index(default_sheet))

    df = load_data(file_source, selected_sheet)

    with st.expander("디버그 확인", expanded=False):
        st.write("사용 파일:", uploaded_file.name if uploaded_file is not None else str(FILE_PATH))
        st.write("시트별 원본 행수")
        st.dataframe(pd.DataFrame(excel_info), use_container_width=True)
        st.write("현재 선택 시트:", selected_sheet)
        st.write("최종 적재 행수:", len(df))
        st.write("중복 행 수(핵심 5개 컬럼 기준):", int(df[["접수일", "채널", "주문번호", "배송비", "교환/반품"]].duplicated().sum()))
        st.write("교환/반품 고유값:", sorted(df["교환/반품"].dropna().astype(str).unique().tolist()))

    with st.sidebar:
        st.header("필터")

        year_options = [int(y) for y in sorted(df["연도"].dropna().unique())]
        selected_years = st.multiselect("연도", year_options, default=year_options)

        type_options = sorted(df["교환/반품"].dropna().unique().tolist())
        selected_types = st.multiselect("구분", type_options, default=type_options)

        channel_options = sorted(df["채널"].dropna().unique().tolist())
        selected_channels = st.multiselect("채널", channel_options, default=channel_options)

        shipping_options = sorted(df["배송비_분류"].dropna().unique().tolist())
        selected_shipping = st.multiselect("배송비 분류", shipping_options, default=shipping_options)

        min_date = df["접수일_dt"].min().date()
        max_date = df["접수일_dt"].max().date()
        date_range = st.date_input("접수일 범위", value=(min_date, max_date), min_value=min_date, max_value=max_date)

    filtered = df.copy()
    if selected_years:
        filtered = filtered[filtered["연도"].isin(selected_years)]
    if selected_types:
        filtered = filtered[filtered["교환/반품"].isin(selected_types)]
    if selected_channels:
        filtered = filtered[filtered["채널"].isin(selected_channels)]
    if selected_shipping:
        filtered = filtered[filtered["배송비_분류"].isin(selected_shipping)]
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered = filtered[(filtered["접수일_dt"].dt.date >= start_date) & (filtered["접수일_dt"].dt.date <= end_date)]

    total_count = len(filtered)
    exchange_count = int((filtered["교환/반품"] == "교환").sum())
    return_count = int((filtered["교환/반품"] == "반품").sum())
    as_count = int((filtered["교환/반품"] == "A/S").sum())

    free_return_count = int((filtered["배송비_분류"] == "첫구매 무료반품").sum())
    company_cost_count = int(filtered["배송비_분류"].isin(["당사부담", "첫구매 무료반품", "첫구매 무료교환"]).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("총 건수", total_count)
    with c2:
        kpi_card("교환", exchange_count, f"비중 {exchange_count / total_count:.1%}" if total_count else "비중 0.0%")
    with c3:
        kpi_card("반품", return_count, f"비중 {return_count / total_count:.1%}" if total_count else "비중 0.0%")
    with c4:
        kpi_card("A/S", as_count, f"비중 {as_count / total_count:.1%}" if total_count else "비중 0.0%")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        exchange_rate = (exchange_count / total_count * 100) if total_count else 0
        kpi_card("교환률", round(exchange_rate, 2), f"{exchange_rate:.2f}%")
    with c6:
        return_rate = (return_count / total_count * 100) if total_count else 0
        kpi_card("반품률", round(return_rate, 2), f"{return_rate:.2f}%")
    with c7:
        kpi_card("첫구매 무료반품", free_return_count, f"비중 {free_return_count / total_count:.1%}" if total_count else "비중 0.0%")
    with c8:
        kpi_card("회사부담성 배송비", company_cost_count, f"비중 {company_cost_count / total_count:.1%}" if total_count else "비중 0.0%")

    col1, col2 = st.columns(2)

    monthly = (
        filtered.groupby(["연월", "교환/반품"], dropna=False)
        .size()
        .reset_index(name="건수")
        .sort_values("연월")
    )
    with col1:
        st.subheader("월별 교환/반품 추이")
        if not monthly.empty:
            fig = px.line(monthly, x="연월", y="건수", color="교환/반품", markers=True)
            fig.update_layout(xaxis_title="연월", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    shipping_summary = (
        filtered.groupby("배송비_분류", dropna=False)
        .size()
        .reset_index(name="건수")
        .sort_values("건수", ascending=False)
    )
    with col2:
        st.subheader("배송비 분류 현황")
        if not shipping_summary.empty:
            fig = px.bar(shipping_summary, x="배송비_분류", y="건수")
            fig.update_layout(xaxis_title="배송비 분류", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    col3, col4 = st.columns(2)

    channel_summary = (
        filtered.groupby(["채널", "교환/반품"], dropna=False)
        .size()
        .reset_index(name="건수")
        .sort_values(["건수", "채널"], ascending=[False, True])
    )
    with col3:
        st.subheader("채널별 현황")
        if not channel_summary.empty:
            fig = px.bar(channel_summary, x="채널", y="건수", color="교환/반품", barmode="group")
            fig.update_layout(xaxis_title="채널", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    daily_summary = (
        filtered.groupby(filtered["접수일_dt"].dt.date)
        .size()
        .reset_index(name="건수")
        .rename(columns={"접수일_dt": "접수일"})
    )
    with col4:
        st.subheader("일자별 접수 건수")
        if not daily_summary.empty:
            fig = px.area(daily_summary, x="접수일", y="건수")
            fig.update_layout(xaxis_title="접수일", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    col5, col6 = st.columns(2)

    cost_by_channel = (
        filtered.groupby(["채널", "배송비_분류"], dropna=False)
        .size()
        .reset_index(name="건수")
    )
    with col5:
        st.subheader("채널별 비용 구조")
        if not cost_by_channel.empty:
            fig = px.bar(cost_by_channel, x="채널", y="건수", color="배송비_분류", barmode="stack")
            fig.update_layout(xaxis_title="채널", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    free_return_trend = (
        filtered[filtered["배송비_분류"] == "첫구매 무료반품"]
        .groupby("연월", dropna=False)
        .size()
        .reset_index(name="건수")
        .sort_values("연월")
    )
    with col6:
        st.subheader("첫구매 무료반품 증가 추이")
        if not free_return_trend.empty:
            fig = px.line(free_return_trend, x="연월", y="건수", markers=True)
            fig.update_layout(xaxis_title="연월", yaxis_title="건수")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("표시할 데이터가 없습니다.")

    st.subheader("요약 테이블")
    summary_table = (
        filtered.pivot_table(
            index="채널",
            columns="교환/반품",
            values="주문번호",
            aggfunc="count",
            fill_value=0,
        )
        .reset_index()
    )
    if not summary_table.empty:
        st.dataframe(summary_table, use_container_width=True)
    else:
        st.info("표시할 데이터가 없습니다.")

    st.subheader("원본 상세 데이터")
    display_cols = ["접수일_dt", "채널", "주문번호", "배송비", "배송비_분류", "교환/반품", "sheet_name"]
    st.dataframe(
        filtered[display_cols].rename(columns={"접수일_dt": "접수일", "sheet_name": "시트명"}),
        use_container_width=True,
        height=420,
    )


if __name__ == "__main__":
    main()
