import re
from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException


st.set_page_config(
    page_title="Google Ads 数据分析工具",
    page_icon="📊",
    layout="wide"
)


def normalize_customer_id(value: str) -> str:
    """把 123-456-7890 变成 1234567890，只保留数字。"""
    return re.sub(r"\D", "", str(value or ""))


def safe_divide(a, b):
    return round(a / b, 4) if b else 0


def check_password() -> bool:
    """简单访问密码，防止工具网址泄露后被别人随便使用。"""
    app_config = st.secrets.get("app", {})
    password = str(app_config.get("password", "")).strip()

    if not password:
        st.warning("你还没有在 secrets.toml 里设置访问密码。上线前建议设置 [app] password。")
        return True

    if st.session_state.get("authenticated", False):
        return True

    st.title("🔐 Google Ads 数据分析工具")
    entered_password = st.text_input("请输入访问密码", type="password")

    if st.button("进入工具", type="primary"):
        if entered_password == password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("密码错误，请重新输入。")

    return False


@st.cache_resource
def get_google_ads_client():
    """从 Streamlit Secrets 读取 Google Ads API 配置。"""
    if "google_ads" not in st.secrets:
        raise RuntimeError("没有找到 [google_ads] 配置，请检查 secrets.toml。")

    config = dict(st.secrets["google_ads"])

    required_keys = [
        "developer_token",
        "client_id",
        "client_secret",
        "refresh_token",
        "login_customer_id",
    ]

    missing = [key for key in required_keys if not str(config.get(key, "")).strip()]
    if missing:
        raise RuntimeError(f"缺少必要配置：{', '.join(missing)}")

    config["login_customer_id"] = normalize_customer_id(config["login_customer_id"])
    config["use_proto_plus"] = True

    return GoogleAdsClient.load_from_dict(config)


def format_google_ads_error(error: GoogleAdsException) -> str:
    messages = []
    for err in error.failure.errors:
        messages.append(err.message)
    return "；".join(messages) if messages else str(error)


def query_campaign_report(client, customer_id: str, start_date: str, end_date: str) -> pd.DataFrame:
    ga_service = client.get_service("GoogleAdsService")

    query = f"""
        SELECT
          customer.currency_code,
          campaign.id,
          campaign.name,
          campaign.status,
          metrics.impressions,
          metrics.clicks,
          metrics.cost_micros,
          metrics.conversions,
          metrics.conversions_value
        FROM campaign
        WHERE
          campaign.status != 'REMOVED'
          AND segments.date BETWEEN '{start_date}' AND '{end_date}'
        ORDER BY metrics.cost_micros DESC
    """

    rows = []

    response = ga_service.search_stream(customer_id=customer_id, query=query)

    for batch in response:
        for row in batch.results:
            impressions = int(row.metrics.impressions or 0)
            clicks = int(row.metrics.clicks or 0)
            cost = float(row.metrics.cost_micros or 0) / 1_000_000
            conversions = float(row.metrics.conversions or 0)
            conversion_value = float(row.metrics.conversions_value or 0)

            rows.append({
                "账户币种": row.customer.currency_code,
                "广告系列ID": str(row.campaign.id),
                "广告系列名称": row.campaign.name,
                "状态": row.campaign.status.name,
                "展示量": impressions,
                "点击量": clicks,
                "花费": round(cost, 2),
                "转化数": round(conversions, 2),
                "转化价值": round(conversion_value, 2),
                "CTR": round(safe_divide(clicks, impressions) * 100, 2),
                "平均CPC": round(safe_divide(cost, clicks), 2),
                "CPA": round(safe_divide(cost, conversions), 2),
                "ROAS": round(safe_divide(conversion_value, cost), 2),
            })

    return pd.DataFrame(rows)


def query_daily_report(client, customer_id: str, start_date: str, end_date: str) -> pd.DataFrame:
    ga_service = client.get_service("GoogleAdsService")

    query = f"""
        SELECT
          segments.date,
          customer.currency_code,
          metrics.impressions,
          metrics.clicks,
          metrics.cost_micros,
          metrics.conversions,
          metrics.conversions_value
        FROM customer
        WHERE segments.date BETWEEN '{start_date}' AND '{end_date}'
        ORDER BY segments.date
    """

    rows = []

    response = ga_service.search_stream(customer_id=customer_id, query=query)

    for batch in response:
        for row in batch.results:
            impressions = int(row.metrics.impressions or 0)
            clicks = int(row.metrics.clicks or 0)
            cost = float(row.metrics.cost_micros or 0) / 1_000_000
            conversions = float(row.metrics.conversions or 0)
            conversion_value = float(row.metrics.conversions_value or 0)

            rows.append({
                "日期": row.segments.date,
                "账户币种": row.customer.currency_code,
                "展示量": impressions,
                "点击量": clicks,
                "花费": round(cost, 2),
                "转化数": round(conversions, 2),
                "转化价值": round(conversion_value, 2),
                "ROAS": round(safe_divide(conversion_value, cost), 2),
            })

    return pd.DataFrame(rows)


def make_excel(campaign_df: pd.DataFrame, daily_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        campaign_df.to_excel(writer, index=False, sheet_name="广告系列数据")
        daily_df.to_excel(writer, index=False, sheet_name="每日趋势")
    return output.getvalue()


if not check_password():
    st.stop()


st.title("📊 Google Ads 数据分析工具")
st.caption("适合个人和小团队查看 Google Ads 广告系列表现，不会修改广告账户。")


try:
    client = get_google_ads_client()
except Exception as e:
    st.error(f"Google Ads API 配置加载失败：{e}")
    st.stop()


app_config = st.secrets.get("app", {})
allowed_customer_ids = [
    normalize_customer_id(x)
    for x in app_config.get("allowed_customer_ids", [])
    if normalize_customer_id(x)
]


with st.sidebar:
    st.header("查询设置")

    if allowed_customer_ids:
        customer_id = st.selectbox(
            "选择广告账户 ID",
            allowed_customer_ids,
            help="这里显示的是 secrets.toml 里 allowed_customer_ids 白名单中的账户。"
        )
    else:
        customer_id = st.text_input(
            "广告账户 ID",
            placeholder="输入10位数字，不要横杠"
        )
        customer_id = normalize_customer_id(customer_id)

    default_start = date.today() - timedelta(days=30)
    default_end = date.today()

    start = st.date_input("开始日期", value=default_start)
    end = st.date_input("结束日期", value=default_end)

    run_button = st.button("刷新数据", type="primary")


if not customer_id:
    st.info("请先在左侧选择或输入广告账户 ID。")
    st.stop()


if allowed_customer_ids and customer_id not in allowed_customer_ids:
    st.error("这个广告账户 ID 不在白名单 allowed_customer_ids 里。")
    st.stop()


if start > end:
    st.error("开始日期不能晚于结束日期。")
    st.stop()


if not run_button:
    st.info("设置好账户和日期后，点击左侧「刷新数据」。")
    st.stop()


start_str = start.isoformat()
end_str = end.isoformat()

with st.spinner("正在从 Google Ads API 拉取数据，请稍等..."):
    try:
        campaign_df = query_campaign_report(client, customer_id, start_str, end_str)
        daily_df = query_daily_report(client, customer_id, start_str, end_str)
    except GoogleAdsException as e:
        st.error("Google Ads API 返回错误：")
        st.code(format_google_ads_error(e))
        st.stop()
    except Exception as e:
        st.error(f"程序运行错误：{e}")
        st.stop()


if campaign_df.empty:
    st.warning("没有查询到广告系列数据。可能原因：日期范围内没有投放，账户 ID 错误，或当前授权账号没有权限。")
    st.stop()


currency = campaign_df["账户币种"].iloc[0] if "账户币种" in campaign_df.columns else ""

total_cost = campaign_df["花费"].sum()
total_clicks = campaign_df["点击量"].sum()
total_impressions = campaign_df["展示量"].sum()
total_conversions = campaign_df["转化数"].sum()
total_conversion_value = campaign_df["转化价值"].sum()
overall_roas = safe_divide(total_conversion_value, total_cost)
overall_ctr = safe_divide(total_clicks, total_impressions) * 100
overall_cpc = safe_divide(total_cost, total_clicks)
overall_cpa = safe_divide(total_cost, total_conversions)


st.subheader("核心指标")

col1, col2, col3, col4 = st.columns(4)
col1.metric(f"总花费（{currency}）", f"{total_cost:,.2f}")
col2.metric("展示量", f"{total_impressions:,}")
col3.metric("点击量", f"{total_clicks:,}")
col4.metric("转化数", f"{total_conversions:,.2f}")

col5, col6, col7, col8 = st.columns(4)
col5.metric("ROAS", f"{overall_roas:.2f}")
col6.metric("CTR", f"{overall_ctr:.2f}%")
col7.metric(f"平均 CPC（{currency}）", f"{overall_cpc:.2f}")
col8.metric(f"CPA（{currency}）", f"{overall_cpa:.2f}")

st.divider()


if not daily_df.empty:
    st.subheader("每日花费趋势")
    fig_daily = px.line(
        daily_df,
        x="日期",
        y="花费",
        markers=True,
        title=f"每日花费趋势（{currency}）"
    )
    st.plotly_chart(fig_daily, use_container_width=True)


st.subheader("广告系列表现")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    cost_df = campaign_df.sort_values("花费", ascending=False)
    fig_cost = px.bar(
        cost_df,
        x="广告系列名称",
        y="花费",
        title=f"广告系列花费排行（{currency}）"
    )
    st.plotly_chart(fig_cost, use_container_width=True)

with col_chart2:
    roas_df = campaign_df.sort_values("ROAS", ascending=False)
    fig_roas = px.bar(
        roas_df,
        x="广告系列名称",
        y="ROAS",
        title="广告系列 ROAS 排行"
    )
    st.plotly_chart(fig_roas, use_container_width=True)


st.divider()

st.subheader("广告系列数据明细")
st.dataframe(campaign_df, use_container_width=True)

st.subheader("每日趋势数据")
st.dataframe(daily_df, use_container_width=True)

excel_data = make_excel(campaign_df, daily_df)

st.download_button(
    label="📥 下载 Excel 报表",
    data=excel_data,
    file_name=f"google_ads_report_{customer_id}_{start_str}_to_{end_str}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)