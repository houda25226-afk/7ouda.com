import io
from datetime import time as dt_time, datetime, date

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# إعدادات
# ==========================================================
MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

ORIGINAL_TEXT_COL = "Notes"
MODEL_TEXT_COL = "الافادة"
CLASSIFICATION_COL = "التصنيف"
WASTED_TIME_COL = "الوقت_المهدر_دقيقة"

SALES_PERSON_CANDIDATES = [
    "Create By", "create by", "CreateBy", "Created By", "created by",
    "Sales Person", "sales person", "Salesperson", "Salesperson Name", "المحصل"
]
CREATED_ON_CANDIDATES = [
    "Created On", "created on", "CreatedOn", "Created Date", "تاريخ الافادة",
    "Date", "date"
]
SUB_STATE_CANDIDATES = ["Sub State", "SubState", "sub state", "الحالة الفرعية"]
FOLLOW_UP_CANDIDATES = [
    "Follow up Due Date", "Follow Up Due Date", "Followup Due Date",
    "Follow Up Due", "Follow up due date", "تاريخ المتابعة"
]
ACCOUNT_CANDIDATES = ["Account Number", "Account No", "Account", "رقم الحساب"]
NET_AMOUNT_CANDIDATES = ["Net Amount", "NetAmount", "صافي المبلغ", "المبلغ الصافي"]

EXCLUDED_SALESPERSONS = {
    "Archive Companies  II Anas",
    "Closed payments  II Anas",
    "Hold Companies  II Anas",
    "Op II Ibrahim Qassem",
    "قانونى -الوطنية",
}

COLOR_SUCCESS = "#34D399"
COLOR_FAIL = "#FB7185"
COLOR_ACCENT = "#5EEAD4"
COLOR_WARN = "#FBBF24"
COLOR_PURPLE = "#A78BFA"
COLOR_BLUE = "#60A5FA"

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color="#E7ECF3",
    font_family="Tajawal, sans-serif",
    margin=dict(t=45, b=35, l=20, r=20),
)
PLOTLY_CONFIG = {"displayModeBar": False, "responsive": True}

st.set_page_config(
    page_title="7ouda Model | لوحة تحليل المكالمات",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================
# Dark Theme
# ==========================================================
CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root{
    --bg:#080D17;
    --surface:#111A2A;
    --surface2:#19263D;
    --surface3:#202F49;
    --accent:#5EEAD4;
    --success:#34D399;
    --danger:#FB7185;
    --warn:#FBBF24;
    --purple:#A78BFA;
    --blue:#60A5FA;
    --text:#F1F5F9;
    --muted:#94A3B8;
    --border:rgba(255,255,255,.075);
}

html,body,[class*="css"]{font-family:'Tajawal',sans-serif;}
.stApp{
    background:
        radial-gradient(circle at 80% 0%, rgba(94,234,212,.055), transparent 28%),
        radial-gradient(circle at 10% 20%, rgba(96,165,250,.04), transparent 25%),
        var(--bg);
    color:var(--text);
}
header[data-testid="stHeader"]{
    background:var(--bg)!important;
    border-bottom:1px solid rgba(255,255,255,.045);
}
div[data-testid="stDecoration"]{background:var(--bg)!important;}
.main .block-container{
    direction:rtl;
    text-align:right;
    max-width:1500px;
    padding-top:2rem;
    padding-bottom:4rem;
}
[data-testid="stSidebar"]{
    background:#090F1A!important;
    border-right:1px solid rgba(255,255,255,.07)!important;
}
[data-testid="stSidebarContent"],[data-testid="stSidebarUserContent"]{
    direction:rtl;
    text-align:right;
}
[data-testid="stSidebar"] *{color:var(--text)!important;}
[data-testid="stSidebar"] .stRadio label{
    background:var(--surface);
    border:1px solid var(--border);
    border-radius:11px;
    padding:.55rem .8rem!important;
    margin-bottom:4px;
}
[data-testid="stSidebar"] .stRadio label:hover{background:var(--surface2);}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"]{gap:4px;}

.sidebar-brand{
    text-align:center;
    padding:.5rem 0 1.2rem;
}
.sidebar-brand .mini{
    font-family:'JetBrains Mono',monospace;
    font-size:.68rem;
    letter-spacing:.14em;
    color:var(--accent);
}
.sidebar-brand .big{
    font-weight:900;
    font-size:1.05rem;
    margin-top:.25rem;
}

.page-eyebrow{
    color:var(--accent);
    font-family:'JetBrains Mono',monospace;
    letter-spacing:.14em;
    font-size:.7rem;
    text-transform:uppercase;
}
.page-title{
    font-size:1.85rem;
    font-weight:900;
    margin:.15rem 0;
}
.page-subtitle{color:var(--muted);font-size:.92rem;}
.divider{
    height:1px;
    margin:1.2rem 0;
    background:linear-gradient(90deg,transparent,rgba(94,234,212,.35),transparent);
}

.card{
    background:linear-gradient(145deg,var(--surface),#0F1726);
    border:1px solid var(--border);
    border-radius:16px;
    padding:1rem 1.15rem;
}
.metric-card{
    background:linear-gradient(145deg,#142239,#0F1726);
    border:1px solid rgba(94,234,212,.10);
    border-radius:16px;
    padding:1rem 1.1rem;
    min-height:125px;
    box-shadow:0 8px 28px rgba(0,0,0,.12);
}
.metric-label{color:var(--muted);font-size:.82rem;}
.metric-value{
    font-family:'JetBrains Mono',monospace;
    font-size:1.55rem;
    font-weight:800;
    color:var(--accent);
    margin-top:.35rem;
}
.metric-sub{color:var(--muted);font-size:.76rem;margin-top:.3rem;}

[data-testid="stVerticalBlockBorderWrapper"]{
    background:linear-gradient(145deg,var(--surface),#0F1726)!important;
    border:1px solid var(--border)!important;
    border-radius:16px!important;
}
.chart-title{
    font-weight:800;
    font-size:.95rem;
    padding-bottom:.5rem;
    margin-bottom:.25rem;
    border-bottom:1px solid var(--border);
}
.highlight{
    background:var(--surface);
    border:1px solid var(--border);
    border-radius:14px;
    padding:1rem;
}
.highlight .label{color:var(--muted);font-size:.8rem;}
.highlight .value{color:var(--accent);font-size:1.15rem;font-weight:900;margin-top:.3rem;}
.highlight .sub{color:var(--muted);font-size:.76rem;margin-top:.15rem;}

[data-testid="stFileUploader"]{
    background:var(--surface)!important;
    border:1.5px dashed rgba(94,234,212,.30)!important;
    border-radius:15px!important;
    padding:.5rem!important;
}
[data-testid="stFileUploader"] section{background:transparent!important;}
[data-testid="stFileUploader"] *{color:var(--text)!important;}
[data-testid="stFileUploader"] small{opacity:1!important;}
[data-testid="stFileUploader"] button{
    background:var(--surface2)!important;
    color:var(--text)!important;
    border:1px solid var(--border)!important;
}
[data-testid="stFileUploaderFile"]{background:var(--surface2)!important;}

[data-testid="stWidgetLabel"] *{color:var(--text)!important;opacity:1!important;}
div[data-baseweb="select"]>div{
    background:var(--surface2)!important;
    color:var(--text)!important;
    border-color:var(--border)!important;
    border-radius:10px!important;
}
div[data-baseweb="popover"],div[data-baseweb="popover"] ul{background:var(--surface2)!important;}
li[role="option"]{color:var(--text)!important;}
li[role="option"]:hover,li[aria-selected="true"]{background:var(--surface3)!important;}
input,textarea{
    background:var(--surface2)!important;
    color:var(--text)!important;
}
[data-testid="stDateInput"] input,[data-testid="stTimeInput"] input{
    background:var(--surface2)!important;
    color:var(--text)!important;
}

.stButton>button,.stDownloadButton>button{
    border-radius:10px!important;
    font-weight:800!important;
    border:1px solid rgba(94,234,212,.18)!important;
}
.stButton>button[kind="primary"]{
    background:linear-gradient(90deg,#5EEAD4,#38BFAE)!important;
    color:#05211C!important;
}
.stDownloadButton>button{
    background:var(--surface2)!important;
    color:var(--text)!important;
}
.stDownloadButton>button:hover{border-color:var(--accent)!important;color:var(--accent)!important;}

[data-testid="stMetric"]{
    background:transparent!important;
    border:0!important;
    padding:0!important;
}
[data-testid="stMetricValue"]{color:var(--accent)!important;font-family:'JetBrains Mono',monospace;}
[data-testid="stDataFrame"]{
    border:1px solid var(--border);
    border-radius:12px;
    overflow:hidden;
}
.stTabs [data-baseweb="tab-list"]{
    gap:5px;
    border-bottom:1px solid var(--border);
}
.stTabs [data-baseweb="tab"]{
    color:var(--muted);
    background:var(--surface);
    border-radius:10px 10px 0 0;
}
.stTabs [aria-selected="true"]{
    color:var(--accent)!important;
    background:var(--surface2)!important;
}
[data-testid="stExpander"]{
    background:var(--surface)!important;
    border:1px solid var(--border)!important;
    border-radius:12px!important;
}
[data-testid="stExpander"] summary{color:var(--text)!important;}
[data-testid="stAlert"]{
    background:var(--surface)!important;
    border-radius:12px!important;
}
[data-testid="stCaptionContainer"]{color:var(--muted)!important;}
#MainMenu,footer{visibility:hidden;}

::-webkit-scrollbar{width:9px;height:9px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--surface3);border-radius:8px;}
::-webkit-scrollbar-thumb:hover{background:var(--accent);}
</style>
"""
st.markdown(CSS_THEME, unsafe_allow_html=True)

# ==========================================================
# Helpers
# ==========================================================
def find_column(df, candidates):
    cols_lower = {str(c).lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in cols_lower:
            return cols_lower[cand.lower().strip()]
    return None


def normalize_text(x):
    return str(x).strip().replace("  ", " ") if pd.notna(x) else ""


def read_uploaded(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(uploaded.getvalue()))
    return pd.read_excel(io.BytesIO(uploaded.getvalue()))


def read_bytes(data, filename):
    if filename.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    return pd.read_excel(io.BytesIO(data))


def page_header(eyebrow, title, subtitle):
    st.markdown(f'<div class="page-eyebrow">{eyebrow}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)


def chart_card(title, fig):
    with st.container(border=True):
        st.markdown(f'<div class="chart-title">{title}</div>', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)


def style_fig(fig, height=390):
    fig.update_layout(**PLOTLY_LAYOUT, height=height)
    return fig


# ==========================================================
# Model
# ==========================================================
@st.cache_resource(show_spinner="جاري تحميل الموديل من Hugging Face...")
def load_model():
    tokenizer = AutoTokenizer.from_pretrained(MODEL_REPO)
    model = AutoModelForSequenceClassification.from_pretrained(MODEL_REPO)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model.to(device)
    model.eval()
    return tokenizer, model, device


def predict_batch(texts, tokenizer, model, device, batch_size=16):
    preds_all, conf_all = [], []
    progress = st.progress(0, text="جاري التصنيف...")
    total = len(texts)
    if total == 0:
        progress.empty()
        return [], []

    for i in range(0, total, batch_size):
        batch = [
            str(x) if pd.notna(x) and str(x).strip() else ""
            for x in texts[i:i + batch_size]
        ]
        inputs = tokenizer(
            batch,
            return_tensors="pt",
            truncation=True,
            padding=True,
            max_length=MAX_LENGTH,
        ).to(device)
        with torch.no_grad():
            probs = torch.softmax(model(**inputs).logits, dim=1)
            preds = torch.argmax(probs, dim=1)
            conf = torch.max(probs, dim=1).values

        preds_all.extend(preds.cpu().tolist())
        conf_all.extend(conf.cpu().tolist())
        done = min(i + batch_size, total)
        progress.progress(done / total, text=f"جاري التصنيف... ({done}/{total})")

    progress.empty()
    return preds_all, conf_all


# ==========================================================
# Wasted time
# ==========================================================
def subtract_break_overlap(prev_time, curr_time, break_start, break_end, gap_minutes):
    if break_start is None or break_end is None or pd.isna(prev_time) or pd.isna(curr_time):
        return gap_minutes
    day = prev_time.date()
    bs = pd.Timestamp.combine(day, break_start)
    be = pd.Timestamp.combine(day, break_end)
    overlap = max(
        (min(curr_time, be) - max(prev_time, bs)).total_seconds() / 60,
        0,
    )
    return max(gap_minutes - overlap, 0)


def calculate_wasted_time(df, sales_col, time_col, break_start, break_end):
    work = df.copy()
    work["_orig_idx"] = work.index
    work[time_col] = pd.to_datetime(work[time_col], errors="coerce")
    work = work.sort_values([sales_col, time_col])
    work["_prev_time"] = work.groupby(sales_col)[time_col].shift(1)

    def calc(row):
        if pd.isna(row["_prev_time"]) or pd.isna(row[time_col]):
            return 0.0
        gap = (row[time_col] - row["_prev_time"]).total_seconds() / 60
        gap = subtract_break_overlap(
            row["_prev_time"], row[time_col],
            break_start, break_end, gap
        )
        return round(max(gap, 0), 1)

    work[WASTED_TIME_COL] = work.apply(calc, axis=1)
    work = work.set_index("_orig_idx").sort_index()
    df[WASTED_TIME_COL] = work[WASTED_TIME_COL]
    return df


# ==========================================================
# Dashboard calculations
# ==========================================================
def classify_label(v):
    try:
        return int(float(v))
    except Exception:
        return None


def covered_mask(df):
    """المكالمة المغطاة = ليست لا يرد وليست مغلق، مع دعم اختلاف الكتابة."""
    if "Sub State" not in df.columns:
        return pd.Series(True, index=df.index)
    s = df["Sub State"].fillna("").astype(str).str.strip().str.lower()
    not_answered = s.str.contains(r"لا\s*يرد|لايرد|no\s*answer|not\s*answer", regex=True, na=False)
    closed = s.str.contains(r"مغلق|closed", regex=True, na=False)
    return ~(not_answered | closed)


def prepare_dashboard_df(df, sales_col, time_col, substate_col, class_col):
    out = df.copy()
    if sales_col and sales_col in out.columns:
        out[sales_col] = out[sales_col].fillna("غير محدد").astype(str).str.strip()
    if substate_col and substate_col in out.columns:
        out["Sub State"] = out[substate_col].fillna("غير محدد").astype(str).str.strip()
    if time_col and time_col in out.columns:
        out["_datetime"] = pd.to_datetime(out[time_col], errors="coerce")
    if class_col and class_col in out.columns:
        out["_class"] = pd.to_numeric(out[class_col], errors="coerce")
    return out


def metric_box(label, value, sub=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_metrics(df, sales_col, class_col):
    covered = covered_mask(df)
    covered_count = int(covered.sum())
    success_count = int((df["_class"] == 1).sum()) if class_col and "_class" in df else 0
    success_rate = round(success_count / covered_count * 100, 1) if covered_count else 0
    agents = df[sales_col].nunique() if sales_col and sales_col in df else 0
    wasted = round(df[WASTED_TIME_COL].sum(), 1) if WASTED_TIME_COL in df else 0

    cols = st.columns(5)
    with cols[0]:
        metric_box("👥 عدد المحصلين", f"{agents:,}", "محصل نشط في البيانات")
    with cols[1]:
        metric_box("⏱️ الوقت المهدر", f"{wasted:,.1f}", "دقيقة")
    with cols[2]:
        metric_box("📞 المكالمات المغطاة", f"{covered_count:,}", "باستثناء لا يرد / مغلق")
    with cols[3]:
        metric_box("✅ المكالمات الناجحة", f"{success_count:,}", "تصنيف الموديل = 1")
    with cols[4]:
        metric_box("📈 نسبة النجاح", f"{success_rate:.1f}%", "ناجحة ÷ المكالمات المغطاة")


def render_dashboard_filters(df, sales_col, substate_col, time_col, class_col):
    st.markdown("### 🎛️ فلاتر الداشبورد")
    c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1, 1, 1])

    sales_values = ["الكل"] + sorted(df[sales_col].dropna().astype(str).unique().tolist()) if sales_col else ["الكل"]
    selected_sales = c1.multiselect("المحصل", sales_values, default=["الكل"])

    sub_values = ["الكل"] + sorted(df[substate_col].dropna().astype(str).unique().tolist()) if substate_col else ["الكل"]
    selected_sub = c2.multiselect("Sub State", sub_values, default=["الكل"])

    dates = pd.to_datetime(df[time_col], errors="coerce").dropna().dt.date if time_col else pd.Series(dtype=object)
    date_options = sorted(dates.unique().tolist()) if len(dates) else []
    selected_dates = c3.multiselect("التاريخ", date_options, default=[])

    start_t = c4.time_input("من الساعة", value=dt_time(0, 0))
    end_t = c5.time_input("إلى الساعة", value=dt_time(23, 59))

    c6, c7 = st.columns([1, 1])
    class_options = ["الكل", "ناجحة", "غير ناجحة"]
    selected_class = c6.multiselect("حالة المكالمة", class_options, default=["الكل"])

    if sales_col and selected_sales and "الكل" not in selected_sales:
        df = df[df[sales_col].astype(str).isin(selected_sales)]
    if substate_col and selected_sub and "الكل" not in selected_sub:
        df = df[df[substate_col].astype(str).isin(selected_sub)]
    if selected_dates and time_col:
        df = df[pd.to_datetime(df[time_col], errors="coerce").dt.date.isin(selected_dates)]

    if time_col:
        dt = pd.to_datetime(df[time_col], errors="coerce")
        tm = dt.dt.time
        if start_t <= end_t:
            df = df[(tm >= start_t) & (tm <= end_t)]
        else:
            df = df[(tm >= start_t) | (tm <= end_t)]

    if class_col and selected_class and "الكل" not in selected_class:
        wanted = [1 if x == "ناجحة" else 0 for x in selected_class]
        df = df[df["_class"].isin(wanted)]

    return df


def render_activity_charts(df, sales_col, substate_col, time_col, class_col):
    # 1) نشاط المحصلين
    if sales_col:
        activity = df.groupby(sales_col).size().reset_index(name="عدد الإفادات")
        activity = activity.sort_values("عدد الإفادات", ascending=True)
        fig = px.bar(activity, x="عدد الإفادات", y=sales_col, orientation="h",
                     color="عدد الإفادات", color_continuous_scale=[COLOR_BLUE, COLOR_ACCENT])
        fig.update_layout(coloraxis_showscale=False, xaxis_title="عدد الإفادات", yaxis_title="")
        chart_card("👥 نشاط المحصلين — عدد الإفادات", style_fig(fig, max(360, len(activity) * 28)))

    # 2) لا يرد ومغلق لكل محصل
    if sales_col and substate_col:
        temp = df.copy()
        s = temp[substate_col].fillna("").astype(str)
        temp["الحالة المختصرة"] = "أخرى"
        temp.loc[s.str.contains(r"لا\s*يرد|لايرد|no\s*answer", case=False, regex=True, na=False), "الحالة المختصرة"] = "لا يرد"
        temp.loc[s.str.contains(r"مغلق|closed", case=False, regex=True, na=False), "الحالة المختصرة"] = "مغلق"
        temp = temp[temp["الحالة المختصرة"].isin(["لا يرد", "مغلق"])]
        if len(temp):
            g = temp.groupby([sales_col, "الحالة المختصرة"]).size().reset_index(name="العدد")
            fig = px.bar(g, x=sales_col, y="العدد", color="الحالة المختصرة",
                         barmode="group", color_discrete_map={"لا يرد": COLOR_WARN, "مغلق": COLOR_FAIL})
            fig.update_layout(xaxis_title="", yaxis_title="عدد الإفادات", legend_title="")
            chart_card("📞 لا يرد ومغلق لكل محصل", style_fig(fig, 410))

    # 3) الوقت المهدر لكل محصل
    if sales_col and WASTED_TIME_COL in df.columns:
        g = df.groupby(sales_col)[WASTED_TIME_COL].sum().reset_index()
        g = g.sort_values(WASTED_TIME_COL, ascending=True)
        fig = px.bar(g, x=WASTED_TIME_COL, y=sales_col, orientation="h",
                     color=WASTED_TIME_COL,
                     color_continuous_scale=[COLOR_ACCENT, COLOR_WARN, COLOR_FAIL])
        fig.update_layout(coloraxis_showscale=False, xaxis_title="دقيقة", yaxis_title="")
        chart_card("⏱️ الوقت المهدر لكل محصل", style_fig(fig, max(360, len(g) * 28)))

    # 4) نشاط المحصلين خلال اليوم
    if sales_col and time_col:
        t = df.copy()
        t["_hour"] = pd.to_datetime(t[time_col], errors="coerce").dt.hour
        t = t.dropna(subset=["_hour"])
        if len(t):
            hourly = t.groupby([sales_col, "_hour"]).size().reset_index(name="عدد الإفادات")
            fig = px.line(hourly, x="_hour", y="عدد الإفادات", color=sales_col,
                          markers=True)
            fig.update_layout(xaxis_title="ساعة اليوم", yaxis_title="عدد الإفادات", legend_title="المحصل")
            chart_card("🕐 نشاط المحصلين على مدار اليوم", style_fig(fig, 430))

    # 5) ناجحة / غير ناجحة
    if class_col and "_class" in df.columns:
        temp = df["_class"].map({1: "ناجحة", 0: "غير ناجحة"}).value_counts().reset_index()
        temp.columns = ["الحالة", "العدد"]
        fig = px.pie(temp, names="الحالة", values="العدد", hole=.62,
                     color="الحالة",
                     color_discrete_map={"ناجحة": COLOR_SUCCESS, "غير ناجحة": COLOR_FAIL})
        fig.update_traces(textinfo="percent+label")
        chart_card("🎯 توزيع المكالمات الناجحة وغير الناجحة", style_fig(fig, 390))


def build_dashboard_html(df, source_name=""):
    """تصدير Snapshot تفاعلي بسيط للداشبورد كصفحة HTML."""
    total = len(df)
    covered = int(covered_mask(df).sum())
    success = int((df["_class"] == 1).sum()) if "_class" in df.columns else 0
    rate = round(success / covered * 100, 1) if covered else 0
    agents = df["_sales"].nunique() if "_sales" in df.columns else 0
    wasted = round(df[WASTED_TIME_COL].sum(), 1) if WASTED_TIME_COL in df.columns else 0

    charts = []

    if "_sales" in df.columns:
        g = df.groupby("_sales").size().reset_index(name="عدد الإفادات")
        fig = px.bar(g.sort_values("عدد الإفادات"), x="عدد الإفادات", y="_sales", orientation="h",
                     color="عدد الإفادات", color_continuous_scale=[COLOR_BLUE, COLOR_ACCENT])
        fig.update_layout(coloraxis_showscale=False, xaxis_title="عدد الإفادات", yaxis_title="")
        charts.append(("👥 نشاط المحصلين", pio.to_html(style_fig(fig, 450), full_html=False, include_plotlyjs=False)))

    if "_sales" in df.columns and WASTED_TIME_COL in df.columns:
        g = df.groupby("_sales")[WASTED_TIME_COL].sum().reset_index()
        fig = px.bar(g.sort_values(WASTED_TIME_COL), x=WASTED_TIME_COL, y="_sales", orientation="h",
                     color=WASTED_TIME_COL,
                     color_continuous_scale=[COLOR_ACCENT, COLOR_WARN, COLOR_FAIL])
        fig.update_layout(coloraxis_showscale=False, xaxis_title="دقيقة", yaxis_title="")
        charts.append(("⏱️ الوقت المهدر", pio.to_html(style_fig(fig, 450), full_html=False, include_plotlyjs=False)))

    if "_datetime" in df.columns:
        h = df.dropna(subset=["_datetime"]).copy()
        h["ساعة"] = h["_datetime"].dt.hour
        h = h.groupby("ساعة").size().reset_index(name="عدد الإفادات")
        fig = px.line(h, x="ساعة", y="عدد الإفادات", markers=True)
        fig.update_layout(xaxis_title="ساعة اليوم", yaxis_title="عدد الإفادات")
        charts.append(("🕐 النشاط خلال اليوم", pio.to_html(style_fig(fig, 450), full_html=False, include_plotlyjs=False)))

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    cards = f"""
    <div class="cards">
      <div><span>👥 المحصلين</span><b>{agents:,}</b></div>
      <div><span>⏱️ الوقت المهدر</span><b>{wasted:,.1f}</b><small>دقيقة</small></div>
      <div><span>📞 المكالمات المغطاة</span><b>{covered:,}</b></div>
      <div><span>✅ الناجحة</span><b>{success:,}</b></div>
      <div><span>📈 نسبة النجاح</span><b>{rate:.1f}%</b></div>
    </div>
    """

    chart_html = "".join(
        f'<section><h3>{title}</h3>{html}</section>' for title, html in charts
    )

    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>7ouda Model - Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
body{{margin:0;background:#080D17;color:#F1F5F9;font-family:Tajawal,Arial,sans-serif}}
.wrap{{max-width:1400px;margin:auto;padding:35px 22px}}
h1{{margin:0;font-size:30px}} .sub{{color:#94A3B8;margin:8px 0 25px}}
.cards{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin:20px 0}}
.cards>div,section{{background:#111A2A;border:1px solid rgba(255,255,255,.07);border-radius:16px;padding:18px}}
.cards span{{display:block;color:#94A3B8;font-size:14px}} .cards b{{display:block;color:#5EEAD4;font-size:27px;margin-top:8px}}
.cards small{{color:#94A3B8}} .charts{{display:grid;grid-template-columns:1fr 1fr;gap:15px}}
section{{overflow:hidden}} h3{{margin:0 0 8px;font-size:17px}}
@media(max-width:900px){{.cards,.charts{{grid-template-columns:1fr}}}}
</style></head>
<body><div class="wrap">
<h1>📊 داشبورد نشاط المحصلين</h1>
<div class="sub">مصدر البيانات: {source_name} · تم الإنشاء {generated}</div>
{cards}
<div class="charts">{chart_html}</div>
</div></body></html>"""


# ==========================================================
# تصدير Excel
# ==========================================================
def dataframe_to_excel(df, sheet_name="النتائج"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


def pending_to_excel(filtered_df):
    """ملف الوعود القائمة + Pivot Summary في شيت منفصل."""
    buffer = io.BytesIO()
    pivot = pd.pivot_table(
        filtered_df,
        index=["Salesperson", "Account Number"],
        values="Net Amount",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="الوعود القائمة")
        pivot.to_excel(writer, index=False, sheet_name="Pivot Table")

        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="17243A")
                cell.alignment = Alignment(horizontal="center")
            for col in range(1, ws.max_column + 1):
                letter = get_column_letter(col)
                max_len = max(
                    len(str(ws.cell(row=r, column=col).value or ""))
                    for r in range(1, min(ws.max_row, 200) + 1)
                )
                ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 35)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    return buffer.getvalue()


# ==========================================================
# Page: Classification
# ==========================================================
def page_classification():
    page_header(
        "CALL QUALITY CLASSIFIER",
        "🎯 تصنيف المكالمات",
        "ارفع الملف وسيتم تصنيف الإفادات: 1 = ناجحة، 0 = غير ناجحة، مع حساب الوقت المهدر.",
    )

    with st.expander("⚙️ إعدادات البريك", expanded=False):
        c1, c2 = st.columns(2)
        break_start = c1.time_input("بداية البريك", value=dt_time(13, 0))
        break_end = c2.time_input("نهاية البريك", value=dt_time(13, 30))

    uploaded = st.file_uploader(
        "ارفع ملف البيانات (CSV / Excel)",
        type=["csv", "xlsx", "xls"],
        key="classification_upload",
    )

    if uploaded is not None:
        st.session_state["raw_file_bytes"] = uploaded.getvalue()
        st.session_state["raw_file_name"] = uploaded.name
        st.session_state["raw_file_size"] = uploaded.size

    if "raw_file_bytes" not in st.session_state:
        st.markdown('<div class="card">📂 ارفع ملفًا للبدء.</div>', unsafe_allow_html=True)
        return

    name = st.session_state["raw_file_name"]
    try:
        df = read_bytes(st.session_state["raw_file_bytes"], name)
    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        return

    # إزالة أول صف بعد العناوين
    if len(df):
        df = df.iloc[1:].reset_index(drop=True)

    if ORIGINAL_TEXT_COL in df.columns:
        df = df.rename(columns={ORIGINAL_TEXT_COL: MODEL_TEXT_COL})

    if MODEL_TEXT_COL not in df.columns:
        st.error(f"الملف يجب أن يحتوي على '{ORIGINAL_TEXT_COL}' أو '{MODEL_TEXT_COL}'.")
        return

    with st.expander("👀 معاينة البيانات", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)

    token = f"{name}_{st.session_state['raw_file_size']}"

    if st.button("🚀 ابدأ التصنيف", type="primary", use_container_width=True):
        tokenizer, model, device = load_model()
        preds, confs = predict_batch(df[MODEL_TEXT_COL].tolist(), tokenizer, model, device)

        result = df.copy()
        result[CLASSIFICATION_COL] = preds
        result["نسبة_الثقة"] = [round(x * 100, 1) for x in confs]

        sales_col = find_column(result, SALES_PERSON_CANDIDATES)
        time_col = find_column(result, CREATED_ON_CANDIDATES)

        if sales_col and time_col:
            result = calculate_wasted_time(
                result, sales_col, time_col, break_start, break_end
            )
        else:
            st.warning("لم أجد عمود المحصل أو التاريخ/الوقت؛ لن يتم حساب الوقت المهدر.")

        result = result.rename(columns={MODEL_TEXT_COL: ORIGINAL_TEXT_COL})
        st.session_state["last_result_df"] = result
        st.session_state["last_sales_col"] = sales_col
        st.session_state["last_time_col"] = time_col
        st.session_state["last_file_token"] = token

    if st.session_state.get("last_file_token") != token:
        return

    result = st.session_state["last_result_df"]
    st.success("تم التصنيف وحساب الوقت المهدر بنجاح ✅")
    st.dataframe(result, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ تحميل نتائج التصنيف Excel",
            dataframe_to_excel(result, "النتائج"),
            "نتائج_التصنيف.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "⬇️ تحميل النتائج CSV",
            result.to_csv(index=False).encode("utf-8-sig"),
            "نتائج_التصنيف.csv",
            "text/csv",
            use_container_width=True,
        )


# ==========================================================
# Page: Dashboard
# ==========================================================
def page_dashboard():
    page_header(
        "ACTIVITY DASHBOARD",
        "📊 داشبورد نشاط المحصلين",
        "داشبورد تفاعلي بالفلاتر والكروت والشارتس الخاصة بالنشاط والنجاح والوقت المهدر.",
    )

    uploaded = st.file_uploader(
        "ارفع الملف المصنّف",
        type=["csv", "xlsx", "xls"],
        key="dash_upload",
    )
    if uploaded is not None:
        st.session_state["dash_raw_bytes"] = uploaded.getvalue()
        st.session_state["dash_raw_name"] = uploaded.name

    if "dash_raw_bytes" not in st.session_state:
        st.markdown('<div class="card">📂 ارفع ملف التصنيف لعرض الداشبورد.</div>', unsafe_allow_html=True)
        return

    try:
        df = read_bytes(
            st.session_state["dash_raw_bytes"],
            st.session_state["dash_raw_name"],
        )
    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        return

    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    time_col = find_column(df, CREATED_ON_CANDIDATES)
    substate_col = find_column(df, SUB_STATE_CANDIDATES)

    if not class_col:
        st.warning("لم أجد عمود التصنيف. اختره يدويًا.")
    cols = ["— بدون —"] + list(df.columns.astype(str))

    with st.expander("⚙️ تأكيد الأعمدة", expanded=not all([class_col, sales_col, time_col])):
        c1, c2, c3, c4 = st.columns(4)
        class_sel = c1.selectbox(
            "التصنيف (1/0)", cols,
            index=cols.index(class_col) if class_col in cols else 0,
        )
        sales_sel = c2.selectbox(
            "المحصل", cols,
            index=cols.index(sales_col) if sales_col in cols else 0,
        )
        time_sel = c3.selectbox(
            "التاريخ والوقت", cols,
            index=cols.index(time_col) if time_col in cols else 0,
        )
        sub_sel = c4.selectbox(
            "Sub State", cols,
            index=cols.index(substate_col) if substate_col in cols else 0,
        )

    class_col = None if class_sel == "— بدون —" else class_sel
    sales_col = None if sales_sel == "— بدون —" else sales_sel
    time_col = None if time_sel == "— بدون —" else time_sel
    substate_col = None if sub_sel == "— بدون —" else sub_sel

    work = prepare_dashboard_df(df, sales_col, time_col, substate_col, class_col)

    # أسماء موحدة للحسابات
    if substate_col:
        work["Sub State"] = work[substate_col]
    if sales_col:
        work["_sales"] = work[sales_col]
    if time_col:
        work["_datetime"] = pd.to_datetime(work[time_col], errors="coerce")
    if class_col:
        work["_class"] = pd.to_numeric(work[class_col], errors="coerce")

    filtered = render_dashboard_filters(
        work, sales_col, substate_col, time_col, class_col
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    render_dashboard_metrics(filtered, sales_col, class_col)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    render_activity_charts(
        filtered, sales_col, substate_col, time_col, class_col
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.download_button(
        "🌐 تحميل Snapshot للداشبورد HTML",
        build_dashboard_html(
            filtered.assign(
                _sales=filtered[sales_col] if sales_col else "",
                _datetime=pd.to_datetime(filtered[time_col], errors="coerce") if time_col else pd.NaT,
            ),
            st.session_state["dash_raw_name"],
        ).encode("utf-8"),
        "داشبورد_النشاط.html",
        "text/html",
        use_container_width=True,
        type="primary",
    )


# ==========================================================
# Page: Pending Promises
# ==========================================================
def page_pending_promises():
    page_header(
        "PENDING PROMISES",
        "📗 الوعود القائمة",
        "ارفع ملف المحفظة، وسيتم استخراج الوعود القائمة لليوم الحالي وتجهيز Excel + Pivot Summary.",
    )

    uploaded = st.file_uploader(
        "ارفع ملف المحفظة (Excel / CSV)",
        type=["xlsx", "xls", "csv"],
        key="pending_upload",
    )
    if uploaded is None:
        st.markdown('<div class="card">📂 ارفع ملف المحفظة للبدء.</div>', unsafe_allow_html=True)
        return

    try:
        wallet = read_uploaded(uploaded)
    except Exception as e:
        st.error(f"خطأ في قراءة ملف المحفظة: {e}")
        return

    salesperson_col = find_column(wallet, ["Salesperson", "Sales Person", "salesperson"])
    substate_col = find_column(wallet, SUB_STATE_CANDIDATES)
    due_col = find_column(wallet, FOLLOW_UP_CANDIDATES)
    account_col = find_column(wallet, ACCOUNT_CANDIDATES)
    amount_col = find_column(wallet, NET_AMOUNT_CANDIDATES)

    missing = []
    if not salesperson_col: missing.append("Salesperson")
    if not substate_col: missing.append("Sub State")
    if not due_col: missing.append("Follow up Due Date")
    if not account_col: missing.append("Account Number")
    if not amount_col: missing.append("Net Amount")

    if missing:
        st.error("الأعمدة التالية غير موجودة: " + "، ".join(missing))
        st.info("الأعمدة الموجودة: " + "، ".join(map(str, wallet.columns)))
        return

    # توحيد الأسماء المطلوبة
    work = wallet.copy()
    work["Salesperson"] = work[salesperson_col].fillna("").astype(str).str.strip()
    work["Sub State"] = work[substate_col].fillna("").astype(str).str.strip()
    work["Follow up Due Date"] = pd.to_datetime(work[due_col], errors="coerce")
    work["Account Number"] = work[account_col]
    work["Net Amount"] = pd.to_numeric(
        work[amount_col].astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    # 1) استبعاد المحصلين المحددين
    excluded_norm = {x.strip().lower() for x in EXCLUDED_SALESPERSONS}
    mask_sales = ~work["Salesperson"].str.lower().isin(excluded_norm)

    # 2) واعد بالسداد فقط
    mask_promise = work["Sub State"].str.strip().str.lower() == "واعد بالسداد"

    # 3) تاريخ اليوم
    today = pd.Timestamp.today().normalize()
    mask_today = work["Follow up Due Date"].dt.normalize() == today

    result = work[mask_sales & mask_promise & mask_today].copy()

    # الاحتفاظ بكل أعمدة المحفظة، مع ضمان الأعمدة المطلوبة في أول الملف
    priority = ["Salesperson", "Account Number", "Net Amount", "Sub State", "Follow up Due Date"]
    others = [c for c in result.columns if c not in priority]
    result = result[priority + others]

    st.markdown("### 📌 نتيجة الوعود القائمة")
    c1, c2, c3 = st.columns(3)
    c1.metric("عدد الوعود", f"{len(result):,}")
    c2.metric("عدد المحصلين", f"{result['Salesperson'].nunique():,}")
    c3.metric("إجمالي Net Amount", f"{result['Net Amount'].sum():,.2f}")

    if len(result):
        st.dataframe(result, use_container_width=True, hide_index=True)

        excel_bytes = pending_to_excel(result)
        st.download_button(
            "⬇️ تحميل الوعود القائمة + Pivot Table",
            excel_bytes,
            "الوعود_القائمة.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        pivot_preview = pd.pivot_table(
            result,
            index=["Salesperson", "Account Number"],
            values="Net Amount",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        st.markdown("### 📊 Pivot Summary")
        st.dataframe(pivot_preview, use_container_width=True, hide_index=True)
    else:
        st.warning(f"لا توجد وعود قائمة مطابقة لشروط اليوم ({today.date()}).")


# ==========================================================
# Placeholders
# ==========================================================
def page_placeholder(eyebrow, title, subtitle, icon):
    page_header(eyebrow, f"{icon} {title}", subtitle)
    st.markdown(
        f'<div class="card" style="text-align:center;padding:3rem;">'
        f'{icon}<br><br>التويب ده جاهز للمنطق القادم.</div>',
        unsafe_allow_html=True,
    )


# ==========================================================
# Navigation
# ==========================================================
PAGES = {
    "🎯 التصنيف": page_classification,
    "📗 الوعود القائمة": page_pending_promises,
    "📕 الوعود المكسورة": lambda: page_placeholder(
        "BROKEN", "الوعود المكسورة", "الوعود التي لم يتم الالتزام بها.", "📕"
    ),
    "⚠️ الإهمال": lambda: page_placeholder(
        "NEGLECT", "الإهمال", "حالات الإهمال في المتابعة.", "⚠️"
    ),
    "🧾 أخطاء الحالات": lambda: page_placeholder(
        "CASE ERRORS", "أخطاء الحالات", "الحالات التي تحتاج مراجعة.", "🧾"
    ),
    "📊 الداشبورد": page_dashboard,
}

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="mini">7OUDA MODEL</div>
            <div class="big">🎙️ لوحة تحليل المكالمات</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    selected_page = st.radio(
        "التنقل",
        list(PAGES.keys()),
        label_visibility="collapsed",
    )

PAGES[selected_page]()
