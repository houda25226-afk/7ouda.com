"""
تطبيق Streamlit لتصنيف إفادات المكالمات + حساب الوقت المهدر + داشبورد
باستخدام الموديل المرفوع على Hugging Face: Mahmoud252002/7oudaModel

طريقة التشغيل محليًا:
    pip install -r requirements.txt --break-system-packages
    streamlit run app.py

============================================================
دليل التخصيص السريع:
- الألوان والخطوط في CSS_THEME تحت.
- كل تويب (صفحة) في دالة منفصلة اسمها page_xxx() — سهل تلاقي مكانك.
- أسماء الأعمدة المتوقعة في الملف (Create By, Created On, Notes)
  متعرّفة في قسم "إعدادات وأسماء الأعمدة" تحت — لو الأسماء اتغيرت غيّرها من هناك.
- التبويبات الأربعة (الوعود القائمة / المكسورة / الإهمال / أخطاء الحالات)
  لسه فاضية (Placeholder) لحد ما نحدد منطق كل واحدة فيها.
============================================================
"""

import io
from datetime import time as dt_time
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# إعدادات وأسماء الأعمدة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

ORIGINAL_TEXT_COL = "Notes"          # اسم العمود الأصلي في الملف
MODEL_TEXT_COL = "الافادة"          # الاسم اللي بيتحول له مؤقتًا عشان الموديل
CLASSIFICATION_COL = "التصنيف"      # عمود النتيجة: 1 = ناجحة / 0 = غير ناجحة
WASTED_TIME_COL = "الوقت_المهدر_دقيقة"

SALES_PERSON_CANDIDATES = ["Create By", "create by", "CreateBy", "Created By", "created by", "Sales Person", "sales person", "المحصل"]
CREATED_ON_CANDIDATES = ["Created On", "created on", "CreatedOn", "تاريخ الافادة"]

st.set_page_config(
    page_title="لوحة تحليل المكالمات | 7oudaModel",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================
# الهوية البصرية (Theme)
# ==========================================================

CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0E1420;
    --surface: #151F30;
    --surface-2: #1B2A42;
    --accent: #5EEAD4;
    --success: #34D399;
    --danger: #FB7185;
    --warn: #FBBF24;
    --text: #E7ECF3;
    --text-dim: #8B96AC;
}

html, body, [class*="css"] {
    font-family: 'Tajawal', sans-serif;
}

.stApp {
    background: radial-gradient(circle at 20% 0%, #10192C 0%, var(--bg) 55%);
    color: var(--text);
}

.main .block-container {
    direction: rtl;
    text-align: right;
    max-width: 1200px;
}
[data-testid="stSidebarContent"], [data-testid="stSidebarUserContent"] {
    direction: rtl;
    text-align: right;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

section[data-testid="stSidebar"] {
    background: #0B111C;
    border-left: none;
    border-right: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: var(--text) !important; }

.sidebar-brand {
    text-align: center;
    padding: 0.6rem 0 1rem 0;
}
.sidebar-brand .title {
    font-weight: 900;
    font-size: 1.15rem;
    color: var(--text);
}
.sidebar-brand .subtitle {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    letter-spacing: 0.12em;
    color: var(--accent);
    text-transform: uppercase;
}

section[data-testid="stSidebar"] .stRadio > div {
    gap: 4px;
}
section[data-testid="stSidebar"] .stRadio label {
    background: var(--surface);
    border-radius: 10px;
    padding: 0.55rem 0.8rem !important;
    margin-bottom: 2px;
    width: 100%;
    border: 1px solid rgba(255,255,255,0.05);
    transition: background 0.15s ease;
}
section[data-testid="stSidebar"] .stRadio label:hover {
    background: var(--surface-2);
}
section[data-testid="stSidebar"] .stRadio input:checked + div {
    color: var(--accent) !important;
}

.page-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.14em;
    color: var(--accent);
    text-transform: uppercase;
    margin-bottom: 0.2rem;
}
.page-title {
    font-weight: 900;
    font-size: 1.75rem;
    margin: 0;
}
.page-subtitle {
    color: var(--text-dim);
    font-size: 0.95rem;
    margin-top: 0.3rem;
}

.waveform {
    display: flex; align-items: center; justify-content: center;
    gap: 4px; height: 30px; margin: 0.8rem 0 1.3rem 0;
}
.waveform span {
    display: inline-block; width: 3px; border-radius: 3px;
    background: linear-gradient(180deg, var(--accent), var(--surface-2));
    animation: wave 1.6s ease-in-out infinite;
}
@keyframes wave {
    0%, 100% { transform: scaleY(0.35); opacity: 0.55; }
    50% { transform: scaleY(1); opacity: 1; }
}

.card {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 14px;
    padding: 1.1rem 1.3rem;
    margin-bottom: 1rem;
}
.placeholder-card {
    background: var(--surface);
    border: 1.5px dashed rgba(255,255,255,0.12);
    border-radius: 14px;
    padding: 2.4rem 1.5rem;
    text-align: center;
    color: var(--text-dim);
}

[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(94, 234, 212, 0.35);
    border-radius: 14px;
    padding: 0.6rem;
}
[data-testid="stFileUploader"] section { background: transparent; }

[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] p {
    color: var(--text) !important;
    opacity: 1 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg {
    fill: var(--text-dim) !important;
}

[data-testid="stFileUploader"] button,
[data-testid="stFileUploaderDropzone"] button {
    background: var(--surface-2) !important;
    color: var(--text) !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    font-weight: 700 !important;
}
[data-testid="stFileUploader"] button:hover {
    border-color: var(--accent) !important;
    color: var(--accent) !important;
}

[data-testid="stFileUploaderFile"] {
    background: var(--surface-2) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploaderFile"] div, [data-testid="stFileUploaderFile"] span {
    color: var(--text) !important;
}
[data-testid="stFileUploaderFileName"] { color: var(--text) !important; }
[data-testid="stFileUploaderFileErrorMessage"] { color: var(--danger) !important; }

[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] span {
    color: var(--text) !important;
    opacity: 1 !important;
}

.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), #3FD9C7);
    color: #06251F;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(94, 234, 212, 0.25);
    color: #06251F;
}

[data-testid="stProgress"] > div > div { background: var(--accent); }

[data-testid="stMetric"] {
    background: var(--surface);
    border-radius: 12px;
    padding: 0.8rem 1rem;
    border: 1px solid rgba(255,255,255,0.06);
}
[data-testid="stMetricValue"] { font-family: 'JetBrains Mono', monospace; color: var(--accent); }

[data-testid="stDataFrame"] {
    border-radius: 12px; overflow: hidden;
    border: 1px solid rgba(255,255,255,0.06);
}

.divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(94,234,212,0.35), transparent);
    margin: 1.4rem 0;
    border: none;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--surface) !important;
    border-radius: 14px !important;
    border: 1px solid rgba(255,255,255,0.07) !important;
}
.chart-card-title {
    font-weight: 700;
    font-size: 0.95rem;
    color: var(--text);
    margin-bottom: 0.4rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid rgba(255,255,255,0.06);
}

.highlight-label {
    font-size: 0.82rem;
    color: var(--text-dim);
    margin-bottom: 0.5rem;
}
.highlight-value {
    font-weight: 900;
    font-size: 1.25rem;
    color: var(--accent);
    line-height: 1.3;
}
.highlight-sub {
    font-size: 0.8rem;
    color: var(--text-dim);
    margin-top: 0.2rem;
}

[data-testid="stExpander"] {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 12px;
    overflow: hidden;
}
[data-testid="stExpander"] summary {
    background: var(--surface-2);
    color: var(--text) !important;
}

div[data-baseweb="select"] > div {
    background: var(--surface-2) !important;
    border-color: rgba(255,255,255,0.1) !important;
    color: var(--text) !important;
    border-radius: 10px !important;
}
div[data-baseweb="popover"] ul {
    background: var(--surface-2) !important;
}
li[role="option"] { color: var(--text) !important; }
li[role="option"]:hover, li[aria-selected="true"] { background: var(--surface) !important; }

[data-testid="stTextInput"] input,
[data-testid="stTimeInput"] input,
[data-testid="stNumberInput"] input {
    background: var(--surface-2) !important;
    color: var(--text) !important;
    border-color: rgba(255,255,255,0.1) !important;
    border-radius: 10px !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
}
.stTabs [data-baseweb="tab"] {
    background: var(--surface);
    border-radius: 10px 10px 0 0;
    color: var(--text-dim);
    padding: 0.5rem 1rem;
}
.stTabs [aria-selected="true"] {
    background: var(--surface-2) !important;
    color: var(--accent) !important;
    font-weight: 700;
}

div[data-testid="stAlert"] {
    background: var(--surface) !important;
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    color: var(--text) !important;
}
div[data-testid="stAlert"] p { color: var(--text) !important; }

[data-testid="stCaptionContainer"] { color: var(--text-dim) !important; }

::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--surface-2); border-radius: 8px; }
::-webkit-scrollbar-thumb:hover { background: var(--accent); }
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


def render_waveform(n_bars: int = 24):
    heights = [14, 22, 30, 18, 26, 30, 20, 26, 16, 22] * (n_bars // 10 + 1)
    bars = "".join(
        f'<span style="height:{heights[i]}px; animation-delay:{(i % 10) * 0.09}s;"></span>'
        for i in range(n_bars)
    )
    st.markdown(f'<div class="waveform">{bars}</div>', unsafe_allow_html=True)


def page_header(eyebrow: str, title: str, subtitle: str, show_wave: bool = False):
    st.markdown(
        f"""
        <div class="page-eyebrow">{eyebrow}</div>
        <p class="page-title">{title}</p>
        <p class="page-subtitle">{subtitle}</p>
        """,
        unsafe_allow_html=True,
    )
    if show_wave:
        render_waveform()
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)


def find_column(df: pd.DataFrame, candidates: list):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in cols_lower:
            return cols_lower[cand.lower().strip()]
    return None


# ==========================================================
# منطق الموديل
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
    all_preds, all_confidences = [], []
    progress_bar = st.progress(0, text="جاري التصنيف...")
    total = len(texts)

    for i in range(0, total, batch_size):
        batch = texts[i : i + batch_size]
        batch = [str(t) if pd.notna(t) and str(t).strip() != "" else "" for t in batch]

        inputs = tokenizer(
            batch, return_tensors="pt", truncation=True, padding=True, max_length=MAX_LENGTH
        ).to(device)

        with torch.no_grad():
            logits = model(**inputs).logits
            probs = torch.softmax(logits, dim=1)
            preds = torch.argmax(probs, dim=1)
            confidences = torch.max(probs, dim=1).values

        all_preds.extend(preds.cpu().tolist())
        all_confidences.extend(confidences.cpu().tolist())

        progress_bar.progress(
            min((i + batch_size) / total, 1.0),
            text=f"جاري التصنيف... ({min(i + batch_size, total)}/{total})",
        )

    progress_bar.empty()
    return all_preds, all_confidences


# ==========================================================
# حساب الوقت المهدر بين المكالمات لكل محصّل
# ==========================================================

def subtract_break_overlap(prev_time, curr_time, break_start, break_end, gap_minutes):
    """بيخصم من الفجوة أي جزء واقع جوه وقت البريك المحدد."""
    if break_start is None or break_end is None or pd.isna(prev_time) or pd.isna(curr_time):
        return gap_minutes
    day = prev_time.date()
    break_start_dt = pd.Timestamp.combine(day, break_start)
    break_end_dt = pd.Timestamp.combine(day, break_end)
    overlap_start = max(prev_time, break_start_dt)
    overlap_end = min(curr_time, break_end_dt)
    overlap_minutes = max((overlap_end - overlap_start).total_seconds() / 60, 0)
    return max(gap_minutes - overlap_minutes, 0)


def calculate_wasted_time(df, sales_col, time_col, break_start, break_end):
    """
    بتحسب الوقت المهدر (بالدقايق) بين كل مكالمة واللي قبلها لنفس المحصّل،
    بعد استبعاد وقت البريك. أول مكالمة لكل محصّل = صفر (مفيش مكالمة قبلها نقيس منها).
    """
    work = df.copy()
    work["_orig_idx"] = work.index
    work[time_col] = pd.to_datetime(work[time_col], errors="coerce")
    work = work.sort_values([sales_col, time_col])
    work["_prev_time"] = work.groupby(sales_col)[time_col].shift(1)

    def compute_row(row):
        if pd.isna(row["_prev_time"]) or pd.isna(row[time_col]):
            return 0.0
        gap_min = (row[time_col] - row["_prev_time"]).total_seconds() / 60
        gap_min = subtract_break_overlap(row["_prev_time"], row[time_col], break_start, break_end, gap_min)
        return round(max(gap_min, 0.0), 1)

    work[WASTED_TIME_COL] = work.apply(compute_row, axis=1)
    work = work.set_index("_orig_idx").sort_index()
    df[WASTED_TIME_COL] = work[WASTED_TIME_COL]
    return df


def highlight_wasted(val):
    if pd.isna(val):
        return ""
    if val > 10:
        return "background-color: #ffb3b3; color: #3a0000;"
    elif val < 1:
        return "background-color: #fff59d; color: #3a3300;"
    return ""


# ==========================================================
# داشبورد مشترك
# ==========================================================

COLOR_SUCCESS = "#34D399"   # أخضر زمردي — ناجحة
COLOR_FAIL = "#FB7185"      # وردي-أحمر — غير ناجحة
COLOR_ACCENT = "#5EEAD4"    # تركواز — لوني أساسي
COLOR_WARN = "#FBBF24"      # كهرماني — تحذيري/متوسط
CHART_COLORS = {"ناجحة": COLOR_SUCCESS, "غير ناجحة": COLOR_FAIL}
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color="#E7ECF3",
    font_family="Tajawal, sans-serif",
    margin=dict(t=42, b=10, l=10, r=10),
)
PLOTLY_CONFIG = {"displayModeBar": False}


def chart_card(title: str, render_fn):
    """كارد موحّد لأي رسم بياني — عنوان صغير + حدود + خلفية متناسقة."""
    with st.container(border=True):
        st.markdown(f'<div class="chart-card-title">{title}</div>', unsafe_allow_html=True)
        render_fn()


def _compute_agent_perf(df, class_col, sales_col):
    agent_perf = (
        df.groupby(sales_col)[class_col]
        .agg(["count", "sum"])
        .rename(columns={"count": "إجمالي المكالمات", "sum": "ناجحة"})
    )
    agent_perf["غير ناجحة"] = agent_perf["إجمالي المكالمات"] - agent_perf["ناجحة"]
    agent_perf["نسبة النجاح %"] = (agent_perf["ناجحة"] / agent_perf["إجمالي المكالمات"] * 100).round(1)
    return agent_perf.sort_values("نسبة النجاح %", ascending=False).reset_index()


def render_metric_cards(df, class_col, sales_col, time_col):
    has_class = class_col and class_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns

    with st.container(border=True):
        st.markdown('<div class="chart-card-title">📌 نظرة عامة</div>', unsafe_allow_html=True)
        metric_cols = st.columns(4)
        metric_cols[0].metric("📞 إجمالي المكالمات", len(df))

        if has_class:
            success_count = int((df[class_col] == 1).sum())
            fail_count = int((df[class_col] == 0).sum())
            success_rate = round(success_count / len(df) * 100, 1) if len(df) else 0
            metric_cols[1].metric("✅ ناجحة", success_count, f"{success_rate}%")
            metric_cols[2].metric("⛔ غير ناجحة", fail_count)
        else:
            metric_cols[1].metric("✅ ناجحة", "—")
            metric_cols[2].metric("⛔ غير ناجحة", "—")

        if has_wasted:
            avg_wasted = round(df[WASTED_TIME_COL].mean(), 1) if len(df) else 0
            metric_cols[3].metric(
                "⏱️ إجمالي الوقت المهدر (دقيقة)",
                round(df[WASTED_TIME_COL].sum(), 1),
                f"متوسط {avg_wasted} د/مكالمة",
            )
        else:
            metric_cols[3].metric("⏱️ إجمالي الوقت المهدر", "—")


def render_pie_chart(df, class_col):
    has_class = class_col and class_col in df.columns

    def _pie():
        if not has_class:
            st.info("مفيش عمود تصنيف عشان نعرض توزيع النتائج.")
            return
        labels_series = df[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
        pie_df = labels_series.value_counts().reset_index()
        pie_df.columns = ["التصنيف", "العدد"]
        success_rate = round((df[class_col] == 1).mean() * 100, 1) if len(df) else 0
        fig = px.pie(
            pie_df, names="التصنيف", values="العدد", hole=0.62,
            color="التصنيف", color_discrete_map=CHART_COLORS,
        )
        fig.update_traces(textinfo="percent", textfont_size=13, marker=dict(line=dict(color="#0E1420", width=3)))
        fig.update_layout(
            **PLOTLY_LAYOUT, showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"),
            annotations=[dict(text=f"{success_rate}%<br><span style='font-size:11px;color:#8B96AC'>نجاح</span>",
                               x=0.5, y=0.5, font_size=22, font_color=COLOR_SUCCESS, showarrow=False)],
        )
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    chart_card("🎯 توزيع نتائج التصنيف", _pie)


def render_wasted_bar(df, sales_col, top_n=10):
    has_wasted = WASTED_TIME_COL in df.columns
    has_sales = sales_col and sales_col in df.columns

    def _wasted_bar():
        if not (has_wasted and has_sales):
            st.info("محتاجين عمود المحصّل + الوقت المهدر عشان نعرض الرسم ده.")
            return
        wasted_by_agent = df.groupby(sales_col)[WASTED_TIME_COL].sum().sort_values(ascending=False)
        if top_n:
            wasted_by_agent = wasted_by_agent.head(top_n)
        wasted_by_agent = wasted_by_agent.reset_index()
        chart_height = max(300, 28 * len(wasted_by_agent))
        fig2 = px.bar(
            wasted_by_agent, x=WASTED_TIME_COL, y=sales_col, orientation="h",
            color=WASTED_TIME_COL, color_continuous_scale=[COLOR_ACCENT, COLOR_WARN, COLOR_FAIL],
        )
        fig2.update_layout(
            **PLOTLY_LAYOUT, yaxis={"categoryorder": "total ascending", "title": ""},
            xaxis_title="الوقت المهدر (دقيقة)", coloraxis_showscale=False, height=chart_height,
        )
        fig2.update_traces(marker_line_width=0)
        st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG)

    title = f"🏆 أعلى {top_n} محصّلين في الوقت المهدر" if top_n else "🏆 كل المحصّلين حسب الوقت المهدر"
    chart_card(title, _wasted_bar)


def render_agent_perf_chart(df, class_col, sales_col, with_table=True):
    has_class = class_col and class_col in df.columns
    has_sales = sales_col and sales_col in df.columns
    if not (has_class and has_sales):
        return
    agent_perf = _compute_agent_perf(df, class_col, sales_col)

    def _agent_perf():
        fig3 = px.bar(
            agent_perf, x=sales_col, y=["ناجحة", "غير ناجحة"],
            color_discrete_sequence=[COLOR_SUCCESS, COLOR_FAIL], barmode="stack",
        )
        fig3.update_layout(**PLOTLY_LAYOUT, legend_title_text="", xaxis_title="", yaxis_title="عدد المكالمات")
        fig3.update_traces(marker_line_width=0)
        st.plotly_chart(fig3, use_container_width=True, config=PLOTLY_CONFIG)
        if with_table:
            with st.expander("📋 جدول ترتيب المحصّلين حسب نسبة النجاح"):
                st.dataframe(agent_perf, use_container_width=True)

    chart_card("📊 أداء كل محصّل (ناجحة مقابل غير ناجحة)", _agent_perf)


def render_wasted_hist(df):
    has_wasted = WASTED_TIME_COL in df.columns

    def _hist():
        if not has_wasted:
            st.info("محتاجين عمود الوقت المهدر عشان نعرض التوزيع ده.")
            return
        fig4 = px.histogram(df, x=WASTED_TIME_COL, nbins=20, color_discrete_sequence=[COLOR_ACCENT])
        fig4.update_layout(**PLOTLY_LAYOUT, bargap=0.08, xaxis_title="الوقت المهدر (دقيقة)", yaxis_title="عدد المرات")
        st.plotly_chart(fig4, use_container_width=True, config=PLOTLY_CONFIG)

    chart_card("⏱️ توزيع الوقت المهدر بين المكالمات", _hist)


def render_trend_chart(df, class_col, time_col):
    has_class = class_col and class_col in df.columns
    has_time = time_col and time_col in df.columns

    def _trend():
        if not has_time:
            st.info("محتاجين عمود التاريخ عشان نعرض اتجاه المكالمات بمرور الوقت.")
            return
        trend_df = df.copy()
        trend_df[time_col] = pd.to_datetime(trend_df[time_col], errors="coerce")
        trend_df["اليوم"] = trend_df[time_col].dt.date
        if has_class:
            trend_df["الحالة"] = trend_df[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
            daily = trend_df.groupby(["اليوم", "الحالة"]).size().reset_index(name="عدد المكالمات")
            fig5 = px.area(daily, x="اليوم", y="عدد المكالمات", color="الحالة", color_discrete_map=CHART_COLORS)
        else:
            daily = trend_df.groupby("اليوم").size().reset_index(name="عدد المكالمات")
            fig5 = px.area(daily, x="اليوم", y="عدد المكالمات", color_discrete_sequence=[COLOR_ACCENT])
        fig5.update_traces(line_width=2)
        fig5.update_layout(**PLOTLY_LAYOUT, legend_title_text="", xaxis_title="", yaxis_title="عدد المكالمات")
        st.plotly_chart(fig5, use_container_width=True, config=PLOTLY_CONFIG)

    chart_card("📅 اتجاه عدد المكالمات يوميًا", _trend)


def render_quick_summary(df, class_col=None, sales_col=None, time_col=None):
    render_metric_cards(df, class_col, sales_col, time_col)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    render_highlights(df, class_col, sales_col, time_col)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        render_pie_chart(df, class_col)
    with col_b:
        render_trend_chart(df, class_col, time_col)

    col_c, col_d = st.columns(2)
    with col_c:
        render_agent_perf_chart(df, class_col, sales_col, with_table=False)
    with col_d:
        render_wasted_bar(df, sales_col)


def render_highlights(df, class_col, sales_col, time_col):
    has_class = class_col and class_col in df.columns
    has_sales = sales_col and sales_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns
    has_time = time_col and time_col in df.columns

    cols = st.columns(3)

    with cols[0]:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">🏅 أفضل محصّل (نسبة نجاح)</div>', unsafe_allow_html=True)
            if has_class and has_sales:
                perf = _compute_agent_perf(df, class_col, sales_col)
                perf = perf[perf["إجمالي المكالمات"] >= 3]
                if len(perf):
                    top = perf.iloc[0]
                    st.markdown(f'<div class="highlight-value">{top[sales_col]}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="highlight-sub">{top["نسبة النجاح %"]}% نجاح ({int(top["إجمالي المكالمات"])} مكالمة)</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="highlight-sub">مفيش بيانات كافية</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="highlight-sub">محتاجين عمود التصنيف والمحصّل</div>', unsafe_allow_html=True)

    with cols[1]:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">🐌 الأكثر إهدارًا للوقت</div>', unsafe_allow_html=True)
            if has_wasted and has_sales:
                wasted_totals = df.groupby(sales_col)[WASTED_TIME_COL].sum().sort_values(ascending=False)
                if len(wasted_totals):
                    st.markdown(f'<div class="highlight-value">{wasted_totals.index[0]}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="highlight-sub">{round(wasted_totals.iloc[0], 1)} دقيقة مهدرة</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="highlight-sub">مفيش بيانات كافية</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="highlight-sub">محتاجين عمود المحصّل والوقت المهدر</div>', unsafe_allow_html=True)

    with cols[2]:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">📅 أكثر يوم نشاطًا</div>', unsafe_allow_html=True)
            if has_time:
                t = pd.to_datetime(df[time_col], errors="coerce")
                if t.notna().any():
                    daily_counts = t.dt.date.value_counts()
                    st.markdown(f'<div class="highlight-value">{daily_counts.index[0]}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="highlight-sub">{int(daily_counts.iloc[0])} مكالمة</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="highlight-sub">مفيش تواريخ صالحة</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="highlight-sub">محتاجين عمود التاريخ</div>', unsafe_allow_html=True)


def render_comparison_matrix(df, class_col, sales_col):
    has_class = class_col and class_col in df.columns
    has_sales = sales_col and sales_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns

    if not (has_class and has_sales):
        st.info("محتاجين عمود التصنيف وعمود المحصّل عشان نبني جدول المقارنة.")
        return

    perf = _compute_agent_perf(df, class_col, sales_col)

    if has_wasted:
        wasted_agg = df.groupby(sales_col)[WASTED_TIME_COL].agg(["sum", "mean"]).round(1)
        wasted_agg.columns = ["إجمالي الوقت المهدر", "متوسط الوقت المهدر"]
        perf = perf.merge(wasted_agg, left_on=sales_col, right_index=True, how="left")

    def _table():
        st.dataframe(
            perf.rename(columns={sales_col: "المحصّل"}),
            use_container_width=True,
            hide_index=True,
        )

    chart_card(f"📋 جدول المقارنة الشامل — كل المحصّلين ({len(perf)})", _table)

    if has_wasted:
        def _scatter():
            fig = px.scatter(
                perf, x="نسبة النجاح %", y="إجمالي الوقت المهدر", size="إجمالي المكالمات",
                text=sales_col, color="نسبة النجاح %",
                color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS],
            )
            avg_success = perf["نسبة النجاح %"].mean()
            avg_wasted = perf["إجمالي الوقت المهدر"].mean()
            fig.add_vline(x=avg_success, line_dash="dot", line_color="#8B96AC", opacity=0.5)
            fig.add_hline(y=avg_wasted, line_dash="dot", line_color="#8B96AC", opacity=0.5)
            fig.update_traces(textposition="top center", textfont_size=10, marker=dict(line=dict(color="#0E1420", width=1)))
            fig.update_layout(
                **PLOTLY_LAYOUT, coloraxis_showscale=False,
                xaxis_title="نسبة النجاح %", yaxis_title="إجمالي الوقت المهدر (دقيقة)",
            )
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
            st.caption(
                "🔵 حجم الفقاعة = عدد المكالمات · الخطوط المنقّطة = المتوسط العام. "
                "أعلى يمين = أداء ممتاز، أسفل يمين = ممتاز بس وقته بيتهدر، أعلى يسار = نجاح قليل بجهد وقت أقل، أسفل يسار = محتاج متابعة."
            )

        chart_card("🧭 خريطة الأداء: نسبة النجاح مقابل الوقت المهدر", _scatter)


def render_full_dashboard(df, class_col=None, sales_col=None, time_col=None):
    render_metric_cards(df, class_col, sales_col, time_col)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    render_highlights(df, class_col, sales_col, time_col)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(
        ["📈 نظرة عامة", "👥 أداء المحصّلين", "⏱️ الوقت المهدر", "🔍 مقارنة شاملة"]
    )

    with sub_tab1:
        col_a, col_b = st.columns(2)
        with col_a:
            render_pie_chart(df, class_col)
        with col_b:
            render_trend_chart(df, class_col, time_col)

    with sub_tab2:
        render_agent_perf_chart(df, class_col, sales_col, with_table=False)

    with sub_tab3:
        col_c, col_d = st.columns(2)
        with col_c:
            render_wasted_bar(df, sales_col, top_n=None)
        with col_d:
            render_wasted_hist(df)

    with sub_tab4:
        render_comparison_matrix(df, class_col, sales_col)


# ==========================================================
# تصدير البيانات (Exporting)
# ==========================================================

def get_download_buffer(df: pd.DataFrame, file_format: str):
    output = io.BytesIO()
    if file_format == "Excel (.xlsx)":
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="النتائج")
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ext = "xlsx"
    else:
        output.write(df.to_csv(index=False).encode("utf-8-sig"))
        mime = "text/csv"
        ext = "csv"
    output.seek(0)
    return output, mime, ext


# ==========================================================
# صفحات التطبيق (Pages)
# ==========================================================

def page_classify():
    page_header("تصنيف الإفادات", "تحليل وتصنيف إفادات المكالمات", "قم برفع ملف الإكسل أو CSV لتشغيل نموذج الذكاء الاصطناعي", show_wave=True)

    if "df" not in st.session_state:
        st.session_state.df = None

    file_upload = st.file_uploader("اختر ملف البيانات (Excel أو CSV)", type=["xlsx", "xls", "csv"])

    if file_upload:
        try:
            if file_upload.name.endswith(".csv"):
                df_loaded = pd.read_csv(file_upload)
            else:
                df_loaded = pd.read_excel(file_upload)
            st.session_state.df = df_loaded
            st.success(f"تم تحميل الملف بنجاح! الإجمالي: {len(df_loaded)} سجل.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    df = st.session_state.df

    if df is not None:
        sales_col = find_column(df, SALES_PERSON_CANDIDATES)
        time_col = find_column(df, CREATED_ON_CANDIDATES)
        text_col = ORIGINAL_TEXT_COL if ORIGINAL_TEXT_COL in df.columns else find_column(df, [ORIGINAL_TEXT_COL, MODEL_TEXT_COL, "إفادة", "تفاصيل"])

        st.markdown("### ⚙️ إعدادات المعالجة")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            break_start = st.time_input("بداية وقت البريك", dt_time(13, 0))
        with col_c2:
            break_end = st.time_input("نهاية وقت البريك", dt_time(14, 0))

        if st.button("🚀 البدء في التحليل والتصنيف"):
            if not text_col:
                st.error(f"لم يتم العثور على عمود النصوص ({ORIGINAL_TEXT_COL}). يرجى التحقق من الملف.")
                return

            tokenizer, model, device = load_model()
            texts = df[text_col].astype(str).tolist()

            preds, confs = predict_batch(texts, tokenizer, model, device)

            df[CLASSIFICATION_COL] = preds
            df["درجة_الثقة"] = [round(c, 4) for c in confs]

            if sales_col and time_col:
                df = calculate_wasted_time(df, sales_col, time_col, break_start, break_end)

            st.session_state.df = df
            st.success("تم التصنيف والحسابات بنجاح!")

        if CLASSIFICATION_COL in df.columns:
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.markdown("### 📊 لمحة سريعة عن النتائج")
            render_quick_summary(df, CLASSIFICATION_COL, sales_col, time_col)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.markdown("### 📥 تحميل البيانات المصنفة")
            exp_format = st.radio("اختر صيغة التحميل", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)
            buf, mime, ext = get_download_buffer(df, exp_format)
            st.download_button(
                label=f"تحميل النتائج ({ext})",
                data=buf,
                file_name=f"classified_calls_{datetime.now().strftime('%Y%m%d_%H%M')}.{ext}",
                mime=mime,
            )


def page_dashboard():
    page_header("لوحة التحليلات", "الداشبورد الشامل", "تحليل تفصيلي للأداء والأوقات المهدرة")
    if st.session_state.get("df") is not None:
        df = st.session_state.df
        sales_col = find_column(df, SALES_PERSON_CANDIDATES)
        time_col = find_column(df, CREATED_ON_CANDIDATES)
        render_full_dashboard(df, CLASSIFICATION_COL, sales_col, time_col)
    else:
        st.warning("يرجى رفع الملف وتشغيل التصنيف في صفحة «تصنيف الإفادات» أولاً.")


def page_promises_active():
    page_header("الوعود القائمة", "متابعة الوعود الفعالة", "سجل الوعود بالسداد القائمة ولم تتجاوز تاريخها")
    st.markdown('<div class="placeholder-card">قيد التطوير — سيتم إضافة المنطق الخاص بها لاحقًا.</div>', unsafe_allow_html=True)


def page_promises_broken():
    page_header("الوعود المكسورة", "متابعة الوعود النكثة", "سجل الوعود التي تم خلفها ولم يتم السداد فيها")
    st.markdown('<div class="placeholder-card">قيد التطوير — سيتم إضافة المنطق الخاص بها لاحقًا.</div>', unsafe_allow_html=True)


def page_neglect():
    page_header("الإهمال", "تحليل الحالات المكتشفة", "متابعة إهمال المتابعة الميدانية أو الترددات")
    st.markdown('<div class="placeholder-card">قيد التطوير — سيتم إضافة المنطق الخاص بها لاحقًا.</div>', unsafe_allow_html=True)


def page_case_errors():
    page_header("أخطاء الحالات", "تدقيق البيانات الإجرائية", "استعراض الأخطاء الواردة في تسجيل البيانات للإفادة")
    st.markdown('<div class="placeholder-card">قيد التطوير — سيتم إضافة المنطق الخاص بها لاحقًا.</div>', unsafe_allow_html=True)


# ==========================================================
# السايدبار وإدارة الصفحات (Navigation)
# ==========================================================

def main():
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
                <div class="title">🎙️ تحليل المكالمات</div>
                <div class="subtitle">7oudaModel Engine</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        selected_page = st.radio(
            "الانتقال إلى:",
            [
                "🎙️ تصنيف الإفادات",
                "📊 الداشبورد الشامل",
                "🤝 الوعود القائمة",
                "⚠️ الوعود المكسورة",
                "🚫 الإهمال",
                "❌ أخطاء الحالات",
            ],
            index=0,
        )

    if selected_page == "🎙️ تصنيف الإفادات":
        page_classify()
    elif selected_page == "📊 الداشبورد الشامل":
        page_dashboard()
    elif selected_page == "🤝 الوعود القائمة":
        page_promises_active()
    elif selected_page == "⚠️ الوعود المكسورة":
        page_promises_broken()
    elif selected_page == "🚫 الإهمال":
        page_neglect()
    elif selected_page == "❌ أخطاء الحالات":
        page_case_errors()


if __name__ == "__main__":
    main()
