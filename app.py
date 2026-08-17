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

ORIGINAL_TEXT_COL = "Notes"         # اسم العمود الأصلي في الملف
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
# ملحوظة مهمة: الـ direction: rtl متطبق بس على محتوى النص
# (الـ block-container والـ sidebar content) مش على هيكل الصفحة كله،
# عشان الـ Sidebar يفضل ثابت فعليًا على الشمال زي ما اتطلب،
# بدل ما ينقلب يمين بسبب انعكاس اتجاه الـ flex layout.

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

/* الاتجاه RTL على محتوى الصفحة الرئيسي والسايدبار بس — مش على الهيكل العام */
.main .block-container {
    direction: rtl;
    text-align: right;
    max-width: 1400px;
    padding-top: 1.5rem;
}
[data-testid="stSidebarContent"], [data-testid="stSidebarUserContent"] {
    direction: rtl;
    text-align: right;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* ===== هيدر ستريملت + التوب بار (كان أبيض) ===== */
header[data-testid="stHeader"] {
    background: #0B111C !important;
    background-color: #0B111C !important;
    border-bottom: 1px solid rgba(255,255,255,0.06);
}
header[data-testid="stHeader"] * {
    color: var(--text) !important;
}
div[data-testid="stToolbar"] {
    background: transparent !important;
    color: var(--text) !important;
}
div[data-testid="stToolbar"] button,
div[data-testid="stToolbar"] svg,
div[data-testid="stToolbar"] span {
    color: var(--text-dim) !important;
    fill: var(--text-dim) !important;
}
div[data-testid="stDecoration"] {
    background: transparent !important;
    display: none !important;
}
/* ستاتوس بار / Deploy button */
[data-testid="stStatusWidget"] {
    background: var(--surface) !important;
    color: var(--text) !important;
}
/* أي نص في التوب بار */
.stDeployButton, .stAppDeployButton {
    color: var(--text-dim) !important;
}
/* خلفية الـ app بالكامل */
.stApp > header {
    background-color: #0B111C !important;
}
section.main > div {
    background: transparent !important;
}

/* ===== الشريط الجانبي (ثابت على الشمال) ===== */
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

/* قائمة التنقل — نخلي الـ radio يبان زي عناصر قائمة جانبية */
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

/* ===== الهيدر ===== */
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

/* موجة صوتية بسيطة */
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

/* ===== بطاقات عامة ===== */
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

/* ===== منطقة رفع الملفات ===== */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(94, 234, 212, 0.35);
    border-radius: 14px;
    padding: 0.6rem;
}
[data-testid="stFileUploader"] section { background: transparent; }

/* نص التعليمات جوه صندوق الرفع (اسحب الملف هنا / الحد الأقصى...) كان بلون باهت جدًا فوق الخلفية الغامقة */
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
/* زرار Browse files */
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
/* اسم الملف بعد الرفع + حجمه + زرار الحذف */
[data-testid="stFileUploaderFile"] {
    background: var(--surface-2) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploaderFile"] div, [data-testid="stFileUploaderFile"] span {
    color: var(--text) !important;
}
[data-testid="stFileUploaderFileName"] { color: var(--text) !important; }
[data-testid="stFileUploaderFileErrorMessage"] { color: var(--danger) !important; }

/* عناوين كل ودجت (label) — كانت أحيانًا بلون باهت جدًا */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] span {
    color: var(--text) !important;
    opacity: 1 !important;
}

/* ===== الأزرار ===== */
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

/* ===== شريط التقدم ===== */
[data-testid="stProgress"] > div > div { background: var(--accent); }

/* ===== المؤشرات (Metrics) ===== */
[data-testid="stMetric"] {
    background: var(--surface);
    border-radius: 12px;
    padding: 0.8rem 1rem;
    border: 1px solid rgba(255,255,255,0.06);
}
[data-testid="stMetricValue"] { font-family: 'JetBrains Mono', monospace; color: var(--accent); }

/* ===== الجداول ===== */
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

/* ===== كروت الشارتس (st.container(border=True)) ===== */
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

/* ===== كروت أبرز النقاط (Highlights) ===== */
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

/* ===== تظليم كامل لباقي عناصر الواجهة (Selectbox / Expander / Tabs / Alerts / Inputs) ===== */
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
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input {
    background: var(--surface-2) !important;
    color: var(--text) !important;
    border-color: rgba(255,255,255,0.1) !important;
    border-radius: 10px !important;
}

/* MultiSelect tags */
[data-baseweb="tag"] {
    background: var(--surface-2) !important;
    color: var(--text) !important;
    border-color: rgba(94,234,212,0.3) !important;
}
[data-baseweb="select"] span {
    color: var(--text) !important;
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

/* Scrollbar متناسق مع الثيم */
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
# داشبورد مشترك (يُستخدم بعد التصنيف مباشرة، وكمان في تويب الداشبورد)
# ==========================================================

# لوحة ألوان موحّدة للداشبورد كله
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
    """لمحة سريعة بتتعرض في تويب «التصنيف» فور ما التصنيف يخلص — دلوقتي فيها 4 شارتس
    + كروت أبرز النقاط، عشان تديك صورة كاملة من غير ما تحتاج تروح تويب الداشبورد."""
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
    """كروت أبرز النقاط — أفضل محصّل / الأكثر إهدارًا للوقت / أكثر يوم نشاطًا."""
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
                perf = perf[perf["إجمالي المكالمات"] >= 3]  # نتجاهل اللي مكالماته قليلة جدًا
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
    """جدول مقارنة شامل لكل المحصّلين + خريطة أداء (نسبة النجاح مقابل الوقت المهدر)."""
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


# ==========================================================
# تحليل سلوك المحصّلين (تعميق)
# ==========================================================

def _build_behavior_table(df, class_col, sales_col, time_col):
    """جدول سلوك شامل لكل محصّل: حجم، نجاح، وقت مهدر، إيقاع، ثبات."""
    has_class = class_col and class_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns
    has_time = time_col and time_col in df.columns

    rows = []
    for agent, g in df.groupby(sales_col):
        n = len(g)
        row = {"المحصّل": agent, "عدد المكالمات": n}

        if has_class:
            succ = int((g[class_col] == 1).sum())
            row["ناجحة"] = succ
            row["غير ناجحة"] = n - succ
            row["نسبة النجاح %"] = round(succ / n * 100, 1) if n else 0
        else:
            row["ناجحة"] = None
            row["غير ناجحة"] = None
            row["نسبة النجاح %"] = None

        if has_wasted:
            row["إجمالي الوقت المهدر"] = round(g[WASTED_TIME_COL].sum(), 1)
            row["متوسط الفجوة (د)"] = round(g[WASTED_TIME_COL].mean(), 1)
            row["وسيط الفجوة (د)"] = round(g[WASTED_TIME_COL].median(), 1)
            row["انحراف الفجوة"] = round(g[WASTED_TIME_COL].std(), 1) if n > 1 else 0
            # نسبة الفجوات الطويلة (>10 د)
            long_gaps = (g[WASTED_TIME_COL] > 10).sum()
            row["% فجوات طويلة"] = round(long_gaps / n * 100, 1) if n else 0
        else:
            row["إجمالي الوقت المهدر"] = None
            row["متوسط الفجوة (د)"] = None
            row["وسيط الفجوة (د)"] = None
            row["انحراف الفجوة"] = None
            row["% فجوات طويلة"] = None

        if has_time:
            t = pd.to_datetime(g[time_col], errors="coerce").dropna()
            if len(t) >= 2:
                span_h = (t.max() - t.min()).total_seconds() / 3600
                row["مدة النشاط (س)"] = round(span_h, 1)
                row["مكالمات/ساعة"] = round(len(t) / span_h, 1) if span_h > 0 else None
            else:
                row["مدة النشاط (س)"] = None
                row["مكالمات/ساعة"] = None
        else:
            row["مدة النشاط (س)"] = None
            row["مكالمات/ساعة"] = None

        # مؤشر كفاءة مبسّط: نسبة النجاح × حجم نسبي − عقوبة الوقت المهدر
        # نطبّعه لاحقًا على كل الجدول
        rows.append(row)

    beh = pd.DataFrame(rows)
    if beh.empty:
        return beh

    # مؤشر الكفاءة (0–100): مزيج من نسبة النجاح + السرعة − الفجوات الطويلة
    if has_class and "نسبة النجاح %" in beh.columns:
        sr = beh["نسبة النجاح %"].fillna(0)
        vol = beh["عدد المكالمات"]
        vol_norm = (vol / vol.max() * 100) if vol.max() else 0
        long_pen = beh["% فجوات طويلة"].fillna(0) if has_wasted else 0
        pace = beh["مكالمات/ساعة"].fillna(0)
        pace_norm = (pace / pace.max() * 100) if pace.max() and pace.max() > 0 else 0
        beh["مؤشر الكفاءة"] = (
            sr * 0.45 + vol_norm * 0.20 + pace_norm * 0.20 - long_pen * 0.15
        ).clip(0, 100).round(1)
    else:
        beh["مؤشر الكفاءة"] = None

    return beh.sort_values("مؤشر الكفاءة", ascending=False, na_position="last").reset_index(drop=True)


def render_behavior_section(df, class_col, sales_col, time_col, sub_col=None):
    """قسم تعميق تحليل سلوك المحصّلين — يُستدعى من الداشبورد."""
    if not sales_col or sales_col not in df.columns or df.empty:
        st.info("محتاجين عمود المحصّل وبيانات عشان تحليل السلوك.")
        return

    st.markdown("### 🧠 تعميق تحليل سلوك المحصّلين")
    st.caption(
        "مؤشر الكفاءة · إيقاع المكالمات · نسبة النجاح حسب الساعة · الفجوات الطويلة · سلاسل النجاح/الفشل"
    )

    beh = _build_behavior_table(df, class_col, sales_col, time_col)
    has_class = class_col and class_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns
    has_time = time_col and time_col in df.columns

    # ---- كروت سلوك سريعة ----
    top_eff = beh.iloc[0] if len(beh) and beh["مؤشر الكفاءة"].notna().any() else None
    slowest = None
    if has_wasted and len(beh):
        slowest = beh.sort_values("متوسط الفجوة (د)", ascending=False).iloc[0]
    steadiest = None
    if has_wasted and len(beh) and beh["انحراف الفجوة"].notna().any():
        steadiest = beh.sort_values("انحراف الفجوة", ascending=True).iloc[0]

    h1, h2, h3 = st.columns(3)
    with h1:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">🏆 أعلى كفاءة</div>', unsafe_allow_html=True)
            if top_eff is not None:
                st.markdown(f'<div class="highlight-value">{top_eff["المحصّل"]}</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="highlight-sub">مؤشر {top_eff["مؤشر الكفاءة"]} · '
                    f'{top_eff.get("نسبة النجاح %", "—")}% نجاح</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown('<div class="highlight-sub">—</div>', unsafe_allow_html=True)
    with h2:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">🐌 أبطأ إيقاع (متوسط فجوة)</div>', unsafe_allow_html=True)
            if slowest is not None:
                st.markdown(f'<div class="highlight-value">{slowest["المحصّل"]}</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="highlight-sub">{slowest["متوسط الفجوة (د)"]} د/مكالمة</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown('<div class="highlight-sub">—</div>', unsafe_allow_html=True)
    with h3:
        with st.container(border=True):
            st.markdown('<div class="highlight-label">🎯 أكثر ثباتًا (أقل تذبذب فجوات)</div>', unsafe_allow_html=True)
            if steadiest is not None:
                st.markdown(f'<div class="highlight-value">{steadiest["المحصّل"]}</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="highlight-sub">انحراف {steadiest["انحراف الفجوة"]}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown('<div class="highlight-sub">—</div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ---- جدول السلوك ----
    def _beh_table():
        show_cols = [c for c in beh.columns if beh[c].notna().any()]
        st.dataframe(beh[show_cols], use_container_width=True, hide_index=True)
        st.caption(
            "مؤشر الكفاءة = 45% نجاح + 20% حجم + 20% سرعة (مكالمات/ساعة) − 15% عقوبة الفجوات الطويلة (>10 د). "
            "كل ما المؤشر أعلى كل ما السلوك أفضل."
        )

    chart_card(f"📋 بطاقة سلوك كل محصّل ({len(beh)})", _beh_table)

    # ---- شارتس سلوك ----
    col1, col2 = st.columns(2)

    with col1:
        def _eff_bar():
            if "مؤشر الكفاءة" not in beh.columns or beh["مؤشر الكفاءة"].isna().all():
                st.info("محتاجين تصنيف عشان نحسب مؤشر الكفاءة")
                return
            plot_df = beh.dropna(subset=["مؤشر الكفاءة"]).sort_values("مؤشر الكفاءة")
            fig = px.bar(
                plot_df, x="مؤشر الكفاءة", y="المحصّل", orientation="h",
                color="مؤشر الكفاءة",
                color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS],
            )
            fig.update_layout(**PLOTLY_LAYOUT, coloraxis_showscale=False,
                              xaxis_title="مؤشر الكفاءة (0–100)", yaxis_title="",
                              height=max(280, 28 * len(plot_df)))
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

        chart_card("🏅 ترتيب المحصّلين بمؤشر الكفاءة", _eff_bar)

    with col2:
        def _pace_scatter():
            if not has_class or not has_wasted:
                st.info("محتاجين تصنيف + وقت مهدر")
                return
            plot_df = beh.dropna(subset=["نسبة النجاح %", "متوسط الفجوة (د)"])
            if plot_df.empty:
                st.info("مفيش بيانات كافية")
                return
            fig = px.scatter(
                plot_df, x="متوسط الفجوة (د)", y="نسبة النجاح %",
                size="عدد المكالمات", text="المحصّل",
                color="مؤشر الكفاءة",
                color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS],
            )
            fig.update_traces(textposition="top center", textfont_size=10,
                              marker=dict(line=dict(color="#0E1420", width=1)))
            fig.update_layout(**PLOTLY_LAYOUT, coloraxis_showscale=False,
                              xaxis_title="متوسط الفجوة بين المكالمات (دقيقة)",
                              yaxis_title="نسبة النجاح %")
            # خطوط المتوسط
            if len(plot_df) > 1:
                fig.add_vline(x=plot_df["متوسط الفجوة (د)"].mean(), line_dash="dot",
                              line_color="#8B96AC", opacity=0.5)
                fig.add_hline(y=plot_df["نسبة النجاح %"].mean(), line_dash="dot",
                              line_color="#8B96AC", opacity=0.5)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
            st.caption("يسار أعلى = سريع وناجح · يمين أعلى = ناجح بس بطيء · يسار أسفل = سريع بس نجاح ضعيف")

        chart_card("🧭 خريطة الإيقاع مقابل النجاح", _pace_scatter)

    # ---- Heatmap: نسبة النجاح حسب الساعة ----
    if has_class and has_time:
        def _hour_heatmap():
            tmp = df.copy()
            tmp["_dt"] = pd.to_datetime(tmp[time_col], errors="coerce")
            tmp = tmp.dropna(subset=["_dt"])
            if tmp.empty:
                st.info("مفيش تواريخ صالحة")
                return
            tmp["الساعة"] = tmp["_dt"].dt.hour
            heat = (
                tmp.groupby([sales_col, "الساعة"])[class_col]
                .agg(["mean", "count"])
                .reset_index()
            )
            heat["نسبة النجاح %"] = (heat["mean"] * 100).round(1)
            # نعرض بس الساعات اللي فيها نشاط
            pivot = heat.pivot(index=sales_col, columns="الساعة", values="نسبة النجاح %")
            # ترتيب الساعات
            pivot = pivot.reindex(sorted(pivot.columns), axis=1)
            fig = px.imshow(
                pivot,
                color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS],
                aspect="auto",
                labels=dict(x="ساعة اليوم", y="المحصّل", color="نجاح %"),
            )
            fig.update_layout(**PLOTLY_LAYOUT, height=max(320, 30 * len(pivot)))
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
            st.caption("كل خلية = نسبة نجاح المحصّل في الساعة دي. الأخضر أعلى · الأحمر أقل.")

        chart_card("🔥 Heatmap — نسبة النجاح حسب الساعة × المحصّل", _hour_heatmap)

    # ---- توزيع الفجوات + سلاسل ----
    col3, col4 = st.columns(2)

    with col3:
        def _gap_box():
            if not has_wasted:
                st.info("محتاجين عمود الوقت المهدر")
                return
            # box plot للفجوات لكل محصّل
            fig = px.box(
                df, x=sales_col, y=WASTED_TIME_COL,
                color_discrete_sequence=[COLOR_ACCENT],
            )
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="الفجوة (دقيقة)")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
            st.caption("الصندوق يوضح توزيع الفجوات: الوسيط، الربع، والقيم الشاذة (فجوات طويلة جدًا).")

        chart_card("📦 توزيع فجوات المكالمات لكل محصّل", _gap_box)

    with col4:
        def _streaks():
            if not has_class or not has_time:
                st.info("محتاجين تصنيف + وقت عشان نحسب السلاسل")
                return
            tmp = df.copy()
            tmp["_dt"] = pd.to_datetime(tmp[time_col], errors="coerce")
            tmp = tmp.dropna(subset=["_dt"]).sort_values([sales_col, "_dt"])

            def max_streak(series, val):
                """أطول سلسلة متتالية بقيمة val."""
                best = cur = 0
                for v in series:
                    if v == val:
                        cur += 1
                        best = max(best, cur)
                    else:
                        cur = 0
                return best

            streak_rows = []
            for agent, g in tmp.groupby(sales_col):
                seq = g[class_col].tolist()
                streak_rows.append({
                    "المحصّل": agent,
                    "أطول سلسلة نجاح": max_streak(seq, 1),
                    "أطول سلسلة فشل": max_streak(seq, 0),
                })
            s_df = pd.DataFrame(streak_rows)
            if s_df.empty:
                st.info("مفيش بيانات")
                return
            fig = px.bar(
                s_df, x="المحصّل", y=["أطول سلسلة نجاح", "أطول سلسلة فشل"],
                barmode="group",
                color_discrete_sequence=[COLOR_SUCCESS, COLOR_FAIL],
            )
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="طول السلسلة",
                              legend_title_text="")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
            st.caption("أطول سلسلة نجاح متتالية vs أطول سلسلة فشل — مؤشر على الاستقرار النفسي/الأسلوبي.")

        chart_card("🔗 سلاسل النجاح والفشل المتتالية", _streaks)

    # ---- نسبة لا يرد/مغلق من إجمالي مكالمات كل محصّل ----
    if sub_col and sub_col in df.columns:
        def _na_rate():
            sub_str = df[sub_col].astype(str)
            mask_na = sub_str.str.contains("لا يرد|لايرد|no answer", case=False, na=False)
            mask_cl = sub_str.str.contains("مغلق|closed|غلق", case=False, na=False)
            tmp = df.copy()
            tmp["_na"] = mask_na.astype(int)
            tmp["_cl"] = mask_cl.astype(int)
            sizes = tmp.groupby(sales_col).size().rename("إجمالي")
            sums = tmp.groupby(sales_col)[["_na", "_cl"]].sum()
            agg = sums.join(sizes).reset_index()
            agg = agg.rename(columns={"_na": "لا_يرد", "_cl": "مغلق"})
            agg["% لا يرد"] = (agg["لا_يرد"] / agg["إجمالي"] * 100).round(1)
            agg["% مغلق"] = (agg["مغلق"] / agg["إجمالي"] * 100).round(1)
            melt = agg.melt(
                id_vars=[sales_col], value_vars=["% لا يرد", "% مغلق"],
                var_name="النوع", value_name="النسبة %",
            )
            fig = px.bar(
                melt, x=sales_col, y="النسبة %", color="النوع", barmode="group",
                color_discrete_map={"% لا يرد": COLOR_WARN, "% مغلق": COLOR_FAIL},
            )
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="% من مكالمات المحصّل",
                              legend_title_text="")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

        chart_card("📵 نسبة «لا يرد» و «مغلق» من إجمالي مكالمات كل محصّل", _na_rate)

    # تحميل جدول السلوك
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        beh.to_excel(writer, index=False, sheet_name="سلوك_المحصلين")
    st.download_button(
        "⬇️ تحميل بطاقة سلوك المحصّلين (Excel)",
        data=buf.getvalue(),
        file_name="تحليل_سلوك_المحصلين.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="beh_dl",
    )


def render_full_dashboard(df, class_col=None, sales_col=None, time_col=None):
    """النسخة الكاملة — تويب «الداشبورد» بس، منظّمة في تبويبات فرعية."""
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
# تصدير الداشبورد كصفحة ويب (HTML) مستقلة — نفس الشكل والألوان
# ==========================================================

def _fig_to_div(fig, div_id, height=420):
    """بيحول الفيجر لكارت HTML — أهم حاجة إننا نديله ارتفاع صريح، لأن من غيره
    الشارت بيتصاغر لحجم شبه صفري جوه صندوق ارتفاعه تلقائي (اللي كان بيحصل قبل كده)."""
    if fig.layout.height is None:
        fig.update_layout(height=height)
    fig.update_layout(**PLOTLY_LAYOUT, autosize=True)
    return pio.to_html(
        fig, full_html=False, include_plotlyjs=False,
        config=PLOTLY_CONFIG, div_id=div_id,
        default_width="100%", default_height=f"{fig.layout.height}px",
    )


def build_dashboard_html(df, class_col, sales_col, time_col, source_name="") -> str:
    """بيبني صفحة HTML مستقلة (تفتح في أي متصفح، تفاعلية زي الأصل)
    فيها نفس تصميم وألوان الداشبورد جوه التطبيق، عشان تتبعت أو تتحفظ كصفحة ويب."""

    has_class = bool(class_col and class_col in df.columns)
    has_sales = bool(sales_col and sales_col in df.columns)
    has_time = bool(time_col and time_col in df.columns)
    has_wasted = WASTED_TIME_COL in df.columns

    total_calls = len(df)
    success_count = int((df[class_col] == 1).sum()) if has_class else None
    fail_count = int((df[class_col] == 0).sum()) if has_class else None
    success_rate = round(success_count / total_calls * 100, 1) if has_class and total_calls else None
    total_wasted = round(df[WASTED_TIME_COL].sum(), 1) if has_wasted else None
    avg_wasted = round(df[WASTED_TIME_COL].mean(), 1) if has_wasted and total_calls else None

    # ---- كروت أبرز النقاط ----
    top_agent_html = '<div class="highlight-sub">محتاجين عمود التصنيف والمحصّل</div>'
    if has_class and has_sales:
        perf_all = _compute_agent_perf(df, class_col, sales_col)
        perf_qualified = perf_all[perf_all["إجمالي المكالمات"] >= 3]
        if len(perf_qualified):
            top = perf_qualified.iloc[0]
            top_agent_html = (
                f'<div class="highlight-value">{top[sales_col]}</div>'
                f'<div class="highlight-sub">{top["نسبة النجاح %"]}% نجاح '
                f'({int(top["إجمالي المكالمات"])} مكالمة)</div>'
            )
        else:
            top_agent_html = '<div class="highlight-sub">مفيش بيانات كافية</div>'

    slow_agent_html = '<div class="highlight-sub">محتاجين عمود المحصّل والوقت المهدر</div>'
    if has_wasted and has_sales:
        wasted_totals = df.groupby(sales_col)[WASTED_TIME_COL].sum().sort_values(ascending=False)
        if len(wasted_totals):
            slow_agent_html = (
                f'<div class="highlight-value">{wasted_totals.index[0]}</div>'
                f'<div class="highlight-sub">{round(wasted_totals.iloc[0], 1)} دقيقة مهدرة</div>'
            )
        else:
            slow_agent_html = '<div class="highlight-sub">مفيش بيانات كافية</div>'

    busiest_day_html = '<div class="highlight-sub">محتاجين عمود التاريخ</div>'
    if has_time:
        t = pd.to_datetime(df[time_col], errors="coerce")
        if t.notna().any():
            daily_counts = t.dt.date.value_counts()
            busiest_day_html = (
                f'<div class="highlight-value">{daily_counts.index[0]}</div>'
                f'<div class="highlight-sub">{int(daily_counts.iloc[0])} مكالمة</div>'
            )
        else:
            busiest_day_html = '<div class="highlight-sub">مفيش تواريخ صالحة</div>'

    # ---- الشارتس (بنبنيها في متغيرات الأول، وبعدين نرتبها بالترتيب اللي هيبقى شكله منظم) ----
    chart_pie = chart_hist = chart_trend = chart_agents = chart_wasted = chart_scatter = None

    if has_class:
        labels_series = df[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
        pie_df = labels_series.value_counts().reset_index()
        pie_df.columns = ["التصنيف", "العدد"]
        fig_pie = px.pie(pie_df, names="التصنيف", values="العدد", hole=0.62,
                          color="التصنيف", color_discrete_map=CHART_COLORS)
        fig_pie.update_traces(textinfo="percent", textfont_size=13,
                               marker=dict(line=dict(color="#0E1420", width=3)))
        fig_pie.update_layout(showlegend=True,
                               legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5, xanchor="center"),
                               annotations=[dict(text=f"{success_rate}%<br><span style='font-size:11px;color:#8B96AC'>نجاح</span>",
                                                  x=0.5, y=0.5, font_size=22, font_color=COLOR_SUCCESS, showarrow=False)])
        chart_pie = ("🎯 توزيع نتائج التصنيف", _fig_to_div(fig_pie, "fig_pie"), False)

    if has_wasted:
        fig_hist = px.histogram(df, x=WASTED_TIME_COL, nbins=20, color_discrete_sequence=[COLOR_ACCENT])
        fig_hist.update_layout(bargap=0.08, xaxis_title="الوقت المهدر (دقيقة)", yaxis_title="عدد المرات")
        chart_hist = ("⏱️ توزيع الوقت المهدر بين المكالمات", _fig_to_div(fig_hist, "fig_hist"), False)

    if has_time:
        trend_df = df.copy()
        trend_df[time_col] = pd.to_datetime(trend_df[time_col], errors="coerce")
        trend_df["اليوم"] = trend_df[time_col].dt.date
        if has_class:
            trend_df["الحالة"] = trend_df[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
            daily = trend_df.groupby(["اليوم", "الحالة"]).size().reset_index(name="عدد المكالمات")
            fig_trend = px.area(daily, x="اليوم", y="عدد المكالمات", color="الحالة", color_discrete_map=CHART_COLORS)
        else:
            daily = trend_df.groupby("اليوم").size().reset_index(name="عدد المكالمات")
            fig_trend = px.area(daily, x="اليوم", y="عدد المكالمات", color_discrete_sequence=[COLOR_ACCENT])
        fig_trend.update_traces(line_width=2)
        fig_trend.update_layout(legend_title_text="", xaxis_title="", yaxis_title="عدد المكالمات")
        chart_trend = ("📅 اتجاه عدد المكالمات يوميًا", _fig_to_div(fig_trend, "fig_trend"), True)

    if has_class and has_sales:
        agent_perf = _compute_agent_perf(df, class_col, sales_col)
        fig_agents = px.bar(agent_perf, x=sales_col, y=["ناجحة", "غير ناجحة"],
                             color_discrete_sequence=[COLOR_SUCCESS, COLOR_FAIL], barmode="stack")
        fig_agents.update_traces(marker_line_width=0)
        fig_agents.update_layout(legend_title_text="", xaxis_title="", yaxis_title="عدد المكالمات")
        chart_agents = ("📊 أداء كل محصّل (ناجحة مقابل غير ناجحة)", _fig_to_div(fig_agents, "fig_agents"), True)

    if has_wasted and has_sales:
        wasted_by_agent = df.groupby(sales_col)[WASTED_TIME_COL].sum().sort_values(ascending=False).reset_index()
        chart_height = max(300, 28 * len(wasted_by_agent))
        fig_wasted = px.bar(wasted_by_agent, x=WASTED_TIME_COL, y=sales_col, orientation="h",
                             color=WASTED_TIME_COL, color_continuous_scale=[COLOR_ACCENT, COLOR_WARN, COLOR_FAIL])
        fig_wasted.update_traces(marker_line_width=0)
        fig_wasted.update_layout(yaxis={"categoryorder": "total ascending", "title": ""},
                                  xaxis_title="الوقت المهدر (دقيقة)", coloraxis_showscale=False, height=chart_height)
        chart_wasted = ("🏆 كل المحصّلين حسب الوقت المهدر", _fig_to_div(fig_wasted, "fig_wasted", height=chart_height), True)


    comparison_table_html = ""
    chart_scatter = None
    if has_class and has_sales:
        perf = _compute_agent_perf(df, class_col, sales_col)
        if has_wasted:
            wasted_agg = df.groupby(sales_col)[WASTED_TIME_COL].agg(["sum", "mean"]).round(1)
            wasted_agg.columns = ["إجمالي الوقت المهدر", "متوسط الوقت المهدر"]
            perf = perf.merge(wasted_agg, left_on=sales_col, right_index=True, how="left")

            fig_scatter = px.scatter(perf, x="نسبة النجاح %", y="إجمالي الوقت المهدر", size="إجمالي المكالمات",
                                      text=sales_col, color="نسبة النجاح %",
                                      color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS])
            fig_scatter.add_vline(x=perf["نسبة النجاح %"].mean(), line_dash="dot", line_color="#8B96AC", opacity=0.5)
            fig_scatter.add_hline(y=perf["إجمالي الوقت المهدر"].mean(), line_dash="dot", line_color="#8B96AC", opacity=0.5)
            fig_scatter.update_traces(textposition="top center", textfont_size=10,
                                       marker=dict(line=dict(color="#0E1420", width=1)))
            fig_scatter.update_layout(coloraxis_showscale=False, xaxis_title="نسبة النجاح %",
                                       yaxis_title="إجمالي الوقت المهدر (دقيقة)")
            chart_scatter = ("🧭 خريطة الأداء: نسبة النجاح مقابل الوقت المهدر", _fig_to_div(fig_scatter, "fig_scatter", height=460), True)

        comparison_table_html = perf.rename(columns={sales_col: "المحصّل"}).to_html(
            index=False, classes="comp-table", border=0, justify="right"
        )

    # ترتيب نهائي منظم: الاتنين نص العرض (دائري + هيستوجرام) جنب بعض، وبعدين كل شارت واسع في صف لوحده
    charts_html = [c for c in [chart_pie, chart_hist, chart_trend, chart_agents, chart_wasted, chart_scatter] if c]

    def metric_card(label, value, sub=""):
        return f'''<div class="metric-box">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>'''

    metrics_html = "".join([
        metric_card("📞 إجمالي المكالمات", total_calls),
        metric_card("✅ ناجحة", success_count if has_class else "—", f"{success_rate}%" if has_class else ""),
        metric_card("⛔ غير ناجحة", fail_count if has_class else "—"),
        metric_card("⏱️ إجمالي الوقت المهدر (دقيقة)", total_wasted if has_wasted else "—",
                     f"متوسط {avg_wasted} د/مكالمة" if has_wasted else ""),
    ])

    charts_grid_html = "".join(
        f'''<div class="chart-box{' wide' if wide else ''}">
            <div class="chart-box-title">{title}</div>
            {div}
        </div>''' for title, div, wide in charts_html
    )

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>داشبورد نشاط المحصّلين | 7oudaModel</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');
:root {{
    --bg: #0E1420; --surface: #151F30; --surface-2: #1B2A42;
    --accent: #5EEAD4; --success: #34D399; --danger: #FB7185; --warn: #FBBF24;
    --text: #E7ECF3; --text-dim: #8B96AC;
}}
* {{ box-sizing: border-box; }}
body {{
    margin: 0; font-family: 'Tajawal', sans-serif; color: var(--text);
    background: radial-gradient(circle at 20% 0%, #10192C 0%, var(--bg) 55%);
    min-height: 100vh;
}}
.wrap {{ max-width: 1280px; margin: 0 auto; padding: 2rem 1.5rem 4rem; }}
.header {{ margin-bottom: 1.5rem; }}
.eyebrow {{
    font-family: 'JetBrains Mono', monospace; font-size: 0.72rem; letter-spacing: 0.14em;
    color: var(--accent); text-transform: uppercase; margin-bottom: 0.3rem;
}}
h1 {{ font-weight: 900; font-size: 1.9rem; margin: 0; }}
.subtitle {{ color: var(--text-dim); font-size: 0.95rem; margin-top: 0.4rem; }}
.divider {{
    height: 1px; background: linear-gradient(90deg, transparent, rgba(94,234,212,0.35), transparent);
    margin: 1.6rem 0; border: none;
}}
.metrics-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; }}
.metric-box {{
    background: var(--surface); border: 1px solid rgba(255,255,255,0.06); border-radius: 14px;
    padding: 1rem 1.1rem;
}}
.metric-label {{ font-size: 0.82rem; color: var(--text-dim); margin-bottom: 0.4rem; }}
.metric-value {{ font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 1.5rem; color: var(--accent); }}
.metric-sub {{ font-size: 0.78rem; color: var(--text-dim); margin-top: 0.3rem; }}
.highlights-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; }}
.highlight-box {{
    background: var(--surface); border: 1px solid rgba(255,255,255,0.06); border-radius: 14px; padding: 1rem 1.1rem;
}}
.highlight-label {{ font-size: 0.82rem; color: var(--text-dim); margin-bottom: 0.5rem; }}
.highlight-value {{ font-weight: 900; font-size: 1.2rem; color: var(--accent); line-height: 1.3; }}
.highlight-sub {{ font-size: 0.8rem; color: var(--text-dim); margin-top: 0.2rem; }}
.charts-grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem; align-items: start; }}
.chart-box {{
    background: var(--surface); border: 1px solid rgba(255,255,255,0.07); border-radius: 14px; padding: 1rem 1.1rem;
    overflow: hidden;
}}
.chart-box.wide {{ grid-column: 1 / -1; }}
.chart-box .plotly-graph-div {{ width: 100% !important; }}
.chart-box-title {{
    font-weight: 700; font-size: 0.95rem; margin-bottom: 0.5rem; padding-bottom: 0.5rem;
    border-bottom: 1px solid rgba(255,255,255,0.06);
}}
.table-box {{
    background: var(--surface); border: 1px solid rgba(255,255,255,0.07); border-radius: 14px;
    padding: 1rem 1.1rem; margin-top: 1rem; overflow-x: auto;
}}
.comp-table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
.comp-table th {{
    background: var(--surface-2); color: var(--accent); padding: 0.55rem 0.7rem; text-align: right;
    position: sticky; top: 0;
}}
.comp-table td {{ padding: 0.5rem 0.7rem; border-bottom: 1px solid rgba(255,255,255,0.05); }}
.comp-table tr:hover td {{ background: rgba(94,234,212,0.05); }}
.footer {{ text-align: center; color: var(--text-dim); font-size: 0.78rem; margin-top: 2.5rem; }}
@media (max-width: 900px) {{
    .metrics-grid, .highlights-grid, .charts-grid {{ grid-template-columns: 1fr; }}
}}
</style>
</head>
<body>
<div class="wrap">
    <div class="header">
        <div class="eyebrow">ACTIVITY DASHBOARD · SNAPSHOT</div>
        <h1>📊 داشبورد نشاط المحصّلين</h1>
        <div class="subtitle">{('مصدر البيانات: ' + source_name) if source_name else ''} · تم إنشاؤه في {generated_at}</div>
    </div>
    <div class="divider"></div>
    <div class="metrics-grid">{metrics_html}</div>
    <div class="divider"></div>
    <div class="highlights-grid">
        <div class="highlight-box"><div class="highlight-label">🏅 أفضل محصّل (نسبة نجاح)</div>{top_agent_html}</div>
        <div class="highlight-box"><div class="highlight-label">🐌 الأكثر إهدارًا للوقت</div>{slow_agent_html}</div>
        <div class="highlight-box"><div class="highlight-label">📅 أكثر يوم نشاطًا</div>{busiest_day_html}</div>
    </div>
    <div class="divider"></div>
    <div class="charts-grid">{charts_grid_html}</div>
    {f'<div class="table-box"><div class="chart-box-title">📋 جدول المقارنة الشامل — كل المحصّلين</div>{comparison_table_html}</div>' if comparison_table_html else ''}
    <div class="footer">تم إنشاء هذه الصفحة تلقائيًا من لوحة تحليل المكالمات · 7oudaModel</div>
</div>
</body>
</html>"""


# ==========================================================
# تويب 1: التصنيف
# ==========================================================

def page_classification():
    page_header(
        "CALL QUALITY CLASSIFIER",
        "🎯 تصنيف المكالمات",
        "ارفع الملف، وهيتصنّف كل صف تلقائيًا (1 = ناجحة، 0 = غير ناجحة) ويتحسب الوقت المهدر لكل محصّل",
        show_wave=True,
    )

    with st.expander("⚙️ إعدادات حساب الوقت المهدر (وقت البريك)", expanded=False):
        c1, c2 = st.columns(2)
        break_start = c1.time_input("بداية البريك", value=dt_time(13, 0))
        break_end = c2.time_input("نهاية البريك", value=dt_time(13, 30))

    uploaded_file = st.file_uploader("ارفع ملف البيانات (CSV أو Excel)", type=["csv", "xlsx", "xls"])

    if uploaded_file is not None:
        st.session_state["raw_file_bytes"] = uploaded_file.getvalue()
        st.session_state["raw_file_name"] = uploaded_file.name
        st.session_state["raw_file_size"] = uploaded_file.size

    has_cached_file = "raw_file_bytes" in st.session_state

    if not has_cached_file:
        st.markdown(
            '<div class="placeholder-card">📂 ارفع ملف عشان تبدأ التصنيف</div>',
            unsafe_allow_html=True,
        )
        return

    file_bytes = st.session_state["raw_file_bytes"]
    file_name = st.session_state["raw_file_name"]
    file_size = st.session_state["raw_file_size"]

    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        return

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    if ORIGINAL_TEXT_COL in df.columns:
        df = df.rename(columns={ORIGINAL_TEXT_COL: MODEL_TEXT_COL})

    st.dataframe(df.head(10), use_container_width=True)

    if MODEL_TEXT_COL not in df.columns:
        st.error(
            f"عمود النص ('{ORIGINAL_TEXT_COL}' أو '{MODEL_TEXT_COL}') مش موجود في الملف. "
            f"الأعمدة الموجودة: {', '.join(df.columns.astype(str))}"
        )
        return

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    file_token = f"{file_name}_{file_size}"

    if st.button("🚀 ابدأ التصنيف", type="primary", use_container_width=True):
        tokenizer, model, device = load_model()

        texts = df[MODEL_TEXT_COL].tolist()
        preds, confidences = predict_batch(texts, tokenizer, model, device)

        result_df = df.copy()
        result_df[CLASSIFICATION_COL] = preds
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]

        sales_col = find_column(result_df, SALES_PERSON_CANDIDATES)
        time_col = find_column(result_df, CREATED_ON_CANDIDATES)

        if sales_col and time_col:
            result_df = calculate_wasted_time(result_df, sales_col, time_col, break_start, break_end)
        else:
            st.warning(
                "مش لاقي عمود المحصّل (Create By) أو عمود التاريخ (Created On) بنفس الاسم المتوقع، "
                "فمش هينحسب الوقت المهدر. الأعمدة الموجودة: " + ", ".join(result_df.columns.astype(str))
            )

        result_df = result_df.rename(columns={MODEL_TEXT_COL: ORIGINAL_TEXT_COL})

        st.session_state["last_result_df"] = result_df
        st.session_state["last_sales_col"] = sales_col
        st.session_state["last_time_col"] = time_col
        st.session_state["last_file_token"] = file_token

    if st.session_state.get("last_file_token") == file_token and "last_result_df" in st.session_state:
        result_df = st.session_state["last_result_df"]
        sales_col = st.session_state.get("last_sales_col")
        time_col = st.session_state.get("last_time_col")

        st.success("تم التصنيف وحساب الوقت المهدر بنجاح ✅")

        if WASTED_TIME_COL in result_df.columns:
            try:
                styled = result_df.style.map(highlight_wasted, subset=[WASTED_TIME_COL])
            except AttributeError:
                styled = result_df.style.applymap(highlight_wasted, subset=[WASTED_TIME_COL])
            st.dataframe(styled, use_container_width=True)
        else:
            st.dataframe(result_df, use_container_width=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        render_quick_summary(result_df, class_col=CLASSIFICATION_COL, sales_col=sales_col, time_col=time_col)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        if file_name.endswith(".csv"):
            output = result_df.to_csv(index=False).encode("utf-8-sig")
            out_name = "نتائج_التصنيف.csv"
            mime = "text/csv"
        else:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="النتائج")
            output = buffer.getvalue()
            out_name = "نتائج_التصنيف.xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        st.download_button(
            "⬇️ تحميل الملف مع التصنيف والوقت المهدر",
            data=output, file_name=out_name, mime=mime, use_container_width=True,
        )


# ==========================================================
# تويب: الوعود القائمة
# ==========================================================

EXCLUDED_SALESPERSONS = [
    "Archive Companies II Anas",
    "Closed payments II Anas",
    "Hold Companies II Anas",
    "Op II Ibrahim Qassem",
    "قانونى -الوطنية",
]

SALESPERSON_CANDS = ["Salesperson", "Sales Person", "salesperson", "Create By", "Created By", "المحصل", "CreateBy"]
SUBSTATE_CANDS = ["Sub State", "SubState", "sub state", "Sub state", "الحالة الفرعية"]
FOLLOWUP_CANDS = ["Follow up Due Date", "Follow Up Due Date", "Followup Due Date", "Follow-up Due Date", "موعد المتابعة"]
ACCOUNT_CANDS = ["Account Number", "Customer Account Number", "Account No", "رقم الحساب", "AccountNumber"]
NET_AMOUNT_CANDS = ["Net Amount", "NetAmount", "Amount", "المبلغ", "صافي المبلغ"]


def page_pending_promises():
    page_header(
        "PENDING PROMISES",
        "📗 الوعود القائمة",
        "ارفع المحفظة → فلترة تلقائية (Salesperson / واعد بالسداد / تاريخ اليوم) → Excel + Pivot",
        show_wave=True,
    )

    uploaded = st.file_uploader(
        "ارفع ملف المحفظة (CSV أو Excel)",
        type=["csv", "xlsx", "xls"],
        key="pending_upload",
    )

    if uploaded is None:
        st.markdown(
            '<div class="placeholder-card">📂 ارفع ملف المحفظة عشان نطلع الوعود القائمة</div>',
            unsafe_allow_html=True,
        )
        return

    try:
        if uploaded.name.endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        return

    # اكتشاف الأعمدة
    sales_col = find_column(df, SALESPERSON_CANDS)
    sub_col = find_column(df, SUBSTATE_CANDS)
    due_col = find_column(df, FOLLOWUP_CANDS)
    account_col = find_column(df, ACCOUNT_CANDS)
    amount_col = find_column(df, NET_AMOUNT_CANDS)

    missing = []
    if not sales_col:
        missing.append("Salesperson")
    if not sub_col:
        missing.append("Sub State")
    if not due_col:
        missing.append("Follow up Due Date")
    if missing:
        st.error(
            f"الأعمدة دي مش موجودة: {', '.join(missing)}. "
            f"الأعمدة الموجودة: {', '.join(df.columns.astype(str))}"
        )
        return

    with st.expander("⚙️ الأعمدة المكتشفة", expanded=False):
        st.write(f"**Salesperson:** `{sales_col}`")
        st.write(f"**Sub State:** `{sub_col}`")
        st.write(f"**Follow up Due Date:** `{due_col}`")
        st.write(f"**Account Number:** `{account_col or '—'}`")
        st.write(f"**Net Amount:** `{amount_col or '—'}`")

    # 1) فلترة Salesperson — استبعاد القائمة
    mask_sales = ~df[sales_col].astype(str).str.strip().isin(EXCLUDED_SALESPERSONS)
    # 2) Sub State = واعد بالسداد
    mask_sub = df[sub_col].astype(str).str.strip() == "واعد بالسداد"
    # 3) Follow up Due Date = تاريخ اليوم
    due_parsed = pd.to_datetime(df[due_col], errors="coerce")
    today = pd.Timestamp.now().normalize()
    mask_due = due_parsed.dt.normalize() == today

    filtered = df[mask_sales & mask_sub & mask_due].copy()

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("📋 إجمالي الصفوف الأصلية", len(df))
    m2.metric("📗 الوعود القائمة (بعد الفلترة)", len(filtered))
    m3.metric("📅 تاريخ اليوم", today.strftime("%Y-%m-%d"))

    if len(filtered) == 0:
        st.warning("مفيش صفوف مطابقة للفلاتر (واعد بالسداد + تاريخ اليوم + استبعاد المحصلين المحددين).")
        return

    st.success(f"تم استخراج {len(filtered)} وعد قائم ✅")
    st.dataframe(filtered, use_container_width=True)

    # تحميل Excel للوعود القائمة
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="الوعود_القائمة")
    st.download_button(
        "⬇️ تحميل الوعود القائمة (Excel)",
        data=buffer.getvalue(),
        file_name=f"الوعود_القائمة_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    # Pivot: Salesperson + Account Number + Net Amount
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("### 📊 Pivot Table — المحصّل × رقم الحساب × صافي المبلغ")

    if not account_col or not amount_col:
        st.info(
            "محتاجين عمودي Account Number و Net Amount عشان نبني الـ Pivot. "
            f"Account: {account_col or 'مش موجود'} · Amount: {amount_col or 'مش موجود'}"
        )
    else:
        pivot_df = filtered.copy()
        pivot_df[amount_col] = pd.to_numeric(pivot_df[amount_col], errors="coerce").fillna(0)
        pivot = (
            pivot_df.groupby([sales_col, account_col], as_index=False)[amount_col]
            .sum()
            .rename(columns={sales_col: "Salesperson", account_col: "Account Number", amount_col: "Net Amount"})
            .sort_values(["Salesperson", "Net Amount"], ascending=[True, False])
        )
        st.dataframe(pivot, use_container_width=True, hide_index=True)

        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
            pivot.to_excel(writer, index=False, sheet_name="Pivot")
            filtered.to_excel(writer, index=False, sheet_name="الوعود_القائمة")
        st.download_button(
            "⬇️ تحميل الـ Pivot + الوعود (Excel)",
            data=buf2.getvalue(),
            file_name=f"pivot_الوعود_القائمة_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="pivot_dl",
        )


def page_placeholder(eyebrow, title, subtitle, icon):
    page_header(eyebrow, f"{icon} {title}", subtitle)
    st.markdown(
        f"""
        <div class="placeholder-card">
            {icon}<br><br>
            التويب ده لسه مننعمهاش — هنحدد سوا منطقها ومصدر بياناتها زي ما اتفقنا "حاجة حاجة"،
            وهتتفعّل هنا أول ما نخلص عليها.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==========================================================
# تويب الداشبورد (محدث)
# ==========================================================

def page_dashboard():
    page_header(
        "ACTIVITY DASHBOARD",
        "📊 داشبورد النشاط",
        "ارفع الملف المصنّف → فلترة (محصّل / Sub State / تاريخ / ناجحة) → كروت + شارتس احترافية",
        show_wave=True,
    )

    dash_file = st.file_uploader(
        "ارفع الملف المصنّف (اللي فيه عمود 'التصنيف')", type=["csv", "xlsx", "xls"], key="dash_upload"
    )

    if dash_file is not None:
        st.session_state["dash_raw_bytes"] = dash_file.getvalue()
        st.session_state["dash_raw_name"] = dash_file.name

    has_cached = "dash_raw_bytes" in st.session_state

    if not has_cached:
        st.markdown(
            '<div class="placeholder-card">📂 ارفع ملف مصنّف عشان يظهر الداشبورد هنا</div>',
            unsafe_allow_html=True,
        )
        return

    file_bytes = st.session_state["dash_raw_bytes"]
    file_name = st.session_state["dash_raw_name"]

    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        return

    # اكتشاف الأعمدة
    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    sales_col = find_column(df, SALES_PERSON_CANDIDATES + SALESPERSON_CANDS)
    time_col = find_column(df, CREATED_ON_CANDIDATES)
    sub_col = find_column(df, SUBSTATE_CANDS + ["Main State", "Final State", "الحالة", "State"])
    has_wasted = WASTED_TIME_COL in df.columns

    with st.expander("⚙️ تأكيد الأعمدة", expanded=(class_col is None or sales_col is None)):
        cols = ["— بدون —"] + list(df.columns.astype(str))
        class_col = st.selectbox(
            "عمود التصنيف (1/0)", cols,
            index=cols.index(class_col) if class_col in cols else 0, key="dash_class_sel"
        )
        sales_col = st.selectbox(
            "عمود المحصّل", cols,
            index=cols.index(sales_col) if sales_col in cols else 0, key="dash_sales_sel"
        )
        time_col = st.selectbox(
            "عمود التاريخ/الوقت", cols,
            index=cols.index(time_col) if time_col in cols else 0, key="dash_time_sel"
        )
        sub_col = st.selectbox(
            "عمود Sub State / الحالة", cols,
            index=cols.index(sub_col) if sub_col in cols else 0, key="dash_sub_sel"
        )
        class_col = None if class_col == "— بدون —" else class_col
        sales_col = None if sales_col == "— بدون —" else sales_col
        time_col = None if time_col == "— بدون —" else time_col
        sub_col = None if sub_col == "— بدون —" else sub_col

    # ========== فلاتر (Slicers) ==========
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 الفلاتر")

    f1, f2, f3, f4 = st.columns(4)

    with f1:
        if sales_col and sales_col in df.columns:
            agents = sorted(df[sales_col].dropna().astype(str).unique().tolist())
            sel_agents = st.multiselect("المحصّلين", agents, default=agents, key="flt_agents")
        else:
            sel_agents = None
            st.caption("مفيش عمود محصّل")

    with f2:
        if sub_col and sub_col in df.columns:
            states = sorted(df[sub_col].dropna().astype(str).unique().tolist())
            sel_states = st.multiselect("Sub State", states, default=states, key="flt_states")
        else:
            sel_states = None
            st.caption("مفيش عمود Sub State")

    with f3:
        if time_col and time_col in df.columns:
            t_series = pd.to_datetime(df[time_col], errors="coerce")
            valid_dates = t_series.dropna().dt.date
            if len(valid_dates):
                min_d, max_d = valid_dates.min(), valid_dates.max()
                date_range = st.date_input(
                    "نطاق التاريخ",
                    value=(min_d, max_d),
                    min_value=min_d,
                    max_value=max_d,
                    key="flt_dates",
                )
            else:
                date_range = None
                st.caption("مفيش تواريخ صالحة")
        else:
            date_range = None
            st.caption("مفيش عمود تاريخ")

    with f4:
        if class_col and class_col in df.columns:
            result_opts = st.multiselect(
                "نتيجة المكالمة",
                ["ناجحة", "غير ناجحة"],
                default=["ناجحة", "غير ناجحة"],
                key="flt_result",
            )
        else:
            result_opts = None
            st.caption("مفيش عمود تصنيف")

    # تطبيق الفلاتر
    filtered = df.copy()
    if sel_agents is not None and sales_col:
        filtered = filtered[filtered[sales_col].astype(str).isin(sel_agents)]
    if sel_states is not None and sub_col:
        filtered = filtered[filtered[sub_col].astype(str).isin(sel_states)]
    if date_range is not None and time_col and isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        t_parsed = pd.to_datetime(filtered[time_col], errors="coerce")
        mask = (t_parsed.dt.date >= date_range[0]) & (t_parsed.dt.date <= date_range[1])
        filtered = filtered[mask]
    if result_opts is not None and class_col:
        want = []
        if "ناجحة" in result_opts:
            want.append(1)
        if "غير ناجحة" in result_opts:
            want.append(0)
        if want:
            filtered = filtered[filtered[class_col].isin(want)]

    st.caption(f"عرض {len(filtered)} من أصل {len(df)} صف بعد الفلترة")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== كروت المؤشرات ==========
    n_agents = filtered[sales_col].nunique() if sales_col and sales_col in filtered.columns else 0
    total_covered = len(filtered)
    success_count = int((filtered[class_col] == 1).sum()) if class_col and class_col in filtered.columns else 0
    success_rate = round(success_count / total_covered * 100, 1) if total_covered else 0
    total_wasted = round(filtered[WASTED_TIME_COL].sum(), 1) if has_wasted else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("👥 عدد المحصّلين", n_agents)
    c2.metric("📞 المكالمات المغطاة", total_covered)
    c3.metric("✅ المكالمات الناجحة", success_count)
    c4.metric("📈 نسبة النجاح", f"{success_rate}%")
    c5.metric("⏱️ إجمالي الوقت المهدر (د)", total_wasted if has_wasted else "—")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== الشارتس ==========
    # صف 1: توزيع التصنيف + نشاط على مدار اليوم
    col_a, col_b = st.columns(2)

    with col_a:
        def _pie():
            if not class_col or class_col not in filtered.columns:
                st.info("مفيش عمود تصنيف")
                return
            labels = filtered[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
            pie_df = labels.value_counts().reset_index()
            pie_df.columns = ["التصنيف", "العدد"]
            fig = px.pie(
                pie_df, names="التصنيف", values="العدد", hole=0.55,
                color="التصنيف", color_discrete_map=CHART_COLORS,
            )
            fig.update_traces(textinfo="percent+label", textfont_size=12,
                              marker=dict(line=dict(color="#0E1420", width=2)))
            fig.update_layout(**PLOTLY_LAYOUT, showlegend=False,
                              annotations=[dict(
                                  text=f"{success_rate}%",
                                  x=0.5, y=0.5, font_size=24, font_color=COLOR_SUCCESS, showarrow=False
                              )])
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
        chart_card("🎯 توزيع نتائج التصنيف", _pie)

    with col_b:
        def _hourly():
            if not time_col or time_col not in filtered.columns:
                st.info("مفيش عمود وقت")
                return
            tmp = filtered.copy()
            tmp["_dt"] = pd.to_datetime(tmp[time_col], errors="coerce")
            tmp = tmp.dropna(subset=["_dt"])
            if tmp.empty:
                st.info("مفيش تواريخ صالحة")
                return
            tmp["الساعة"] = tmp["_dt"].dt.hour
            hourly = tmp.groupby("الساعة").size().reset_index(name="عدد المكالمات")
            fig = px.bar(hourly, x="الساعة", y="عدد المكالمات",
                         color_discrete_sequence=[COLOR_ACCENT])
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="ساعة اليوم", yaxis_title="عدد المكالمات",
                              xaxis=dict(dtick=1))
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
        chart_card("🕐 نشاط المحصّلين على مدار اليوم (بالساعة)", _hourly)

    # صف 2: أداء المحصلين + الوقت المهدر
    col_c, col_d = st.columns(2)

    with col_c:
        def _agent_calls():
            if not sales_col or sales_col not in filtered.columns:
                st.info("مفيش عمود محصّل")
                return
            if class_col and class_col in filtered.columns:
                perf = _compute_agent_perf(filtered, class_col, sales_col)
                fig = px.bar(
                    perf, x=sales_col, y=["ناجحة", "غير ناجحة"],
                    color_discrete_sequence=[COLOR_SUCCESS, COLOR_FAIL], barmode="stack",
                )
            else:
                counts = filtered.groupby(sales_col).size().reset_index(name="عدد المكالمات")
                fig = px.bar(counts, x=sales_col, y="عدد المكالمات",
                             color_discrete_sequence=[COLOR_ACCENT])
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="عدد المكالمات",
                              legend_title_text="")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
        chart_card("📊 نشاط المحصّلين (عدد المكالمات)", _agent_calls)

    with col_d:
        def _wasted_chart():
            if not (has_wasted and sales_col and sales_col in filtered.columns):
                st.info("محتاجين عمود الوقت المهدر + المحصّل")
                return
            w = (filtered.groupby(sales_col)[WASTED_TIME_COL]
                 .sum().sort_values(ascending=True).reset_index())
            h = max(280, 26 * len(w))
            fig = px.bar(
                w, x=WASTED_TIME_COL, y=sales_col, orientation="h",
                color=WASTED_TIME_COL,
                color_continuous_scale=[COLOR_ACCENT, COLOR_WARN, COLOR_FAIL],
            )
            fig.update_layout(**PLOTLY_LAYOUT, height=h, coloraxis_showscale=False,
                              xaxis_title="الوقت المهدر (دقيقة)", yaxis_title="")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
        chart_card("⏱️ الوقت المهدر لكل محصّل", _wasted_chart)

    # صف 3: لا يرد / مغلق حسب المحصل
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    def _noanswer_closed():
        if not (sales_col and sub_col and sales_col in filtered.columns and sub_col in filtered.columns):
            st.info("محتاجين عمود المحصّل + Sub State عشان نعرض لا يرد / مغلق")
            return
        # نبحث عن قيم فيها "لا يرد" أو "مغلق" (أو مشابه)
        sub_str = filtered[sub_col].astype(str)
        mask_na = sub_str.str.contains("لا يرد|لايرد|no answer|لا يرد", case=False, na=False)
        mask_cl = sub_str.str.contains("مغلق|closed|غلق", case=False, na=False)
        target = filtered[mask_na | mask_cl].copy()
        if target.empty:
            st.info("مفيش صفوف بحالة «لا يرد» أو «مغلق» في البيانات المفلترة")
            return
        target["_نوع"] = "أخرى"
        target.loc[mask_na.loc[target.index], "_نوع"] = "لا يرد"
        target.loc[mask_cl.loc[target.index], "_نوع"] = "مغلق"
        # لو الصف فيه الاتنين، نفضّل اللي اتحدد أخير
        counts = target.groupby([sales_col, "_نوع"]).size().reset_index(name="العدد")
        fig = px.bar(
            counts, x=sales_col, y="العدد", color="_نوع", barmode="group",
            color_discrete_map={"لا يرد": COLOR_WARN, "مغلق": COLOR_FAIL, "أخرى": COLOR_ACCENT},
        )
        fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="عدد الإفادات",
                          legend_title_text="")
        fig.update_traces(marker_line_width=0)
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    chart_card("📵 المحصّلين × إفادات «لا يرد» و «مغلق»", _noanswer_closed)

    # صف 4: اتجاه يومي
    def _daily_trend():
        if not time_col or time_col not in filtered.columns:
            st.info("مفيش عمود تاريخ")
            return
        tmp = filtered.copy()
        tmp["_dt"] = pd.to_datetime(tmp[time_col], errors="coerce")
        tmp = tmp.dropna(subset=["_dt"])
        tmp["اليوم"] = tmp["_dt"].dt.date
        if class_col and class_col in tmp.columns:
            tmp["الحالة"] = tmp[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
            daily = tmp.groupby(["اليوم", "الحالة"]).size().reset_index(name="عدد المكالمات")
            fig = px.area(daily, x="اليوم", y="عدد المكالمات", color="الحالة",
                          color_discrete_map=CHART_COLORS)
        else:
            daily = tmp.groupby("اليوم").size().reset_index(name="عدد المكالمات")
            fig = px.area(daily, x="اليوم", y="عدد المكالمات", color_discrete_sequence=[COLOR_ACCENT])
        fig.update_traces(line_width=2)
        fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="عدد المكالمات",
                          legend_title_text="")
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    chart_card("📅 اتجاه عدد المكالمات يوميًا", _daily_trend)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ===== تعميق تحليل سلوك المحصّلين =====
    render_behavior_section(
        filtered,
        class_col=class_col,
        sales_col=sales_col,
        time_col=time_col,
        sub_col=sub_col,
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # تحميل HTML + بيانات
    dashboard_html = build_dashboard_html(
        filtered, class_col=class_col, sales_col=sales_col, time_col=time_col, source_name=file_name
    )
    st.download_button(
        "🌐 تحميل الداشبورد كصفحة ويب (HTML)",
        data=dashboard_html.encode("utf-8"),
        file_name="داشبورد_النشاط.html",
        mime="text/html",
        use_container_width=True,
        key="dash_html_download",
        type="primary",
    )
    st.caption("صفحة ويب مستقلة بنفس الكروت والشارتس — تفتح في أي متصفح بدون تشغيل التطبيق.")

    with st.expander("⬇️ تحميل البيانات المفلترة", expanded=False):
        if file_name.endswith(".csv"):
            output = filtered.to_csv(index=False).encode("utf-8-sig")
            out_name = "بيانات_مفلترة.csv"
            mime = "text/csv"
        else:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                filtered.to_excel(writer, index=False, sheet_name="البيانات")
            output = buffer.getvalue()
            out_name = "بيانات_مفلترة.xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        st.download_button(
            "⬇️ تحميل البيانات المفلترة",
            data=output, file_name=out_name, mime=mime, use_container_width=True, key="dash_download",
        )


# ==========================================================
# التنقل (Sidebar Navigation)
# ==========================================================

PAGES = {
    "🎯 التصنيف": page_classification,
    "📗 الوعود القائمة": page_pending_promises,
    "📕 الوعود المكسورة": lambda: page_placeholder(
        "BROKEN", "الوعود المكسورة", "المكالمات اللي فيها وعد سداد اتكسر", "📕"
    ),
    "⚠️ الإهمال": lambda: page_placeholder(
        "NEGLECT", "الإهمال", "حالات الإهمال في المتابعة", "⚠️"
    ),
    "🧾 أخطاء الحالات": lambda: page_placeholder(
        "CASE ERRORS", "أخطاء الحالات", "الحالات اللي فيها أخطاء في التسجيل أو المتابعة", "🧾"
    ),
    "📊 الداشبورد": page_dashboard,
}

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="subtitle">7OUDA MODEL</div>
            <div class="title">🎙️ لوحة تحليل المكالمات</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    selected_page = st.radio("التنقل", list(PAGES.keys()), label_visibility="collapsed")

PAGES[selected_page]()
