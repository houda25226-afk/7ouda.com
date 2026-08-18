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
import base64
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
    max-width: 1200px;
}
[data-testid="stSidebarContent"], [data-testid="stSidebarUserContent"] {
    direction: rtl;
    text-align: right;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

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

/* Scrollbar متناسق مع الثيم */
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--surface-2); border-radius: 8px; }
::-webkit-scrollbar-thumb:hover { background: var(--accent); }
.actual-logo, .actual-company-logo { overflow:hidden !important; padding:5px !important; }
.actual-logo img, .actual-company-logo img { width:100% !important; height:100% !important; object-fit:contain !important; display:block !important; border-radius:10px !important; }
.upload-status { margin:.7rem 0 1rem; padding:.7rem .9rem; border-radius:10px; background:rgba(94,234,212,.08); border:1px solid rgba(94,234,212,.18); color:var(--text) !important; }
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


# تحسينات إضافية للـ Dark Theme + كروت الشركات والفترات
DARK_UI_FIX = """
<style>
/* نصوص Streamlit الافتراضية */
.stApp, .stApp p, .stApp span, .stApp label, .stApp small,
.stApp div, .stApp [data-testid="stMarkdownContainer"] {
    color: var(--text);
}
.stApp [data-testid="stCaptionContainer"],
.stApp .stCaption,
.stApp [data-testid="stHelp"] {
    color: var(--text-dim) !important;
}
.stApp input, .stApp textarea {
    color: var(--text) !important;
}
.stApp input::placeholder, .stApp textarea::placeholder {
    color: #9AA6BA !important;
    opacity: 1 !important;
}
.stApp [data-baseweb="select"] span {
    color: var(--text) !important;
}
.stApp [role="option"] {
    color: var(--text) !important;
}
.stApp [data-testid="stRadio"] label,
.stApp [data-testid="stCheckbox"] label {
    color: var(--text) !important;
}
.stApp [data-testid="stFileUploader"] section {
    color: var(--text) !important;
}
.stApp [data-testid="stFileUploader"] section * {
    color: var(--text) !important;
}
.stApp [data-testid="stDataFrame"] * {
    color: var(--text);
}
.stApp [data-testid="stMetricLabel"] {
    color: var(--text-dim) !important;
}
.stApp [data-testid="stMetricValue"] {
    color: var(--accent) !important;
}

/* أزرار عامة */
.stButton > button, .stDownloadButton > button {
    min-height: 44px;
    box-shadow: 0 5px 18px rgba(0,0,0,.18);
}
.stButton > button[kind="secondary"] {
    background: #17243A !important;
    color: #E7ECF3 !important;
    border: 1px solid rgba(94,234,212,.16) !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--accent) !important;
    color: var(--accent) !important;
    background: #1D304C !important;
}

/* عنوانات الأقسام */
.section-kicker {
    font-size: 1.05rem;
    font-weight: 900;
    color: #F3F7FB !important;
    margin: .25rem 0 .15rem;
}
.section-help {
    color: #AAB6C9 !important;
    font-size: .88rem;
    margin-bottom: .9rem;
}

/* كروت الشركات */
.company-card {
    display:flex;
    align-items:center;
    gap:14px;
    min-height:100px;
    padding:18px;
    background:linear-gradient(135deg,#17243A,#121C2E);
    border:1px solid rgba(255,255,255,.08);
    border-radius:18px;
    box-shadow:0 10px 30px rgba(0,0,0,.16);
    margin-bottom:9px;
}
.company-mark, .company-logo {
    display:flex;
    align-items:center;
    justify-content:center;
    border:2px solid;
    border-radius:16px;
    flex-shrink:0;
    font-size:1.45rem;
    font-weight:900;
    box-shadow:0 0 24px rgba(0,0,0,.2);
}
.company-mark img, .company-logo img {
    width:100%;
    height:100%;
    object-fit:contain;
    border-radius:inherit;
    display:block;
}
.company-mark {
    width:76px;
    height:76px;
    padding:6px;
    overflow:hidden;
}
.company-logo {
    padding:6px;
    overflow:hidden;
}
.company-name {
    font-size:1.05rem;
    font-weight:900;
    color:#F4F7FB !important;
}
.company-sub, .selected-company-sub {
    font-size:.78rem;
    color:#94A3B8 !important;
    margin-top:4px;
}
.selected-company {
    display:flex;
    align-items:center;
    gap:16px;
    margin:15px 0 8px;
    padding:14px 18px;
    border:1px solid rgba(94,234,212,.18);
    background:linear-gradient(90deg,#132A2C,#151F30);
    border-radius:16px;
}
.company-logo.large {
    width:70px !important;
    height:70px !important;
    border-radius:18px;
}
.company-logo.large img {
    width:100%;
    height:100%;
    object-fit:contain;
}
.selected-company-label {
    font-size:.75rem;
    color:#94A3B8 !important;
}
.selected-company-name {
    font-size:1.25rem;
    font-weight:900;
    color:#F8FAFC !important;
    margin-top:2px;
}

/* بانر الفترة */
.period-banner {
    display:flex;
    align-items:center;
    gap:14px;
    padding:15px 18px;
    margin:12px 0;
    background:linear-gradient(135deg,#182A43,#132136);
    border:1px solid rgba(110,168,254,.18);
    border-radius:16px;
}
.period-banner.daily {
    border-color:rgba(94,234,212,.2);
    background:linear-gradient(135deg,#142F2E,#142237);
}
.period-icon {
    font-size:1.6rem;
}
.period-label {
    color:#F5F8FC !important;
    font-weight:900;
    font-size:1.05rem;
}
.period-desc {
    color:#9EACC0 !important;
    font-size:.8rem;
    margin-top:3px;
}
.schedule-summary {
    margin:8px 0 14px;
    padding:10px 14px;
    background:#111B2C;
    border:1px solid rgba(255,255,255,.06);
    border-radius:10px;
    color:#C8D2E0 !important;
    font-size:.86rem;
}
.schedule-summary span {
    color:#8FA0B7 !important;
    margin:0 4px;
}
.schedule-summary b {
    color:var(--accent) !important;
}

/* المجمع اليومي */
.daily-total-card {
    display:grid;
    grid-template-columns:1fr auto auto;
    align-items:center;
    gap:18px;
    padding:18px 22px;
    margin:12px 0 18px;
    background:linear-gradient(135deg,#122B2A,#17243A);
    border:1px solid rgba(94,234,212,.2);
    border-radius:18px;
}
.daily-total-title {
    color:#F7FAFC !important;
    font-size:1.15rem;
    font-weight:900;
}
.daily-total-sub {
    color:#94A3B8 !important;
    font-size:.8rem;
    margin-top:4px;
}
.daily-total-number {
    color:var(--accent) !important;
    font-family:'JetBrains Mono',monospace;
    font-size:1.65rem;
    font-weight:900;
}
.daily-total-label {
    color:#8FA0B7 !important;
    font-size:.75rem;
}
@media (max-width: 800px) {
    .daily-total-card { grid-template-columns:1fr; }
}
</style>
"""
st.markdown(DARK_UI_FIX, unsafe_allow_html=True)


# ==========================================================
# إعدادات الـ Plotly الموحدة
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
    hoverlabel=dict(bgcolor="#151F30", font_color="#E7ECF3"),
)
PLOTLY_CONFIG = {"displayModeBar": False}


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

COMPANY_LOGO_TREE = "data:image/p
