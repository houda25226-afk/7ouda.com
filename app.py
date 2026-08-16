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
- أسماء الأعمدة المتوقعة في الملف (Sales Person, Created On, Note)
  متعرّفة في قسم "إعدادات وأسماء الأعمدة" تحت — لو الأسماء اتغيرت غيّرها من هناك.
- التبويبات الأربعة (الوعود القائمة / المكسورة / الإهمال / أخطاء الحالات)
  لسه فاضية (Placeholder) لحد ما نحدد منطق كل واحدة فيها.
============================================================
"""

import io
from datetime import time as dt_time

import pandas as pd
import plotly.express as px
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# إعدادات وأسماء الأعمدة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

ORIGINAL_TEXT_COL = "Note"          # اسم العمود الأصلي في الملف
MODEL_TEXT_COL = "الافادة"          # الاسم اللي بيتحول له مؤقتًا عشان الموديل
CLASSIFICATION_COL = "التصنيف"      # عمود النتيجة: 1 = ناجحة / 0 = غير ناجحة
WASTED_TIME_COL = "الوقت_المهدر_دقيقة"

SALES_PERSON_CANDIDATES = ["Sales Person", "sales person", "SalesPerson", "المحصل"]
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
[data-testid="stToolbar"] {visibility: hidden;}
/* الزرار اللي بيفتح/يقفل الشريط الجانبي لازم يفضل ظاهر */
[data-testid="collapsedControl"] {visibility: visible !important;}

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

def render_dashboard(df: pd.DataFrame, class_col: str = None, sales_col: str = None):
    has_class = class_col and class_col in df.columns
    has_wasted = WASTED_TIME_COL in df.columns
    has_sales = sales_col and sales_col in df.columns

    # ----- بطاقات (Metrics) -----
    metric_cols = st.columns(4)
    metric_cols[0].metric("📞 إجمالي المكالمات", len(df))

    if has_class:
        success_count = int((df[class_col] == 1).sum())
        fail_count = int((df[class_col] == 0).sum())
        metric_cols[1].metric("✅ ناجحة", success_count)
        metric_cols[2].metric("⛔ غير ناجحة", fail_count)
    else:
        metric_cols[1].metric("✅ ناجحة", "—")
        metric_cols[2].metric("⛔ غير ناجحة", "—")

    if has_wasted:
        metric_cols[3].metric("⏱️ إجمالي الوقت المهدر (دقيقة)", round(df[WASTED_TIME_COL].sum(), 1))
    else:
        metric_cols[3].metric("⏱️ إجمالي الوقت المهدر", "—")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    chart_colors = {"ناجحة": "#34D399", "غير ناجحة": "#FB7185"}

    col_a, col_b = st.columns(2)

    with col_a:
        if has_class:
            labels_series = df[class_col].map({1: "ناجحة", 0: "غير ناجحة"})
            pie_df = labels_series.value_counts().reset_index()
            pie_df.columns = ["التصنيف", "العدد"]
            fig = px.pie(
                pie_df, names="التصنيف", values="العدد", hole=0.55,
                color="التصنيف", color_discrete_map=chart_colors,
                title="توزيع نتائج التصنيف",
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font_color="#E7ECF3", legend_title_text="",
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("مفيش عمود تصنيف في الملف عشان نعرض توزيع النتائج.")

    with col_b:
        if has_wasted and has_sales:
            wasted_by_agent = (
                df.groupby(sales_col)[WASTED_TIME_COL].sum().sort_values(ascending=False).head(10).reset_index()
            )
            fig2 = px.bar(
                wasted_by_agent, x=WASTED_TIME_COL, y=sales_col, orientation="h",
                title="أعلى 10 محصّلين في الوقت المهدر (دقيقة)",
                color=WASTED_TIME_COL, color_continuous_scale=["#5EEAD4", "#FB7185"],
            )
            fig2.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font_color="#E7ECF3", yaxis={"categoryorder": "total ascending"},
                coloraxis_showscale=False,
            )
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("محتاجين عمود المحصّل + الوقت المهدر عشان نعرض الرسم ده.")

    if has_class and has_sales:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        agent_perf = (
            df.groupby(sales_col)[class_col]
            .agg(["count", "sum"])
            .rename(columns={"count": "إجمالي المكالمات", "sum": "ناجحة"})
        )
        agent_perf["غير ناجحة"] = agent_perf["إجمالي المكالمات"] - agent_perf["ناجحة"]
        agent_perf["نسبة النجاح %"] = (agent_perf["ناجحة"] / agent_perf["إجمالي المكالمات"] * 100).round(1)
        agent_perf = agent_perf.sort_values("نسبة النجاح %", ascending=False).reset_index()

        fig3 = px.bar(
            agent_perf, x=sales_col, y=["ناجحة", "غير ناجحة"],
            title="أداء كل محصّل (عدد المكالمات الناجحة/غير الناجحة)",
            color_discrete_sequence=["#34D399", "#FB7185"], barmode="stack",
        )
        fig3.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font_color="#E7ECF3", legend_title_text="",
        )
        st.plotly_chart(fig3, use_container_width=True)


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

    if uploaded_file is None:
        st.markdown(
            '<div class="placeholder-card">📂 ارفع ملف عشان تبدأ التصنيف</div>',
            unsafe_allow_html=True,
        )
        return

    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        return

    # ------------------------------------------------------
    # معالجة تلقائية بعد الرفع مباشرة:
    #   1) حذف أول صف بيانات بعد صف العناوين
    #   2) تغيير اسم عمود Note لـ الافادة (عشان الموديل)
    # ------------------------------------------------------
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    if ORIGINAL_TEXT_COL in df.columns:
        df = df.rename(columns={ORIGINAL_TEXT_COL: MODEL_TEXT_COL})

    st.markdown(
        f'<div class="card">✅ تم تحميل الملف — عدد الصفوف بعد الحذف: <b>{len(df)}</b></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(df.head(10), use_container_width=True)

    if MODEL_TEXT_COL not in df.columns:
        st.error(
            f"عمود النص ('{ORIGINAL_TEXT_COL}' أو '{MODEL_TEXT_COL}') مش موجود في الملف. "
            f"الأعمدة الموجودة: {', '.join(df.columns.astype(str))}"
        )
        return

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if st.button("🚀 ابدأ التصنيف", type="primary", use_container_width=True):
        tokenizer, model, device = load_model()

        texts = df[MODEL_TEXT_COL].tolist()
        preds, confidences = predict_batch(texts, tokenizer, model, device)

        result_df = df.copy()
        result_df[CLASSIFICATION_COL] = preds  # 1 = ناجحة / 0 = غير ناجحة
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]

        # ---- حساب الوقت المهدر ----
        sales_col = find_column(result_df, SALES_PERSON_CANDIDATES)
        time_col = find_column(result_df, CREATED_ON_CANDIDATES)

        if sales_col and time_col:
            result_df = calculate_wasted_time(result_df, sales_col, time_col, break_start, break_end)
        else:
            st.warning(
                "مش لاقي عمود المحصّل (Sales Person) أو عمود التاريخ (Created On) بنفس الاسم المتوقع، "
                "فمش هينحسب الوقت المهدر. الأعمدة الموجودة: " + ", ".join(result_df.columns.astype(str))
            )

        # ---- رجّع اسم العمود لـ Note قبل العرض والتحميل ----
        result_df = result_df.rename(columns={MODEL_TEXT_COL: ORIGINAL_TEXT_COL})

        st.session_state["last_result_df"] = result_df
        st.session_state["last_sales_col"] = sales_col

        st.success("تم التصنيف وحساب الوقت المهدر بنجاح ✅")

        # ---- عرض الجدول مع تلوين الوقت المهدر ----
        if WASTED_TIME_COL in result_df.columns:
            styled = result_df.style.applymap(highlight_wasted, subset=[WASTED_TIME_COL])
            st.dataframe(styled, use_container_width=True)
        else:
            st.dataframe(result_df, use_container_width=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # ---- الداشبورد السريع بعد التصنيف ----
        render_dashboard(result_df, class_col=CLASSIFICATION_COL, sales_col=sales_col)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # ---- زرار التحميل ----
        if uploaded_file.name.endswith(".csv"):
            output = result_df.to_csv(index=False).encode("utf-8-sig")
            file_name = "نتائج_التصنيف.csv"
            mime = "text/csv"
        else:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="النتائج")
            output = buffer.getvalue()
            file_name = "نتائج_التصنيف.xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        st.download_button(
            "⬇️ تحميل الملف مع التصنيف والوقت المهدر",
            data=output, file_name=file_name, mime=mime, use_container_width=True,
        )


# ==========================================================
# تويبات 2-5: هنبنيها واحدة واحدة لما نحدد منطق كل واحدة
# ==========================================================

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
# تويب 6: الداشبورد (رفع ملف مُصنّف جاهز وبناء الداشبورد تلقائي)
# ==========================================================

def page_dashboard():
    page_header(
        "ACTIVITY DASHBOARD",
        "📊 داشبورد النشاط",
        "ارفع الملف بعد ما يتصنّف، والداشبورد هيتبني تلقائي",
        show_wave=True,
    )

    dash_file = st.file_uploader(
        "ارفع الملف المصنّف (اللي فيه عمود 'التصنيف')", type=["csv", "xlsx", "xls"], key="dash_upload"
    )

    df = None
    if dash_file is not None:
        try:
            if dash_file.name.endswith(".csv"):
                df = pd.read_csv(dash_file)
            else:
                df = pd.read_excel(dash_file)
        except Exception as e:
            st.error(f"مش قادر أقرأ الملف: {e}")
            return
    elif "last_result_df" in st.session_state:
        st.info("مفيش ملف مرفوع دلوقتي — ده الداشبورد بتاع آخر تصنيف عملته في تويب «التصنيف».")
        df = st.session_state["last_result_df"]

    if df is None:
        st.markdown(
            '<div class="placeholder-card">📂 ارفع ملف مصنّف عشان يظهر الداشبورد هنا</div>',
            unsafe_allow_html=True,
        )
        return

    # اختيار الأعمدة (بيحاول يتعرف عليها لوحده، وتقدر تعدّل لو عايز)
    class_col_guess = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    sales_col_guess = find_column(df, SALES_PERSON_CANDIDATES)

    with st.expander("⚙️ تأكيد الأعمدة", expanded=(class_col_guess is None or sales_col_guess is None)):
        cols = ["— بدون —"] + list(df.columns.astype(str))
        class_col_sel = st.selectbox(
            "عمود التصنيف (1/0)", cols, index=cols.index(class_col_guess) if class_col_guess in cols else 0
        )
        sales_col_sel = st.selectbox(
            "عمود المحصّل", cols, index=cols.index(sales_col_guess) if sales_col_guess in cols else 0
        )

    class_col = None if class_col_sel == "— بدون —" else class_col_sel
    sales_col = None if sales_col_sel == "— بدون —" else sales_col_sel

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    render_dashboard(df, class_col=class_col, sales_col=sales_col)


# ==========================================================
# التنقل (Sidebar Navigation)
# ==========================================================

PAGES = {
    "🎯 التصنيف": page_classification,
    "📗 الوعود القائمة": lambda: page_placeholder(
        "PENDING", "الوعود القائمة", "المكالمات اللي فيها وعد سداد لسه قائم", "📗"
    ),
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
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**الموديل المستخدم**")
    st.code(MODEL_REPO, language=None)
    st.markdown(f"[عرض الموديل على Hugging Face ↗](https://huggingface.co/{MODEL_REPO})")

PAGES[selected_page]()
