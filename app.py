"""
تطبيق Streamlit — قاعدة أساسية (Base) بسايدبار وتنقل بين أقسام.
هنبني عليها بعدين قسم قسم حسب المطلوب.

طريقة التشغيل محليًا:
    pip install streamlit pandas openpyxl plotly --break-system-packages
    streamlit run app.py

============================================================
دليل التخصيص السريع:
- الألوان والخطوط كلها في متغير CSS_THEME تحت — عدّل من هناك بس.
- كل قسم من أقسام السايدبار له دالة render_* منفصلة تحت — سهل تضيف/تعدّل قسم من غير ما تلخبط الباقي.
- قائمة الأقسام نفسها في NAV_ITEMS تحت — ضيف أو شيل منها وهيتحدث السايدبار تلقائيًا.
- قسم "النشاط" فيه: رفع ملف -> حذف أول صف بعد الهيدر -> اختيار عمود التصنيف
  -> زرار "ابدأ التصنيف" (يحول ناجحة/غير ناجحة لـ 1/0) -> تشارتات -> تحميل الملف.
- قسم "لوحة النشاط" بيعرض نفس البيانات بعد التصنيف كداشبورد مستقل.
============================================================
"""

import io

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==========================================================
# إعداد الصفحة
# ==========================================================

st.set_page_config(
    page_title="النشاط",
    page_icon="🧭",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ==========================================================
# الهوية البصرية (Theme)
# ==========================================================
# لوحة الألوان:
#   خلفية كحلي غامق جدًا (#0B1220) هادية ومريحة
#   سطوح البطاقات (#131C2E) بحواف ناعمة وظل خفيف
#   لون أساسي أخضر-زمردي (#34D399) — إحساس نشاط وحيوية
#   لون تكميلي كهرماني دافئ (#FBBF24) للتباين والتنبيهات
#   نص أساسي فاتح (#F1F5F9) ونص ثانوي رمادي مزرق (#8A93A6)

ACCENT = "#34D399"
ACCENT_2 = "#FBBF24"
DANGER = "#F87171"
TEXT = "#F1F5F9"
TEXT_DIM = "#8A93A6"
SURFACE = "#131C2E"

CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0B1220;
    --surface: #131C2E;
    --surface-2: #1A2740;
    --accent: #34D399;
    --accent-2: #FBBF24;
    --danger: #F87171;
    --text: #F1F5F9;
    --text-dim: #8A93A6;
}

html, body, [class*="css"] {
    font-family: 'Cairo', sans-serif;
}

/* الحاوية الكبيرة LTR عشان السايدبار يفضل شمال، والنص جواه RTL */
[data-testid="stAppViewContainer"] {
    direction: ltr;
}
.main .block-container,
[data-testid="stMain"] {
    direction: rtl;
}

.stApp {
    background:
        radial-gradient(circle at 90% -10%, rgba(52,211,153,0.08) 0%, transparent 45%),
        radial-gradient(circle at 5% 105%, rgba(251,191,36,0.06) 0%, transparent 45%),
        var(--bg);
    color: var(--text);
}

/* نضمن إن كل نصوص المتن (مش بس الكروت المخصصة) واضحة فوق الخلفية الغامقة */
.stApp, .stApp p, .stApp span, .stApp label, .stApp li,
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stMarkdown, .stMarkdown p, .stText, .stCaption,
[data-testid="stFileUploaderDropzoneInstructions"],
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stWidgetLabel"] p,
[data-testid="stMetricLabel"] {
    color: var(--text) !important;
}
.stApp small, [data-testid="stCaptionContainer"] {
    color: var(--text-dim) !important;
}

#MainMenu, footer {visibility: hidden;}
header[data-testid="stHeader"] {
    background: transparent;
}
/* زرار فتح/قفل السايدبار — نضمن إنه دايمًا ظاهر وبلون واضح */
[data-testid="collapsedControl"] {
    visibility: visible !important;
    color: var(--text) !important;
}
[data-testid="collapsedControl"] svg {
    fill: var(--text) !important;
}

/* ===== السايدبار ===== */
section[data-testid="stSidebar"] {
    direction: rtl;
    background: #080C16;
    border-left: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: var(--text) !important; }

.brand-block {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    padding: 1rem 0.2rem 1.1rem 0.2rem;
    margin-bottom: 0.6rem;
    border-bottom: 1px solid rgba(255,255,255,0.08);
}
.brand-logo {
    width: 44px;
    height: 44px;
    border-radius: 13px;
    background: linear-gradient(135deg, var(--accent), var(--accent-2));
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.35rem;
    flex-shrink: 0;
    box-shadow: 0 4px 16px rgba(52, 211, 153, 0.25);
}
.brand-name {
    font-weight: 900;
    font-size: 1.15rem;
    color: var(--text) !important;
    line-height: 1.2;
}
.brand-sub {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.12em;
    color: var(--text-dim) !important;
    text-transform: uppercase;
}

/* قائمة التنقل في السايدبار (radio نخليها تبان كأزرار كبيرة) */
section[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 0.5rem;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 0.95rem 1.1rem;
    width: 100%;
    min-height: 3.2rem;
    display: flex;
    align-items: center;
    transition: background 0.15s ease, border-color 0.15s ease;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background: var(--surface-2);
    border-color: rgba(52, 211, 153, 0.35);
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
    display: none;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label div[data-testid="stMarkdownContainer"] p {
    font-size: 1.15rem !important;
    font-weight: 700 !important;
}

/* ===== الهيدر / Hero ===== */
.hero-wrap {
    text-align: center;
    padding: 1rem 0 0.3rem 0;
}
.hero-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.18em;
    color: var(--accent);
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}
.hero-title {
    font-weight: 900;
    font-size: 1.9rem;
    color: var(--text);
    margin: 0;
    line-height: 1.35;
}
.hero-subtitle {
    color: var(--text-dim);
    font-size: 0.95rem;
    margin-top: 0.55rem;
}

/* ===== البطاقات العامة ===== */
.card {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 16px;
    padding: 1.15rem 1.35rem;
    margin-bottom: 1rem;
}

.stat-card {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 16px;
    padding: 1rem 1.2rem;
    text-align: center;
    height: 100%;
}
.stat-card .stat-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 1.6rem;
    font-weight: 700;
    color: var(--accent);
}
.stat-card .stat-value.danger { color: var(--danger); }
.stat-card .stat-value.warn { color: var(--accent-2); }
.stat-card .stat-label {
    color: var(--text-dim);
    font-size: 0.82rem;
    margin-top: 0.3rem;
}

.section-title {
    font-weight: 900;
    font-size: 1.05rem;
    color: var(--text);
    margin: 1.4rem 0 0.6rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* ===== منطقة رفع الملفات ===== */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(52, 211, 153, 0.4);
    border-radius: 16px;
    padding: 0.7rem;
}
[data-testid="stFileUploader"] section {
    background: transparent;
}
/* زرار "Browse files" جوه صندوق الرفع — نديله نفس ستايل باقي الأزرار عشان يبان */
[data-testid="stFileUploader"] button {
    background: linear-gradient(90deg, var(--accent), #22B888) !important;
    color: #05170F !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 10px !important;
    opacity: 1 !important;
}
[data-testid="stFileUploader"] button p,
[data-testid="stFileUploader"] button span {
    color: #05170F !important;
}
[data-testid="stFileUploader"] button:hover {
    box-shadow: 0 6px 16px rgba(52, 211, 153, 0.3);
}

/* ===== الأزرار ===== */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), #22B888);
    color: #05170F;
    font-weight: 700;
    border: none;
    border-radius: 11px;
    padding: 0.65rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
    width: 100%;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 20px rgba(52, 211, 153, 0.3);
    color: #05170F;
}

/* ===== شريط التقدم ===== */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
}

/* ===== المؤشرات (Metrics) ===== */
[data-testid="stMetric"] {
    background: var(--surface);
    border-radius: 14px;
    padding: 0.85rem 1rem;
    border: 1px solid rgba(255,255,255,0.06);
}
[data-testid="stMetricValue"] {
    font-family: 'JetBrains Mono', monospace;
    color: var(--accent);
}

/* ===== الجداول ===== */
[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.06);
}

/* ===== تنبيهات النجاح/الخطأ ===== */
div[data-baseweb="notification"] {
    border-radius: 12px;
}

/* فاصل بسيط بدل الخط الافتراضي */
.divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(52,211,153,0.35), transparent);
    margin: 1.4rem 0;
    border: none;
}
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


# ==========================================================
# دوال مساعدة عامة (قراءة الملف - التصنيف - التشارتات - التحميل)
# ==========================================================

SUCCESS_LABEL = "ناجحة"
FAIL_LABEL = "غير ناجحة"


def read_and_clean_file(uploaded_file) -> pd.DataFrame:
    """يقرأ الملف (CSV/Excel) وبيشيل أول صف بعد صف العواوين
    (اللي بيكون فيه بيانات زيادة زي أسامي بدل ما يبدأ البيانات الحقيقية)."""
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # حذف أول صف بعد العواميد (الهيدر) لو موجود
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    return df


def detect_classification_column(df: pd.DataFrame):
    """بيحاول يلاقي عمود التصنيف تلقائيًا (اللي فيه كلمة تصنيف أو نتيجة أو حالة)."""
    keywords = ["تصنيف", "التصنيف", "نتيجة", "النتيجة", "حالة", "الحالة"]
    for col in df.columns:
        col_str = str(col)
        if any(k in col_str for k in keywords):
            return col
    return df.columns[0] if len(df.columns) else None


def detect_date_column(df: pd.DataFrame):
    """بيحاول يلاقي عمود تاريخ تلقائيًا."""
    keywords = ["تاريخ", "التاريخ", "وقت", "الوقت", "date", "Date", "DATE"]
    for col in df.columns:
        col_str = str(col)
        if any(k in col_str for k in keywords):
            try:
                converted = pd.to_datetime(df[col], errors="coerce")
                if converted.notna().sum() > 0:
                    return col
            except Exception:
                continue
    return None


def classify_dataframe(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """بيحول عمود التصنيف: ناجحة -> 1 وأي حاجة تانية (غير ناجحة/فاضية) -> 0.
    بيحتفظ بعمود نصي 'الحالة' عشان نستخدمه في التشارتات."""
    out = df.copy()

    def norm(v):
        s = str(v).strip()
        return SUCCESS_LABEL if s == SUCCESS_LABEL else FAIL_LABEL

    out["الحالة"] = out[col].apply(norm)
    out[col] = out["الحالة"].map({SUCCESS_LABEL: 1, FAIL_LABEL: 0})
    return out


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="النشاط")
    buffer.seek(0)
    return buffer.getvalue()


def _plotly_dark_layout(fig, title=""):
    fig.update_layout(
        title=dict(text=title, font=dict(color=TEXT, size=15, family="Cairo")),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TEXT, family="Cairo"),
        legend=dict(font=dict(color=TEXT)),
        margin=dict(t=50, b=30, l=20, r=20),
        height=320,
    )
    return fig


def build_gauge_chart(success_rate: float):
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=success_rate,
            number={"suffix": "%", "font": {"color": ACCENT, "size": 36}},
            gauge={
                "axis": {"range": [0, 100], "tickcolor": TEXT_DIM},
                "bar": {"color": ACCENT},
                "bgcolor": "rgba(255,255,255,0.04)",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 50], "color": "rgba(248,113,113,0.18)"},
                    {"range": [50, 80], "color": "rgba(251,191,36,0.18)"},
                    {"range": [80, 100], "color": "rgba(52,211,153,0.18)"},
                ],
            },
        )
    )
    return _plotly_dark_layout(fig, "نسبة النجاح")


def build_donut_chart(success_count: int, fail_count: int):
    fig = go.Figure(
        go.Pie(
            labels=[SUCCESS_LABEL, FAIL_LABEL],
            values=[success_count, fail_count],
            hole=0.62,
            marker=dict(colors=[ACCENT, DANGER]),
            textinfo="percent",
            textfont=dict(color="#05170F", size=13),
        )
    )
    return _plotly_dark_layout(fig, "توزيع نتائج النشاط")


def build_trend_or_group_chart(df: pd.DataFrame, date_col, group_col):
    if date_col:
        tmp = df.copy()
        tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
        tmp = tmp.dropna(subset=[date_col])
        tmp["الشهر"] = tmp[date_col].dt.to_period("M").astype(str)
        grouped = tmp.groupby(["الشهر", "الحالة"]).size().unstack(fill_value=0)
        fig = go.Figure()
        if SUCCESS_LABEL in grouped:
            fig.add_bar(x=grouped.index, y=grouped[SUCCESS_LABEL], name=SUCCESS_LABEL, marker_color=ACCENT)
        if FAIL_LABEL in grouped:
            fig.add_bar(x=grouped.index, y=grouped[FAIL_LABEL], name=FAIL_LABEL, marker_color=DANGER)
        fig.update_layout(barmode="stack")
        return _plotly_dark_layout(fig, "الاتجاه الشهري للنشاط")

    if group_col:
        tmp = df.copy()
        top_values = tmp[group_col].astype(str).value_counts().head(10).index
        tmp = tmp[tmp[group_col].astype(str).isin(top_values)]
        grouped = tmp.groupby([group_col, "الحالة"]).size().unstack(fill_value=0)
        fig = go.Figure()
        if SUCCESS_LABEL in grouped:
            fig.add_bar(x=grouped.index.astype(str), y=grouped[SUCCESS_LABEL], name=SUCCESS_LABEL, marker_color=ACCENT)
        if FAIL_LABEL in grouped:
            fig.add_bar(x=grouped.index.astype(str), y=grouped[FAIL_LABEL], name=FAIL_LABEL, marker_color=DANGER)
        fig.update_layout(barmode="stack")
        return _plotly_dark_layout(fig, f"النشاط حسب {group_col}")

    counts = df["الحالة"].value_counts()
    fig = go.Figure(
        go.Bar(
            x=[SUCCESS_LABEL, FAIL_LABEL],
            y=[counts.get(SUCCESS_LABEL, 0), counts.get(FAIL_LABEL, 0)],
            marker_color=[ACCENT, DANGER],
        )
    )
    return _plotly_dark_layout(fig, "أعداد النتائج")


def render_kpi_row(classified_df: pd.DataFrame):
    total = len(classified_df)
    success_count = int((classified_df["الحالة"] == SUCCESS_LABEL).sum())
    fail_count = total - success_count
    success_rate = round((success_count / total) * 100, 1) if total else 0.0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value">{total}</div>'
            f'<div class="stat-label">إجمالي السجلات</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value">{success_count}</div>'
            f'<div class="stat-label">ناجحة</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value danger">{fail_count}</div>'
            f'<div class="stat-label">غير ناجحة</div></div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value warn">{success_rate}%</div>'
            f'<div class="stat-label">نسبة النجاح</div></div>',
            unsafe_allow_html=True,
        )

    return total, success_count, fail_count, success_rate


def render_charts_block(classified_df: pd.DataFrame, key_prefix: str, allow_group_pick: bool = True):
    total, success_count, fail_count, success_rate = render_kpi_row(classified_df)

    st.markdown('<hr class="divider" />', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(build_gauge_chart(success_rate), use_container_width=True, key=f"{key_prefix}_gauge")
    with col2:
        st.plotly_chart(build_donut_chart(success_count, fail_count), use_container_width=True, key=f"{key_prefix}_donut")

    date_col = detect_date_column(classified_df)
    group_col = None
    if not date_col and allow_group_pick:
        other_cols = [c for c in classified_df.columns if c not in ["الحالة"]]
        group_options = ["بدون تقسيم إضافي"] + other_cols
        chosen = st.selectbox(
            "تحليل النتائج حسب عمود إضافي (اختياري)",
            group_options,
            key=f"{key_prefix}_group_col",
        )
        if chosen != "بدون تقسيم إضافي":
            group_col = chosen

    st.plotly_chart(
        build_trend_or_group_chart(classified_df, date_col, group_col),
        use_container_width=True,
        key=f"{key_prefix}_trend",
    )


# ==========================================================
# تعريف أقسام السايدبار — ضيف/شيل من هنا وهيتحدث التنقل تلقائيًا
# ==========================================================

NAV_ITEMS = {
    "النشاط": "📡",
    "لوحة النشاط": "📊",
    "الوعود": "🤝",
    "الاهمال": "🗂️",
}


def render_nashat():
    """قسم النشاط — رفع الملف، تنضيفه، التصنيف، التشارتات، والتحميل."""
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-eyebrow">ACTIVITY</div>
            <p class="hero-title">النشاط</p>
            <p class="hero-subtitle">ارفع ملف Excel أو CSV عشان تبدأ</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "ارفع ملف البيانات (CSV أو Excel)", type=["csv", "xlsx", "xls"]
    )

    if uploaded_file is not None:
        file_key = (uploaded_file.name, uploaded_file.size)
        if st.session_state.get("nashat_file_key") != file_key:
            try:
                df = read_and_clean_file(uploaded_file)
            except Exception as e:
                st.error(f"مش قادر أقرأ الملف: {e}")
                st.stop()
            st.session_state["nashat_file_key"] = file_key
            st.session_state["nashat_raw_df"] = df
            st.session_state.pop("nashat_classified_df", None)

        df = st.session_state["nashat_raw_df"]

        st.markdown(
            f'<div class="card">✅ تم تحميل الملف وحذف أول صف بعد العواوين — '
            f'عدد الصفوف: <b>{len(df)}</b> | عدد الأعمدة: <b>{len(df.columns)}</b></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(df.head(20), use_container_width=True)

        st.markdown('<div class="section-title">🎯 التصنيف</div>', unsafe_allow_html=True)

        default_col = detect_classification_column(df)
        default_index = list(df.columns).index(default_col) if default_col in df.columns else 0
        classification_col = st.selectbox(
            "اختار عمود التصنيف (اللي فيه ناجحة / غير ناجحة)",
            list(df.columns),
            index=default_index,
            key="nashat_class_col_select",
        )

        if st.button("🚀 ابدأ التصنيف", key="nashat_classify_btn"):
            classified = classify_dataframe(df, classification_col)
            st.session_state["nashat_classified_df"] = classified
            st.session_state["nashat_classification_col"] = classification_col

        classified_df = st.session_state.get("nashat_classified_df")
        if classified_df is not None:
            st.markdown(
                '<div class="card">✅ تم التصنيف بنجاح — العمود اتحول لـ 1 (ناجحة) و 0 (غير ناجحة)</div>',
                unsafe_allow_html=True,
            )

            st.markdown('<div class="section-title">📊 نظرة سريعة</div>', unsafe_allow_html=True)
            render_charts_block(classified_df, key_prefix="nashat")

            st.markdown('<div class="section-title">📄 البيانات بعد التصنيف</div>', unsafe_allow_html=True)
            st.dataframe(classified_df.head(50), use_container_width=True)

            st.download_button(
                "⬇️ تحميل الملف بعد التصنيف",
                data=to_excel_bytes(classified_df),
                file_name="النشاط_بعد_التصنيف.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="nashat_download_btn",
            )
    else:
        st.markdown(
            '<div class="card" style="text-align:center; color: var(--text-dim);">'
            "📂 ارفع ملف عشان تبدأ."
            "</div>",
            unsafe_allow_html=True,
        )


def render_nashat_dashboard():
    """لوحة النشاط — داشبورد مستقل بيعرض بيانات النشاط بعد التصنيف."""
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-eyebrow">ACTIVITY DASHBOARD</div>
            <p class="hero-title">لوحة النشاط</p>
            <p class="hero-subtitle">نظرة شاملة على أداء النشاط بعد التصنيف</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    classified_df = st.session_state.get("nashat_classified_df")

    if classified_df is None:
        st.markdown(
            '<div class="card" style="text-align:center; color: var(--text-dim);">'
            "📭 لسه مفيش بيانات مصنّفة. روح تبويب <b>النشاط</b> الأول، "
            "ارفع الملف واعمل تصنيف، وهتلاقي الداشبورد هنا اتملى أوتوماتيك."
            "</div>",
            unsafe_allow_html=True,
        )
        return

    render_charts_block(classified_df, key_prefix="nashat_dashboard", allow_group_pick=True)

    st.markdown('<div class="section-title">📄 كل البيانات</div>', unsafe_allow_html=True)
    st.dataframe(classified_df, use_container_width=True)

    st.download_button(
        "⬇️ تحميل الملف بعد التصنيف",
        data=to_excel_bytes(classified_df),
        file_name="النشاط_بعد_التصنيف.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="nashat_dashboard_download_btn",
    )


def render_waeed():
    """قسم الوعود — لسه تحت الإنشاء."""
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-eyebrow">PROMISES</div>
            <p class="hero-title">الوعود</p>
            <p class="hero-subtitle">القسم ده لسه تحت الإنشاء</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="card" style="text-align:center; color: var(--text-dim);">'
        "🤝 قريبًا — قولّي المطلوب هنا بالظبط ونبنيه."
        "</div>",
        unsafe_allow_html=True,
    )


def render_ihmal():
    """قسم الاهمال — لسه تحت الإنشاء."""
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-eyebrow">NEGLECT</div>
            <p class="hero-title">الاهمال</p>
            <p class="hero-subtitle">القسم ده لسه تحت الإنشاء</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="card" style="text-align:center; color: var(--text-dim);">'
        "🗂️ قريبًا — قولّي المطلوب هنا بالظبط ونبنيه."
        "</div>",
        unsafe_allow_html=True,
    )


# ==========================================================
# الشريط الجانبي — الهوية + التنقل
# ==========================================================

with st.sidebar:
    st.markdown(
        """
        <div class="brand-block">
            <div class="brand-logo">🧭</div>
            <div>
                <div class="brand-name">النشاط</div>
                <div class="brand-sub">Dashboard</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    labels = [f"{icon}  {name}" for name, icon in NAV_ITEMS.items()]
    selected_label = st.radio(
        "التنقل", labels, label_visibility="collapsed"
    )
    selected_section = selected_label.split("  ", 1)[1]


# ==========================================================
# عرض القسم المختار
# ==========================================================

if selected_section == "النشاط":
    render_nashat()
elif selected_section == "لوحة النشاط":
    render_nashat_dashboard()
elif selected_section == "الوعود":
    render_waeed()
elif selected_section == "الاهمال":
    render_ihmal()
