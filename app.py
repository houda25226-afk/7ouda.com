"""
نقطة الدخول الرئيسية للتطبيق — تطبيق متعدد الصفحات.

الصفحات:
    - النشاط              (pages/nashat.py)      -> تصنيف إفادات المكالمات (ناجحة/غير ناجحة)
    - الوعود القائمة        (pages/waeed_qaema.py)  -> لسه تحت الإنشاء
    - الوعود المكسورة       (pages/maksora.py)      -> لسه تحت الإنشاء
    - الإهمالي              (pages/ihmali.py)       -> لسه تحت الإنشاء

طريقة التشغيل محليًا:
    pip install streamlit transformers torch pandas openpyxl --break-system-packages
    streamlit run app.py

============================================================
دليل التخصيص السريع:
- الألوان والخطوط كلها في متغير CSS_THEME تحت — عدّل من هناك بس.
- كل صفحة ملفها منفصل جوه فولدر pages/ — أي تعديل خاص بصفحة معينة يكون هناك.
- عايز تضيف صفحة جديدة؟ اعمل ملف جديد في pages/ وضيفه في قائمة PAGES تحت
  مع عنوان وأيقونة، وهيظهر تلقائيًا كزرار في السايدبار.
============================================================
"""

import streamlit as st

# ==========================================================
# إعداد الصفحة (لازم يتنفذ مرة واحدة بس، وهنا في نقطة الدخول)
# ==========================================================

st.set_page_config(
    page_title="النشاط",
    page_icon="📡",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ==========================================================
# الهوية البصرية (Theme) — مشتركة بين كل الصفحات
# ==========================================================
# لوحة الألوان:
#   خلفية داكنة رمادية-بنفسجي (#0D0F1A) بإحساس لوحة تحكم عصرية
#   سطوح البطاقات (#161927) بدرجة أفتح شوية من الخلفية
#   لون أساسي بنفسجي-إندجو (#7C5CFC) مع تركواز مكمّل (#22D3C5)
#   أخضر للنجاح (#34D399) ووردي-أحمر لعدم النجاح (#FB7185)
#   نص أساسي فاتح (#EDEFF7) ونص ثانوي رمادي مزرق (#8B90A8)

CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0D0F1A;
    --surface: #161927;
    --surface-2: #1F2436;
    --accent: #7C5CFC;
    --accent-2: #22D3C5;
    --success: #34D399;
    --danger: #FB7185;
    --text: #EDEFF7;
    --text-dim: #8B90A8;
}

html, body, [class*="css"] {
    font-family: 'Tajawal', sans-serif;
}

/* ------------------------------------------------------------
   مهم: بنخلي النص RTL من غير ما نقلب ترتيب السايدبار مع المحتوى.
   لو سبنا direction: rtl على الحاوية الكبيرة، السايدبار بيتقلب يمين.
   عشان كدا الحاوية الكبيرة بتفضل LTR، والـ RTL بيتطبق بس جوه
   المحتوى نفسه (السايدبار + المتن) عشان السايدبار يفضل شمال دايمًا.
------------------------------------------------------------ */
[data-testid="stAppViewContainer"] {
    direction: ltr;
}
.main .block-container,
[data-testid="stMain"] {
    direction: rtl;
}
section[data-testid="stSidebar"] {
    direction: rtl;
    background: #0A0C16;
    border-left: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: var(--text) !important; }

.stApp {
    background: radial-gradient(circle at 15% -5%, #1B1F35 0%, var(--bg) 55%);
    color: var(--text);
}

/* إخفاء العناصر الافتراضية الزيادة */
#MainMenu, footer, header {visibility: hidden;}

/* ===== الهيدر / Hero ===== */
.hero-wrap {
    text-align: center;
    padding: 1rem 0 0.4rem 0;
}
.hero-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.16em;
    color: var(--accent-2);
    text-transform: uppercase;
    margin-bottom: 0.4rem;
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
    margin-top: 0.5rem;
}

/* الموجة الصوتية — العنصر المميز */
.waveform {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    height: 32px;
    margin: 1rem 0 1.5rem 0;
}
.waveform span {
    display: inline-block;
    width: 3px;
    border-radius: 3px;
    background: linear-gradient(180deg, var(--accent-2), var(--accent));
    animation: wave 1.6s ease-in-out infinite;
}
@keyframes wave {
    0%, 100% { transform: scaleY(0.35); opacity: 0.55; }
    50% { transform: scaleY(1); opacity: 1; }
}

/* ===== البطاقات العامة ===== */
.card {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 14px;
    padding: 1.1rem 1.3rem;
    margin-bottom: 1rem;
}

/* ===== منطقة رفع الملفات ===== */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(124, 92, 252, 0.4);
    border-radius: 14px;
    padding: 0.6rem;
}
[data-testid="stFileUploader"] section {
    background: transparent;
}

/* ===== الأزرار ===== */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
    color: #0A0C16;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(124, 92, 252, 0.3);
    color: #0A0C16;
}

/* ===== أزرار التنقل بين الصفحات في السايدبار ===== */
[data-testid="stSidebarNav"] a,
div[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
    border-radius: 10px;
}

/* ===== شريط التقدم ===== */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
}

/* ===== المؤشرات (Metrics) ===== */
[data-testid="stMetric"] {
    background: var(--surface);
    border-radius: 12px;
    padding: 0.8rem 1rem;
    border: 1px solid rgba(255,255,255,0.06);
}
[data-testid="stMetricValue"] {
    font-family: 'JetBrains Mono', monospace;
    color: var(--accent-2);
}

/* ===== الجداول ===== */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.06);
}

/* ===== تنبيهات النجاح/الخطأ ===== */
div[data-baseweb="notification"] {
    border-radius: 10px;
}

/* فاصل بسيط بدل الخط الافتراضي */
.divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(124,92,252,0.4), transparent);
    margin: 1.4rem 0;
    border: none;
}
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


# ==========================================================
# تعريف الصفحات والتنقل — كل صفحة بتظهر كزرار في السايدبار
# ==========================================================

nashat_page = st.Page(
    "pages/nashat.py",
    title="النشاط",
    icon="📡",
    default=True,
)
waeed_qaema_page = st.Page(
    "pages/waeed_qaema.py",
    title="الوعود القائمة",
    icon="🕒",
)
maksora_page = st.Page(
    "pages/maksora.py",
    title="الوعود المكسورة",
    icon="⚠️",
)
ihmali_page = st.Page(
    "pages/ihmali.py",
    title="الإهمالي",
    icon="🗂️",
)

pg = st.navigation(
    [nashat_page, waeed_qaema_page, maksora_page, ihmali_page],
    position="sidebar",
)
pg.run()
