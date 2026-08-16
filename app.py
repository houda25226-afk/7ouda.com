"""
تطبيق Streamlit — قاعدة أساسية (Base) بسايدبار وتنقل بين أقسام.
هنبني عليها بعدين قسم قسم حسب المطلوب.

طريقة التشغيل محليًا:
    pip install streamlit pandas --break-system-packages
    streamlit run app.py

============================================================
دليل التخصيص السريع:
- الألوان والخطوط كلها في متغير CSS_THEME تحت — عدّل من هناك بس.
- كل قسم من أقسام السايدبار له دالة render_* منفصلة تحت — سهل تضيف/تعدّل قسم من غير ما تلخبط الباقي.
- قائمة الأقسام نفسها في NAV_ITEMS تحت — ضيف أو شيل منها وهيتحدث السايدبار تلقائيًا.
============================================================
"""

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

#MainMenu, footer, header {visibility: hidden;}

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

/* قائمة التنقل في السايدبار (radio نخليها تبان كأزرار) */
section[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 0.35rem;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 11px;
    padding: 0.55rem 0.8rem;
    width: 100%;
    transition: background 0.15s ease, border-color 0.15s ease;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background: var(--surface-2);
    border-color: rgba(52, 211, 153, 0.35);
}
section[data-testid="stSidebar"] div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
    display: none;
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
}
.stat-card .stat-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 1.6rem;
    font-weight: 700;
    color: var(--accent);
}
.stat-card .stat-label {
    color: var(--text-dim);
    font-size: 0.82rem;
    margin-top: 0.3rem;
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

/* ===== الأزرار ===== */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), #22B888);
    color: #05170F;
    font-weight: 700;
    border: none;
    border-radius: 11px;
    padding: 0.65rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
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
# تعريف أقسام السايدبار — ضيف/شيل من هنا وهيتحدث التنقل تلقائيًا
# ==========================================================

NAV_ITEMS = {
    "الرئيسية": "🏠",
    # هنضيف هنا أقسام تانية بعدين (زي: النشاط، الوعود القائمة... إلخ)
}


def render_home():
    """القسم الرئيسي — نقطة بداية بسيطة، هنستبدلها بمحتوى فعلي بعدين."""
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-eyebrow">DASHBOARD</div>
            <p class="hero-title">أهلاً بيك 👋</p>
            <p class="hero-subtitle">دي نقطة البداية للتطبيق — هنضيف الأقسام واحد واحد من هنا</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            '<div class="stat-card"><div class="stat-value">0</div>'
            '<div class="stat-label">قسم شغال</div></div>',
            unsafe_allow_html=True,
        )
    with col2:
        st.markdown(
            '<div class="stat-card"><div class="stat-value">—</div>'
            '<div class="stat-label">آخر تحديث</div></div>',
            unsafe_allow_html=True,
        )
    with col3:
        st.markdown(
            '<div class="stat-card"><div class="stat-value">✓</div>'
            '<div class="stat-label">الحالة</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="card">'
        "🧭 قولّي إيه القسم اللي عايز نضيفه بعدين (زي صفحة تصنيف المكالمات، "
        "أو أي قسم تاني) ونبنيه هنا خطوة خطوة."
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

if selected_section == "الرئيسية":
    render_home()
