import streamlit as st
import pandas as pd

# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(
    page_title="لوحة النشاط",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================
# THEME
# ==========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&display=swap');

:root {
    --bg: #07111F;
    --sidebar: #0B1626;
    --card: #101D30;
    --card2: #14243A;

    --green: #34D399;
    --green-dark: #10B981;
    --yellow: #FBBF24;

    --white: #F8FAFC;
    --text: #E5E7EB;
    --muted: #94A3B8;

    --border: rgba(255,255,255,0.08);
}

/* =========================
   GENERAL
========================= */

html, body, [class*="css"] {
    font-family: 'Cairo', sans-serif !important;
}

.stApp {
    background:
        radial-gradient(
            circle at 85% 0%,
            rgba(52,211,153,0.08),
            transparent 35%
        ),
        var(--bg);
}

/* Main */

[data-testid="stAppViewContainer"] {
    direction: ltr;
}

[data-testid="stMain"] {
    direction: rtl;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    max-width: 1450px !important;
}

/* =========================
   SIDEBAR
========================= */

section[data-testid="stSidebar"] {
    background: var(--sidebar) !important;
    border-right: 1px solid var(--border);
    direction: rtl !important;
}

section[data-testid="stSidebar"] > div {
    padding: 1.5rem 1rem !important;
}

/* كل النصوص في السايدبار */

section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] div {
    color: var(--text) !important;
}

/* Brand */

.sidebar-brand {
    background: linear-gradient(
        145deg,
        #12243A,
        #0D1A2B
    );

    border: 1px solid rgba(52,211,153,0.18);

    border-radius: 18px;

    padding: 18px;

    margin-bottom: 25px;

    text-align: center;

    box-shadow:
        0 10px 30px rgba(0,0,0,0.2);
}

.sidebar-logo {
    font-size: 2.5rem;
    margin-bottom: 5px;
}

.sidebar-title {
    color: var(--white) !important;
    font-size: 1.35rem;
    font-weight: 900;
}

.sidebar-subtitle {
    color: var(--muted) !important;
    font-size: 0.75rem;
    margin-top: 3px;
}

/* Navigation */

section[data-testid="stSidebar"]
div[role="radiogroup"] {
    gap: 10px !important;
}

/* نخفي الـ radio الحقيقي */

section[data-testid="stSidebar"]
div[role="radiogroup"] label
[data-baseweb="radio"] {
    display: none !important;
}

/* Navigation buttons */

section[data-testid="stSidebar"]
div[role="radiogroup"] label {

    background: #101D30 !important;

    border: 1px solid rgba(255,255,255,0.06) !important;

    border-radius: 13px !important;

    padding: 14px 16px !important;

    min-height: 52px !important;

    display: flex !important;

    align-items: center !important;

    justify-content: flex-start !important;

    transition: all 0.2s ease;

    cursor: pointer;

}

/* Hover */

section[data-testid="stSidebar"]
div[role="radiogroup"] label:hover {

    background: #16283F !important;

    border-color: rgba(52,211,153,0.4) !important;

    transform: translateX(-3px);

}

/* النص */

section[data-testid="stSidebar"]
div[role="radiogroup"] label p {

    color: #F8FAFC !important;

    font-size: 1rem !important;

    font-weight: 700 !important;

    margin: 0 !important;

}

/* selected */

section[data-testid="stSidebar"]
div[role="radiogroup"]
label:has(input:checked) {

    background:
        linear-gradient(
            90deg,
            rgba(52,211,153,0.18),
            rgba(52,211,153,0.06)
        ) !important;

    border-color: rgba(52,211,153,0.65) !important;

}

/* =========================
   HERO
========================= */

.hero {

    background:
        linear-gradient(
            135deg,
            rgba(52,211,153,0.10),
            rgba(20,36,58,0.8)
        );

    border: 1px solid rgba(52,211,153,0.16);

    border-radius: 22px;

    padding: 30px;

    margin-bottom: 25px;

    text-align: right;

}

.hero-icon {
    font-size: 2.8rem;
}

.hero-title {

    color: #FFFFFF !important;

    font-size: 2.2rem;

    font-weight: 900;

    margin: 0;

}

.hero-subtitle {

    color: #CBD5E1 !important;

    font-size: 1rem;

    margin-top: 8px;

}

/* =========================
   CARDS
========================= */

.card {

    background: var(--card);

    border: 1px solid var(--border);

    border-radius: 18px;

    padding: 22px;

    margin-bottom: 20px;

}

/* =========================
   UPLOADER
========================= */

[data-testid="stFileUploader"] {

    background: #0D1A2B !important;

    border: 1px dashed rgba(52,211,153,0.45) !important;

    border-radius: 18px !important;

    padding: 15px !important;

}

[data-testid="stFileUploaderDropzone"] {

    background: transparent !important;

}

[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small {

    color: #CBD5E1 !important;

}

/* Browse button */

[data-testid="stFileUploader"] button {

    background: var(--green) !important;

    color: #052016 !important;

    border: none !important;

    border-radius: 10px !important;

    font-weight: 800 !important;

}

/* =========================
   BUTTONS
========================= */

.stButton > button {

    width: 100%;

    background:
        linear-gradient(
            90deg,
            var(--green),
            var(--green-dark)
        ) !important;

    color: #032016 !important;

    border: none !important;

    border-radius: 12px !important;

    padding: 0.75rem 1rem !important;

    font-weight: 900 !important;

    font-size: 1rem !important;

    transition: all 0.2s ease;

}

.stButton > button:hover {

    transform: translateY(-2px);

    box-shadow:
        0 10px 25px rgba(52,211,153,0.25);

}

/* =========================
   METRICS
========================= */

[data-testid="stMetric"] {

    background: var(--card);

    border: 1px solid var(--border);

    border-radius: 16px;

    padding: 18px;

}

[data-testid="stMetricLabel"] {

    color: var(--muted) !important;

}

[data-testid="stMetricValue"] {

    color: var(--green) !important;

    font-weight: 900 !important;

}

/* =========================
   DATAFRAME
========================= */

[data-testid="stDataFrame"] {

    border-radius: 15px;

    overflow: hidden;

}

/* =========================
   DOWNLOAD
========================= */

.stDownloadButton > button {

    width: 100%;

    background: #16283F !important;

    color: #F8FAFC !important;

    border: 1px solid rgba(52,211,153,0.35) !important;

    border-radius: 12px !important;

    font-weight: 700 !important;

}

.stDownloadButton > button:hover {

    border-color: var(--green) !important;

}

/* =========================
   DIVIDER
========================= */

.divider {

    height: 1px;

    background:
        linear-gradient(
            90deg,
            transparent,
            rgba(52,211,153,0.4),
            transparent
        );

    margin: 25px 0;

}

</style>
""", unsafe_allow_html=True)


# ==========================================================
# NAVIGATION
# ==========================================================

NAV_ITEMS = {
    "النشاط": "📡",
    "الوعود": "🤝",
    "الإهمال": "🗂️",
}


# ==========================================================
# SIDEBAR
# ==========================================================

with st.sidebar:

    st.markdown("""
    <div class="sidebar-brand">

        <div class="sidebar-logo">
            📡
        </div>

        <div class="sidebar-title">
            لوحة النشاط
        </div>

        <div class="sidebar-subtitle">
            Activity Dashboard
        </div>

    </div>
    """, unsafe_allow_html=True)

    labels = [
        f"{icon}  {name}"
        for name, icon in NAV_ITEMS.items()
    ]

    selected_label = st.radio(
        "القائمة",
        labels,
        label_visibility="collapsed"
    )

    selected_section = selected_label.split("  ", 1)[1]

    st.markdown(
        "<div style='height:25px'></div>",
        unsafe_allow_html=True
    )

    st.caption("Dashboard v2.0")


# ==========================================================
# ACTIVITY
# ==========================================================

def render_nashat():

    # Hero
    st.markdown("""
    <div class="hero">

        <div class="hero-icon">
            📡
        </div>

        <div class="hero-title">
            لوحة النشاط
        </div>

        <div class="hero-subtitle">
            ارفع ملف البيانات وشغّل نموذج التصنيف للحصول على تحليل كامل للنشاط.
        </div>

    </div>
    """, unsafe_allow_html=True)


    # Upload
    st.markdown("""
    <div class="card">

        <h3 style="color:#F8FAFC !important;">
            📁 رفع ملف البيانات
        </h3>

        <p style="color:#94A3B8 !important;">
            ارفع ملف Excel أو CSV ثم اضغط على
            <b style="color:#34D399;">ابدأ التصنيف</b>
        </p>

    </div>
    """, unsafe_allow_html=True)


    uploaded_file = st.file_uploader(
        "رفع ملف البيانات",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed"
    )


    if uploaded_file is None:

        st.markdown("""
        <div class="card" style="text-align:center;">

            <div style="
                font-size:3rem;
                margin-bottom:10px;
            ">
                📂
            </div>

            <h3 style="color:#F8FAFC !important;">
                لم يتم رفع ملف بعد
            </h3>

            <p style="color:#94A3B8 !important;">
                ارفع ملف البيانات من الأعلى لبدء التحليل.
            </p>

        </div>
        """, unsafe_allow_html=True)

        return


    # ======================================================
    # READ FILE
    # ======================================================

    try:

        if uploaded_file.name.lower().endswith(".csv"):

            df = pd.read_csv(uploaded_file)

        else:

            df = pd.read_excel(uploaded_file)

    except Exception as e:

        st.error(
            f"حدث خطأ أثناء قراءة الملف: {e}"
        )

        return


    # ======================================================
    # REMOVE FIRST DATA ROW
    # ======================================================

    if len(df) > 0:

        df = df.iloc[1:].reset_index(drop=True)


    # ======================================================
    # FILE INFO
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.markdown("### 📊 معلومات الملف")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric(
            "عدد الصفوف",
            f"{len(df):,}"
        )

    with c2:
        st.metric(
            "عدد الأعمدة",
            f"{len(df.columns):,}"
        )

    with c3:
        st.metric(
            "حالة الملف",
            "جاهز"
        )

    with c4:
        st.metric(
            "اسم الملف",
            uploaded_file.name[:18]
        )


    # ======================================================
    # START CLASSIFICATION
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.markdown("### 🤖 التصنيف")

    if "classified" not in st.session_state:

        st.session_state.classified = False


    start_classification = st.button(
        "🚀 ابدأ التصنيف",
        key="start_classification"
    )


    if start_classification:

        with st.spinner("جاري تصنيف البيانات..."):

            # ==============================================
            # هنا هنحط استدعاء موديل HuggingFace
            # ==============================================

            # مثال مؤقت:
            #
            # predictions = model(...)
            #
            # وبعدها:
            #
            # df["التصنيف"] = predictions

            # مؤقت للتجربة فقط
            df["التصنيف"] = 1

            st.session_state.result_df = df.copy()

            st.session_state.classified = True

        st.success("✅ تم الانتهاء من التصنيف بنجاح")


    # ======================================================
    # DASHBOARD
    # ======================================================

    if st.session_state.classified:

        result_df = st.session_state.result_df

        st.markdown(
            '<div class="divider"></div>',
            unsafe_allow_html=True
        )

        st.markdown("## 📈 Dashboard")

        total = len(result_df)

        success = int(
            (result_df["التصنيف"] == 1).sum()
        )

        failed = int(
            (result_df["التصنيف"] == 0).sum()
        )

        success_rate = (
            success / total * 100
            if total > 0 else 0
        )


        # ==============================================
        # KPIs
        # ==============================================

        k1, k2, k3, k4 = st.columns(4)

        with k1:
            st.metric(
                "إجمالي الحالات",
                f"{total:,}"
            )

        with k2:
            st.metric(
                "ناجحة",
                f"{success:,}"
            )

        with k3:
            st.metric(
                "غير ناجحة",
                f"{failed:,}"
            )

        with k4:
            st.metric(
                "نسبة النجاح",
                f"{success_rate:.1f}%"
            )


        st.markdown(
            '<div class="divider"></div>',
            unsafe_allow_html=True
        )


        # ==============================================
        # CHARTS
        # ==============================================

        import plotly.express as px


        col1, col2 = st.columns(2)


        # -----------------------------
        # Pie
        # -----------------------------

        with col1:

            chart_df = pd.DataFrame({
                "الحالة": [
                    "ناجحة",
                    "غير ناجحة"
                ],
                "العدد": [
                    success,
                    failed
                ]
            })

            fig = px.pie(
                chart_df,
                names="الحالة",
                values="العدد",
                hole=0.55,
                title="توزيع نتائج التصنيف"
            )

            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#F8FAFC",
                legend_title_text=""
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


        # -----------------------------
        # Bar
        # -----------------------------

        with col2:

            fig2 = px.bar(
                chart_df,
                x="الحالة",
                y="العدد",
                text="العدد",
                title="عدد الحالات حسب التصنيف"
            )

            fig2.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#F8FAFC",
                xaxis_title="",
                yaxis_title="عدد الحالات"
            )

            fig2.update_traces(
                textposition="outside"
            )

            st.plotly_chart(
                fig2,
                use_container_width=True
            )


        # ==============================================
        # DATA
        # ==============================================

        st.markdown("### 📋 البيانات بعد التصنيف")

        st.dataframe(
            result_df,
            use_container_width=True,
            height=450
        )


        # ==============================================
        # DOWNLOAD
        # ==============================================

        st.markdown(
            '<div class="divider"></div>',
            unsafe_allow_html=True
        )

        csv_data = result_df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ تحميل ملف النتائج",
            data=csv_data,
            file_name="activity_classified.csv",
            mime="text/csv"
        )


# ==========================================================
# PROMISES
# ==========================================================

def render_waeed():

    st.markdown("""
    <div class="hero">

        <div class="hero-icon">
            🤝
        </div>

        <div class="hero-title">
            الوعود
        </div>

        <div class="hero-subtitle">
            قسم متابعة وتحليل الوعود.
        </div>

    </div>
    """, unsafe_allow_html=True)

    st.info("🚧 القسم تحت الإنشاء")


# ==========================================================
# NEGLECT
# ==========================================================

def render_ihmal():

    st.markdown("""
    <div class="hero">

        <div class="hero-icon">
            🗂️
        </div>

        <div class="hero-title">
            الإهمال
        </div>

        <div class="hero-subtitle">
            قسم تحليل حالات الإهمال.
        </div>

    </div>
    """, unsafe_allow_html=True)

    st.info("🚧 القسم تحت الإنشاء")


# ==========================================================
# ROUTING
# ==========================================================

if selected_section == "النشاط":

    render_nashat()

elif selected_section == "الوعود":

    render_waeed()

elif selected_section == "الإهمال":

    render_ihmal()
