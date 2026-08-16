import io
import os
import tempfile
import zipfile

import pandas as pd
import streamlit as st
import torch

from transformers import AutoTokenizer, AutoModelForSequenceClassification


# ==========================================================
# إعدادات التطبيق
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

# مهم:
# 0 = غير ناجحة
# 1 = ناجحة
LABEL_MAP = {
    0: 0,
    1: 1
}


# ==========================================================
# إعداد الصفحة
# ==========================================================

st.set_page_config(
    page_title="لوحة النشاط",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================================
# CSS
# ==========================================================

CSS_THEME = """
<style>

@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&display=swap');

:root {
    --bg: #08111F;
    --surface: #101C2D;
    --surface2: #14243A;

    --green: #34D399;
    --green-dark: #1FAF78;

    --yellow: #FBBF24;
    --red: #F87171;
    --blue: #60A5FA;

    --text: #F8FAFC;
    --muted: #94A3B8;

    --border: rgba(255,255,255,0.08);
}

html,
body,
[class*="css"] {
    font-family: 'Cairo', sans-serif !important;
}

.stApp {
    background:
        radial-gradient(
            circle at 90% 0%,
            rgba(52,211,153,0.10),
            transparent 30%
        ),
        radial-gradient(
            circle at 0% 100%,
            rgba(96,165,250,0.07),
            transparent 30%
        ),
        var(--bg);
    color: var(--text);
}


/* =========================
   Main
========================= */

[data-testid="stMain"] {
    direction: rtl;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
}


/* =========================
   Sidebar
========================= */

section[data-testid="stSidebar"] {
    background: #07101C !important;
    border-right: 1px solid rgba(255,255,255,0.06);
    direction: rtl;
}

section[data-testid="stSidebar"] > div {
    padding: 1.2rem 1rem;
}

section[data-testid="stSidebar"] * {
    font-family: 'Cairo', sans-serif !important;
}


/* Sidebar Brand */

.sidebar-brand {
    background:
        linear-gradient(
            135deg,
            rgba(52,211,153,0.16),
            rgba(96,165,250,0.08)
        );

    border: 1px solid rgba(52,211,153,0.18);

    border-radius: 18px;

    padding: 18px;

    margin-bottom: 20px;

    text-align: center;
}

.sidebar-logo {
    width: 58px;
    height: 58px;

    margin: auto;

    border-radius: 17px;

    display: flex;
    align-items: center;
    justify-content: center;

    font-size: 28px;

    background:
        linear-gradient(
            135deg,
            var(--green),
            var(--blue)
        );

    box-shadow:
        0 8px 30px rgba(52,211,153,0.22);
}

.sidebar-title {
    margin-top: 10px;

    font-size: 21px;

    font-weight: 900;

    color: var(--text);
}

.sidebar-subtitle {
    margin-top: 2px;

    color: var(--muted);

    font-size: 12px;
}


/* Radio navigation */

section[data-testid="stSidebar"]
div[role="radiogroup"] {
    gap: 9px;
}

section[data-testid="stSidebar"]
div[role="radiogroup"]
label {

    background: var(--surface);

    border: 1px solid var(--border);

    border-radius: 13px;

    padding: 12px 14px;

    min-height: 48px;

    transition: 0.2s;
}

section[data-testid="stSidebar"]
div[role="radiogroup"]
label:hover {

    background: var(--surface2);

    border-color: rgba(52,211,153,0.35);
}

section[data-testid="stSidebar"]
div[role="radiogroup"]
label p {

    font-size: 15px !important;

    font-weight: 700 !important;

    color: var(--text) !important;
}


/* =========================
   Hero
========================= */

.hero {

    background:
        linear-gradient(
            135deg,
            rgba(16,28,45,0.95),
            rgba(20,36,58,0.85)
        );

    border:

        1px solid

        rgba(255,255,255,0.07);

    border-radius: 22px;

    padding: 30px;

    margin-bottom: 22px;

    text-align: center;

    box-shadow:
        0 15px 40px rgba(0,0,0,0.15);
}

.hero-icon {
    font-size: 42px;
}

.hero-title {

    font-size: 32px;

    font-weight: 900;

    margin-top: 4px;

    color: var(--text);
}

.hero-subtitle {

    color: var(--muted);

    font-size: 15px;

    margin-top: 5px;
}


/* =========================
   Cards
========================= */

.info-card {

    background: var(--surface);

    border: 1px solid var(--border);

    border-radius: 17px;

    padding: 18px;

    margin-bottom: 18px;
}

.info-card-title {

    color: var(--muted);

    font-size: 13px;

    margin-bottom: 6px;
}

.info-card-value {

    color: var(--text);

    font-size: 28px;

    font-weight: 900;
}

.info-card-green {

    border-top: 3px solid var(--green);
}

.info-card-red {

    border-top: 3px solid var(--red);
}

.info-card-blue {

    border-top: 3px solid var(--blue);
}

.info-card-yellow {

    border-top: 3px solid var(--yellow);
}


/* =========================
   Buttons
========================= */

.stButton > button {

    background:
        linear-gradient(
            135deg,
            var(--green),
            var(--green-dark)
        ) !important;

    color: #052016 !important;

    border: none !important;

    border-radius: 12px !important;

    font-weight: 800 !important;

    min-height: 48px;

    transition: 0.2s;
}

.stButton > button:hover {

    transform: translateY(-2px);

    box-shadow:
        0 10px 25px rgba(52,211,153,0.22);
}

.stDownloadButton > button {

    background:
        linear-gradient(
            135deg,
            var(--blue),
            #3B82F6
        ) !important;

    color: white !important;

    border: none !important;

    border-radius: 12px !important;

    font-weight: 800 !important;

    min-height: 48px;
}


/* =========================
   Upload
========================= */

[data-testid="stFileUploader"] {

    background: var(--surface);

    border: 1px dashed rgba(52,211,153,0.35);

    border-radius: 17px;

    padding: 12px;
}

[data-testid="stFileUploaderDropzone"] {

    background: transparent !important;
}


/* =========================
   Metrics
========================= */

[data-testid="stMetric"] {

    background: var(--surface);

    border: 1px solid var(--border);

    border-radius: 16px;

    padding: 15px;
}

[data-testid="stMetricValue"] {

    color: var(--green) !important;

    font-weight: 900 !important;
}


/* =========================
   Dataframe
========================= */

[data-testid="stDataFrame"] {

    border-radius: 15px;

    overflow: hidden;

    border: 1px solid var(--border);
}


/* =========================
   Section title
========================= */

.section-title {

    font-size: 22px;

    font-weight: 900;

    color: var(--text);

    margin-top: 15px;

    margin-bottom: 12px;
}

.section-subtitle {

    color: var(--muted);

    font-size: 13px;

    margin-bottom: 15px;
}


/* =========================
   Status
========================= */

.status-success {

    background: rgba(52,211,153,0.10);

    border: 1px solid rgba(52,211,153,0.25);

    color: #86EFAC;

    padding: 12px 16px;

    border-radius: 12px;

    margin-bottom: 15px;
}

.status-warning {

    background: rgba(251,191,36,0.10);

    border: 1px solid rgba(251,191,36,0.25);

    color: #FDE68A;

    padding: 12px 16px;

    border-radius: 12px;

    margin-bottom: 15px;
}


/* =========================
   Hide Streamlit
========================= */

#MainMenu,
footer,
header {
    visibility: hidden;
}

</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


# ==========================================================
# Session State
# ==========================================================

if "classified_df" not in st.session_state:
    st.session_state.classified_df = None

if "classified_file_name" not in st.session_state:
    st.session_state.classified_file_name = None

if "model_loaded" not in st.session_state:
    st.session_state.model_loaded = False


# ==========================================================
# Helpers
# ==========================================================

def read_uploaded_file(uploaded_file):

    if uploaded_file.name.lower().endswith(".csv"):

        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


def prepare_data(df):

    df = df.copy()

    # --------------------------------------------
    # حذف أول صف بعد العناوين
    # --------------------------------------------

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    # --------------------------------------------
    # Note -> الافادة
    # --------------------------------------------

    if "Note" in df.columns:

        df = df.rename(
            columns={
                "Note": "الافادة"
            }
        )

    return df


def restore_note_column(df):

    df = df.copy()

    if "الافادة" in df.columns:

        df = df.rename(
            columns={
                "الافادة": "Note"
            }
        )

    return df


# ==========================================================
# Model
# ==========================================================

@st.cache_resource(show_spinner=False)
def load_model():

    tokenizer = AutoTokenizer.from_pretrained(
        MODEL_REPO
    )

    model = AutoModelForSequenceClassification.from_pretrained(
        MODEL_REPO
    )

    device = (
        "cuda"
        if torch.cuda.is_available()
        else "cpu"
    )

    model.to(device)

    model.eval()

    return tokenizer, model, device


def predict_batch(
    texts,
    tokenizer,
    model,
    device,
    batch_size=16
):

    predictions = []

    confidences = []

    total = len(texts)

    progress = st.progress(
        0,
        text="جاري تصنيف البيانات..."
    )

    for i in range(
        0,
        total,
        batch_size
    ):

        batch = texts[
            i:i + batch_size
        ]

        batch = [

            str(x)
            if pd.notna(x)
            else ""

            for x in batch

        ]

        inputs = tokenizer(
            batch,
            return_tensors="pt",
            truncation=True,
            padding=True,
            max_length=MAX_LENGTH
        )

        inputs = {
            k: v.to(device)
            for k, v in inputs.items()
        }

        with torch.no_grad():

            outputs = model(
                **inputs
            )

            probabilities = torch.softmax(
                outputs.logits,
                dim=1
            )

            preds = torch.argmax(
                probabilities,
                dim=1
            )

            conf = torch.max(
                probabilities,
                dim=1
            ).values

        predictions.extend(
            preds.cpu().tolist()
        )

        confidences.extend(
            conf.cpu().tolist()
        )

        done = min(
            i + batch_size,
            total
        )

        progress.progress(
            done / total,
            text=f"جاري التصنيف... {done} / {total}"
        )

    progress.empty()

    return predictions, confidences


# ==========================================================
# Export
# ==========================================================

def create_download_file(
    df,
    original_name
):

    if original_name.lower().endswith(".csv"):

        output = df.to_csv(
            index=False
        ).encode("utf-8-sig")

        return (
            output,
            "نتائج_التصنيف.csv",
            "text/csv"
        )

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="النتائج"
        )

    return (
        buffer.getvalue(),
        "نتائج_التصنيف.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================================================
# Dashboard Charts
# ==========================================================

def render_activity_charts(df):

    st.markdown(
        '<div class="section-title">📊 تحليل النشاط</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------
    # Classification distribution
    # --------------------------------------------

    if "التصنيف_المتوقع" in df.columns:

        chart_data = (
            df["التصنيف_المتوقع"]
            .value_counts()
            .sort_index()
        )

        chart_df = pd.DataFrame(
            {
                "عدد الإفادات": chart_data
            }
        )

        chart_df.index = [
            "غير ناجحة (0)"
            if x == 0
            else "ناجحة (1)"
            for x in chart_df.index
        ]

        st.markdown(
            "### 📌 توزيع نتائج التصنيف"
        )

        st.bar_chart(
            chart_df,
            use_container_width=True
        )

    # --------------------------------------------
    # Confidence
    # --------------------------------------------

    if "نسبة_الثقة" in df.columns:

        st.markdown(
            "### 🎯 توزيع مستوى الثقة"
        )

        confidence_series = pd.to_numeric(
            df["نسبة_الثقة"],
            errors="coerce"
        ).dropna()

        if len(confidence_series) > 0:

            bins = [
                0,
                50,
                60,
                70,
                80,
                90,
                100
            ]

            labels = [
                "أقل من 50%",
                "50-60%",
                "60-70%",
                "70-80%",
                "80-90%",
                "90-100%"
            ]

            confidence_groups = pd.cut(
                confidence_series,
                bins=bins,
                labels=labels,
                include_lowest=True
            )

            confidence_chart = (
                confidence_groups
                .value_counts()
                .sort_index()
            )

            st.bar_chart(
                confidence_chart,
                use_container_width=True
            )

    # --------------------------------------------
    # Date analysis
    # --------------------------------------------

    date_columns = [

        col for col in df.columns

        if any(
            word in str(col).lower()
            for word in [
                "date",
                "التاريخ",
                "تاريخ"
            ]
        )

    ]

    if date_columns:

        date_col = date_columns[0]

        temp = df.copy()

        temp[date_col] = pd.to_datetime(
            temp[date_col],
            errors="coerce"
        )

        temp = temp.dropna(
            subset=[date_col]
        )

        if len(temp) > 0:

            temp["اليوم"] = (
                temp[date_col]
                .dt.date
            )

            daily = (
                temp.groupby("اليوم")
                .size()
            )

            st.markdown(
                "### 📅 النشاط اليومي"
            )

            st.line_chart(
                daily,
                use_container_width=True
            )

    # --------------------------------------------
    # Collector analysis
    # --------------------------------------------

    collector_columns = [

        col for col in df.columns

        if any(
            word in str(col).lower()
            for word in [
                "collector",
                "محصل",
                "المحصل",
                "اسم المحصل",
                "user",
                "agent"
            ]
        )

    ]

    if collector_columns:

        collector_col = collector_columns[0]

        collector_stats = (
            df[collector_col]
            .astype(str)
            .value_counts()
            .head(15)
        )

        st.markdown(
            "### 👤 نشاط المحصلين"
        )

        st.bar_chart(
            collector_stats,
            use_container_width=True
        )


# ==========================================================
# Sidebar
# ==========================================================

NAV_ITEMS = {
    "النشاط": "📡",
    "Dashboard النشاط": "📊",
    "الوعود المكسورة": "🤝",
    "الإهمال": "🗂️"
}


with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">

            <div class="sidebar-logo">
                📡
            </div>

            <div class="sidebar-title">
                لوحة النشاط
            </div>

            <div class="sidebar-subtitle">
                Activity Intelligence
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    labels = [
        f"{icon}  {name}"
        for name, icon
        in NAV_ITEMS.items()
    ]

    selected = st.radio(
        "القائمة",
        labels,
        label_visibility="collapsed"
    )

    selected_section = selected.split(
        "  ",
        1
    )[1]

    st.markdown("---")

    st.markdown(
        "### 🤖 الموديل"
    )

    st.caption(
        "الموديل المستخدم من Hugging Face"
    )

    st.code(
        MODEL_REPO,
        language=None
    )

    if st.button(
        "🤖 تحميل الموديل",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "جاري تحميل الموديل..."
            ):

                load_model()

            st.session_state.model_loaded = True

            st.success(
                "الموديل جاهز ✅"
            )

        except Exception as e:

            st.error(
                f"حصل خطأ أثناء تحميل الموديل: {e}"
            )

    if st.session_state.model_loaded:

        st.success(
            "الموديل محمل وجاهز"
        )


# ==========================================================
# TAB 1 — النشاط
# ==========================================================

def render_activity():

    st.markdown(
        """
        <div class="hero">

            <div class="hero-icon">
                📡
            </div>

            <div class="hero-title">
                تصنيف نشاط المكالمات
            </div>

            <div class="hero-subtitle">
                ارفع ملف النشاط، وبعدها ابدأ التصنيف بالموديل
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "📂 ارفع ملف النشاط",
        type=[
            "csv",
            "xlsx",
            "xls"
        ],
        key="activity_upload"
    )

    if uploaded_file is None:

        st.markdown(
            """
            <div class="info-card">

            <b>طريقة الاستخدام</b>

            <br><br>

            1️⃣ ارفع ملف Excel أو CSV

            <br>

            2️⃣ أول صف بعد أسماء الأعمدة سيتم حذفه تلقائيًا

            <br>

            3️⃣ عمود Note سيتم استخدامه كـ الافادة

            <br>

            4️⃣ اضغط "ابدأ التصنيف"

            <br>

            5️⃣ الناجحة = 1 وغير الناجحة = 0

            </div>
            """,
            unsafe_allow_html=True
        )

        return

    # --------------------------------------------
    # Read
    # --------------------------------------------

    try:

        df = read_uploaded_file(
            uploaded_file
        )

    except Exception as e:

        st.error(
            f"مش قادر أقرأ الملف: {e}"
        )

        return

    original_rows = len(df)

    # --------------------------------------------
    # Prepare
    # --------------------------------------------

    df = prepare_data(df)

    st.markdown(
        '<div class="status-success">'
        'تم تحميل الملف وتجهيزه بنجاح ✅'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------
    # Metrics
    # --------------------------------------------

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📄 الصفوف الأصلية",
        original_rows
    )

    c2.metric(
        "📄 الصفوف بعد التنظيف",
        len(df)
    )

    c3.metric(
        "📋 عدد الأعمدة",
        len(df.columns)
    )

    c4.metric(
        "📝 عمود الإفادة",
        "موجود"
        if "الافادة" in df.columns
        else "غير موجود"
    )

    # --------------------------------------------
    # Validate
    # --------------------------------------------

    if "الافادة" not in df.columns:

        st.error(
            "❌ مش لاقي عمود Note في الملف."
        )

        st.info(
            "لازم يكون عندك عمود اسمه Note."
        )

        return

    # --------------------------------------------
    # Preview
    # --------------------------------------------

    with st.expander(
        "👀 معاينة البيانات قبل التصنيف",
        expanded=False
    ):

        st.dataframe(
            df.head(20),
            use_container_width=True
        )

    st.markdown(
        '<div style="height:10px"></div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------
    # Start Classification
    # --------------------------------------------

    if st.button(
        "🚀 ابدأ التصنيف",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "جاري تجهيز الموديل..."
            ):

                tokenizer, model, device = load_model()

            texts = df[
                "الافادة"
            ].tolist()

            preds, confidences = predict_batch(
                texts,
                tokenizer,
                model,
                device
            )

            result_df = df.copy()

            # ------------------------------------
            # Numeric classification
            # 0 = غير ناجحة
            # 1 = ناجحة
            # ------------------------------------

            result_df[
                "التصنيف_المتوقع"
            ] = [
                LABEL_MAP[int(p)]
                for p in preds
            ]

            result_df[
                "نسبة_الثقة"
            ] = [
                round(
                    float(c) * 100,
                    1
                )
                for c in confidences
            ]

            # ------------------------------------
            # Return الافادة -> Note
            # ------------------------------------

            result_df = restore_note_column(
                result_df
            )

            # ------------------------------------
            # Save
            # ------------------------------------

            st.session_state.classified_df = result_df

            st.session_state.classified_file_name = (
                uploaded_file.name
            )

            st.success(
                "تم تصنيف الملف بنجاح ✅"
            )

        except Exception as e:

            st.error(
                f"حصل خطأ أثناء التصنيف: {e}"
            )

            return

    # ======================================================
    # Results
    # ======================================================

    if st.session_state.classified_df is not None:

        result_df = st.session_state.classified_df

        st.markdown(
            '<div class="section-title">'
            '📊 نتائج التصنيف'
            '</div>',
            unsafe_allow_html=True
        )

        # --------------------------------------------
        # Counts
        # --------------------------------------------

        if "التصنيف_المتوقع" in result_df.columns:

            success_count = int(
                (
                    result_df[
                        "التصنيف_المتوقع"
                    ] == 1
                ).sum()
            )

            failed_count = int(
                (
                    result_df[
                        "التصنيف_المتوقع"
                    ] == 0
                ).sum()
            )

            total_count = len(
                result_df
            )

            success_rate = (
                success_count /
                total_count *
                100
                if total_count
                else 0
            )

            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "📊 إجمالي الإفادات",
                total_count
            )

            c2.metric(
                "🟢 ناجحة",
                success_count
            )

            c3.metric(
                "🔴 غير ناجحة",
                failed_count
            )

            c4.metric(
                "📈 نسبة النجاح",
                f"{success_rate:.1f}%"
            )

        # --------------------------------------------
        # Preview result
        # --------------------------------------------

        st.markdown(
            "### 📋 عينة من النتائج"
        )

        st.dataframe(
            result_df.head(30),
            use_container_width=True
        )

        # --------------------------------------------
        # Charts
        # --------------------------------------------

        render_activity_charts(
            result_df
        )

        # --------------------------------------------
        # Download
        # --------------------------------------------

        output, file_name, mime = create_download_file(
            result_df,
            st.session_state.classified_file_name
        )

        st.markdown(
            '<div class="section-title">'
            '⬇️ تحميل النتائج'
            '</div>',
            unsafe_allow_html=True
        )

        st.download_button(
            "⬇️ تحميل الملف المصنف",
            data=output,
            file_name=file_name,
            mime=mime,
            use_container_width=True
        )


# ==========================================================
# TAB 2 — Dashboard النشاط
# ==========================================================

def render_activity_dashboard():

    st.markdown(
        """
        <div class="hero">

            <div class="hero-icon">
                📊
            </div>

            <div class="hero-title">
                Dashboard النشاط
            </div>

            <div class="hero-subtitle">
                ارفع الملف بعد التصنيف وشوف تحليل النشاط بشكل تفاعلي
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    dashboard_file = st.file_uploader(
        "📂 ارفع ملف النشاط بعد التصنيف",
        type=[
            "csv",
            "xlsx",
            "xls"
        ],
        key="dashboard_upload"
    )

    if dashboard_file is None:

        st.markdown(
            """
            <div class="info-card">

            📌 <b>هنا ترفع الملف المصنف يدويًا.</b>

            <br><br>

            Dashboard هتقرأ البيانات الموجودة في الملف
            وتعمل الإحصائيات والرسوم تلقائيًا.

            </div>
            """,
            unsafe_allow_html=True
        )

        return

    try:

        df = read_uploaded_file(
            dashboard_file
        )

    except Exception as e:

        st.error(
            f"مش قادر أقرأ الملف: {e}"
        )

        return

    st.success(
        f"تم تحميل الملف: {dashboard_file.name} ✅"
    )

    # ======================================================
    # Main metrics
    # ======================================================

    total = len(df)

    success = 0
    failed = 0

    if "التصنيف_المتوقع" in df.columns:

        classification = pd.to_numeric(
            df["التصنيف_المتوقع"],
            errors="coerce"
        )

        success = int(
            (classification == 1).sum()
        )

        failed = int(
            (classification == 0).sum()
        )

    success_rate = (
        success / total * 100
        if total
        else 0
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📄 إجمالي النشاط",
        total
    )

    c2.metric(
        "🟢 ناجحة",
        success
    )

    c3.metric(
        "🔴 غير ناجحة",
        failed
    )

    c4.metric(
        "📈 نسبة النجاح",
        f"{success_rate:.1f}%"
    )

    st.markdown(
        "---"
    )

    # ======================================================
    # Charts
    # ======================================================

    render_activity_charts(
        df
    )

    # ======================================================
    # Raw data
    # ======================================================

    st.markdown(
        "### 📋 البيانات"
    )

    st.dataframe(
        df,
        use_container_width=True,
        height=450
    )


# ==========================================================
# TAB 3 — الوعود المكسورة
# ==========================================================

def render_broken_promises():

    st.markdown(
        """
        <div class="hero">

            <div class="hero-icon">
                🤝
            </div>

            <div class="hero-title">
                الوعود المكسورة
            </div>

            <div class="hero-subtitle">
                القسم جاهز — هنضيف منطق الوعود المكسورة بعدين
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="info-card">

        🚧 <b>القسم تحت التجهيز</b>

        <br><br>

        لما تحددلي قواعد الوعود المكسورة،
        هنضيف التصنيف والتحليل والـ Dashboard هنا.

        </div>
        """,
        unsafe_allow_html=True
    )


# ==========================================================
# TAB 4 — الإهمال
# ==========================================================

def render_neglect():

    st.markdown(
        """
        <div class="hero">

            <div class="hero-icon">
                🗂️
            </div>

            <div class="hero-title">
                الإهمال
            </div>

            <div class="hero-subtitle">
                القسم جاهز — هنضيف منطق الإهمال بعدين
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="info-card">

        🚧 <b>القسم تحت التجهيز</b>

        <br><br>

        هنضيف قواعد الإهمال والتصنيف
        والرسوم البيانية لما تحددلي المطلوب.

        </div>
        """,
        unsafe_allow_html=True
    )


# ==========================================================
# Navigation
# ==========================================================

if selected_section == "النشاط":

    render_activity()

elif selected_section == "Dashboard النشاط":

    render_activity_dashboard()

elif selected_section == "الوعود المكسورة":

    render_broken_promises()

elif selected_section == "الإهمال":

    render_neglect()
