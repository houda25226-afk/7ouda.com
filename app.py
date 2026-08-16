import io
import re
import numpy as np
import pandas as pd
import streamlit as st

from transformers import (
    AutoTokenizer,
    AutoModelForSequenceClassification
)
import torch


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="لوحة النشاط",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# SETTINGS
# =========================================================

# ضع هنا اسم موديل Hugging Face
# مثال:
# HF_MODEL_ID = "username/my-arabic-classifier"

HF_MODEL_ID = "Mahmoud252002/7oudaModel"

MAX_LENGTH = 256


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>

    /* =========================
       GENERAL
    ========================= */

    .stApp {
        background: #08111F;
    }

    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    h1, h2, h3, h4, p, label {
        font-family: Arial, sans-serif !important;
    }

    h1, h2, h3, h4 {
        color: #F8FAFC !important;
    }

    p, label {
        color: #CBD5E1 !important;
    }


    /* =========================
       SIDEBAR
    ========================= */

    section[data-testid="stSidebar"] {
        background: #07101D;
        border-right: 1px solid #1E293B;
    }

    section[data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    .sidebar-brand {
        background: #0F1C2E;
        border: 1px solid #24344D;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 25px;
        text-align: center;
    }

    .sidebar-brand-icon {
        font-size: 38px;
        margin-bottom: 8px;
    }

    .sidebar-brand-title {
        font-size: 22px;
        font-weight: 800;
        color: #F8FAFC;
    }

    .sidebar-brand-sub {
        font-size: 12px;
        color: #94A3B8;
        margin-top: 4px;
    }


    /* =========================
       HERO
    ========================= */

    .hero {
        background: linear-gradient(
            135deg,
            #102A31,
            #0E1B2D
        );

        border: 1px solid #1E8069;
        border-radius: 22px;
        padding: 32px;
        margin-bottom: 25px;
    }

    .hero-title {
        font-size: 36px;
        font-weight: 900;
        color: #F8FAFC;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        font-size: 16px;
        color: #94A3B8;
    }


    /* =========================
       CARDS
    ========================= */

    .info-card {
        background: #101C2E;
        border: 1px solid #24344D;
        border-radius: 18px;
        padding: 22px;
        margin-bottom: 18px;
    }

    .info-title {
        color: #F8FAFC;
        font-size: 18px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .info-text {
        color: #94A3B8;
        font-size: 14px;
    }


    /* =========================
       METRICS
    ========================= */

    div[data-testid="stMetric"] {
        background: #101C2E;
        border: 1px solid #24344D;
        border-radius: 18px;
        padding: 18px;
    }

    div[data-testid="stMetricLabel"] {
        color: #94A3B8 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #34D399 !important;
        font-weight: 800;
    }


    /* =========================
       BUTTON
    ========================= */

    .stButton > button {
        width: 100%;
        background: linear-gradient(
            90deg,
            #34D399,
            #10B981
        );

        color: #062017 !important;
        border: none;
        border-radius: 12px;
        font-weight: 800;
        padding: 12px 20px;
        min-height: 48px;
    }

    .stButton > button:hover {
        background: linear-gradient(
            90deg,
            #6EE7B7,
            #34D399
        );

        color: #052E1B !important;
    }


    /* =========================
       FILE UPLOADER
    ========================= */

    [data-testid="stFileUploader"] {
        background: #101C2E;
        border: 1px dashed #34D399;
        border-radius: 18px;
        padding: 18px;
    }


    /* =========================
       DATAFRAME
    ========================= */

    [data-testid="stDataFrame"] {
        border: 1px solid #24344D;
        border-radius: 14px;
        overflow: hidden;
    }


    /* =========================
       DIVIDER
    ========================= */

    .section-line {
        height: 1px;
        background: #24344D;
        margin: 28px 0;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# MODEL
# =========================================================

@st.cache_resource
def load_model():

    if (
        HF_MODEL_ID == ""
        or HF_MODEL_ID == "PUT_YOUR_HUGGINGFACE_MODEL_HERE"
    ):
        return None, None

    tokenizer = AutoTokenizer.from_pretrained(
        HF_MODEL_ID
    )

    model = AutoModelForSequenceClassification.from_pretrained(
        HF_MODEL_ID
    )

    device = torch.device(
        "cuda" if torch.cuda.is_available() else "cpu"
    )

    model.to(device)
    model.eval()

    return tokenizer, model


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_text(text):

    if pd.isna(text):
        return ""

    text = str(text)

    text = re.sub(
        r"https?://\S+|www\.\S+",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# CLASSIFICATION
# =========================================================

def classify_texts(
    texts,
    tokenizer,
    model
):

    device = next(model.parameters()).device

    predictions = []
    probabilities = []

    batch_size = 16

    for start in range(
        0,
        len(texts),
        batch_size
    ):

        batch = texts[
            start:start + batch_size
        ]

        encoded = tokenizer(
            batch,
            padding=True,
            truncation=True,
            max_length=MAX_LENGTH,
            return_tensors="pt"
        )

        encoded = {
            key: value.to(device)
            for key, value in encoded.items()
        }

        with torch.no_grad():

            outputs = model(**encoded)

            probs = torch.softmax(
                outputs.logits,
                dim=1
            )

        probs = probs.cpu().numpy()

        preds = np.argmax(
            probs,
            axis=1
        )

        predictions.extend(
            preds.tolist()
        )

        probabilities.extend(
            probs.tolist()
        )

    return (
        predictions,
        probabilities
    )


# =========================================================
# EXCEL DOWNLOAD
# =========================================================

def create_excel(
    df,
    summary
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Results"
        )

        summary.to_excel(
            writer,
            index=False,
            sheet_name="Dashboard"
        )

    output.seek(0)

    return output


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">

            <div class="sidebar-brand-icon">
                📡
            </div>

            <div class="sidebar-brand-title">
                لوحة النشاط
            </div>

            <div class="sidebar-brand-sub">
                Activity Dashboard
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    selected_section = st.radio(
        "القائمة",
        [
            "📡 النشاط",
            "🤝 الوعود",
            "🗂️ الإهمال"
        ],
        label_visibility="collapsed"
    )

    st.markdown("---")

    st.caption(
        "Dashboard v3.0"
    )


# =========================================================
# PROMISES PAGE
# =========================================================

if "الوعود" in selected_section:

    st.title("🤝 الوعود")

    st.info(
        "قسم الوعود جاهز نضيف فيه تحليل الوعود "
        "بعد الانتهاء من قسم النشاط."
    )

    st.stop()


# =========================================================
# NEGLECT PAGE
# =========================================================

if "الإهمال" in selected_section:

    st.title("🗂️ الإهمال")

    st.info(
        "قسم الإهمال سيتم بناؤه بعد قسم النشاط."
    )

    st.stop()


# =========================================================
# ACTIVITY PAGE
# =========================================================

st.markdown(
    """
    <div class="hero">

        <div class="hero-title">
            📡 لوحة النشاط
        </div>

        <div class="hero-subtitle">
            ارفع ملف البيانات، شغّل التصنيف،
            وشوف تحليل النشاط بالكامل في Dashboard ديناميكية.
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# UPLOAD
# =========================================================

st.subheader("📂 رفع ملف البيانات")

uploaded_file = st.file_uploader(
    "ارفع ملف Excel أو CSV",
    type=[
        "csv",
        "xlsx",
        "xls"
    ],
    help="سيتم حذف أول صف بعد صف أسماء الأعمدة تلقائيًا."
)


if uploaded_file is None:

    st.info(
        "📂 ارفع ملف Excel أو CSV عشان تبدأ."
    )

    st.stop()


# =========================================================
# READ FILE
# =========================================================

try:

    if uploaded_file.name.lower().endswith(".csv"):

        df = pd.read_csv(
            uploaded_file
        )

    else:

        df = pd.read_excel(
            uploaded_file
        )

except Exception as e:

    st.error(
        f"حدث خطأ أثناء قراءة الملف: {e}"
    )

    st.stop()


# =========================================================
# REMOVE FIRST DATA ROW
# =========================================================

if len(df) > 0:

    df = df.iloc[1:].reset_index(
        drop=True
    )


# =========================================================
# BASIC INFO
# =========================================================

st.success(
    f"تم تحميل الملف بنجاح: {len(df):,} صف"
)


col1, col2, col3 = st.columns(3)

with col1:

    st.metric(
        "عدد الصفوف",
        f"{len(df):,}"
    )

with col2:

    st.metric(
        "عدد الأعمدة",
        f"{len(df.columns):,}"
    )

with col3:

    st.metric(
        "حجم الملف",
        f"{uploaded_file.size / 1024:.1f} KB"
    )


# =========================================================
# SELECT TEXT COLUMN
# =========================================================

st.markdown(
    '<div class="section-line"></div>',
    unsafe_allow_html=True
)

st.subheader(
    "📝 عمود النص المستخدم في التصنيف"
)


text_columns = list(
    df.select_dtypes(
        include=["object", "string"]
    ).columns
)


if not text_columns:

    st.error(
        "لم يتم العثور على أي عمود نصي في الملف."
    )

    st.stop()


default_index = 0

for i, col in enumerate(text_columns):

    name = str(col).lower()

    if any(
        word in name
        for word in [
            "text",
            "comment",
            "note",
            "call",
            "conversation",
            "نص",
            "مكالمة",
            "ملاحظ",
            "تعليق",
            "كلام"
        ]
    ):

        default_index = i
        break


text_column = st.selectbox(
    "اختار العمود اللي يحتوي على نص المكالمة",
    text_columns,
    index=default_index
)


# =========================================================
# PREVIEW
# =========================================================

with st.expander(
    "👁️ معاينة البيانات"
):

    st.dataframe(
        df.head(10),
        use_container_width=True
    )


# =========================================================
# START CLASSIFICATION
# =========================================================

st.markdown(
    '<div class="section-line"></div>',
    unsafe_allow_html=True
)

st.subheader(
    "🚀 التصنيف"
)

st.write(
    "بعد الضغط على الزر، سيتم تصنيف جميع الصفوف "
    "وحساب احتمال كل تصنيف."
)


start_classification = st.button(
    "🚀 ابدأ التصنيف",
    type="primary"
)


# =========================================================
# CLASSIFICATION
# =========================================================

if start_classification:

    tokenizer, model = load_model()

    if model is None:

        st.error(
            "لم يتم تحديد موديل Hugging Face. "
            "افتح الكود وضع اسم الـ Model Repository "
            "في المتغير HF_MODEL_ID."
        )

        st.stop()


    progress = st.progress(0)

    status = st.empty()


    # ---------------------------------------------
    # CLEAN TEXT
    # ---------------------------------------------

    texts = (
        df[text_column]
        .fillna("")
        .astype(str)
        .apply(clean_text)
        .tolist()
    )


    status.write(
        "⏳ جاري تصنيف البيانات..."
    )


    predictions, probabilities = classify_texts(
        texts,
        tokenizer,
        model
    )


    # ---------------------------------------------
    # SAVE PREDICTION
    # ---------------------------------------------

    df["التصنيف"] = predictions

    # 0 = غير ناجحة
    # 1 = ناجحة

    df["التصنيف"] = df[
        "التصنيف"
    ].astype(int)


    # ---------------------------------------------
    # PROBABILITIES
    # ---------------------------------------------

    probabilities = np.array(
        probabilities
    )


    if probabilities.shape[1] >= 2:

        df["احتمال غير ناجحة"] = (
            probabilities[:, 0] * 100
        ).round(2)

        df["احتمال ناجحة"] = (
            probabilities[:, 1] * 100
        ).round(2)

    else:

        df["احتمال غير ناجحة"] = 0.0

        df["احتمال ناجحة"] = 0.0


    df["احتمال التصنيف"] = np.where(
        df["التصنيف"] == 1,
        df["احتمال ناجحة"],
        df["احتمال غير ناجحة"]
    )


    progress.progress(100)

    status.success(
        "✅ تم الانتهاء من التصنيف."
    )


    # =====================================================
    # DASHBOARD
    # =====================================================

    st.markdown(
        '<div class="section-line"></div>',
        unsafe_allow_html=True
    )

    st.header(
        "📊 Dashboard النشاط"
    )


    # =====================================================
    # KPIs
    # =====================================================

    total = len(df)

    successful = int(
        (df["التصنيف"] == 1).sum()
    )

    unsuccessful = int(
        (df["التصنيف"] == 0).sum()
    )


    success_rate = (
        successful / total * 100
        if total > 0
        else 0
    )


    avg_probability = (
        df["احتمال التصنيف"].mean()
        if total > 0
        else 0
    )


    k1, k2, k3, k4, k5 = st.columns(5)


    with k1:

        st.metric(
            "إجمالي الحالات",
            f"{total:,}"
        )


    with k2:

        st.metric(
            "ناجحة",
            f"{successful:,}"
        )


    with k3:

        st.metric(
            "غير ناجحة",
            f"{unsuccessful:,}"
        )


    with k4:

        st.metric(
            "نسبة النجاح",
            f"{success_rate:.1f}%"
        )


    with k5:

        st.metric(
            "متوسط الثقة",
            f"{avg_probability:.1f}%"
        )


    # =====================================================
    # CHARTS
    # =====================================================

    st.markdown(
        '<div class="section-line"></div>',
        unsafe_allow_html=True
    )


    chart1, chart2 = st.columns(2)


    # ---------------------------------------------
    # DISTRIBUTION
    # ---------------------------------------------

    with chart1:

        st.subheader(
            "📈 توزيع التصنيفات"
        )

        distribution = pd.DataFrame(
            {
                "الحالة": [
                    "ناجحة",
                    "غير ناجحة"
                ],
                "العدد": [
                    successful,
                    unsuccessful
                ]
            }
        )

        st.bar_chart(
            distribution.set_index(
                "الحالة"
            ),
            use_container_width=True
        )


    # ---------------------------------------------
    # PERCENTAGE
    # ---------------------------------------------

    with chart2:

        st.subheader(
            "🥧 نسبة النجاح"
        )

        percentage = pd.DataFrame(
            {
                "الحالة": [
                    "ناجحة",
                    "غير ناجحة"
                ],
                "النسبة": [
                    success_rate,
                    100 - success_rate
                ]
            }
        )

        st.bar_chart(
            percentage.set_index(
                "الحالة"
            ),
            use_container_width=True
        )


    # =====================================================
    # PROBABILITY
    # =====================================================

    st.markdown(
        '<div class="section-line"></div>',
        unsafe_allow_html=True
    )

    st.subheader(
        "🎯 توزيع احتمالات النموذج"
    )

    probability_df = pd.DataFrame(
        {
            "احتمال التصنيف": df[
                "احتمال التصنيف"
            ]
        }
    )

    st.line_chart(
        probability_df,
        use_container_width=True
    )


    # =====================================================
    # COLLECTOR ANALYSIS
    # =====================================================

    possible_collector_columns = [
        col
        for col in df.columns
        if any(
            word in str(col).lower()
            for word in [
                "collector",
                "agent",
                "collector name",
                "المحصل",
                "اسم المحصل",
                "المندوب",
                "الموظف"
            ]
        )
    ]


    if possible_collector_columns:

        collector_column = (
            possible_collector_columns[0]
        )

        st.markdown(
            '<div class="section-line"></div>',
            unsafe_allow_html=True
        )

        st.subheader(
            "👥 نشاط المحصلين"
        )

        collector_stats = (
            df.groupby(
                collector_column
            )
            .agg(
                إجمالي=("التصنيف", "count"),
                ناجحة=("التصنيف", "sum")
            )
            .reset_index()
        )

        collector_stats["غير ناجحة"] = (
            collector_stats["إجمالي"]
            - collector_stats["ناجحة"]
        )

        collector_stats["نسبة النجاح"] = (
            collector_stats["ناجحة"]
            / collector_stats["إجمالي"]
            * 100
        ).round(1)


        st.dataframe(
            collector_stats.sort_values(
                "نسبة النجاح",
                ascending=False
            ),
            use_container_width=True,
            hide_index=True
        )


        st.bar_chart(
            collector_stats.set_index(
                collector_column
            )[[
                "ناجحة",
                "غير ناجحة"
            ]],
            use_container_width=True
        )


    # =====================================================
    # RESULTS
    # =====================================================

    st.markdown(
        '<div class="section-line"></div>',
        unsafe_allow_html=True
    )

    st.subheader(
        "📋 نتائج التصنيف"
    )

    st.dataframe(
        df,
        use_container_width=True,
        height=450
    )


    # =====================================================
    # DOWNLOAD
    # =====================================================

    st.markdown(
        '<div class="section-line"></div>',
        unsafe_allow_html=True
    )

    st.subheader(
        "⬇️ تحميل النتائج"
    )


    summary = pd.DataFrame(
        {
            "المؤشر": [
                "إجمالي الحالات",
                "ناجحة",
                "غير ناجحة",
                "نسبة النجاح",
                "متوسط الثقة"
            ],
            "القيمة": [
                total,
                successful,
                unsuccessful,
                round(success_rate, 2),
                round(avg_probability, 2)
            ]
        }
    )


    excel_file = create_excel(
        df,
        summary
    )


    d1, d2 = st.columns(2)


    with d1:

        st.download_button(
            label="📥 تحميل النتائج Excel",
            data=excel_file,
            file_name="activity_dashboard.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )


    with d2:

        csv_data = df.to_csv(
            index=False
        ).encode("utf-8-sig")


        st.download_button(
            label="📥 تحميل النتائج CSV",
            data=csv_data,
            file_name="activity_results.csv",
            mime="text/csv"
        )


    st.success(
        "🎉 الـ Dashboard اتحدثت بناءً على الملف الحالي."
    )
