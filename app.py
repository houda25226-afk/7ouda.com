"""
Streamlit App
تصنيف إفادات المكالمات + Dashboard للنشاط

Model:
Mahmoud252002/7oudaModel
"""

import io
import pandas as pd
import streamlit as st
import torch

from transformers import (
    AutoTokenizer,
    AutoModelForSequenceClassification
)

import plotly.express as px
import plotly.graph_objects as go


# ==========================================================
# إعدادات عامة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

# 1 = ناجحة
# 0 = غير ناجحة
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

@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0B1220;
    --surface: #131C2E;
    --surface-2: #1A2740;

    --accent: #34D399;
    --accent-2: #FBBF24;

    --success: #34D399;
    --danger: #F87171;

    --text: #F1F5F9;
    --text-dim: #94A3B8;
}

html,
body,
[class*="css"] {
    font-family: 'Tajawal', sans-serif !important;
}

/* =========================
   Background
========================= */

.stApp {
    background:
        radial-gradient(
            circle at 90% -10%,
            rgba(52,211,153,0.08) 0%,
            transparent 40%
        ),
        radial-gradient(
            circle at 10% 100%,
            rgba(251,191,36,0.05) 0%,
            transparent 40%
        ),
        var(--bg);

    color: var(--text);
}


/* =========================
   Main direction
========================= */

[data-testid="stAppViewContainer"] {
    direction: ltr;
}

[data-testid="stMain"] {
    direction: rtl;
}

.main .block-container {
    direction: rtl;
    padding-top: 2rem;
}


/* =========================
   Text
========================= */

.stApp p,
.stApp span,
.stApp label,
.stApp li,
.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4,
.stApp h5,
.stApp h6 {

    color: var(--text) !important;
}

.stApp small {
    color: var(--text-dim) !important;
}


/* =========================
   Hide default
========================= */

#MainMenu,
footer {
    visibility: hidden;
}

header[data-testid="stHeader"] {
    background: transparent;
}


/* =========================
   Sidebar
========================= */

section[data-testid="stSidebar"] {

    direction: rtl;

    background:
        linear-gradient(
            180deg,
            #080C16 0%,
            #0B1220 100%
        );

    border-left: 1px solid rgba(255,255,255,0.06);
}

section[data-testid="stSidebar"] * {
    color: var(--text) !important;
}


/* Sidebar brand */

.sidebar-brand {

    background:
        linear-gradient(
            135deg,
            rgba(52,211,153,0.12),
            rgba(251,191,36,0.05)
        );

    border: 1px solid rgba(52,211,153,0.15);

    border-radius: 18px;

    padding: 1.2rem;

    margin-bottom: 1.2rem;

    text-align: center;
}

.sidebar-logo {
    font-size: 2.2rem;
}

.sidebar-title {
    font-size: 1.25rem;
    font-weight: 900;
    margin-top: 0.4rem;
}

.sidebar-subtitle {
    color: var(--text-dim) !important;
    font-size: 0.75rem;
}


/* =========================
   Hero
========================= */

.hero {

    background:
        linear-gradient(
            135deg,
            rgba(52,211,153,0.08),
            rgba(19,28,46,0.9)
        );

    border: 1px solid rgba(52,211,153,0.15);

    border-radius: 22px;

    padding: 2rem;

    margin-bottom: 1.5rem;

    text-align: center;
}

.hero-icon {
    font-size: 3rem;
}

.hero-title {

    font-size: 2.3rem;

    font-weight: 900;

    margin-top: 0.4rem;

    margin-bottom: 0.3rem;
}

.hero-subtitle {

    color: var(--text-dim) !important;

    font-size: 1rem;
}


/* =========================
   Cards
========================= */

.card {

    background: var(--surface);

    border: 1px solid rgba(255,255,255,0.06);

    border-radius: 18px;

    padding: 1.3rem;

    margin-bottom: 1rem;
}


/* =========================
   Metrics
========================= */

[data-testid="stMetric"] {

    background: var(--surface);

    border: 1px solid rgba(255,255,255,0.06);

    border-radius: 16px;

    padding: 1rem;
}

[data-testid="stMetricValue"] {

    color: var(--accent) !important;

    font-family: 'JetBrains Mono', monospace;
}


/* =========================
   Buttons
========================= */

.stButton > button,
.stDownloadButton > button {

    background:
        linear-gradient(
            90deg,
            var(--accent),
            #22B888
        ) !important;

    color: #05170F !important;

    font-weight: 800 !important;

    border: none !important;

    border-radius: 12px !important;

    padding: 0.7rem 1.2rem !important;

    transition: all 0.2s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover {

    transform: translateY(-2px);

    box-shadow:
        0 8px 20px
        rgba(52,211,153,0.25);
}


/* =========================
   File uploader
========================= */

[data-testid="stFileUploader"] {

    background: var(--surface);

    border: 1.5px dashed
        rgba(52,211,153,0.4);

    border-radius: 18px;

    padding: 0.8rem;
}

[data-testid="stFileUploader"] section {
    background: transparent;
}


/* =========================
   Dataframe
========================= */

[data-testid="stDataFrame"] {

    border-radius: 16px;

    overflow: hidden;

    border: 1px solid
        rgba(255,255,255,0.06);
}


/* =========================
   Progress
========================= */

[data-testid="stProgress"] > div > div {

    background:
        linear-gradient(
            90deg,
            var(--accent),
            var(--accent-2)
        );
}


/* =========================
   Divider
========================= */

.divider {

    height: 1px;

    background:
        linear-gradient(
            90deg,
            transparent,
            rgba(52,211,153,0.35),
            transparent
        );

    margin: 1.5rem 0;
}

</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


# ==========================================================
# Sidebar
# ==========================================================

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
                Activity Dashboard
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### ⚙️ الإعدادات")

    text_column_input = st.text_input(
        "اسم عمود الإفادة",
        value="الافادة"
    )

    st.markdown("---")

    st.markdown("### 🤖 الموديل")

    st.code(
        MODEL_REPO,
        language=None
    )

    st.caption(
        "الموديل يتم تحميله مباشرة من Hugging Face"
    )


# ==========================================================
# تحميل الموديل
# ==========================================================

@st.cache_resource(
    show_spinner="جاري تحميل الموديل من Hugging Face..."
)
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


# ==========================================================
# Prediction
# ==========================================================

def predict_batch(
    texts,
    tokenizer,
    model,
    device,
    batch_size=16
):

    all_preds = []
    all_confidences = []

    progress_bar = st.progress(
        0,
        text="جاري التصنيف..."
    )

    total = len(texts)

    if total == 0:
        progress_bar.empty()
        return [], []

    for i in range(
        0,
        total,
        batch_size
    ):

        batch = texts[
            i:i + batch_size
        ]

        batch = [
            str(t)
            if pd.notna(t)
            and str(t).strip() != ""
            else ""
            for t in batch
        ]

        inputs = tokenizer(
            batch,
            return_tensors="pt",
            truncation=True,
            padding=True,
            max_length=MAX_LENGTH
        ).to(device)

        with torch.no_grad():

            logits = model(
                **inputs
            ).logits

            probs = torch.softmax(
                logits,
                dim=1
            )

            preds = torch.argmax(
                probs,
                dim=1
            )

            confidences = torch.max(
                probs,
                dim=1
            ).values

        all_preds.extend(
            preds.cpu().tolist()
        )

        all_confidences.extend(
            confidences.cpu().tolist()
        )

        done = min(
            i + batch_size,
            total
        )

        progress_bar.progress(
            done / total,
            text=f"جاري التصنيف... {done}/{total}"
        )

    progress_bar.empty()

    return (
        all_preds,
        all_confidences
    )


# ==========================================================
# Helpers
# ==========================================================

def find_column(
    df,
    keywords
):

    for col in df.columns:

        col_lower = str(col).lower()

        for keyword in keywords:

            if keyword.lower() in col_lower:
                return col

    return None


# ==========================================================
# Dashboard
# ==========================================================

def render_dashboard(
    result_df
):

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="card">

        <h2 style="margin:0;">
        📊 Dashboard النشاط
        </h2>

        <p style="color:#94A3B8 !important;">
        تحليل نتائج التصنيف والنشاط الموجود في الملف
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    # ======================================================
    # Metrics
    # ======================================================

    total = len(result_df)

    successful = int(
        (result_df["التصنيف"] == 1).sum()
    )

    unsuccessful = int(
        (result_df["التصنيف"] == 0).sum()
    )

    success_rate = (
        successful / total * 100
        if total > 0
        else 0
    )

    avg_confidence = (
        result_df["نسبة_الثقة"].mean()
        if total > 0
        else 0
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📞 إجمالي الإفادات",
        f"{total:,}"
    )

    c2.metric(
        "✅ ناجحة",
        f"{successful:,}"
    )

    c3.metric(
        "❌ غير ناجحة",
        f"{unsuccessful:,}"
    )

    c4.metric(
        "📈 نسبة النجاح",
        f"{success_rate:.1f}%"
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ======================================================
    # Charts row 1
    # ======================================================

    col1, col2 = st.columns(2)

    # -------------------------
    # Success / Failure
    # -------------------------

    with col1:

        chart_df = pd.DataFrame({
            "الحالة": [
                "ناجحة",
                "غير ناجحة"
            ],
            "العدد": [
                successful,
                unsuccessful
            ]
        })

        fig = px.pie(
            chart_df,
            names="الحالة",
            values="العدد",
            hole=0.55,
            title="توزيع الإفادات"
        )

        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(
                family="Tajawal"
            ),
            legend_title=""
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # -------------------------
    # Confidence
    # -------------------------

    with col2:

        fig = px.histogram(
            result_df,
            x="نسبة_الثقة",
            nbins=20,
            title="توزيع نسبة الثقة",
            labels={
                "نسبة_الثقة":
                "نسبة الثقة %"
            }
        )

        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(
                family="Tajawal"
            )
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # ======================================================
    # Detect Collector
    # ======================================================

    collector_col = find_column(
        result_df,
        [
            "المحصل",
            "اسم المحصل",
            "collector",
            "collector_name",
            "agent",
            "agent_name"
        ]
    )

    # ======================================================
    # Collector activity
    # ======================================================

    if collector_col:

        collector_stats = (
            result_df
            .groupby(collector_col)
            .agg(
                إجمالي=("التصنيف", "count"),
                ناجحة=("التصنيف", "sum")
            )
            .reset_index()
        )

        collector_stats["غير ناجحة"] = (
            collector_stats["إجمالي"]
            -
            collector_stats["ناجحة"]
        )

        collector_stats["نسبة النجاح"] = (
            collector_stats["ناجحة"]
            /
            collector_stats["إجمالي"]
            *
            100
        )

        collector_stats = (
            collector_stats
            .sort_values(
                "إجمالي",
                ascending=False
            )
            .head(15)
        )

        st.markdown(
            '<div class="divider"></div>',
            unsafe_allow_html=True
        )

        st.subheader(
            "👤 نشاط المحصلين"
        )

        col1, col2 = st.columns(2)

        # عدد الإفادات لكل محصل
        with col1:

            fig = px.bar(
                collector_stats,
                x=collector_col,
                y="إجمالي",
                title="عدد الإفادات لكل محصل",
                text="إجمالي"
            )

            fig.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="المحصل",
                yaxis_title="عدد الإفادات",
                font=dict(
                    family="Tajawal"
                )
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # نسبة النجاح لكل محصل
        with col2:

            collector_rate = (
                collector_stats
                .sort_values(
                    "نسبة النجاح",
                    ascending=False
                )
            )

            fig = px.bar(
                collector_rate,
                x=collector_col,
                y="نسبة النجاح",
                title="نسبة النجاح لكل محصل",
                text=collector_rate[
                    "نسبة النجاح"
                ].round(1)
            )

            fig.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="المحصل",
                yaxis_title="نسبة النجاح %",
                font=dict(
                    family="Tajawal"
                )
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # جدول المحصلين

        st.dataframe(
            collector_stats,
            use_container_width=True,
            hide_index=True
        )

    # ======================================================
    # Detect Date Column
    # ======================================================

    date_col = find_column(
        result_df,
        [
            "التاريخ",
            "تاريخ",
            "date",
            "datetime",
            "created_at",
            "created"
        ]
    )

    if date_col:

        try:

            date_data = result_df.copy()

            date_data[date_col] = pd.to_datetime(
                date_data[date_col],
                errors="coerce"
            )

            date_data = date_data.dropna(
                subset=[date_col]
            )

            if len(date_data) > 0:

                daily = (
                    date_data
                    .groupby(
                        date_data[date_col].dt.date
                    )
                    .agg(
                        إجمالي=("التصنيف", "count"),
                        ناجحة=("التصنيف", "sum")
                    )
                    .reset_index()
                )

                daily["غير ناجحة"] = (
                    daily["إجمالي"]
                    -
                    daily["ناجحة"]
                )

                daily["نسبة النجاح"] = (
                    daily["ناجحة"]
                    /
                    daily["إجمالي"]
                    *
                    100
                )

                st.markdown(
                    '<div class="divider"></div>',
                    unsafe_allow_html=True
                )

                st.subheader(
                    "📅 النشاط عبر الزمن"
                )

                fig = go.Figure()

                fig.add_trace(
                    go.Scatter(
                        x=daily[date_col],
                        y=daily["إجمالي"],
                        mode="lines+markers",
                        name="إجمالي النشاط"
                    )
                )

                fig.add_trace(
                    go.Scatter(
                        x=daily[date_col],
                        y=daily["ناجحة"],
                        mode="lines+markers",
                        name="ناجحة"
                    )
                )

                fig.add_trace(
                    go.Scatter(
                        x=daily[date_col],
                        y=daily["غير ناجحة"],
                        mode="lines+markers",
                        name="غير ناجحة"
                    )
                )

                fig.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis_title="التاريخ",
                    yaxis_title="عدد الإفادات",
                    font=dict(
                        family="Tajawal"
                    )
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

    # ======================================================
    # Confidence / Classification table
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.subheader(
        "📋 نتائج التصنيف"
    )

    st.dataframe(
        result_df,
        use_container_width=True,
        hide_index=True
    )


# ==========================================================
# Hero
# ==========================================================

st.markdown(
    """
    <div class="hero">

        <div class="hero-icon">
            📡
        </div>

        <div class="hero-title">
            لوحة النشاط
        </div>

        <div class="hero-subtitle">
            ارفع ملف البيانات، صنّف الإفادات،
            وشوف تحليل النشاط بالكامل
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ==========================================================
# Upload
# ==========================================================

st.markdown(
    """
    <div class="card">

    <h3>
    📂 رفع ملف البيانات
    </h3>

    <p style="color:#94A3B8 !important;">
    ارفع ملف Excel أو CSV وسيتم تجهيز البيانات تلقائيًا.
    </p>

    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "ارفع ملف البيانات",
    type=[
        "csv",
        "xlsx",
        "xls"
    ],
    label_visibility="collapsed"
)


# ==========================================================
# Process File
# ==========================================================

if uploaded_file is not None:

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
            f"❌ مش قادر أقرأ الملف: {e}"
        )

        st.stop()


    # ======================================================
    # Rename Note -> الافادة
    # ======================================================

    if "Note" in df.columns:

        df = df.rename(
            columns={
                "Note": "الافادة"
            }
        )


    # ======================================================
    # Remove first data row
    # ======================================================

    if len(df) > 0:

        df = (
            df
            .iloc[1:]
            .reset_index(drop=True)
        )


    # ======================================================
    # File info
    # ======================================================

    st.success(
        f"✅ تم تحميل الملف بنجاح — "
        f"{len(df):,} صف"
    )

    st.dataframe(
        df.head(10),
        use_container_width=True,
        hide_index=True
    )


    # ======================================================
    # Check Text Column
    # ======================================================

    if text_column_input not in df.columns:

        st.error(
            f"""
            ❌ عمود الإفادة غير موجود.

            العمود المطلوب:
            `{text_column_input}`

            الأعمدة الموجودة:
            {", ".join(
                df.columns.astype(str)
            )}
            """
        )

        st.stop()


    # ======================================================
    # Start Classification
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="card">

        <h3>
        🤖 جاهز للتصنيف
        </h3>

        <p style="color:#94A3B8 !important;">
        اضغط على الزر لبدء تصنيف جميع الإفادات.
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    start_classification = st.button(
        "🚀 ابدأ التصنيف",
        type="primary",
        use_container_width=True
    )


    if start_classification:

        # ==================================================
        # Load model
        # ==================================================

        tokenizer, model, device = load_model()


        # ==================================================
        # Texts
        # ==================================================

        texts = (
            df[text_column_input]
            .fillna("")
            .astype(str)
            .tolist()
        )


        # ==================================================
        # Prediction
        # ==================================================

        preds, confidences = predict_batch(
            texts,
            tokenizer,
            model,
            device
        )


        # ==================================================
        # Result
        # ==================================================

        result_df = df.copy()


        # 1 = ناجحة
        # 0 = غير ناجحة

        result_df["التصنيف"] = [
            1 if int(p) == 1 else 0
            for p in preds
        ]


        result_df["نسبة_الثقة"] = [
            round(
                float(c) * 100,
                1
            )
            for c in confidences
        ]


        # ==================================================
        # Save result in session
        # ==================================================

        st.session_state["result_df"] = result_df


        st.success(
            "✅ تم التصنيف بنجاح"
        )


# ==========================================================
# Show Dashboard
# ==========================================================

if "result_df" in st.session_state:

    result_df = st.session_state[
        "result_df"
    ]

    render_dashboard(
        result_df
    )


    # ======================================================
    # Download
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )

    st.subheader(
        "⬇️ تحميل النتائج"
    )


    col1, col2 = st.columns(2)


    # CSV
    with col1:

        csv_output = (
            result_df
            .to_csv(
                index=False,
                encoding="utf-8-sig"
            )
            .encode("utf-8-sig")
        )

        st.download_button(
            "⬇️ تحميل CSV",
            data=csv_output,
            file_name="نتائج_النشاط.csv",
            mime="text/csv",
            use_container_width=True
        )


    # Excel
    with col2:

        buffer = io.BytesIO()

        with pd.ExcelWriter(
            buffer,
            engine="openpyxl"
        ) as writer:

            result_df.to_excel(
                writer,
                index=False,
                sheet_name="نتائج النشاط"
            )

        excel_output = buffer.getvalue()

        st.download_button(
            "⬇️ تحميل Excel",
            data=excel_output,
            file_name="نتائج_النشاط.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )
