import pandas as pd
import streamlit as st
import torch
import plotly.express as px
import plotly.graph_objects as go

from transformers import (
    AutoTokenizer,
    AutoModelForSequenceClassification
)

from io import BytesIO
from datetime import datetime


# ==========================================================
# إعداد الصفحة
# ==========================================================

st.set_page_config(
    page_title="Activity Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================================
# إعدادات الموديل
# ==========================================================

# ==========================================================
# مهم جدًا:
# اكتب هنا Hugging Face Repository
#
# مثال:
# HF_MODEL_ID = "Mahmoud/model-name"
# ==========================================================

HF_MODEL_ID = "Mahmoud252002/7oudaModel"

BATCH_SIZE = 16
MAX_LENGTH = 256


# ==========================================================
# Theme
# ==========================================================

CSS_THEME = """
<style>

@import url(
'https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap'
);

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

html,
body,
[class*="css"] {
    font-family: 'Cairo', sans-serif;
}


/* ==========================================================
   Main
   ========================================================== */

[data-testid="stAppViewContainer"] {
    direction: ltr;
}

.main .block-container,
[data-testid="stMain"] {
    direction: rtl;
    max-width: 1600px;
}

.stApp {

    background:

        radial-gradient(
            circle at 90% -10%,
            rgba(52,211,153,0.08) 0%,
            transparent 45%
        ),

        radial-gradient(
            circle at 5% 105%,
            rgba(251,191,36,0.06) 0%,
            transparent 45%
        ),

        var(--bg);

    color: var(--text);
}


/* ==========================================================
   Text
   ========================================================== */

.stApp,
.stApp p,
.stApp span,
.stApp label,
.stApp li,
.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4,
.stApp h5,
.stApp h6,
.stMarkdown,
.stMarkdown p,
.stText,
.stCaption,
[data-testid="stWidgetLabel"] p,
[data-testid="stMetricLabel"] {

    color:
        var(--text) !important;
}

.stApp small,
[data-testid="stCaptionContainer"] {

    color:
        var(--text-dim) !important;
}


/* ==========================================================
   Header
   ========================================================== */

#MainMenu,
footer {

    visibility:
        hidden;
}

header[data-testid="stHeader"] {

    background:
        transparent;
}


/* ==========================================================
   Sidebar
   ========================================================== */

section[data-testid="stSidebar"] {

    direction:
        rtl;

    background:
        #080C16;

    border-left:
        1px solid rgba(255,255,255,0.06);
}

section[data-testid="stSidebar"] * {

    color:
        var(--text) !important;
}


/* ==========================================================
   Brand
   ========================================================== */

.brand-block {

    display:
        flex;

    align-items:
        center;

    gap:
        0.7rem;

    padding:
        1rem 0.2rem 1.1rem 0.2rem;

    margin-bottom:
        0.6rem;

    border-bottom:
        1px solid rgba(255,255,255,0.08);
}

.brand-logo {

    width:
        44px;

    height:
        44px;

    border-radius:
        13px;

    background:
        linear-gradient(
            135deg,
            var(--accent),
            var(--accent-2)
        );

    display:
        flex;

    align-items:
        center;

    justify-content:
        center;

    font-size:
        1.35rem;

    box-shadow:
        0 4px 16px
        rgba(52,211,153,0.25);
}

.brand-name {

    font-weight:
        900;

    font-size:
        1.15rem;
}

.brand-sub {

    font-family:
        'JetBrains Mono',
        monospace;

    font-size:
        0.65rem;

    letter-spacing:
        0.12em;

    color:
        var(--text-dim) !important;
}


/* ==========================================================
   Sidebar Navigation
   ========================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"] {

    gap:
        0.5rem;
}

section[data-testid="stSidebar"]
div[role="radiogroup"] label {

    background:
        var(--surface);

    border:
        1px solid
        rgba(255,255,255,0.06);

    border-radius:
        12px;

    padding:
        0.95rem 1.1rem;

    width:
        100%;

    min-height:
        3.2rem;

    display:
        flex;

    align-items:
        center;

    transition:
        background 0.15s ease,
        border-color 0.15s ease;
}

section[data-testid="stSidebar"]
div[role="radiogroup"] label:hover {

    background:
        var(--surface-2);

    border-color:
        rgba(52,211,153,0.35);
}

section[data-testid="stSidebar"]
div[role="radiogroup"]
label[data-baseweb="radio"] > div:first-child {

    display:
        none;
}

section[data-testid="stSidebar"]
div[role="radiogroup"]
label div[data-testid="stMarkdownContainer"] p {

    font-size:
        1.15rem !important;

    font-weight:
        700 !important;
}


/* ==========================================================
   Hero
   ========================================================== */

.hero-wrap {

    text-align:
        center;

    padding:
        1rem 0 0.8rem 0;
}

.hero-eyebrow {

    font-family:
        'JetBrains Mono',
        monospace;

    font-size:
        0.7rem;

    letter-spacing:
        0.18em;

    color:
        var(--accent);

    text-transform:
        uppercase;

    margin-bottom:
        0.5rem;
}

.hero-title {

    font-weight:
        900;

    font-size:
        2.2rem;

    color:
        var(--text);

    margin:
        0;

    line-height:
        1.35;
}

.hero-subtitle {

    color:
        var(--text-dim);

    font-size:
        0.95rem;

    margin-top:
        0.55rem;
}


/* ==========================================================
   Cards
   ========================================================== */

.card {

    background:
        var(--surface);

    border:
        1px solid
        rgba(255,255,255,0.06);

    border-radius:
        16px;

    padding:
        1.15rem 1.35rem;

    margin-bottom:
        1rem;
}


/* ==========================================================
   Upload
   ========================================================== */

[data-testid="stFileUploader"] {

    background:
        var(--surface);

    border:
        1.5px dashed
        rgba(52,211,153,0.4);

    border-radius:
        16px;

    padding:
        0.7rem;
}

[data-testid="stFileUploader"] section {

    background:
        transparent;
}


/* ==========================================================
   Buttons
   ========================================================== */

.stButton > button,
.stDownloadButton > button {

    background:
        linear-gradient(
            90deg,
            var(--accent),
            #22B888
        );

    color:
        #05170F;

    font-weight:
        700;

    border:
        none;

    border-radius:
        11px;

    padding:
        0.7rem 1.2rem;

    transition:
        transform 0.15s ease,
        box-shadow 0.15s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover {

    transform:
        translateY(-1px);

    box-shadow:
        0 8px 20px
        rgba(52,211,153,0.3);

    color:
        #05170F;
}


/* ==========================================================
   Metrics
   ========================================================== */

[data-testid="stMetric"] {

    background:
        var(--surface);

    border-radius:
        14px;

    padding:
        1rem;

    border:
        1px solid
        rgba(255,255,255,0.06);
}

[data-testid="stMetricValue"] {

    font-family:
        'JetBrains Mono',
        monospace;

    color:
        var(--accent);
}


/* ==========================================================
   Progress
   ========================================================== */

[data-testid="stProgress"] > div > div {

    background:
        linear-gradient(
            90deg,
            var(--accent),
            var(--accent-2)
        );
}


/* ==========================================================
   DataFrame
   ========================================================== */

[data-testid="stDataFrame"] {

    border-radius:
        14px;

    overflow:
        hidden;

    border:
        1px solid
        rgba(255,255,255,0.06);
}


/* ==========================================================
   Divider
   ========================================================== */

.divider {

    height:
        1px;

    background:
        linear-gradient(
            90deg,
            transparent,
            rgba(52,211,153,0.35),
            transparent
        );

    margin:
        1.4rem 0;

    border:
        none;
}


/* ==========================================================
   Section title
   ========================================================== */

.section-title {

    font-size:
        1.25rem;

    font-weight:
        900;

    margin:
        1rem 0;
}

</style>
"""

st.markdown(
    CSS_THEME,
    unsafe_allow_html=True
)


# ==========================================================
# Load Hugging Face Model
# ==========================================================

@st.cache_resource
def load_classifier():

    tokenizer = AutoTokenizer.from_pretrained(
        HF_MODEL_ID
    )

    model = AutoModelForSequenceClassification.from_pretrained(
        HF_MODEL_ID
    )

    device = torch.device(
        "cuda"
        if torch.cuda.is_available()
        else "cpu"
    )

    model.to(device)

    model.eval()

    return (
        tokenizer,
        model,
        device
    )


# ==========================================================
# تحديد الـ Label
# ==========================================================

def normalize_label(label):

    label = str(label).strip().lower()

    replacements = {
        "ناجحة": "1",
        "ناجح": "1",
        "successful": "1",
        "success": "1",
        "positive": "1",
        "1": "1",

        "غير ناجحة": "0",
        "غير ناجح": "0",
        "unsuccessful": "0",
        "failure": "0",
        "failed": "0",
        "negative": "0",
        "0": "0",
    }

    return replacements.get(
        label,
        label
    )


def prediction_to_binary(
    model,
    prediction
):

    id2label = model.config.id2label

    original_label = id2label.get(
        int(prediction),
        str(prediction)
    )

    normalized = normalize_label(
        original_label
    )

    # لو الموديل أصلاً بيرجع 0 أو 1
    if normalized in ["0", "1"]:
        return int(normalized)

    # fallback
    if int(prediction) in [0, 1]:
        return int(prediction)

    return int(prediction)


# ==========================================================
# Classification
# ==========================================================

def classify_texts(
    texts,
    tokenizer,
    model,
    device
):

    predictions = []

    probability_0 = []

    probability_1 = []

    confidence = []

    total = len(texts)

    progress_bar = st.progress(0)

    status = st.empty()


    for start in range(
        0,
        total,
        BATCH_SIZE
    ):

        batch = texts[
            start:
            start + BATCH_SIZE
        ]


        inputs = tokenizer(

            batch,

            padding=True,

            truncation=True,

            max_length=MAX_LENGTH,

            return_tensors="pt"
        )


        inputs = {

            key:
                value.to(device)

            for key, value
            in inputs.items()
        }


        with torch.no_grad():

            outputs = model(
                **inputs
            )


        # ==================================================
        # Softmax
        # ==================================================

        probs = torch.softmax(
            outputs.logits,
            dim=-1
        )


        batch_pred = torch.argmax(
            probs,
            dim=-1
        )


        batch_conf = torch.max(
            probs,
            dim=-1
        ).values


        # ==================================================
        # الاحتمالات
        # ==================================================

        if probs.shape[1] == 2:

            p0 = probs[:, 0]

            p1 = probs[:, 1]

        else:

            # لو الموديل مش Binary
            p0 = torch.zeros(
                len(batch),
                device=device
            )

            p1 = torch.zeros(
                len(batch),
                device=device
            )


        predictions.extend(
            batch_pred.cpu().numpy()
        )

        probability_0.extend(
            p0.cpu().numpy()
        )

        probability_1.extend(
            p1.cpu().numpy()
        )

        confidence.extend(
            batch_conf.cpu().numpy()
        )


        processed = min(
            start + BATCH_SIZE,
            total
        )


        progress = (
            processed /
            total
        )


        progress_bar.progress(
            progress
        )


        status.markdown(
            f"""
            **جاري التصنيف...**

            {processed:,} / {total:,}
            """
        )


    progress_bar.empty()

    status.empty()


    return (
        predictions,
        probability_0,
        probability_1,
        confidence
    )


# ==========================================================
# Excel Export
# ==========================================================

def create_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Classification"
        )


    output.seek(0)

    return output


# ==========================================================
# Dashboard HTML
# ==========================================================

def create_dashboard_html(
    df,
    charts
):

    generated_at = datetime.now().strftime(
        "%Y-%m-%d %H:%M"
    )


    total = len(df)

    successful = (
        df["Prediction"] == 1
    ).sum()

    unsuccessful = (
        df["Prediction"] == 0
    ).sum()

    success_rate = (
        successful / total * 100
        if total > 0
        else 0
    )

    avg_probability = (
        df["Probability_1"].mean()
        if total > 0
        else 0
    )


    chart_html = ""


    for chart in charts:

        chart_html += chart.to_html(
            full_html=False,
            include_plotlyjs=False
        )


    html = f"""

<!DOCTYPE html>

<html lang="ar" dir="rtl">

<head>

<meta charset="UTF-8">

<title>
Activity Dashboard
</title>

<script src="https://cdn.plot.ly/plotly-2.35.2.min.js">
</script>

<style>

body {{

    margin: 0;

    padding: 30px;

    background:
        #0B1220;

    color:
        #F1F5F9;

    font-family:
        Arial,
        sans-serif;
}}

h1 {{

    text-align:
        center;
}}

.subtitle {{

    text-align:
        center;

    color:
        #8A93A6;

    margin-bottom:
        30px;
}}

.kpis {{

    display:
        grid;

    grid-template-columns:
        repeat(4, 1fr);

    gap:
        15px;

    margin-bottom:
        25px;
}}

.kpi {{

    background:
        #131C2E;

    border:
        1px solid #25314A;

    border-radius:
        15px;

    padding:
        25px;

    text-align:
        center;
}}

.value {{

    font-size:
        28px;

    font-weight:
        bold;

    color:
        #34D399;
}}

.label {{

    color:
        #8A93A6;

    margin-top:
        8px;
}}

.chart {{

    background:
        #131C2E;

    border-radius:
        15px;

    padding:
        10px;

    margin-bottom:
        20px;
}}

.footer {{

    text-align:
        center;

    color:
        #8A93A6;

    margin-top:
        30px;
}}

</style>

</head>


<body>


<h1>
📊 Activity Dashboard
</h1>


<div class="subtitle">

تقرير نشاط وتصنيف البيانات

<br>

تم إنشاء التقرير:
{generated_at}

</div>


<div class="kpis">


<div class="kpi">

<div class="value">
{total:,}
</div>

<div class="label">
إجمالي الحالات
</div>

</div>


<div class="kpi">

<div class="value">
{successful:,}
</div>

<div class="label">
الحالات الناجحة
</div>

</div>


<div class="kpi">

<div class="value">
{unsuccessful:,}
</div>

<div class="label">
الحالات غير الناجحة
</div>

</div>


<div class="kpi">

<div class="value">
{success_rate:.1f}%
</div>

<div class="label">
نسبة النجاح
</div>

</div>


</div>


{chart_html}


<div class="footer">

Activity Classification Dashboard

</div>


</body>

</html>

"""

    return html.encode(
        "utf-8"
    )


# ==========================================================
# Dashboard
# ==========================================================

def render_dashboard(
    df,
    date_column=None,
    collector_column=None
):

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <div class="hero-wrap">

            <div class="hero-eyebrow">
                ACTIVITY ANALYTICS
            </div>

            <p class="hero-title">
                📊 Dashboard النشاط
            </p>

            <p class="hero-subtitle">
                تحليل ديناميكي لنتائج التصنيف
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ======================================================
    # Filters
    # ======================================================

    st.markdown(
        "### 🎛️ الفلاتر"
    )


    filter_col1, filter_col2 = st.columns(2)


    with filter_col1:

        if collector_column:

            collectors = sorted(
                df[collector_column]
                .fillna("غير محدد")
                .astype(str)
                .unique()
                .tolist()
            )

            selected_collectors = st.multiselect(
                "👤 المحصل",
                collectors,
                default=collectors
            )

        else:

            selected_collectors = None


    with filter_col2:

        selected_predictions = st.multiselect(

            "🎯 التصنيف",

            [0, 1],

            default=[0, 1],

            format_func=lambda x:
                "1 - ناجحة"
                if x == 1
                else "0 - غير ناجحة"
        )


    # ======================================================
    # Apply filters
    # ======================================================

    filtered_df = df.copy()


    if collector_column:

        filtered_df = filtered_df[
            filtered_df[
                collector_column
            ]
            .fillna("غير محدد")
            .astype(str)
            .isin(
                selected_collectors
            )
        ]


    filtered_df = filtered_df[
        filtered_df["Prediction"]
        .isin(selected_predictions)
    ]


    # ======================================================
    # KPIs
    # ======================================================

    total = len(filtered_df)


    successful = (
        filtered_df["Prediction"] == 1
    ).sum()


    unsuccessful = (
        filtered_df["Prediction"] == 0
    ).sum()


    success_rate = (
        successful / total * 100
        if total
        else 0
    )


    avg_probability = (

        filtered_df["Probability_1"].mean()

        if total

        else 0
    )


    high_confidence = (

        filtered_df["Confidence"] >= 80
    ).sum()


    c1, c2, c3, c4, c5 = st.columns(5)


    with c1:

        st.metric(
            "📊 إجمالي الحالات",
            f"{total:,}"
        )


    with c2:

        st.metric(
            "✅ ناجحة",
            f"{successful:,}"
        )


    with c3:

        st.metric(
            "❌ غير ناجحة",
            f"{unsuccessful:,}"
        )


    with c4:

        st.metric(
            "📈 نسبة النجاح",
            f"{success_rate:.1f}%"
        )


    with c5:

        st.metric(
            "🎯 متوسط Probability",
            f"{avg_probability:.1f}%"
        )


    # ======================================================
    # Charts list
    # ======================================================

    charts = []


    # ======================================================
    # Chart 1 - Success Distribution
    # ======================================================

    col1, col2 = st.columns(2)


    with col1:

        pie_data = pd.DataFrame({

            "Prediction": [
                "ناجحة",
                "غير ناجحة"
            ],

            "Count": [
                successful,
                unsuccessful
            ]
        })


        fig_pie = px.pie(

            pie_data,

            names="Prediction",

            values="Count",

            hole=0.55,

            title="🎯 نسبة النجاح"
        )


        fig_pie.update_layout(

            template="plotly_dark",

            paper_bgcolor="rgba(0,0,0,0)",

            plot_bgcolor="rgba(0,0,0,0)",

            font=dict(
                family="Cairo"
            )
        )


        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )


        charts.append(
            fig_pie
        )


    # ======================================================
    # Chart 2 - Probability Distribution
    # ======================================================

    with col2:

        if total > 0:

            fig_prob = px.histogram(

                filtered_df,

                x="Probability_1",

                nbins=20,

                title="📊 توزيع Probability للنجاح",

                labels={
                    "Probability_1":
                        "Probability of Success %"
                }
            )


            fig_prob.update_layout(

                template="plotly_dark",

                paper_bgcolor="rgba(0,0,0,0)",

                plot_bgcolor="rgba(0,0,0,0)",

                font=dict(
                    family="Cairo"
                )
            )


            st.plotly_chart(
                fig_prob,
                use_container_width=True
            )


            charts.append(
                fig_prob
            )


    # ======================================================
    # Chart 3 - Success by Collector
    # ======================================================

    if collector_column and total > 0:

        collector_data = (

            filtered_df

            .groupby(
                collector_column
            )

            .agg(

                Total=(
                    "Prediction",
                    "count"
                ),

                Successful=(
                    "Prediction",
                    "sum"
                ),

                Avg_Probability=(
                    "Probability_1",
                    "mean"
                )

            )

            .reset_index()
        )


        collector_data[
            "Success_Rate"
        ] = (

            collector_data["Successful"]

            /
            collector_data["Total"]

            * 100
        )


        collector_data[
            "Success_Rate"
        ] = collector_data[
            "Success_Rate"
        ].round(1)


        col1, col2 = st.columns(2)


        # --------------------------------------------------
        # Collector Activity
        # --------------------------------------------------

        with col1:

            top_collectors = (

                collector_data

                .sort_values(
                    "Total",
                    ascending=False
                )

                .head(15)
            )


            fig_collector = px.bar(

                top_collectors,

                x=collector_column,

                y="Total",

                text="Total",

                title="👥 أكثر المحصلين نشاطًا",

                labels={
                    "Total":
                        "عدد الحالات"
                }
            )


            fig_collector.update_layout(

                template="plotly_dark",

                paper_bgcolor="rgba(0,0,0,0)",

                plot_bgcolor="rgba(0,0,0,0)",

                font=dict(
                    family="Cairo"
                )
            )


            st.plotly_chart(

                fig_collector,

                use_container_width=True
            )


            charts.append(
                fig_collector
            )


        # --------------------------------------------------
        # Success Rate by Collector
        # --------------------------------------------------

        with col2:

            top_success = (

                collector_data

                .sort_values(
                    "Success_Rate",
                    ascending=False
                )

                .head(15)
            )


            fig_success = px.bar(

                top_success,

                x=collector_column,

                y="Success_Rate",

                text="Success_Rate",

                title="🏆 نسبة النجاح حسب المحصل",

                labels={
                    "Success_Rate":
                        "نسبة النجاح %"
                }
            )


            fig_success.update_layout(

                template="plotly_dark",

                paper_bgcolor="rgba(0,0,0,0)",

                plot_bgcolor="rgba(0,0,0,0)",

                font=dict(
                    family="Cairo"
                )
            )


            st.plotly_chart(

                fig_success,

                use_container_width=True
            )


            charts.append(
                fig_success
            )


        # --------------------------------------------------
        # Collector Table
        # --------------------------------------------------

        st.markdown(
            "### 👥 تفاصيل نشاط المحصلين"
        )


        display_collector = (
            collector_data
            .sort_values(
                "Total",
                ascending=False
            )
        )


        st.dataframe(

            display_collector,

            use_container_width=True,

            hide_index=True
        )


    # ======================================================
    # Daily Activity
    # ======================================================

    if date_column and total > 0:

        temp = filtered_df.copy()


        temp[date_column] = pd.to_datetime(

            temp[date_column],

            errors="coerce"
        )


        temp = temp.dropna(
            subset=[date_column]
        )


        if not temp.empty:

            daily = (

                temp

                .groupby(
                    temp[date_column].dt.date
                )

                .agg(

                    Total=(
                        "Prediction",
                        "count"
                    ),

                    Successful=(
                        "Prediction",
                        "sum"
                    )

                )

                .reset_index()
            )


            daily["Success_Rate"] = (

                daily["Successful"]

                /
                daily["Total"]

                * 100
            )


            # ------------------------------------------------
            # Daily Activity
            # ------------------------------------------------

            st.markdown(
                "### 📅 النشاط اليومي"
            )


            fig_daily = px.line(

                daily,

                x=date_column,

                y="Total",

                markers=True,

                title="📈 حجم النشاط يوميًا",

                labels={
                    "Total":
                        "عدد الحالات"
                }
            )


            fig_daily.update_layout(

                template="plotly_dark",

                paper_bgcolor="rgba(0,0,0,0)",

                plot_bgcolor="rgba(0,0,0,0)",

                font=dict(
                    family="Cairo"
                )
            )


            st.plotly_chart(

                fig_daily,

                use_container_width=True
            )


            charts.append(
                fig_daily
            )


            # ------------------------------------------------
            # Daily Success Rate
            # ------------------------------------------------

            fig_daily_success = px.line(

                daily,

                x=date_column,

                y="Success_Rate",

                markers=True,

                title="🎯 نسبة النجاح اليومية",

                labels={
                    "Success_Rate":
                        "نسبة النجاح %"
                }
            )


            fig_daily_success.update_layout(

                template="plotly_dark",

                paper_bgcolor="rgba(0,0,0,0)",

                plot_bgcolor="rgba(0,0,0,0)",

                font=dict(
                    family="Cairo"
                )
            )


            st.plotly_chart(

                fig_daily_success,

                use_container_width=True
            )


            charts.append(
                fig_daily_success
            )


    # ======================================================
    # Confidence
    # ======================================================

    if total > 0:

        high_confidence = (

            filtered_df[
                "Confidence"
            ] >= 80
        ).sum()


        low_confidence = (

            filtered_df[
                "Confidence"
            ] < 60
        ).sum()


        st.markdown(
            "### 🤖 ثقة الموديل"
        )


        confidence_data = pd.DataFrame({

            "المستوى": [
                "ثقة عالية ≥ 80%",
                "ثقة متوسطة 60-79%",
                "ثقة منخفضة < 60%"
            ],

            "العدد": [

                (
                    filtered_df["Confidence"] >= 80
                ).sum(),

                (
                    (
                        filtered_df["Confidence"] >= 60
                    )
                    &
                    (
                        filtered_df["Confidence"] < 80
                    )
                ).sum(),

                (
                    filtered_df["Confidence"] < 60
                ).sum()
            ]
        })


        fig_conf = px.bar(

            confidence_data,

            x="المستوى",

            y="العدد",

            text="العدد",

            title="🧠 توزيع ثقة الموديل"
        )


        fig_conf.update_layout(

            template="plotly_dark",

            paper_bgcolor="rgba(0,0,0,0)",

            plot_bgcolor="rgba(0,0,0,0)",

            font=dict(
                family="Cairo"
            )
        )


        st.plotly_chart(

            fig_conf,

            use_container_width=True
        )


        charts.append(
            fig_conf
        )


    # ======================================================
    # Download Dashboard
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )


    dashboard_html = create_dashboard_html(

        filtered_df,

        charts
    )


    st.download_button(

        label="📥 تحميل الـ Dashboard HTML",

        data=dashboard_html,

        file_name="activity_dashboard.html",

        mime="text/html",

        use_container_width=True
    )


    return filtered_df


# ==========================================================
# Activity Page
# ==========================================================

def render_nashat():

    st.markdown(
        """
        <div class="hero-wrap">

            <div class="hero-eyebrow">
                ACTIVITY
            </div>

            <p class="hero-title">
                📡 النشاط
            </p>

            <p class="hero-subtitle">
                ارفع الملف، صنّف البيانات، وشوف Dashboard كاملة
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ======================================================
    # Upload
    # ======================================================

    uploaded_file = st.file_uploader(

        "📂 ارفع ملف البيانات",

        type=[
            "csv",
            "xlsx",
            "xls"
        ]
    )


    if uploaded_file is None:

        st.markdown(
            """
            <div class="card"
                 style="text-align:center;">

                📂 ارفع ملف Excel أو CSV عشان تبدأ

                <br><br>

                بعد الرفع سيتم حذف أول صف من البيانات
                تلقائيًا.

            </div>
            """,
            unsafe_allow_html=True
        )

        return


    # ======================================================
    # Read File
    # ======================================================

    try:

        if uploaded_file.name.lower().endswith(
            ".csv"
        ):

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

        return


    # ======================================================
    # حذف أول صف بعد الـ Headers
    # ======================================================

    if len(df) > 0:

        df = df.iloc[1:].reset_index(
            drop=True
        )


    # ======================================================
    # File Information
    # ======================================================

    st.markdown(
        f"""
        <div class="card">

            ✅ تم تحميل الملف

            <br><br>

            🗂️ اسم الملف:
            <b>{uploaded_file.name}</b>

            <br><br>

            📊 عدد الصفوف بعد حذف أول صف:
            <b>{len(df):,}</b>

            &nbsp;&nbsp; | &nbsp;&nbsp;

            📋 عدد الأعمدة:
            <b>{len(df.columns)}</b>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ======================================================
    # Column Selection
    # ======================================================

    st.markdown(
        "### ⚙️ إعدادات التصنيف"
    )


    text_column = st.selectbox(

        "📝 اختار عمود النص للموديل",

        options=list(
            df.columns
        )
    )


    # ======================================================
    # Detect Date / Collector
    # ======================================================

    date_candidates = [

        col

        for col in df.columns

        if any(

            word in str(col).lower()

            for word in [
                "date",
                "تاريخ",
                "التاريخ"
            ]
        )
    ]


    collector_candidates = [

        col

        for col in df.columns

        if any(

            word in str(col).lower()

            for word in [
                "collector",
                "محصل",
                "المحصل",
                "collector_name"
            ]
        )
    ]


    col1, col2 = st.columns(2)


    with col1:

        date_options = [
            "لا يوجد"
        ] + list(
            df.columns
        )


        date_default = (

            date_options.index(
                date_candidates[0]
            )

            if date_candidates

            else 0
        )


        date_column = st.selectbox(

            "📅 عمود التاريخ",

            date_options,

            index=date_default
        )


        if date_column == "لا يوجد":

            date_column = None


    with col2:

        collector_options = [
            "لا يوجد"
        ] + list(
            df.columns
        )


        collector_default = (

            collector_options.index(
                collector_candidates[0]
            )

            if collector_candidates

            else 0
        )


        collector_column = st.selectbox(

            "👤 عمود المحصل",

            collector_options,

            index=collector_default
        )


        if collector_column == "لا يوجد":

            collector_column = None


    # ======================================================
    # Preview
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )


    st.markdown(
        "### 👀 معاينة البيانات بعد حذف أول صف"
    )


    st.dataframe(

        df.head(20),

        use_container_width=True,

        hide_index=True
    )


    # ======================================================
    # Start Classification Button
    # ======================================================

    st.markdown(
        '<div class="divider"></div>',
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <div class="card"
             style="text-align:center;">

            <h3>
                🤖 جاهز للتصنيف؟
            </h3>

            <p>
                اضغط الزر لتشغيل الموديل على كامل البيانات
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    start_classification = st.button(

        "🚀 ابدأ التصنيف",

        use_container_width=True
    )


    # ======================================================
    # Start
    # ======================================================

    if start_classification:

        if df.empty:

            st.error(
                "❌ مفيش بيانات بعد حذف أول صف."
            )

            return


        # --------------------------------------------------
        # Text
        # --------------------------------------------------

        texts = (

            df[text_column]

            .fillna("")

            .astype(str)

            .tolist()
        )


        # --------------------------------------------------
        # Load Model
        # --------------------------------------------------

        try:

            with st.spinner(
                "🤖 جاري تحميل الموديل من Hugging Face..."
            ):

                tokenizer, model, device = (
                    load_classifier()
                )

        except Exception as e:

            st.error(
                f"""
                ❌ حصل خطأ أثناء تحميل الموديل.

                تأكد من قيمة:

                HF_MODEL_ID

                والخاصة بـ Hugging Face Repository.

                الخطأ:

                {e}
                """
            )

            return


        # --------------------------------------------------
        # Classification
        # --------------------------------------------------

        st.markdown(
            "### 🚀 جاري التصنيف"
        )


        try:

            (
                predictions,
                p0,
                p1,
                confidence
            ) = classify_texts(

                texts,

                tokenizer,

                model,

                device
            )

        except Exception as e:

            st.error(
                f"❌ حصل خطأ أثناء التصنيف: {e}"
            )

            return


        # --------------------------------------------------
        # Binary Prediction
        # --------------------------------------------------

        binary_predictions = [

            prediction_to_binary(
                model,
                prediction
            )

            for prediction
            in predictions
        ]


        # --------------------------------------------------
        # Results
        # --------------------------------------------------

        df["Prediction"] = (
            binary_predictions
        )


        df["Probability_0"] = [

            round(
                float(x) * 100,
                2
            )

            for x in p0
        ]


        df["Probability_1"] = [

            round(
                float(x) * 100,
                2
            )

            for x in p1
        ]


        df["Confidence"] = [

            round(
                float(x) * 100,
                2
            )

            for x in confidence
        ]


        # --------------------------------------------------
        # Prediction Label
        # --------------------------------------------------

        df["Prediction_Label"] = (

            df["Prediction"]

            .map({

                0:
                    "غير ناجحة",

                1:
                    "ناجحة"
            })
        )


        # --------------------------------------------------
        # Success
        # --------------------------------------------------

        st.success(
            f"✅ تم تصنيف {len(df):,} حالة بنجاح"
        )


        # ==================================================
        # Dashboard
        # ==================================================

        filtered_df = render_dashboard(

            df,

            date_column=date_column,

            collector_column=collector_column
        )


        # ==================================================
        # Results
        # ==================================================

        st.markdown(
            '<div class="divider"></div>',
            unsafe_allow_html=True
        )


        st.markdown(
            "### 📋 نتائج التصنيف"
        )


        st.dataframe(

            filtered_df,

            use_container_width=True,

            height=500,

            hide_index=True
        )


        # ==================================================
        # Download Excel
        # ==================================================

        excel_file = create_excel(
            filtered_df
        )


        st.download_button(

            label="⬇️ تحميل نتائج التصنيف Excel",

            data=excel_file,

            file_name="activity_classification.xlsx",

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True
        )


# ==========================================================
# Promises
# ==========================================================

def render_waeed():

    st.markdown(
        """
        <div class="hero-wrap">

            <div class="hero-eyebrow">
                PROMISES
            </div>

            <p class="hero-title">
                🤝 الوعود
            </p>

            <p class="hero-subtitle">
                القسم ده لسه تحت الإنشاء
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <div class="card"
             style="text-align:center;">

            🤝 قريبًا

        </div>
        """,
        unsafe_allow_html=True
    )


# ==========================================================
# Neglect
# ==========================================================

def render_ihmal():

    st.markdown(
        """
        <div class="hero-wrap">

            <div class="hero-eyebrow">
                NEGLECT
            </div>

            <p class="hero-title">
                🗂️ الاهمال
            </p>

            <p class="hero-subtitle">
                القسم ده لسه تحت الإنشاء
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <div class="card"
             style="text-align:center;">

            🗂️ قريبًا

        </div>
        """,
        unsafe_allow_html=True
    )


# ==========================================================
# Navigation
# ==========================================================

NAV_ITEMS = {

    "النشاط":
        "📡",

    "الوعود":
        "🤝",

    "الاهمال":
        "🗂️",
}


# ==========================================================
# Sidebar
# ==========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="brand-block">

            <div class="brand-logo">
                🧭
            </div>

            <div>

                <div class="brand-name">
                    النشاط
                </div>

                <div class="brand-sub">
                    Dashboard
                </div>

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


    selected_label = st.radio(

        "التنقل",

        labels,

        label_visibility="collapsed"
    )


    selected_section = (
        selected_label
        .split("  ", 1)[1]
    )


# ==========================================================
# Render
# ==========================================================

if selected_section == "النشاط":

    render_nashat()

elif selected_section == "الوعود":

    render_waeed()

elif selected_section == "الاهمال":

    render_ihmal()
