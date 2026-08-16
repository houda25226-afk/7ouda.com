import io
import os
import time
import requests
import pandas as pd
import streamlit as st
import plotly.express as px


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Activity Dashboard",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# THEME
# =========================================================

CSS = """
<style>

@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&display=swap');

:root {
    --bg: #08111f;
    --bg2: #0d1728;
    --card: #111d30;
    --card2: #16243a;
    --green: #34d399;
    --green2: #10b981;
    --yellow: #fbbf24;
    --red: #f87171;
    --blue: #60a5fa;
    --text: #f8fafc;
    --muted: #94a3b8;
    --border: rgba(255,255,255,0.08);
}

html, body, [class*="css"] {
    font-family: 'Cairo', sans-serif;
}

.stApp {
    background:
        radial-gradient(
            circle at 85% 0%,
            rgba(52,211,153,0.10),
            transparent 30%
        ),
        radial-gradient(
            circle at 10% 100%,
            rgba(96,165,250,0.08),
            transparent 30%
        ),
        var(--bg);
}

/* Hide default */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    background: transparent !important;
}

/* Sidebar */

section[data-testid="stSidebar"] {
    background: #070d18;
    border-right: 1px solid var(--border);
}

section[data-testid="stSidebar"] * {
    font-family: 'Cairo', sans-serif;
}

.sidebar-brand {
    padding: 18px 8px 22px 8px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 20px;
}

.sidebar-logo {
    width: 48px;
    height: 48px;
    border-radius: 15px;

    background:
        linear-gradient(
            135deg,
            var(--green),
            var(--yellow)
        );

    display: flex;
    align-items: center;
    justify-content: center;

    font-size: 24px;

    box-shadow:
        0 8px 30px
        rgba(52,211,153,0.22);
}

.sidebar-title {
    font-size: 20px;
    font-weight: 900;
    color: var(--text);
    margin-top: 10px;
}

.sidebar-subtitle {
    color: var(--muted);
    font-size: 12px;
}

/* Main */

.main-title {
    font-size: 36px;
    font-weight: 900;
    color: var(--text);
    margin-bottom: 3px;
}

.main-subtitle {
    color: var(--muted);
    font-size: 15px;
    margin-bottom: 28px;
}

/* Cards */

.dashboard-card {
    background:
        linear-gradient(
            145deg,
            rgba(22,36,58,0.96),
            rgba(13,23,40,0.96)
        );

    border:
        1px solid var(--border);

    border-radius: 18px;

    padding: 22px;

    box-shadow:
        0 10px 40px
        rgba(0,0,0,0.18);

    margin-bottom: 18px;
}

/* KPI */

.kpi {
    background:
        linear-gradient(
            145deg,
            #122139,
            #0d1829
        );

    border: 1px solid var(--border);

    border-radius: 18px;

    padding: 20px;

    min-height: 125px;

    position: relative;

    overflow: hidden;
}

.kpi::after {
    content: "";

    position: absolute;

    width: 90px;
    height: 90px;

    right: -30px;
    top: -30px;

    background:
        radial-gradient(
            circle,
            rgba(52,211,153,0.15),
            transparent 70%
        );
}

.kpi-label {
    color: var(--muted);
    font-size: 13px;
}

.kpi-value {
    color: var(--text);
    font-size: 30px;
    font-weight: 900;
    margin-top: 6px;
}

.kpi-icon {
    font-size: 25px;
}

/* Upload */

.upload-card {
    background:
        linear-gradient(
            145deg,
            rgba(17,29,48,0.98),
            rgba(12,22,37,0.98)
        );

    border:
        1px dashed
        rgba(52,211,153,0.45);

    border-radius: 20px;

    padding: 24px;

    margin-bottom: 20px;
}

/* Buttons */

.stButton > button {
    width: 100%;

    border: none !important;

    border-radius: 12px !important;

    padding: 11px 18px !important;

    font-family: 'Cairo', sans-serif !important;

    font-weight: 800 !important;

    background:
        linear-gradient(
            90deg,
            var(--green),
            var(--green2)
        ) !important;

    color: #04130d !important;

    transition: 0.2s;
}

.stButton > button:hover {
    transform: translateY(-2px);

    box-shadow:
        0 8px 25px
        rgba(52,211,153,0.25);
}

/* Download */

.stDownloadButton > button {
    width: 100%;

    border-radius: 12px !important;

    font-family: 'Cairo', sans-serif !important;

    font-weight: 800 !important;
}

/* Dataframe */

[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
}

/* File uploader */

[data-testid="stFileUploader"] {
    background: transparent;
}

[data-testid="stFileUploaderDropzone"] {
    background: #0c1728 !important;
    border-radius: 14px !important;
    border: 1px dashed rgba(52,211,153,0.35) !important;
}

/* Progress */

[data-testid="stProgress"] > div > div {
    background:
        linear-gradient(
            90deg,
            var(--green),
            var(--yellow)
        );
}

/* Alert */

[data-testid="stAlert"] {
    border-radius: 12px;
}

/* RTL */

.block-container {
    direction: rtl;
}

</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


# =========================================================
# SETTINGS
# =========================================================

# ضع Model ID الخاص بك هنا إذا لم تضعه في Streamlit Secrets
DEFAULT_MODEL = ""

try:
    HF_TOKEN = st.secrets["HF_TOKEN"]
except Exception:
    HF_TOKEN = os.getenv("HF_TOKEN", "")

try:
    HF_MODEL = st.secrets["HF_MODEL"]
except Exception:
    HF_MODEL = os.getenv("HF_MODEL", DEFAULT_MODEL)


# =========================================================
# SESSION STATE
# =========================================================

if "df_original" not in st.session_state:
    st.session_state.df_original = None

if "df_result" not in st.session_state:
    st.session_state.df_result = None

if "file_name" not in st.session_state:
    st.session_state.file_name = None

if "classified" not in st.session_state:
    st.session_state.classified = False


# =========================================================
# FUNCTIONS
# =========================================================

def clean_dataframe(df):
    """
    حذف أول صف بعد أسماء الأعمدة.
    """

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    return df


def detect_text_column(df):
    """
    محاولة اكتشاف عمود النص تلقائيًا.
    """

    possible_names = [
        "text",
        "Text",
        "النص",
        "الشكوى",
        "الملاحظة",
        "الملاحظات",
        "الوصف",
        "Description",
        "comment",
        "Comment",
        "message",
        "Message",
        "المعاملة",
        "الحالة",
    ]

    for col in possible_names:
        if col in df.columns:
            return col

    # اختيار أول عمود نصي
    object_cols = df.select_dtypes(
        include=["object"]
    ).columns.tolist()

    if object_cols:
        return object_cols[0]

    return None


def detect_collector_column(df):
    possible = [
        "المحصل",
        "اسم المحصل",
        "collector",
        "Collector",
        "collector_name",
        "Collector Name",
        "اسم_المحصل",
    ]

    for col in possible:
        if col in df.columns:
            return col

    return None


def detect_date_column(df):

    possible = [
        "التاريخ",
        "تاريخ",
        "date",
        "Date",
        "transaction_date",
        "created_at",
    ]

    for col in possible:
        if col in df.columns:
            return col

    return None


def classify_text(text):

    if not HF_TOKEN:
        raise ValueError(
            "لم يتم العثور على HF_TOKEN في Streamlit Secrets."
        )

    if not HF_MODEL:
        raise ValueError(
            "لم يتم تحديد HF_MODEL في Streamlit Secrets."
        )

    url = (
        "https://router.huggingface.co/"
        f"hf-inference/models/{HF_MODEL}"
    )

    headers = {
        "Authorization": f"Bearer {HF_TOKEN}",
        "Content-Type": "application/json",
    }

    payload = {
        "inputs": str(text),
        "parameters": {
            "top_k": 10
        }
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=90,
    )

    if response.status_code != 200:

        raise RuntimeError(
            f"Hugging Face Error "
            f"{response.status_code}: "
            f"{response.text[:500]}"
        )

    result = response.json()

    # بعض نماذج HF ترجع قائمة مباشرة
    if isinstance(result, list):

        # لو nested list
        if len(result) > 0 and isinstance(result[0], list):
            result = result[0]

        if len(result) == 0:
            return 0, 0.0, "UNKNOWN"

        best = max(
            result,
            key=lambda x: float(
                x.get("score", 0)
            )
        )

        label = str(
            best.get("label", "")
        )

        score = float(
            best.get("score", 0)
        )

        # محاولة تحويل labels إلى 0 / 1
        label_lower = label.lower()

        if label_lower in [
            "1",
            "label_1",
            "positive",
            "successful",
            "success",
            "ناجحة",
            "ناجح",
        ]:
            prediction = 1

        elif label_lower in [
            "0",
            "label_0",
            "negative",
            "unsuccessful",
            "failure",
            "failed",
            "غير ناجحة",
            "غير ناجح",
        ]:
            prediction = 0

        else:
            # لو النموذج مدرب على LABEL_0 / LABEL_1
            if "1" in label_lower:
                prediction = 1
            else:
                prediction = 0

        return prediction, score, label

    raise RuntimeError(
        f"Unexpected Hugging Face response: {result}"
    )


def run_classification(df, text_column):

    predictions = []
    probabilities = []
    raw_labels = []

    total = len(df)

    progress = st.progress(0)

    status = st.empty()

    for i, text in enumerate(
        df[text_column].fillna("").astype(str)
    ):

        # إعادة المحاولة البسيطة
        last_error = None

        for attempt in range(3):

            try:

                pred, prob, raw = classify_text(text)

                predictions.append(pred)
                probabilities.append(prob)
                raw_labels.append(raw)

                break

            except Exception as e:

                last_error = e

                if attempt < 2:
                    time.sleep(2)

        else:

            raise RuntimeError(
                f"فشل تصنيف الصف رقم {i + 1}: "
                f"{last_error}"
            )

        percent = int(
            ((i + 1) / total) * 100
        )

        progress.progress(percent)

        status.markdown(
            f"**جاري التصنيف:** "
            f"{i + 1} / {total}"
        )

    progress.empty()
    status.empty()

    result = df.copy()

    result["Prediction"] = predictions

    result["Probability"] = probabilities

    result["Model_Label"] = raw_labels

    # 1 = ناجحة
    # 0 = غير ناجحة

    result["Prediction_Label"] = result[
        "Prediction"
    ].map({
        1: "ناجحة",
        0: "غير ناجحة"
    })

    return result


def make_excel_download(df):

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

    output.seek(0)

    return output.getvalue()


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">

            <div class="sidebar-logo">
                📡
            </div>

            <div class="sidebar-title">
                النشاط
            </div>

            <div class="sidebar-subtitle">
                Activity Dashboard
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )

    section = st.radio(
        "الأقسام",
        [
            "📡 النشاط",
            "🤝 الوعود",
            "🗂️ الإهمال",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    st.caption(
        "Dashboard v2.0"
    )


# =========================================================
# ACTIVITY PAGE
# =========================================================

if section == "📡 النشاط":

    st.markdown(
        """
        <div class="main-title">
            📡 لوحة النشاط
        </div>

        <div class="main-subtitle">
            ارفع ملف البيانات وشغّل نموذج التصنيف للحصول
            على تحليل كامل للنشاط.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -----------------------------------------------------
    # Upload
    # -----------------------------------------------------

    st.markdown(
        """
        <div class="upload-card">

        <h3>
        📂 رفع ملف البيانات
        </h3>

        <p style="color:#94a3b8;">
        ارفع ملف Excel أو CSV ثم اضغط على
        "ابدأ التصنيف".
        </p>

        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "اختار الملف",
        type=[
            "csv",
            "xlsx",
            "xls"
        ],
        label_visibility="collapsed",
    )

    if uploaded_file is not None:

        # لو المستخدم غير الملف
        if (
            st.session_state.file_name
            != uploaded_file.name
        ):

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

                # حذف أول صف
                df = clean_dataframe(df)

                st.session_state.df_original = df

                st.session_state.df_result = None

                st.session_state.file_name = (
                    uploaded_file.name
                )

                st.session_state.classified = False

            except Exception as e:

                st.error(
                    f"حدث خطأ أثناء قراءة الملف: {e}"
                )
                st.stop()

        df = st.session_state.df_original

        # -------------------------------------------------
        # File information
        # -------------------------------------------------

        col1, col2, col3 = st.columns(3)

        with col1:

            st.markdown(
                f"""
                <div class="kpi">

                    <div class="kpi-icon">
                    📄
                    </div>

                    <div class="kpi-label">
                    اسم الملف
                    </div>

                    <div style="
                        color:#f8fafc;
                        font-weight:700;
                        margin-top:5px;
                    ">
                    {uploaded_file.name}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

        with col2:

            st.markdown(
                f"""
                <div class="kpi">

                    <div class="kpi-icon">
                    📊
                    </div>

                    <div class="kpi-label">
                    عدد الصفوف
                    </div>

                    <div class="kpi-value">
                    {len(df):,}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

        with col3:

            st.markdown(
                f"""
                <div class="kpi">

                    <div class="kpi-icon">
                    🧱
                    </div>

                    <div class="kpi-label">
                    عدد الأعمدة
                    </div>

                    <div class="kpi-value">
                    {len(df.columns):,}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

        st.write("")

        # -------------------------------------------------
        # Detect text column
        # -------------------------------------------------

        text_column = detect_text_column(df)

        if text_column:

            st.info(
                f"📝 عمود النص المستخدم في التصنيف: "
                f"**{text_column}**"
            )

        else:

            st.error(
                "لم أستطع تحديد عمود النص تلقائيًا."
            )

        # -------------------------------------------------
        # Preview
        # -------------------------------------------------

        with st.expander(
            "👀 معاينة البيانات",
            expanded=False
        ):

            st.dataframe(
                df.head(20),
                use_container_width=True,
                height=350,
            )

        # -------------------------------------------------
        # Start Classification
        # -------------------------------------------------

        st.markdown("### 🚀 التصنيف")

        if not HF_TOKEN:

            st.warning(
                "⚠️ لم يتم إعداد HF_TOKEN."
                " أضفه في Streamlit Secrets."
            )

        if not HF_MODEL:

            st.warning(
                "⚠️ لم يتم إعداد HF_MODEL."
                " أضفه في Streamlit Secrets."
            )

        if text_column:

            if st.button(
                "🚀 ابدأ التصنيف",
                type="primary",
                use_container_width=True,
            ):

                try:

                    with st.spinner(
                        "جاري تشغيل الموديل..."
                    ):

                        result = run_classification(
                            df,
                            text_column
                        )

                    st.session_state.df_result = result

                    st.session_state.classified = True

                    st.success(
                        "✅ تم الانتهاء من التصنيف بنجاح."
                    )

                except Exception as e:

                    st.error(
                        f"❌ حصل خطأ أثناء التصنيف:\n\n{e}"
                    )

        # =================================================
        # DASHBOARD
        # =================================================

        if (
            st.session_state.classified
            and
            st.session_state.df_result is not None
        ):

            result = (
                st.session_state.df_result
            )

            st.markdown("---")

            st.markdown(
                """
                <div class="main-title"
                     style="font-size:28px;">
                    📊 Dashboard
                </div>

                <div class="main-subtitle">
                    تحليل ديناميكي لنتائج التصنيف.
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ---------------------------------------------
            # KPIs
            # ---------------------------------------------

            total = len(result)

            successful = int(
                (result["Prediction"] == 1).sum()
            )

            unsuccessful = int(
                (result["Prediction"] == 0).sum()
            )

            success_rate = (
                successful / total * 100
                if total > 0
                else 0
            )

            avg_probability = (
                result["Probability"].mean() * 100
                if total > 0
                else 0
            )

            k1, k2, k3, k4, k5 = st.columns(5)

            kpis = [
                (
                    k1,
                    "📦",
                    "إجمالي العمليات",
                    f"{total:,}"
                ),
                (
                    k2,
                    "✅",
                    "ناجحة",
                    f"{successful:,}"
                ),
                (
                    k3,
                    "❌",
                    "غير ناجحة",
                    f"{unsuccessful:,}"
                ),
                (
                    k4,
                    "📈",
                    "نسبة النجاح",
                    f"{success_rate:.1f}%"
                ),
                (
                    k5,
                    "🤖",
                    "متوسط الثقة",
                    f"{avg_probability:.1f}%"
                ),
            ]

            for col, icon, label, value in kpis:

                with col:

                    st.markdown(
                        f"""
                        <div class="kpi">

                            <div class="kpi-icon">
                            {icon}
                            </div>

                            <div class="kpi-label">
                            {label}
                            </div>

                            <div class="kpi-value">
                            {value}
                            </div>

                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            st.write("")

            # ---------------------------------------------
            # Charts Row 1
            # ---------------------------------------------

            c1, c2 = st.columns(2)

            # Distribution
            with c1:

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
                    title="توزيع نتائج التصنيف",
                )

                fig.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font_family="Cairo",
                    legend_title="",
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

            # Probability
            with c2:

                fig2 = px.histogram(
                    result,
                    x="Probability",
                    nbins=20,
                    title="توزيع احتمالية التصنيف",
                )

                fig2.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font_family="Cairo",
                    xaxis_title="Probability",
                    yaxis_title="عدد العمليات",
                )

                st.plotly_chart(
                    fig2,
                    use_container_width=True
                )

            # ---------------------------------------------
            # Collector Dashboard
            # ---------------------------------------------

            collector_col = detect_collector_column(
                result
            )

            if collector_col:

                st.markdown(
                    "### 👥 أداء المحصلين"
                )

                collector_stats = (
                    result
                    .groupby(collector_col)
                    .agg(
                        إجمالي=(
                            "Prediction",
                            "count"
                        ),
                        ناجحة=(
                            "Prediction",
                            lambda x:
                            (x == 1).sum()
                        ),
                        غير_ناجحة=(
                            "Prediction",
                            lambda x:
                            (x == 0).sum()
                        ),
                        متوسط_الثقة=(
                            "Probability",
                            "mean"
                        ),
                    )
                    .reset_index()
                )

                collector_stats[
                    "نسبة_النجاح"
                ] = (
                    collector_stats["ناجحة"]
                    /
                    collector_stats["إجمالي"]
                    * 100
                )

                collector_stats = (
                    collector_stats
                    .sort_values(
                        "نسبة_النجاح",
                        ascending=False
                    )
                )

                left, right = st.columns(2)

                with left:

                    fig3 = px.bar(
                        collector_stats.head(15),
                        x="نسبة_النجاح",
                        y=collector_col,
                        orientation="h",
                        title="نسبة النجاح حسب المحصل",
                    )

                    fig3.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        font_family="Cairo",
                    )

                    st.plotly_chart(
                        fig3,
                        use_container_width=True
                    )

                with right:

                    st.dataframe(
                        collector_stats,
                        use_container_width=True,
                        hide_index=True,
                    )

            # ---------------------------------------------
            # Date Dashboard
            # ---------------------------------------------

            date_col = detect_date_column(
                result
            )

            if date_col:

                temp = result.copy()

                temp[date_col] = pd.to_datetime(
                    temp[date_col],
                    errors="coerce"
                )

                temp = temp.dropna(
                    subset=[date_col]
                )

                if len(temp) > 0:

                    daily = (
                        temp
                        .groupby(
                            temp[date_col].dt.date
                        )
                        .agg(
                            إجمالي=(
                                "Prediction",
                                "count"
                            ),
                            ناجحة=(
                                "Prediction",
                                lambda x:
                                (x == 1).sum()
                            ),
                            غير_ناجحة=(
                                "Prediction",
                                lambda x:
                                (x == 0).sum()
                            ),
                        )
                        .reset_index()
                    )

                    daily[
                        "نسبة_النجاح"
                    ] = (
                        daily["ناجحة"]
                        /
                        daily["إجمالي"]
                        * 100
                    )

                    st.markdown(
                        "### 📅 النشاط اليومي"
                    )

                    fig4 = px.line(
                        daily,
                        x=date_col,
                        y=[
                            "إجمالي",
                            "ناجحة",
                            "غير_ناجحة"
                        ],
                        markers=True,
                        title="حركة النشاط يوميًا",
                    )

                    fig4.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        font_family="Cairo",
                    )

                    st.plotly_chart(
                        fig4,
                        use_container_width=True
                    )

            # ---------------------------------------------
            # Result table
            # ---------------------------------------------

            st.markdown(
                "### 📋 نتائج التصنيف"
            )

            st.dataframe(
                result,
                use_container_width=True,
                height=500,
            )

            # ---------------------------------------------
            # Download
            # ---------------------------------------------

            st.markdown(
                "### ⬇️ تحميل النتائج"
            )

            excel_data = make_excel_download(
                result
            )

            st.download_button(
                label="⬇️ تحميل ملف النتائج Excel",
                data=excel_data,
                file_name="classification_results.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

            # Reset

            if st.button(
                "🔄 بدء تحليل ملف جديد",
                use_container_width=True,
            ):

                st.session_state.df_original = None
                st.session_state.df_result = None
                st.session_state.file_name = None
                st.session_state.classified = False

                st.rerun()


# =========================================================
# PROMISES
# =========================================================

elif section == "🤝 الوعود":

    st.markdown(
        """
        <div class="main-title">
            🤝 الوعود
        </div>

        <div class="main-subtitle">
            قسم الوعود — جاهز للإضافة لاحقًا.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "القسم ده تحت الإنشاء."
    )


# =========================================================
# NEGLECT
# =========================================================

elif section == "🗂️ الإهمال":

    st.markdown(
        """
        <div class="main-title">
            🗂️ الإهمال
        </div>

        <div class="main-subtitle">
            قسم الإهمال — جاهز للإضافة لاحقًا.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "القسم ده تحت الإنشاء."
    )
