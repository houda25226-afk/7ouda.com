"""
تطبيق Streamlit لتصنيف إفادات المكالمات (ناجحة / غير ناجحة)
باستخدام الموديل المرفوع على Hugging Face: Mahmoud252002/7oudaModel

طريقة التشغيل محليًا:
    pip install streamlit transformers torch pandas openpyxl --break-system-packages
    streamlit run app.py

============================================================
دليل التخصيص السريع (لو عايز تضيف/تعدّل حاجة بعدين):
- الألوان والخطوط كلها في متغير CSS_THEME تحت — عدّل من هناك بس.
- منطق التصنيف كله في قسم "منطق الموديل" — منفصل عن الواجهة.
- كل قسم في الواجهة متعلّم بعنوان تعليقي واضح عشان تلاقي مكانك بسرعة.
============================================================
"""

import io
import pandas as pd
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# إعدادات عامة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256
# التصنيف رقمي: 1 = ناجحة، 0 = غير ناجحة
LABEL_MAP = {0: 0, 1: 1}

st.set_page_config(
    page_title="النشاط",
    page_icon="🎯",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ==========================================================
# الهوية البصرية (Theme)
# ==========================================================
# لوحة الألوان:
#   خلفية كحلي-فحمي غامق (#0F1115) هادية ومريحة للعين
#   سطوح البطاقات (#181B22) بدرجة أفتح شوية، بحواف ناعمة
#   لون أساسي كهرماني دافئ (#F2A93B) — إحساس تنبيه/حيوية
#   لون تكميلي أزرق-تركواز هادي (#3FB6A8) للتباين
#   أخضر للنجاح (#3DD68C) ووردي-أحمر لعدم النجاح (#F0616D)
#   نص أساسي فاتح (#F4F5F7) ونص ثانوي رمادي (#8B92A0)

CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0F1115;
    --surface: #181B22;
    --surface-2: #20242D;
    --accent: #F2A93B;
    --accent-2: #3FB6A8;
    --success: #3DD68C;
    --danger: #F0616D;
    --text: #F4F5F7;
    --text-dim: #8B92A0;
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
        radial-gradient(circle at 85% -10%, rgba(242,169,59,0.08) 0%, transparent 45%),
        radial-gradient(circle at 10% 110%, rgba(63,182,168,0.08) 0%, transparent 45%),
        var(--bg);
    color: var(--text);
}

#MainMenu, footer, header {visibility: hidden;}

/* ===== السايدبار ===== */
section[data-testid="stSidebar"] {
    direction: rtl;
    background: #0B0D11;
    border-left: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: var(--text) !important; }

.brand-block {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    padding: 1rem 0.2rem 1.1rem 0.2rem;
    margin-bottom: 0.4rem;
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
    box-shadow: 0 4px 16px rgba(242, 169, 59, 0.28);
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

/* ===== الهيدر / Hero ===== */
.hero-wrap {
    text-align: center;
    padding: 1rem 0 0.3rem 0;
}
.hero-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.18em;
    color: var(--accent-2);
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

/* نقاط زخرفية نابضة تحت العنوان */
.dots-row {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin: 1.1rem 0 1.6rem 0;
}
.dots-row span {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: var(--accent);
    opacity: 0.35;
    animation: pulse 1.8s ease-in-out infinite;
}
.dots-row span:nth-child(2) { background: var(--accent-2); animation-delay: 0.3s; }
.dots-row span:nth-child(3) { background: var(--accent); animation-delay: 0.6s; }
.dots-row span:nth-child(4) { background: var(--accent-2); animation-delay: 0.9s; }
.dots-row span:nth-child(5) { background: var(--accent); animation-delay: 1.2s; }
@keyframes pulse {
    0%, 100% { opacity: 0.3; transform: scale(1); }
    50% { opacity: 1; transform: scale(1.5); }
}

/* ===== البطاقات العامة ===== */
.card {
    background: var(--surface);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 16px;
    padding: 1.15rem 1.35rem;
    margin-bottom: 1rem;
}

/* ===== منطقة رفع الملفات ===== */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(242, 169, 59, 0.4);
    border-radius: 16px;
    padding: 0.7rem;
}
[data-testid="stFileUploader"] section {
    background: transparent;
}

/* ===== الأزرار ===== */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), #E08F1F);
    color: #201202;
    font-weight: 700;
    border: none;
    border-radius: 11px;
    padding: 0.65rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 20px rgba(242, 169, 59, 0.3);
    color: #201202;
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
    background: linear-gradient(90deg, transparent, rgba(242,169,59,0.35), transparent);
    margin: 1.4rem 0;
    border: none;
}
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


def render_dots():
    """صف نقاط زخرفية بسيطة تحت العنوان."""
    st.markdown(
        '<div class="dots-row"><span></span><span></span><span></span><span></span><span></span></div>',
        unsafe_allow_html=True,
    )


# ==========================================================
# الهيدر
# ==========================================================

st.markdown(
    """
    <div class="hero-wrap">
        <div class="hero-eyebrow">CALL QUALITY CLASSIFIER</div>
        <p class="hero-title">تصنيف إفادات المكالمات</p>
        <p class="hero-subtitle">ارفع ملف Excel أو CSV، وهيتصنّف كل صف تلقائيًا لـ 1 (ناجحة) أو 0 (غير ناجحة)</p>
    </div>
    """,
    unsafe_allow_html=True,
)
render_dots()


# ==========================================================
# الشريط الجانبي — الهوية + الإعدادات
# ==========================================================

with st.sidebar:
    st.markdown(
        """
        <div class="brand-block">
            <div class="brand-logo">🎯</div>
            <div>
                <div class="brand-name">النشاط</div>
                <div class="brand-sub">Call Activity</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("### ⚙️ الإعدادات")
    text_column_input = st.text_input("اسم عمود النص (الإفادة)", value="الافادة")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**الموديل المستخدم**")
    st.code(MODEL_REPO, language=None)
    st.markdown(f"[عرض الموديل على Hugging Face ↗](https://huggingface.co/{MODEL_REPO})")


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
    """بيصنف قائمة نصوص على دفعات (batches) عشان الأداء يبقى أسرع."""
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
# رفع الملف
# ==========================================================

uploaded_file = st.file_uploader("ارفع ملف البيانات (CSV أو Excel)", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        st.stop()

    # ------------------------------------------------------
    # معالجة تلقائية للملف بعد رفعه مباشرة:
    #   1) تغيير اسم عمود "Note" لـ "الافادة" (لو موجود)
    #   2) حذف أول صف بيانات بعد صف العناوين (index 0)
    # ------------------------------------------------------
    if "Note" in df.columns:
        df = df.rename(columns={"Note": "الافادة"})

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    st.markdown(
        f'<div class="card">✅ تم تحميل الملف بنجاح — عدد الصفوف: <b>{len(df)}</b></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(df.head(10), use_container_width=True)

    if text_column_input not in df.columns:
        st.error(
            f"عمود '{text_column_input}' مش موجود في الملف. "
            f"الأعمدة الموجودة فعلاً: {', '.join(df.columns.astype(str))}"
        )
        st.stop()

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if st.button("🚀 ابدأ التصنيف", type="primary", use_container_width=True):
        tokenizer, model, device = load_model()

        texts = df[text_column_input].tolist()
        preds, confidences = predict_batch(texts, tokenizer, model, device)

        result_df = df.copy()
        result_df["التصنيف_المتوقع"] = [LABEL_MAP[p] for p in preds]
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]

        # بعد ما التصنيف يخلص، رجّع اسم عمود النص لـ "Note" في ملف النتيجة
        result_df = result_df.rename(columns={text_column_input: "Note"})

        st.success("تم التصنيف بنجاح ✅")
        st.dataframe(result_df, use_container_width=True)

        counts = result_df["التصنيف_المتوقع"].value_counts()
        col1, col2 = st.columns(2)
        col1.metric("✅ إفادات ناجحة (1)", int(counts.get(1, 0)))
        col2.metric("⛔ إفادات غير ناجحة (0)", int(counts.get(0, 0)))

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

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
            label="⬇️ تحميل الملف مع التصنيف",
            data=output,
            file_name=file_name,
            mime=mime,
            use_container_width=True,
        )
else:
    st.markdown(
        '<div class="card" style="text-align:center; color: var(--text-dim);">'
        "📂 ارفع ملف عشان تبدأ — لازم يحتوي على عمود بالنص المراد تصنيفه."
        "</div>",
        unsafe_allow_html=True,
    )
