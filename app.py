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
LABEL_MAP = {0: "غير ناجحة", 1: "ناجحة"}

st.set_page_config(
    page_title="تصنيف المكالمات | 7oudaModel",
    page_icon="🎙️",
    layout="centered",
)

# ==========================================================
# الهوية البصرية (Theme)
# ==========================================================
# لوحة الألوان:
#   خلفية داكنة (#0E1420) تحاكي لوحة تحكم مركز اتصالات ليلي
#   سطوح البطاقات (#151F30) بدرجة أفتح شوية من الخلفية
#   لون أساسي تركواز (#5EEAD4) — يرمز للصوت/الموجة الصوتية
#   أخضر للنجاح (#34D399) ووردي-أحمر لعدم النجاح (#FB7185)
#   نص أساسي فاتح (#E7ECF3) ونص ثانوي رمادي مزرق (#8B96AC)

CSS_THEME = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&family=JetBrains+Mono:wght@500;700&display=swap');

:root {
    --bg: #0E1420;
    --surface: #151F30;
    --surface-2: #1B2A42;
    --accent: #5EEAD4;
    --success: #34D399;
    --danger: #FB7185;
    --text: #E7ECF3;
    --text-dim: #8B96AC;
}

html, body, [class*="css"]  {
    font-family: 'Tajawal', sans-serif;
    direction: rtl;
}

.stApp {
    background: radial-gradient(circle at 20% 0%, #10192C 0%, var(--bg) 55%);
    color: var(--text);
}

/* إخفاء العناصر الافتراضية الزيادة */
#MainMenu, footer, header {visibility: hidden;}

/* ===== الهيدر / Hero ===== */
.hero-wrap {
    text-align: center;
    padding: 1.2rem 0 0.4rem 0;
}
.hero-eyebrow {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    letter-spacing: 0.15em;
    color: var(--accent);
    text-transform: uppercase;
    margin-bottom: 0.4rem;
}
.hero-title {
    font-weight: 900;
    font-size: 2.1rem;
    color: var(--text);
    margin: 0;
    line-height: 1.3;
}
.hero-subtitle {
    color: var(--text-dim);
    font-size: 0.98rem;
    margin-top: 0.5rem;
}

/* الموجة الصوتية — العنصر المميز */
.waveform {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    height: 34px;
    margin: 1.1rem 0 1.6rem 0;
}
.waveform span {
    display: inline-block;
    width: 3px;
    border-radius: 3px;
    background: linear-gradient(180deg, var(--accent), var(--surface-2));
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

/* ===== الشريط الجانبي ===== */
section[data-testid="stSidebar"] {
    background: #0B111C;
    border-left: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: var(--text) !important; }

/* ===== منطقة رفع الملفات ===== */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1.5px dashed rgba(94, 234, 212, 0.35);
    border-radius: 14px;
    padding: 0.6rem;
}
[data-testid="stFileUploader"] section {
    background: transparent;
}

/* ===== الأزرار ===== */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, var(--accent), #3FD9C7);
    color: #06251F;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(94, 234, 212, 0.25);
    color: #06251F;
}

/* ===== شريط التقدم ===== */
[data-testid="stProgress"] > div > div {
    background: var(--accent);
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
    color: var(--accent);
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
    background: linear-gradient(90deg, transparent, rgba(94,234,212,0.35), transparent);
    margin: 1.4rem 0;
    border: none;
}
</style>
"""

st.markdown(CSS_THEME, unsafe_allow_html=True)


def render_waveform(n_bars: int = 28):
    """موجة صوتية متحركة بسيطة — العنصر البصري المميز للتطبيق."""
    bars = ""
    heights = [14, 22, 30, 18, 26, 34, 20, 28, 16, 24] * (n_bars // 10 + 1)
    for i in range(n_bars):
        delay = (i % 10) * 0.09
        bars += f'<span style="height:{heights[i]}px; animation-delay:{delay}s;"></span>'
    st.markdown(f'<div class="waveform">{bars}</div>', unsafe_allow_html=True)


# ==========================================================
# الهيدر
# ==========================================================

st.markdown(
    """
    <div class="hero-wrap">
        <div class="hero-eyebrow">CALL QUALITY CLASSIFIER</div>
        <p class="hero-title">🎙️ تصنيف إفادات المكالمات</p>
        <p class="hero-subtitle">ارفع ملف Excel أو CSV، وهيتصنّف كل صف تلقائيًا لـ «ناجحة» أو «غير ناجحة»</p>
    </div>
    """,
    unsafe_allow_html=True,
)
render_waveform()


# ==========================================================
# الشريط الجانبي — الإعدادات
# ==========================================================

with st.sidebar:
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

        st.success("تم التصنيف بنجاح ✅")
        st.dataframe(result_df, use_container_width=True)

        counts = result_df["التصنيف_المتوقع"].value_counts()
        col1, col2 = st.columns(2)
        col1.metric("✅ إفادات ناجحة", int(counts.get("ناجحة", 0)))
        col2.metric("⛔ إفادات غير ناجحة", int(counts.get("غير ناجحة", 0)))

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
