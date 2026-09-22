"""
تطبيق Streamlit لتصنيف إفادات المكالمات + حساب الوقت المهدر + داشبورد
باستخدام الموديل: Mahmoud252002/7oudaModel

التشغيل:
    pip install -r requirements.txt
    streamlit run app.py

الصفحات:
  - تصنيف المكالمات (حسب الشركة والفترة)
  - الوعود (قائمة + مكسورة)
  - الإهمال ومتابعة الإهمال
  - الجدولة المتعثرة (محفظة + سدادات)
  - تحليل نشاط المحصّلين (Dashboard + تصدير HTML)

أسماء الأعمدة قابلة للتعديل من قسم الإعدادات أعلى الملف.
"""

import io
import hashlib
import re
import zipfile
from io import BytesIO
from datetime import time as dt_time
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
import torch
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# تهيئة إعدادات الصفحة والمظهر (مبكرًا لمنع أخطاء المتغيرات)
# ==========================================================

st.set_page_config(
    page_title="إيجادة — إدارة المحفظة ونشاط المحصلين",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

THEMES = {
    # هوية الوطنية للتأمين (Wataniya)
    "dark": {
        "bg": "#0A1F14",
        "bg_glow": "#0D2818",
        "surface": "#122A1C",
        "surface_2": "#183524",
        "surface_3": "#0E2318",
        "surface_hover": "#1C3F2A",
        "sidebar_bg": "#081A10",
        "accent_surface": "#143528",
        "accent": "#1A8A4A",
        "accent_strong": "#126D3C",
        "on_accent": "#FFFFFF",
        "success": "#1A8A4A",
        "danger": "#C45C5C",
        "warn": "#F19F29",
        "text": "#F3F9F5",
        "text_dim": "#B8D4C4",
        "text_muted": "#8FB8A0",
        "placeholder": "#7AA890",
        "border": "rgba(26, 138, 74, 0.28)",
        "border_soft": "rgba(148, 180, 160, 0.18)",
        "input_bg": "#183524",
        "chart_marker": "#0A1F14",
        "chart_text": "#FFFFFF",
        "danger_soft": "#3B1A1A", "danger_text": "#FECDD3",
        "warn_soft": "#3A2A08", "warn_text": "#FDE68A",
    },
    "light": {
        "bg": "#F7F9EF",
        "bg_glow": "#EEF3E4",
        "surface": "#FFFFFF",
        "surface_2": "#F0F4E8",
        "surface_3": "#F9FBF4",
        "surface_hover": "#E6EDD8",
        "sidebar_bg": "#FFFFFF",
        "accent_surface": "#E8F5EC",
        "accent": "#126D3C",
        "accent_strong": "#0E5A32",
        "on_accent": "#FFFFFF",
        "success": "#126D3C",
        "danger": "#C45C5C",
        "warn": "#F19F29",
        "text": "#1A2E22",
        "text_dim": "#4A6354",
        "text_muted": "#6B8574",
        "placeholder": "#7A9484",
        "border": "rgba(18, 109, 60, 0.22)",
        "border_soft": "rgba(71, 100, 80, 0.14)",
        "input_bg": "#FFFFFF",
        "chart_marker": "#D4DEC8",
        "chart_text": "#1A2E22",
        "danger_soft": "#FDE8E8", "danger_text": "#8B2E2E",
        "warn_soft": "#FEF3C7", "warn_text": "#92400E",
    },
}


def _detect_native_streamlit_theme() -> str:
    """بيقرأ الوضع (Light/Dark) المختار في Streamlit."""
    try:
        ctx_theme = st.context.theme
        theme_type = getattr(ctx_theme, "type", None)
        if theme_type is None and hasattr(ctx_theme, "get"):
            theme_type = ctx_theme.get("type")
        if theme_type in ("light", "dark"):
            return theme_type
    except Exception:
        pass
    return st.session_state.get("theme_mode", "dark")


THEME_NAME = _detect_native_streamlit_theme()
st.session_state["theme_mode"] = THEME_NAME
THEME = THEMES.get(THEME_NAME, THEMES["dark"])


def _inject_wataniya_identity_css():
    """حقن هوية الوطنية والتنسيقات."""
    t = THEME
    css = f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Tajawal', sans-serif !important;
        background-color: {t["bg"]} !important;
        color: {t["text"]} !important;
        direction: rtl;
    }}

    .stApp {{
        background: radial-gradient(circle at 50% 0%, {t["bg_glow"]} 0%, {t["bg"]} 75%) !important;
    }}

    .stButton>button {{
        background-color: {t["accent"]} !important;
        color: {t["on_accent"]} !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 700 !important;
    }}
    .stButton>button:hover {{
        background-color: {t["accent_strong"]} !important;
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


_inject_wataniya_identity_css()

# ==========================================================
# إعدادات وأسماء الأعمدة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

ORIGINAL_TEXT_COL = "Notes"         # اسم العمود الأصلي في الملف
NOTES_CANDIDATES = [
    "Notes", "notes", "NOTE", "Note", "الافادة", "الإفادة", "افادة", "إفادة",
    "Call Notes", "call notes", "Comment", "Comments", "ملاحظات",
]
MODEL_TEXT_COL = "الافادة"          # الاسم اللي بيتحول له مؤقتًا لـ الموديل
CLASSIFICATION_COL = "التصنيف"      # عمود النتيجة: 1 = ناجحة / 0 = غير ناجحة
WASTED_TIME_COL = "الوقت_المهدر_دقيقة"

ID_CANDIDATES = ["Account ID", "account id", "AccountID", "ID", "id", "رقم الحساب", "الرقم التعريفي", "Account No", "account no", "Account Number"]
SALES_PERSON_CANDIDATES = ["Create By", "create by", "CreateBy", "Created By", "created by", "Sales Person", "sales person", "المحصّل", "Salesperson", "salesperson", "SalesPerson"]
SALES_TEAM_CANDIDATES = ["Sales Team", "sales team", "SalesTeam", "Team", "team", "فريق المبيعات", "الفريق", "فريق", "Sales team"]
COLLECTED_BY_CANDIDATES = ["Collected by", "collected by", "Collected By", "COLLECTED BY", "Created by", "created by", "Created By", "CREATED BY", "المحصل", "المحصّل", "Collector", "collector"]
ACCOUNT_NUMBER_CANDIDATES = ["Customer Account number", "Customer Account Number", "customer account number", "Customer Account No", "Account Number", "account number", "Account No", "رقم حساب العميل", "رقم الحساب"]
CREATED_ON_CANDIDATES = ["Created On", "created on", "CreatedOn", "تاريخ الافادة"]
CLAIM_CANDIDATES = ["Claim", "claim", "CLAIM", "رقم المطالبة", "رقم المطالبه"]
DUPLICATE_WINDOW_MINUTES = 20
DURATION_CANDIDATES = ["Call Duration", "call duration", "CallDuration", "Duration", "duration", "مدة المكالمة", "Call Time", "call time", "Talk Time", "talk time", "Duration (min)", "مدة"]

# ========================================================== 
# إعدادات تويب الوعود القائمة (المحفظة)
# ==========================================================
PROMISE_SUB_STATE_CANDIDATES = ["Sub State", "sub state", "SubState", "الحالة الفرعية"]
PROMISE_DUE_DATE_CANDIDATES = ["Follow up Due Date", "follow up due date", "FollowUpDueDate", "تاريخ المتابعة", "Due Date"]
PROMISE_NET_AMOUNT_CANDIDATES = ["Net Amount", "net amount", "NetAmount", "صافي المبلغ", "مبلغ المديونية"]

PROMISE_SUB_STATE_VALUE = "واعد بالسداد"

PROMISE_EXCLUDED_SALES = [
    "Archive Companies  II Anas",
    "Closed payments  II Anas",
    "Hold Companies  II Anas",
    "Op II Ibrahim Qassem",
    "قانونى -الوطنية",
]
PROMISES_EXCLUDED_SALES_KEY = "promises_excluded_sales"
PROMISES_AVAILABLE_SALES_KEY = "promises_available_sales"


def init_promises_sales_filter():
    """تهيئة قائمة المحصلين المستبعدين."""
    if PROMISES_EXCLUDED_SALES_KEY not in st.session_state:
        st.session_state[PROMISES_EXCLUDED_SALES_KEY] = list(PROMISE_EXCLUDED_SALES)
    if PROMISES_AVAILABLE_SALES_KEY not in st.session_state:
        st.session_state[PROMISES_AVAILABLE_SALES_KEY] = []


def _promises_excluded_sales():
    init_promises_sales_filter()
    return [str(x).strip() for x in st.session_state.get(PROMISES_EXCLUDED_SALES_KEY, []) if str(x).strip()]


def _promises_exclusion_token():
    """توكن لتحديث الكاش عند تغيير فلتر المحصلين."""
    return hashlib.sha256("|".join(sorted(_promises_excluded_sales())).encode("utf-8")).hexdigest()[:16]


def _extract_promises_sales_from_upload(uploaded):
    """قراءة أسماء المحصلين من ملف المحفظة بدون تطبيق فلاتر الوعود."""
    if uploaded is None:
        return []
    try:
        raw_df = read_uploaded_dataframe(uploaded)
    except Exception:
        return []
    df = raw_df.iloc[1:].copy() if len(raw_df) > 0 else raw_df
    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    if not sales_col:
        return []
    vals = df[sales_col].astype(str).str.strip()
    return sorted({v for v in vals.tolist() if v and v.lower() not in {"nan", "none", "null", ""}})


def _sync_promises_available_sales(uploaded=None, result_keys=None):
    """تحديث قائمة المحصلين المتاحة من الملف أو من الكاش."""
    init_promises_sales_filter()
    found = []
    if uploaded is not None:
        found = _extract_promises_sales_from_upload(uploaded)
    if not found and result_keys:
        for key in result_keys:
            cached = st.session_state.get(key) or {}
            found = list(cached.get("available_sales") or [])
            if found:
                break
            df = cached.get("df")
            sales_col = cached.get("sales_col")
            if df is not None and sales_col and sales_col in getattr(df, "columns", []):
                vals = df[sales_col].astype(str).str.strip()
                found = sorted({v for v in vals.tolist() if v and v.lower() not in {"nan", "none", "null", ""}})
                if found:
                    break
    if found:
        merged = sorted(set(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, [])) | set(found))
        st.session_state[PROMISES_AVAILABLE_SALES_KEY] = merged
    return list(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, []))


# ==========================================================
# إعدادات تويب الإهمال
# ==========================================================
NEGLECT_SUB_STATES_DEFAULT = [
    "تم ابلاغ العميل - اتصال",
    "لايرد مع التكرار",
    "جدولة",
    "واعد بالسداد",
    "تم ابلاغ العميل - واتسب",
    "لا يرد",
    "إعفاء || بإنتظار المستند",
    "متوفي",
    "مغلق مع التكرار"
]

NEGLECT_LAST_DATE_CANDIDATES = ["Follow up Last Date", "follow up last date", "Followup Last Date", "FollowUpLastDate", "Last Follow Up", "Last Follow-up", "تاريخ آخر متابعة", "تاريخ اخر متابعة", "آخر متابعة"]
NEGLECT_RESULT_KEY = "neglect_result"
APP_DATA_CACHE_KEY = "app_uploaded_data_cache"
DASHBOARD_SOURCE_HASH_KEY = "dashboard_source_hash"

DASH_WALLET_CACHE_KEY = "dashboard_wallet_cache"
DASH_PAYMENTS_CACHE_KEY = "dashboard_payments_cache"
DASH_ACTIVE_PAGE_KEY = "dashboard_active_page"

# ==========================================================
# أعمدة إضافية لتحليل المحفظة داخل داشبورد النشاط
# ==========================================================
WALLET_ASSIGN_DATE_CANDIDATES = [
    "Assign Date", "assign date", "AssignDate", "Assignment Date", "assignment date",
    "تاريخ الاسناد", "تاريخ الإسناد", "Assign date",
]
WALLET_DEBIT_DATE_CANDIDATES = [
    "Debit Date", "debit date", "DebitDate", "Date of Debit", "Accident Date",
    "تاريخ الحادث", "تاريخ الدين",
]
WALLET_NATIONALITY_CANDIDATES = [
    "Nationality", "nationality", "Customer Nationality", "client nationality",
    "جنسية العميل", "الجنسية", "جنسية",
]
WALLET_CUSTOMER_STATE_CANDIDATES = [
    "Customer State", "customer state", "Client State", "client state",
    "حالة العميل", "Status", "status", "State", "state",
    "Sub State", "sub state", "SubState", "الحالة الفرعية",
]
WALLET_CUSTOMER_ID_CANDIDATES = [
    "Customer ID", "customer id", "CustomerId", "Client ID", "client id",
    "Debitor", "debitor", "رقم العميل", "رقم المدين", "Customer Number",
    "Customer Account number", "Customer Account Number", "رقم حساب العميل",
]
WALLET_AGING_BUCKETS = [
    (0, 30, "0 — 30 يوم"),
    (31, 60, "31 — 60 يوم"),
    (61, 90, "61 — 90 يوم"),
    (91, 180, "91 — 180 يوم"),
    (181, None, "أكثر من 180 يوم"),
]
WALLET_AGING_SOURCE_CANDIDATES = [
    "عمر الاسناد", "عمر الإسناد", "عمر_الاسناد", "عمر الاسناد ",
    "Assignment Age", "assignment age", "Aging", "aging", "Age Bucket",
    "Aging Bucket", "عمر الإسناد بالأيام", "عمر الاسناد بالايام",
]
WALLET_AGING_COL = "عمر_الاسناد"
WALLET_COLLECTED_COL = "تم_التحصيل"
WALLET_REMAINING_COL = "باقي_المديونية"

# ==========================================================
# إعدادات تويب الجدولة المتعثرة
# ==========================================================
SCHEDULE_RESULT_KEY = "schedule_stalled_result"
SCHEDULE_SUB_STATE_VALUE = "جدولة"
SCHEDULE_DEBITOR_CANDIDATES = [
    "Debitor", "debitor", "DEBITOR", "Debtor", "debtor",
    "رقم المدين", "المدين", "Debitor ID", "Debitor Id",
]
SCHEDULE_PAYMENT_DATE_CANDIDATES = [
    "Date of Creation", "date of creation", "Date Of Creation",
    "Creation Date", "creation date", "تاريخ الإنشاء", "تاريخ السداد",
    "Payment Date", "payment date",
]
SCHEDULE_REGULAR_MAX_DAYS = 60
SCHEDULE_STATUS_COL = "حالة_الجدولة"
SCHEDULE_LAST_PAYMENT_COL = "تاريخ_آخر_سداد"
SCHEDULE_DAYS_SINCE_COL = "أيام_منذ_آخر_سداد"

# ==========================================================
# إعدادات تويب أخطاء الحالات
# ==========================================================
CASE_ERRORS_RESULT_KEY = "case_errors_result"
CASE_ERROR_REASON_COL = "سبب_الخطأ"
CASE_ERROR_RULE_COL = "قاعدة_الخطأ"
CASE_PAYMENT_CANDIDATES = [
    "Payment", "payment", "PAYMENT", "Paid Amount", "paid amount",
    "Amount Paid", "المبلغ المدفوع", "مدفوع", "قيمة السداد", "Payment Amount",
]
CASE_PAYMENT_INDICATING_STATES = [
    "جدولة",
    "سدد كامل المديونية",
    "جدولة مقفلة",
    "جدولة مقفلة بخصم",
    "سدد كامل المديونية بخصم",
]
CASE_PAYMENT_MUST_BE_POSITIVE = ["جدولة", "سدد كامل المديونية"]
CASE_ORIGINAL_DEBT_CANDIDATES = [
    "أصل المديونية", "اصل المديونية", "Original Amount", "original amount",
    "Original Debt", "original debt", "Principal", "principal",
    "Principal Amount", "مبلغ أصل المديونية", "الأصل", "الاصل",
]
CASE_NET_AMOUNT_ERROR_MIN = 50.0


def uploaded_file_hash(uploaded_file):
    if uploaded_file is None:
        return None
    return hashlib.sha256(uploaded_file.getvalue()).hexdigest()


def _clear_cached_results(result_keys):
    for key in result_keys:
        if key.startswith("period_results:"):
            period_key = key.split(":", 1)[1]
            st.session_state.setdefault("period_results", {}).pop(period_key, None)
        elif key == "dashboard_source":
            st.session_state.pop("dashboard_source", None)
            st.session_state.pop(DASHBOARD_SOURCE_HASH_KEY, None)
        else:
            st.session_state.pop(key, None)


def sync_file_cache(widget_key, cache_scope, result_keys):
    """يحافظ على النتائج عبر rerun ويمسحها فقط عند إزالة الملف أو تغييره."""
    uploaded = st.session_state.get(widget_key)
    cache = st.session_state.setdefault(APP_DATA_CACHE_KEY, {})
    previous = cache.get(cache_scope)
    if uploaded is None:
        _clear_cached_results(result_keys)
        cache.pop(cache_scope, None)
        return
    current_hash = uploaded_file_hash(uploaded)
    if previous and previous.get("file_hash") != current_hash:
        _clear_cached_results(result_keys)
    cache[cache_scope] = {"file_hash": current_hash, "filename": uploaded.name}


def init_neglect_state():
    if "neglect_sub_states" not in st.session_state:
        st.session_state["neglect_sub_states"] = NEGLECT_SUB_STATES_DEFAULT.copy()
    if "neglect_available_states" not in st.session_state:
        st.session_state["neglect_available_states"] = []
    if "neglect_mode" not in st.session_state:
        st.session_state["neglect_mode"] = "neglect"


TODAY_KEY = "promises_today"


def _init_promises_today():
    """تحديد تاريخ اليوم مرة واحدة عند أول تشغيل."""
    if TODAY_KEY not in st.session_state:
        st.session_state[TODAY_KEY] = datetime.now().date()


def parse_date_cell(val):
    """بيحول خلية التاريخ لـ date مهما كان شكلها."""
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, pd.Timestamp):
        return val.date()
    if hasattr(val, "date"):
        return val.date()
    txt = str(val).strip()
    if not txt:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y", "%d %B %Y"):
        try:
            return datetime.strptime(txt, fmt).date()
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(txt, errors="coerce").date()
    except Exception:
        return None


PROMISES_RESULT_KEY = "promises_result"
BROKEN_RESULT_KEY = "promises_broken_result"


def find_column(df, candidates):
    """البحث عن أول عمود يطابق أحد الأسماء المرشحة"""
    for col in candidates:
        if col in df.columns:
            return col
    return None


def read_uploaded_dataframe(uploaded):
    """قراءة الملف المرفوع سواء Excel أو CSV"""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded)
        except UnicodeDecodeError:
            uploaded.seek(0)
            return pd.read_csv(uploaded, encoding="utf-8-sig")
    else:
        return pd.read_excel(uploaded)


def _run_promises_pipeline(
    uploaded,
    result_key,
    due_mode,
    count_label,
):
    """معالجة ملف المحفظة وحفظ النتيجة في الكاش."""
    init_promises_sales_filter()
    _file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    _excl_token = _promises_exclusion_token()
    cached = st.session_state.get(result_key)
    if (
        cached
        and cached.get("file_hash") == _file_hash
        and cached.get("exclusion_token") == _excl_token
    ):
        avail = cached.get("available_sales") or []
        if avail:
            merged = sorted(set(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, [])) | set(avail))
            st.session_state[PROMISES_AVAILABLE_SALES_KEY] = merged
        return False

    try:
        raw_df = read_uploaded_dataframe(uploaded)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return False

    total_in_file = len(raw_df) - 1
    df = raw_df

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    substate_col = find_column(df, PROMISE_SUB_STATE_CANDIDATES)
    duedate_col = find_column(df, PROMISE_DUE_DATE_CANDIDATES)
    net_col = find_column(df, PROMISE_NET_AMOUNT_CANDIDATES)

    missing = [n for n, c in [
        ("المحصّل (Salesperson)", sales_col),
        ("الحالة الفرعية (Sub State)", substate_col),
        ("تاريخ المتابعة (Follow up Due Date)", duedate_col),
    ] if not c]
    if missing:
        st.error(
            "تعذر العثور على أعمدة مهمة في الملف. الأعمدة المطلوبة: "
            f"{', '.join(missing)}\n\nالأعمدة الموجودة في الملف: {', '.join(df.columns.astype(str))}"
        )
        return False

    sales_vals = df[sales_col].astype(str).str.strip()
    available_sales = sorted({v for v in sales_vals.tolist() if v and v.lower() not in {"nan", "none", "null"}})
    merged_available = sorted(set(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, [])) | set(available_sales))
    st.session_state[PROMISES_AVAILABLE_SALES_KEY] = merged_available

    excluded = set(_promises_excluded_sales())
    keep_sales = ~sales_vals.isin(excluded)
    dropped_sales = int((~keep_sales).sum())
    df = df[keep_sales].copy()

    if substate_col:
        sub_vals = df[substate_col].astype(str).str.strip()
        keep_sub = sub_vals == PROMISE_SUB_STATE_VALUE
        dropped_sub = int((~keep_sub).sum())
        df = df[keep_sub].copy()
    else:
        dropped_sub = 0

    target_date = st.session_state[TODAY_KEY]
    due_vals = pd.Series([parse_date_cell(v) for v in df[duedate_col]], index=df.index)
    if due_mode == "before":
        keep_due = due_vals.apply(lambda d: d is not None and d < target_date)
        due_desc = f"التاريخ قبل اليوم (< {target_date.strftime('%Y-%m-%d')})"
    else:
        keep_due = due_vals == target_date
        due_desc = f"التاريخ يساوي اليوم ({target_date.strftime('%Y-%m-%d')})"
    dropped_due = int((~keep_due).sum())
    df = df[keep_due].copy()

    if net_col and net_col in df.columns:
        summary_df = df.groupby(sales_col).agg(
            **{
                count_label: (duedate_col, "count"),
                "صافي المديونية (Net Amount)": (net_col, "sum"),
            }
        )
    else:
        summary_df = df.groupby(sales_col).agg(
            **{count_label: (duedate_col, "count")}
        )
    summary_df = summary_df.sort_values(count_label, ascending=False).reset_index()
    summary_df.columns = ["المحصّل " + str(sales_col), count_label] + (
        ["صافي المديونية (Net Amount)"] if net_col and net_col in df.columns else []
    )

    st.session_state[result_key] = {
        "df": df,
        "summary_df": summary_df,
        "target_date": target_date,
        "filename": uploaded.name,
        "file_hash": _file_hash,
        "exclusion_token": _excl_token,
        "available_sales": available_sales,
        "sales_col": sales_col,
        "substate_col": substate_col,
        "duedate_col": duedate_col,
        "net_col": net_col,
        "total_in_file": total_in_file,
        "dropped_sales": dropped_sales,
        "dropped_sub": dropped_sub,
        "dropped_due": dropped_due,
        "due_mode": due_mode,
        "due_desc": due_desc,
    }
    return True


PROMISES_COMPANY_KEY = "promises_selected_company"
PROMISES_COMPANY_OPTIONS = ["الوطنية للتأمين", "تري للتأمين"]
PROMISE_COMPANY_CANDIDATES = ["Company", "company", "Company Name", "اسم الشركة", "الشركة"]
PROMISES_AGENT_FILTER_KEY = "promises_selected_agent_filter"

OPS_SCALE = ["#1A8A4A", "#126D3C", "#0E5A32"]
PLOTLY_TEMPLATE = "plotly_dark" if THEME_NAME == "dark" else "plotly_white"
PLOTLY_CONFIG = {"responsive": True, "displayModeBar": False}


def page_header(badge, title, subtitle):
    st.markdown(f"### {title}")
    st.caption(subtitle)
    st.divider()


def render_promises_filter_notice():
    selected_agent = st.session_state.get(PROMISES_AGENT_FILTER_KEY)
    if selected_agent:
        c1, c2 = st.columns([4, 1])
        with c1:
            st.info(f"🔍 يتم التصفية حالياً حسب المحصّل: **{selected_agent}**")
        with c2:
            if st.button("❌ إلغاء التصفية", key="clear_promises_agent_filter_btn", use_container_width=True):
                st.session_state.pop(PROMISES_AGENT_FILTER_KEY, None)
                st.rerun()


def get_promises_view(df, sales_col):
    selected_agent = st.session_state.get(PROMISES_AGENT_FILTER_KEY)
    if selected_agent and sales_col and sales_col in df.columns:
        return df[df[sales_col] == selected_agent].copy()
    return df


def render_selectable_chart(fig, key, filter_key):
    event = st.plotly_chart(
        fig,
        use_container_width=True,
        config=PLOTLY_CONFIG,
        key=key,
        on_select="rerun",
        selection_mode="points",
    )
    if event and getattr(event, "selection", None):
        points = event.selection.get("points", [])
        if points:
            customdata = points[0].get("customdata")
            if customdata:
                selected_val = customdata[0] if isinstance(customdata, (list, tuple)) else customdata
                if st.session_state.get(filter_key) != selected_val:
                    st.session_state[filter_key] = selected_val
                    st.rerun()


def _apply_ops_chart_style(fig, title_text, height=400, xaxis_title="", show_legend=True, margin=None, extra=None):
    layout_args = dict(
        title=dict(text=title_text, font=dict(size=18, color=THEME["text"])),
        height=height,
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Tajawal, sans-serif", color=THEME["text"]),
        showlegend=show_legend,
        margin=margin or dict(t=50, b=50, l=50, r=50),
        xaxis=dict(title=xaxis_title, gridcolor=THEME["border_soft"], color=THEME["text_dim"]),
        yaxis=dict(gridcolor=THEME["border_soft"], color=THEME["text_dim"]),
    )
    if extra:
        for k, v in extra.items():
            if isinstance(v, dict) and k in layout_args and isinstance(layout_args[k], dict):
                layout_args[k].update(v)
            else:
                layout_args[k] = v
    fig.update_layout(**layout_args)


def _rounded_rect_path(x0, x1, y0, y1, radius=0.02):
    return (
        f"M {x0+radius},{y0} "
        f"L {x1-radius},{y0} Q {x1},{y0} {x1},{y0+radius} "
        f"L {x1},{y1-radius} Q {x1},{y1} {x1-radius},{y1} "
        f"L {x0+radius},{y1} Q {x0},{y1} {x0},{y1-radius} "
        f"L {x0},{y0+radius} Q {x0},{y0} {x0+radius},{y0} Z"
    )


def _build_promises_agent_summary(df, sales_col, net_col):
    if not sales_col or sales_col not in df.columns:
        return pd.DataFrame()
    summary = df.groupby(sales_col).size().reset_index(name="عدد الوعود")
    if net_col and net_col in df.columns:
        amounts = pd.to_numeric(df[net_col], errors="coerce").fillna(0)
        amount_summary = df.assign(_promise_amount=amounts).groupby(sales_col)["_promise_amount"].sum().reset_index(name="إجمالي المديونية")
        summary = summary.merge(amount_summary, on=sales_col, how="left")
    return summary.sort_values("عدد الوعود", ascending=False).reset_index(drop=True)


def render_promises_dashboard(df, summary, mode_label):
    sales_col = summary.get("sales_col")
    net_col = summary.get("net_col")
    render_promises_filter_notice()
    df = get_promises_view(df, sales_col)
    total = len(df)
    total_amount = pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum() if net_col and net_col in df.columns else 0
    agent_summary = _build_promises_agent_summary(df, sales_col, net_col)
    agent_count = len(agent_summary)
    avg_amount = total_amount / total if total else 0

    st.subheader(f"📊 ملخص الوعود — {mode_label}")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🤝 إجمالي الوعود", f"{total:,}")
    k2.metric("👥 عدد المحصّلين", f"{agent_count:,}")
    k3.metric("💰 إجمالي المديونية", f"{total_amount:,.0f}" if net_col else "—")
    k4.metric("📈 متوسط المديونية", f"{avg_amount:,.0f}" if net_col else "—")

    if total and sales_col and not agent_summary.empty:
        st.markdown("#### 📈 تحليلات الوعود التفاعلية")
        left, right = st.columns(2)
        with left:
            count_df = agent_summary.head(15).sort_values("عدد الوعود")
            count_fig = px.bar(
                count_df,
                x="عدد الوعود",
                y=sales_col,
                orientation="h",
                text="عدد الوعود",
                color="عدد الوعود",
                color_continuous_scale=OPS_SCALE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                count_fig, "عدد الوعود حسب المحصّل",
                height=430, xaxis_title="عدد الوعود", show_legend=False,
                margin=dict(t=70, b=55, l=160, r=55),
            )
            count_fig.update_traces(
                texttemplate="%{x:,.0f}",
                textposition="outside",
                cliponaxis=False,
                customdata=count_df[sales_col],
                hovertemplate="<b>%{y}</b><br>عدد الوعود: %{x:,}<extra></extra>",
            )
            with st.container(border=True):
                render_selectable_chart(count_fig, f"promises_count_{mode_label}", filter_key=PROMISES_AGENT_FILTER_KEY)
        with right:
            if net_col and "إجمالي المديونية" in agent_summary.columns:
                amount_df = agent_summary.head(15).sort_values("إجمالي المديونية")
                amount_fig = px.bar(
                    amount_df,
                    x="إجمالي المديونية",
                    y=sales_col,
                    orientation="h",
                    text="إجمالي المديونية",
                    color="إجمالي المديونية",
                    color_continuous_scale=OPS_SCALE,
                    template=PLOTLY_TEMPLATE,
                )
                _apply_ops_chart_style(
                    amount_fig, "إجمالي المديونية حسب المحصّل",
                    height=430, xaxis_title="إجمالي المديونية", show_legend=False,
                    margin=dict(t=70, b=55, l=160, r=70),
                )
                amount_fig.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                    cliponaxis=False,
                    customdata=amount_df[sales_col],
                    hovertemplate="<b>%{y}</b><br>إجمالي المديونية: %{x:,.0f}<extra></extra>",
                )
                with st.container(border=True):
                    render_selectable_chart(amount_fig, f"promises_amount_{mode_label}", filter_key=PROMISES_AGENT_FILTER_KEY)
            else:
                st.info("لا يوجد عمود صافي المديونية لعرض الرسم المالي.")

    display_cols = [c for c in [summary.get("sales_col"), summary.get("substate_col"), summary.get("duedate_col"), summary.get("net_col")] if c and c in df.columns]
    st.subheader(f"📋 بيانات الوعود — {mode_label}")
    if display_cols:
        st.dataframe(df[display_cols], use_container_width=True, hide_index=True)
    if not df.empty:
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="الوعود")
        st.download_button(
            f"⬇️ تحميل بيانات الوعود — {mode_label}",
            data=out_excel.getvalue(),
            file_name=f"الوعود_{mode_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"promises_unified_download_{mode_label}",
            type="primary",
        )


def _filter_promises_by_company(df, company_label):
    """يفلتر عمود الشركة إن وُجد."""
    company_col = find_column(df, PROMISE_COMPANY_CANDIDATES)
    if not company_col or company_col not in df.columns:
        return df.copy()
    short_name = "الوطنية" if "الوطنية" in company_label else "تري"
    values = df[company_col].astype(str).str.strip()
    mask = values.str.contains(company_label, case=False, na=False) | values.str.contains(short_name, case=False, na=False)
    return df.loc[mask].copy() if mask.any() else df.iloc[0:0].copy()


def _combine_promises_cached_results(company_label, result_keys=None):
    parts = []
    result_keys = result_keys or (PROMISES_RESULT_KEY, BROKEN_RESULT_KEY)
    for key, promise_type in zip(result_keys, ("الوعود القائمة", "الوعود المكسورة")):
        cached = st.session_state.get(key)
        if not cached or cached.get("df") is None:
            continue
        part = _filter_promises_by_company(cached["df"], company_label)
        if part.empty:
            continue
        part = part.copy()
        part["نوع الوعد"] = promise_type
        parts.append(part)
    if not parts:
        return pd.DataFrame(), None
    combined = pd.concat(parts, ignore_index=True, sort=False)
    meta = st.session_state.get(result_keys[0]) or st.session_state.get(result_keys[1])
    return combined, meta


def render_promises_kpi_dashboard(total, standing_count, broken_count, agent_count, total_amount):
    OPS_POSITIVE = "#1A8A4A"
    OPS_NEGATIVE = "#C45C5C"
    OPS_MID = "#F19F29"
    cards = [
        ("🤝<br>إجمالي الوعود", total, {"valueformat": ",d"}, THEME["text"]),
        ("📗<br>الوعود القائمة", standing_count, {"valueformat": ",d"}, OPS_POSITIVE),
        ("📕<br>الوعود المكسورة", broken_count, {"valueformat": ",d"}, OPS_NEGATIVE),
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("💰<br>إجمالي المديونية", total_amount, {"valueformat": ",.0f"}, OPS_MID),
    ]
    figure = go.Figure()
    count = len(cards)
    gap = 0.018
    width = (1 - gap * (count + 1)) / count
    for index, (label, value, number_format, number_color) in enumerate(cards):
        x0 = gap + index * (width + gap)
        x1 = x0 + width
        figure.add_shape(
            type="path",
            xref="paper",
            yref="paper",
            path=_rounded_rect_path(x0, x1, 0.06, 0.94, radius=0.022),
            line={"color": THEME["border"], "width": 1},
            fillcolor=THEME["surface"],
            layer="below",
        )
        figure.add_trace(
            go.Indicator(
                mode="number",
                value=float(value or 0),
                domain={"x": [x0 + 0.012, x1 - 0.012], "y": [0.12, 0.88]},
                title={"text": label, "font": {"size": 18, "color": THEME["text_dim"]}, "align": "center"},
                number={"font": {"size": 32, "color": number_color}, **number_format},
            )
        )
    figure.update_layout(
        height=200,
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Tajawal, sans-serif", "color": THEME["text"]},
        margin={"t": 8, "b": 8, "l": 8, "r": 8},
    )
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG, key="promises_kpi_dashboard")


def render_combined_promises_dashboard(df, meta, company_label):
    OPS_POSITIVE = "#1A8A4A"
    OPS_NEGATIVE = "#C45C5C"
    sales_col = meta.get("sales_col") if meta else None
    net_col = meta.get("net_col") if meta else None
    render_promises_filter_notice()
    df = get_promises_view(df, sales_col)
    total = len(df)
    standing_count = int((df["نوع الوعد"] == "الوعود القائمة").sum()) if "نوع الوعد" in df.columns else 0
    broken_count = int((df["نوع الوعد"] == "الوعود المكسورة").sum()) if "نوع الوعد" in df.columns else 0
    total_amount = pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum() if net_col and net_col in df.columns else 0
    agent_count = int(df[sales_col].nunique()) if sales_col and sales_col in df.columns else 0

    st.subheader(f"📊 ملخص الوعود — {company_label}")
    render_promises_kpi_dashboard(total, standing_count, broken_count, agent_count, total_amount if net_col else 0)

    if total and sales_col and sales_col in df.columns:
        st.markdown("#### 📈 تحليلات الوعود القائمة والمكسورة")
        agent_type = df.groupby([sales_col, "نوع الوعد"]).size().reset_index(name="عدد الوعود")
        agent_order = agent_type.groupby(sales_col)["عدد الوعود"].sum().sort_values(ascending=False).head(15).index.tolist()
        agent_type = agent_type[agent_type[sales_col].isin(agent_order)]
        agent_type["_sort"] = agent_type[sales_col].map({name: i for i, name in enumerate(agent_order)})
        agent_type = agent_type.sort_values(["_sort", "نوع الوعد"], ascending=[True, True])
        left, right = st.columns(2)
        with left:
            fig = px.bar(
                agent_type,
                x="عدد الوعود",
                y=sales_col,
                color="نوع الوعد",
                barmode="group",
                orientation="h",
                text="عدد الوعود",
                color_discrete_map={"الوعود القائمة": OPS_POSITIVE, "الوعود المكسورة": OPS_NEGATIVE},
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig, "القائمة والمكسورة حسب المحصّل",
                height=430, xaxis_title="عدد الوعود",
                margin=dict(t=70, b=70, l=170, r=55),
                extra={"xaxis": dict(tickformat=",.0f", automargin=True)},
            )
            fig.update_traces(
                texttemplate="%{x:,.0f}",
                textposition="outside",
                textfont=dict(size=13, color=THEME["text"]),
                cliponaxis=False,
                customdata=agent_type[sales_col],
                hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f}<extra></extra>",
            )
            with st.container(border=True):
                render_selectable_chart(fig, "promises_combined_by_agent", filter_key=PROMISES_AGENT_FILTER_KEY)
        with right:
            if net_col and net_col in df.columns:
                amount_work = df.assign(_amount=pd.to_numeric(df[net_col], errors="coerce").fillna(0))
                agent_amount = amount_work.groupby([sales_col, "نوع الوعد"])["_amount"].sum().reset_index(name="إجمالي المديونية")
                agent_amount = agent_amount[agent_amount[sales_col].isin(agent_order)]
                fig = px.bar(
                    agent_amount,
                    x="إجمالي المديونية",
                    y=sales_col,
                    color="نوع الوعد",
                    barmode="group",
                    orientation="h",
                    text="إجمالي المديونية",
                    color_discrete_map={"الوعود القائمة": OPS_POSITIVE, "الوعود المكسورة": OPS_NEGATIVE},
                    template=PLOTLY_TEMPLATE,
                )
                _apply_ops_chart_style(
                    fig, "إجمالي المديونية حسب المحصّل",
                    height=430, xaxis_title="إجمالي المديونية",
                    margin=dict(t=70, b=70, l=170, r=70),
                    extra={"xaxis": dict(tickformat=",.0f", separatethousands=True, automargin=True)},
                )
                fig.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                    textfont=dict(size=13, color=THEME["text"]),
                    cliponaxis=False,
                    customdata=agent_amount[sales_col],
                    hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f} جنيه<extra></extra>",
                )
                with st.container(border=True):
                    render_selectable_chart(fig, "promises_combined_amount_by_agent", filter_key=PROMISES_AGENT_FILTER_KEY)
            else:
                st.info("لا يوجد عمود صافي المديونية لعرض الرسم المالي.")

        type_counts = df["نوع الوعد"].value_counts().rename_axis("نوع الوعد").reset_index(name="عدد الوعود")
        fig = px.pie(
            type_counts,
            values="عدد الوعود",
            names="نوع الوعد",
            hole=0.55,
            color="نوع الوعد",
            color_discrete_map={"الوعود القائمة": OPS_POSITIVE, "الوعود المكسورة": OPS_NEGATIVE},
            template=PLOTLY_TEMPLATE,
        )
        _apply_ops_chart_style(
            fig, "توزيع الوعود القائمة والمكسورة",
            height=430,
            margin=dict(t=70, b=55, l=40, r=40),
        )
        fig.update_traces(
            texttemplate="%{value:,.0f}<br>%{percent:.1%}",
            textfont=dict(size=15, color=THEME["text"]),
            textinfo="text",
            hovertemplate="<b>%{label}</b><br>عدد الوعود: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
        )
        with st.container(border=True):
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="promises_combined_type_share")

    if sales_col and sales_col in df.columns and "نوع الوعد" in df.columns:
        agent_promise_table = (
            df.groupby([sales_col, "نوع الوعد"]).size()
            .unstack(fill_value=0)
            .reset_index()
        )
        if "الوعود القائمة" not in agent_promise_table.columns:
            agent_promise_table["الوعود القائمة"] = 0
        if "الوعود المكسورة" not in agent_promise_table.columns:
            agent_promise_table["الوعود المكسورة"] = 0
        agent_promise_table["الإجمالي"] = agent_promise_table["الوعود القائمة"] + agent_promise_table["الوعود المكسورة"]
        agent_promise_table = agent_promise_table.rename(columns={sales_col: "المحصّل"})
        agent_promise_table = agent_promise_table[["المحصّل", "الوعود القائمة", "الوعود المكسورة", "الإجمالي"]].sort_values("الإجمالي", ascending=False)
        st.subheader("📊 ملخص الوعود حسب المحصّل")
        st.dataframe(agent_promise_table, use_container_width=True, hide_index=True)

    display_cols = [c for c in [sales_col, "نوع الوعد", meta.get("substate_col") if meta else None, meta.get("duedate_col") if meta else None, net_col] if c and c in df.columns]
    st.subheader("📋 تفاصيل الوعود القائمة والمكسورة")
    if display_cols:
        st.dataframe(df[display_cols], use_container_width=True, hide_index=True)

    report_date = datetime.now().strftime("%Y-%m-%d")
    company_file_name = company_label.replace(" ", "_")
    standing_df = df[df["نوع الوعد"] == "الوعود القائمة"].copy() if "نوع الوعد" in df.columns else pd.DataFrame()
    broken_df = df[df["نوع الوعد"] == "الوعود المكسورة"].copy() if "نوع الوعد" in df.columns else pd.DataFrame()

    def _excel_bytes(report_df, sheet_name):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            report_df.to_excel(writer, index=False, sheet_name=sheet_name)
        return buffer.getvalue()

    download_left, download_right = st.columns(2)
    with download_left:
        st.download_button(
            "⬇️ تحميل تقرير الوعود القائمة",
            data=_excel_bytes(standing_df, "الوعود القائمة"),
            file_name=f"تقرير_الوعود_القائمة_{company_file_name}_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="promises_standing_download",
            type="primary",
            disabled=standing_df.empty,
        )
    with download_right:
        st.download_button(
            "⬇️ تحميل تقرير الوعود المكسورة",
            data=_excel_bytes(broken_df, "الوعود المكسورة"),
            file_name=f"تقرير_الوعود_المكسورة_{company_file_name}_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="promises_broken_download",
            type="primary",
            disabled=broken_df.empty,
        )


def page_promises():
    """صفحة الوعود القائمة والمكسورة."""
    _init_promises_today()
    init_promises_sales_filter()
    page_header(
        "PROMISES",
        "📚 الوعود",
        "لكل شركة محفظة مستقلة؛ اختر الشركة وارفع ملفها لعرض الوعود القائمة والمكسورة معًا",
    )
    company_label = st.radio("اختر الشركة", PROMISES_COMPANY_OPTIONS, horizontal=True, key=PROMISES_COMPANY_KEY)
    company_slug = "wataniya" if company_label == "الوطنية للتأمين" else "tary"
    standing_key = f"promises_{company_slug}_standing"
    broken_key = f"promises_{company_slug}_broken"
    upload_key = f"promises_upload_{company_slug}"
    cache_scope = f"promises_upload:{company_slug}"
    result_keys = (standing_key, broken_key)

    def _clear_promises_company_cache():
        for key in result_keys:
            st.session_state.pop(key, None)

    uploaded = st.file_uploader(
        f"📂 ارفع محفظة {company_label} (Excel أو CSV)",
        type=["xlsx", "xls", "csv"],
        key=upload_key,
        on_change=sync_file_cache,
        args=(upload_key, cache_scope, result_keys),
    )
    if uploaded is not None:
        st.caption(f"المحفظة المختارة: {company_label} · الملف: {uploaded.name}")
        _run_promises_pipeline(uploaded, standing_key, due_mode="today", count_label="عدد الوعود القائمة")
        _run_promises_pipeline(uploaded, broken_key, due_mode="before", count_label="عدد الوعود المكسورة")
        _sync_promises_available_sales(uploaded=uploaded, result_keys=result_keys)
    else:
        cached_file = st.session_state.get(APP_DATA_CACHE_KEY, {}).get(cache_scope, {})
        cached_result = st.session_state.get(standing_key) or st.session_state.get(broken_key)
        if cached_result or cached_file:
            saved_name = (cached_result or cached_file).get("filename", "المحفظة المحفوظة")
            st.success(f"✅ محفظة {company_label} محفوظة: {saved_name}. لن تُحذف عند التنقل بين التبويبات.")
            _sync_promises_available_sales(uploaded=None, result_keys=result_keys)

    available = list(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, []))
    excluded = list(st.session_state.get(PROMISES_EXCLUDED_SALES_KEY, []))
    included = [s for s in available if s not in excluded]

    with st.expander("⚙️ إدارة المحصلين (Sales Person) — إضافة / استبعاد", expanded=bool(available)):
        st.caption("المحصل اللي هتشيله من القائمة هتتشال بياناته من تحليل الوعود، واللي هتضيفه هترجع بياناته.")
        if not available:
            st.info("💡 ارفع ملف المحفظة أولاً عشان أقدر أستخرج قائمة المحصلين من عمود Sales Person.")
        else:
            to_add_options = [s for s in available if s in excluded]
            selected_to_add = st.multiselect(
                "محصلين مستبعدين — اختر لإضافتهم للتحليل:",
                options=to_add_options,
                key=f"promises_add_sales_{company_slug}",
            )
            c_add, c_reset = st.columns(2)
            with c_add:
                if st.button("➕ إضافة المحصلين المختارين", use_container_width=True, key=f"promises_btn_add_{company_slug}"):
                    if selected_to_add:
                        for name in selected_to_add:
                            if name in st.session_state[PROMISES_EXCLUDED_SALES_KEY]:
                                st.session_state[PROMISES_EXCLUDED_SALES_KEY].remove(name)
                        _clear_promises_company_cache()
                        st.success(f"تمت إضافة {len(selected_to_add)} محصل للتحليل.")
                        st.rerun()
            with c_reset:
                if st.button("↩️ استعادة الاستبعاد الافتراضي", use_container_width=True, key=f"promises_btn_reset_{company_slug}"):
                    st.session_state[PROMISES_EXCLUDED_SALES_KEY] = list(PROMISE_EXCLUDED_SALES)
                    _clear_promises_company_cache()
                    st.rerun()

            st.divider()
            st.write(f"📋 المحصلين المشمولين حالياً في التحليل ({len(included)}):")
            if not included:
                st.warning("لا يوجد محصلين مشمولين حالياً — أضف من قائمة المستبعدين أو راجع الاستبعاد.")
            else:
                cols = st.columns(3)
                for i, name in enumerate(included):
                    with cols[i % 3]:
                        if st.button(f"❌ {name}", key=f"promises_del_{company_slug}_{i}", use_container_width=True):
                            if name not in st.session_state[PROMISES_EXCLUDED_SALES_KEY]:
                                st.session_state[PROMISES_EXCLUDED_SALES_KEY].append(name)
                            _clear_promises_company_cache()
                            st.rerun()

            if excluded:
                with st.expander(f"محصلين مستبعدين ({len(excluded)})", expanded=False):
                    st.write(" · ".join(excluded))

    combined, meta = _combine_promises_cached_results(company_label, result_keys=result_keys)
    if combined.empty or meta is None:
        st.info(f"📂 ارفع محفظة {company_label} لعرض الوعود القائمة والمكسورة معًا.")
        return
    render_combined_promises_dashboard(combined, meta, company_label)


# ==========================================================
# التشغيل الرئيسي للتطبيق
# ==========================================================
if __name__ == "__main__":
    page_promises()
