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
# أسماء الأعمدة المتوقعة في ملف المحفظة — لو اتغيرت غيّرها من هنا.
PROMISE_SUB_STATE_CANDIDATES = ["Sub State", "sub state", "SubState", "الحالة الفرعية"]
PROMISE_DUE_DATE_CANDIDATES = ["Follow up Due Date", "follow up due date", "FollowUpDueDate", "تاريخ المتابعة", "Due Date"]
PROMISE_NET_AMOUNT_CANDIDATES = ["Net Amount", "net amount", "NetAmount", "صافي المبلغ", "مبلغ المديونية"]

# قيمة Sub State اللي بتمثل وعد قائم
PROMISE_SUB_STATE_VALUE = "واعد بالسداد"

# المحصّلين اللي بنستبعدهم من الوعود القائمة
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
    """تهيئة قائمة المحصلين المستبعدين (قابلة للتعديل من واجهة الوعود)."""
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

# صفحات إضافية اختيارية في تحليل نشاط المحصلين: المحفظة + السداد + الربط بينهما
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
# عمود عمر الإسناد الجاهز في ملف المحفظة (مش بنحسبه من Assign Date)
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
# حد الشهرين بالأيام: ≤ 60 يوم = منتظمة، أكثر من 60 يوم = متعثرة
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
# حالات تدل على سداد / جدولة (بتستبعد من فلاتر "غير السداد")
CASE_PAYMENT_INDICATING_STATES = [
    "جدولة",
    "سدد كامل المديونية",
    "جدولة مقفلة",
    "جدولة مقفلة بخصم",
    "سدد كامل المديونية بخصم",
]
# حالات قاعدة Payment = 0 (لازم يكون في Payment)
CASE_PAYMENT_MUST_BE_POSITIVE = ["جدولة", "سدد كامل المديونية"]
# عمود أصل المديونية — لو = 0 مع Net Amount = 0 مش بنعتبرها خطأ
CASE_ORIGINAL_DEBT_CANDIDATES = [
    "أصل المديونية", "اصل المديونية", "Original Amount", "original amount",
    "Original Debt", "original debt", "Principal", "principal",
    "Principal Amount", "مبلغ أصل المديونية", "الأصل", "الاصل",
]
# حد المتبقي المقبول عند الإقفال (سدد كامل / جدولة مقفلة)
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
            st.session_state.pop(DASHBOARD_SOURCE_KEY, None)
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
        st.session_state["neglect_mode"] = "neglect"  # 'neglect' or 'followup'

# التاريخ المستهدف: تاريخ اليوم (بيتم تحديده مرة واحدة عند أول عرض للصفحة)
TODAY_KEY = "promises_today"


def _init_promises_today():
    """بنحدد تاريخ اليوم مرة واحدة في أول رن للصفحة لـ ميترفرش مع كل إعادة تشغيل."""
    if TODAY_KEY not in st.session_state:
        st.session_state[TODAY_KEY] = datetime.now().date()


def parse_date_cell(val):
    """بيحول خلية التاريخ لـ date مهما كان شكلها (datetime / date / نص)."""
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


PROMISES_RESULT_KEY = "promises_result"  # كاش نتائج الوعود القائمة (فلتر: اليوم)
BROKEN_RESULT_KEY = "promises_broken_result"  # كاش نتائج الوعود المكسورة (فلتر: قبل اليوم)


def _run_promises_pipeline(
    uploaded,
    result_key,
    due_mode,
    count_label,
):
    """معالجة ملف المحفظة (فلترة + تجميع) وحفظ النتيجة في كاش الصفحة.

    - ``due_mode='today'``: Follow up Due Date == تاريخ اليوم (وعود قائمة).
    - ``due_mode='before'``: Follow up Due Date < تاريخ اليوم (وعود مكسورة).
    - ``count_label``: نص ملصق التجميع (مثلًا \"الوعود القائمة\" أو \"الوعود المكسورة\").

    بترجع ``True`` لو تم الحفظ في الكاش، و``False`` لو اتعرضت من الكاش مباشرة.
    """
    init_promises_sales_filter()
    _file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    _excl_token = _promises_exclusion_token()
    cached = st.session_state.get(result_key)
    if (
        cached
        and cached.get("file_hash") == _file_hash
        and cached.get("exclusion_token") == _excl_token
    ):
        # تحديث قائمة المحصلين المتاحة من الكاش لو موجودة
        avail = cached.get("available_sales") or []
        if avail:
            merged = sorted(set(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, [])) | set(avail))
            st.session_state[PROMISES_AVAILABLE_SALES_KEY] = merged
        return False  # النتيجة موجودة في الكاش من رفع الملف ده — مش يلزم وجود إعادة معالجة

    try:
        raw_df = read_uploaded_dataframe(uploaded)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return False

    total_in_file = len(raw_df) - 1  # قبل حذف أول صف
    df = raw_df

    # 1) حذف أول صف بعد العناوين (زي قاعدة باقي الملفات في التطبيق)
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

    # 2) فلترة Salesperson — حسب قائمة الاستبعاد القابلة للتعديل من الواجهة
    sales_vals = df[sales_col].astype(str).str.strip()
    available_sales = sorted({v for v in sales_vals.tolist() if v and v.lower() not in {"nan", "none", "null"}})
    merged_available = sorted(set(st.session_state.get(PROMISES_AVAILABLE_SALES_KEY, [])) | set(available_sales))
    st.session_state[PROMISES_AVAILABLE_SALES_KEY] = merged_available

    excluded = set(_promises_excluded_sales())
    keep_sales = ~sales_vals.isin(excluded)
    dropped_sales = int((~keep_sales).sum())
    df = df[keep_sales].copy()

    # 3) فلترة Sub State = واعد بالسداد
    if substate_col:
        sub_vals = df[substate_col].astype(str).str.strip()
        keep_sub = sub_vals == PROMISE_SUB_STATE_VALUE
        dropped_sub = int((~keep_sub).sum())
        df = df[keep_sub].copy()
    else:
        dropped_sub = 0

    # 4) فلترة Follow up Due Date حسب وضع التبويبة
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

    # 5) الجدول التجميعي لكل محصّل
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

    # 💾 حفظ النتائج في الكاش — تفضل موجودة لحد ما نعمل reload أو نشيل الملف
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
    """يفلتر عمود الشركة إن وُجد؛ وإذا لم يوجد فكل الملف يُعامل كملف الشركة المختارة."""
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
    """صفحة واحدة تجمع الوعود القائمة والمكسورة داخل محفظة الشركة المختارة فقط."""
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

    # إدارة المحصلين بعد معالجة الملف عشان القائمة تبقى جاهزة
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


st.set_page_config(
    page_title="الوطنية — لوحة تحليل المكالمات",
    page_icon="🟢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================
# الهوية البصرية (Theme)
# ==========================================================
# ملحوظة مهمة: الـ direction: rtl متطبق بس على محتوى النص
# (الـ block-container والـ sidebar content) مش على هيكل الصفحة كله،
# لـ الـ Sidebar يفضل ثابت فعليًا على الشمال زي ما اتطلب،
# بدل ما ينقلب يمين بسبب انعكاس اتجاه الـ flex layout.


# ==========================================================
# إعدادات المظهر (Light / Dark Mode)
# ==========================================================
THEMES = {
    # هوية الوطنية للتأمين (Wataniya) — ألوان مستخرجة من واجهة Export History
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
    """
    بيقرأ الوضع (Light/Dark) اللي المستخدم مختاره فعليًا من قائمة
    إعدادات Streamlit نفسها ("⋮" ← Settings ← Choose app theme)
    عن طريق st.context.theme.type (متاحة من Streamlit 1.46+).
    لو مش متاحة لأي سبب (نسخة قديمة من Streamlit)، بيرجع لآخر قيمة
    محفوظة في session_state، ولو لا يوجد هيستخدم "dark" كافتراضي.
    """
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
    """حقن هوية الوطنية: نفس سكايلاين الصورة بالظبط."""
    t = THEME
    is_light = THEME_NAME == "light"

    _SKYLINE_B64 = "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"
    skyline = f'url("data:image/png;base64,{_SKYLINE_B64}")'

    if is_light:
        bg_image = skyline
        bg_pos = "center bottom"
        bg_size = "72% auto"
        bg_repeat = "no-repeat"
    else:
        bg_image = "none"
        bg_pos = "center"
        bg_size = "auto"
        bg_repeat = "no-repeat"

    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');

html, body, [class*="css"]  {{
  font-family: 'Tajawal', sans-serif !important;
}}

/* خلفية الصفحة — نفس زخرفة الوطنية بالظبط */
.stApp {{
  background-color: {t["bg"]} !important;
  background-image: {bg_image} !important;
  background-position: {bg_pos} !important;
  background-size: {bg_size} !important;
  background-repeat: {bg_repeat} !important;
  background-attachment: fixed !important;
}}

/* الحاوية الرئيسية — نازلة في النص + استريتش */
.block-container {{
  background: {t["surface"]} !important;
  border: 1px solid {t["border"]} !important;
  border-radius: 18px !important;
  box-shadow: 0 8px 32px rgba(18, 109, 60, 0.08) !important;
  padding-top: 1.75rem !important;
  padding-bottom: 2.25rem !important;
  padding-left: 2.25rem !important;
  padding-right: 2.25rem !important;
  max-width: min(1320px, 88vw) !important;
  margin-left: auto !important;
  margin-right: auto !important;
  margin-top: 3rem !important;
  margin-bottom: 3rem !important;
}}

/* الشريط الجانبي */
section[data-testid="stSidebar"] {{
  background: {t["sidebar_bg"]} !important;
  border-left: 1px solid {t["border_soft"]} !important;
}}
section[data-testid="stSidebar"] .block-container {{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}}

/* أزرار Primary — أخضر الوطنية */
div.stButton > button[kind="primary"],
div.stButton > button[data-testid="baseButton-primary"],
button[kind="primary"] {{
  background: {t["accent_strong"]} !important;
  background-image: linear-gradient(180deg, {t["accent"]} 0%, {t["accent_strong"]} 100%) !important;
  color: {t["on_accent"]} !important;
  border: none !important;
  border-radius: 999px !important;
  font-weight: 700 !important;
  box-shadow: 0 2px 8px rgba(18, 109, 60, 0.25) !important;
}}
div.stButton > button[kind="primary"]:hover,
button[kind="primary"]:hover {{
  filter: brightness(1.06);
  box-shadow: 0 4px 14px rgba(18, 109, 60, 0.35) !important;
}}

/* أزرار Secondary */
div.stButton > button[kind="secondary"],
button[kind="secondary"] {{
  background: {t["surface"]} !important;
  color: {t["accent_strong"]} !important;
  border: 1.5px solid {t["accent"]} !important;
  border-radius: 999px !important;
  font-weight: 600 !important;
}}
div.stButton > button[kind="secondary"]:hover {{
  background: {t["accent_surface"]} !important;
}}

/* حقول الإدخال */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
.stTextInput input, .stNumberInput input, .stDateInput input {{
  border-radius: 10px !important;
  border-color: {t["border"]} !important;
  background: {t["input_bg"]} !important;
}}

/* تبويبات / شرائح */
.stTabs [data-baseweb="tab-list"] {{
  gap: 0.25rem;
  background: {t["surface_2"]};
  border-radius: 999px;
  padding: 4px;
}}
.stTabs [data-baseweb="tab"] {{
  border-radius: 999px !important;
  color: {t["text_dim"]} !important;
}}
.stTabs [aria-selected="true"] {{
  background: {t["accent_strong"]} !important;
  color: {t["on_accent"]} !important;
}}

/* بطاقات الحاويات */
div[data-testid="stVerticalBlockBorderWrapper"] {{
  border-radius: 14px !important;
  border-color: {t["border"]} !important;
  background: {t["surface"]} !important;
}}

/* Metrics */
div[data-testid="stMetric"] {{
  background: {t["surface_2"]};
  border: 1px solid {t["border_soft"]};
  border-radius: 14px;
  padding: 0.75rem 1rem;
}}
div[data-testid="stMetricValue"] {{
  color: {t["accent_strong"]} !important;
}}

/* جداول */
div[data-testid="stDataFrame"] {{
  border-radius: 12px !important;
  overflow: hidden;
  border: 1px solid {t["border_soft"]};
}}

/* عناوين */
h1, h2, h3 {{
  color: {t["text"]} !important;
  font-weight: 800 !important;
}}
p, span, label {{
  color: {t["text"]};
}}

/* شريط التحميل / Progress */
.stProgress > div > div > div > div {{
  background-color: {t["accent_strong"]} !important;
}}

/* Download buttons */
div.stDownloadButton > button {{
  border-radius: 999px !important;
  font-weight: 700 !important;
}}
</style>
""",
        unsafe_allow_html=True,
    )


_inject_wataniya_identity_css()

# Streamlit native theme is the source of truth for the app UI.
# Plotly receives the matching palette below; custom CSS applies Wataniya identity.

def page_header(eyebrow: str, title: str, subtitle: str, centered: bool = False):
    if eyebrow:
        st.caption(eyebrow.upper())
    st.title(title)
    st.caption(subtitle)
    st.divider()


def find_column(df: pd.DataFrame, candidates: list):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in cols_lower:
            return cols_lower[cand.lower().strip()]
    return None


# ==========================================================
# منطق الموديل
# ==========================================================

@st.cache_resource(show_spinner="جارٍ تحميل النموذج من Hugging Face...")
def load_model():
    tokenizer = AutoTokenizer.from_pretrained(MODEL_REPO)
    model = AutoModelForSequenceClassification.from_pretrained(MODEL_REPO)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model.to(device)
    model.eval()
    return tokenizer, model, device


def predict_batch(texts, tokenizer, model, device, batch_size=16):
    all_preds, all_confidences = [], []
    progress_bar = st.progress(0, text="جارٍ التصنيف...")
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

        done = min(i + batch_size, total)
        pct = int(done / total * 100) if total else 100
        progress_bar.progress(
            done / total if total else 1.0,
            text=f"جارٍ التصنيف... {pct}%  ({done}/{total})",
        )

    progress_bar.empty()
    return all_preds, all_confidences


def remove_claim_duplicates(df, claim_col, time_col, classification_col=CLASSIFICATION_COL):
    """حذف تكرارات Claim وفق اليوم والتصنيف والوقت، مع الاحتفاظ بالصف الأحدث."""
    stats = {
        "input_rows": len(df),
        "output_rows": len(df),
        "removed_rows": 0,
        "duplicate_groups": 0,
        "success_priority_groups": 0,
        "non_success_window_groups": 0,
        "skipped_rows": 0,
    }
    if not claim_col or claim_col not in df.columns or not time_col or time_col not in df.columns:
        stats["skipped_rows"] = len(df)
        return df.copy(), stats
    if classification_col not in df.columns:
        stats["skipped_rows"] = len(df)
        return df.copy(), stats

    work = df.copy()
    work["_dedup_time"] = pd.to_datetime(work[time_col], errors="coerce")
    work["_dedup_claim"] = work[claim_col].astype("string").str.strip()
    work["_dedup_day"] = work["_dedup_time"].dt.date
    work["_dedup_order"] = range(len(work))
    valid = (
        work["_dedup_time"].notna()
        & work["_dedup_day"].notna()
        & work["_dedup_claim"].notna()
        & work["_dedup_claim"].ne("")
    )
    stats["skipped_rows"] = int((~valid).sum())
    keep_indices = set(work.index[~valid])
    valid_work = work.loc[valid]

    for (_, _), group in valid_work.groupby(["_dedup_claim", "_dedup_day"], sort=False):
        group = group.sort_values(["_dedup_time", "_dedup_order"])
        if len(group) == 1:
            keep_indices.add(group.index[0])
            continue

        stats["duplicate_groups"] += 1
        successful = pd.to_numeric(group[classification_col], errors="coerce").eq(1)
        if successful.any():
            # وجود ناجحة يلغي غير الناجحة، ونحتفظ بآخر إفادة ناجحة.
            keep_indices.add(group.loc[successful].index[-1])
            stats["success_priority_groups"] += 1
            continue

        # عند كون كل التكرارات غير ناجحة: نكوّن مجموعات متجاورة بفارق أقل من 20 دقيقة.
        cluster = [group.index[0]]
        previous_time = group.iloc[0]["_dedup_time"]
        for row_index, row in group.iloc[1:].iterrows():
            current_time = row["_dedup_time"]
            if current_time - previous_time < pd.Timedelta(minutes=DUPLICATE_WINDOW_MINUTES):
                cluster.append(row_index)
            else:
                keep_indices.add(cluster[-1])
                cluster = [row_index]
            previous_time = current_time
        keep_indices.add(cluster[-1])
        stats["non_success_window_groups"] += 1

    result = df.loc[sorted(keep_indices, key=lambda index: work.loc[index, "_dedup_order"])].copy()
    stats["output_rows"] = len(result)
    stats["removed_rows"] = stats["input_rows"] - stats["output_rows"]
    return result.reset_index(drop=True), stats


def render_duplicate_summary(stats):
    """عرض نتيجة تنظيف التكرارات بعد التصنيف."""
    if not stats:
        return
    if stats.get("removed_rows", 0) > 0:
        st.success(
            f"🧹 تم حذف {stats['removed_rows']:,} تكرار من أصل {stats['input_rows']:,} صف "
            f"والاحتفاظ بـ {stats['output_rows']:,} صف."
        )
        st.caption(
            f"مجموعات بأولوية ناجحة: {stats.get('success_priority_groups', 0)} · "
            f"مجموعات غير ناجحة ضمن نافذة {DUPLICATE_WINDOW_MINUTES} دقيقة: "
            f"{stats.get('non_success_window_groups', 0)}"
        )
    elif stats.get("skipped_rows", 0) == stats.get("input_rows", 0):
        st.warning("لم تُطبَّق إزالة التكرارات: يلزم وجود عمود Claim وعمود التاريخ والوقت.")
    else:
        st.info("لم يتم العثور على تكرارات مطابقة وفق قواعد Claim واليوم والتصنيف.")


# ==========================================================
# حساب الوقت المهدر بين المكالمات لكل محصّل
# ==========================================================

def subtract_break_overlap(prev_time, curr_time, break_start, break_end, gap_minutes):
    """بيخصم من الفجوة أي جزء واقع جوه وقت الاستراحة المحدد."""
    if break_start is None or break_end is None or pd.isna(prev_time) or pd.isna(curr_time):
        return gap_minutes
    day = prev_time.date()
    break_start_dt = pd.Timestamp.combine(day, break_start)
    break_end_dt = pd.Timestamp.combine(day, break_end)
    overlap_start = max(prev_time, break_start_dt)
    overlap_end = min(curr_time, break_end_dt)
    overlap_minutes = max((overlap_end - overlap_start).total_seconds() / 60, 0)
    return max(gap_minutes - overlap_minutes, 0)


def calculate_wasted_time(df, sales_col, time_col, break_start, break_end):
    """
    بتحسب الوقت المهدر (بالدقايق) بين كل مكالمة واللي قبلها لنفس المحصّل،
    بعد استبعاد وقت الاستراحة. أول مكالمة لكل محصّل = صفر (لا يوجد مكالمة قبلها نقيس منها).
    """
    work = df.copy()
    work["_orig_idx"] = work.index
    work[time_col] = pd.to_datetime(work[time_col], errors="coerce")
    work = work.sort_values([sales_col, time_col])
    work["_prev_time"] = work.groupby(sales_col)[time_col].shift(1)

    def compute_row(row):
        if pd.isna(row["_prev_time"]) or pd.isna(row[time_col]):
            return 0.0
        gap_min = (row[time_col] - row["_prev_time"]).total_seconds() / 60
        gap_min = subtract_break_overlap(row["_prev_time"], row[time_col], break_start, break_end, gap_min)
        return round(max(gap_min, 0.0), 1)

    work[WASTED_TIME_COL] = work.apply(compute_row, axis=1)
    work = work.set_index("_orig_idx").sort_index()
    df[WASTED_TIME_COL] = work[WASTED_TIME_COL]
    return df


# Tables use Streamlit's native rendering; no cell-level CSS is injected.

# ==========================================================
# داشبورد مشترك (يُستخدم بعد التصنيف مباشرة، وكمان في تويب الداشبورد)
# ==========================================================

# لوحة ألوان موحّدة للداشبورد كله — تتبدل حسب المود المختار
COLOR_SUCCESS = THEMES[THEME_NAME]["success"]
COLOR_FAIL = THEMES[THEME_NAME]["danger"]
COLOR_ACCENT = THEMES[THEME_NAME]["accent"]
COLOR_WARN = THEMES[THEME_NAME]["warn"]
CHART_COLORS = {"ناجحة": COLOR_SUCCESS, "غير ناجحة": COLOR_FAIL}
# لوحة التصنيف: ثلاثة ألوان أساسية فقط مع درجاتها للحفاظ على هوية بصرية موحّدة.
CLASSIFICATION_PRIMARY = COLOR_ACCENT
CLASSIFICATION_SECONDARY = THEMES[THEME_NAME]["accent_strong"]
CLASSIFICATION_HIGHLIGHT = COLOR_SUCCESS
CLASSIFICATION_SCALE = [CLASSIFICATION_SECONDARY, CLASSIFICATION_PRIMARY, CLASSIFICATION_HIGHLIGHT]

# ==========================================================
# لوحة موحّدة لتويبات: الوعود + الإهمال + الجدولة + أخطاء الحالات
# لونان أساسيان (تيل) + درجة ثالثة أقرب (أفتح) فقط — بدون أحمر/أصفر منفصل
# ==========================================================
OPS_DARK = "#126D3C"      # أخضر الوطنية الغامق
OPS_MID = "#3A9B5C"       # متوسط
OPS_LIGHT = "#8BC49A"     # فاتح
OPS_SCALE = [OPS_DARK, OPS_MID, OPS_LIGHT]
OPS_POSITIVE = OPS_DARK   # قائمة / تم التغطية / منتظمة
OPS_NEGATIVE = OPS_MID    # مكسورة / لم يتم / متعثرة
OPS_NEUTRAL = OPS_LIGHT   # محايد / بدون سداد / أخرى

# Palette النشاط (نفس العائلة عشان متفرقش عن باقي الداشبورد)
ACTIVITY_PRIMARY = OPS_DARK
ACTIVITY_SECONDARY = OPS_MID
ACTIVITY_MUTED = OPS_LIGHT
ACTIVITY_AGENT_PALETTE = [OPS_DARK, "#1A8A4A", OPS_MID, "#5BB37A", OPS_LIGHT, "#B8D9C4", "#0E5A32", "#8BC49A", "#D4EBD9"]
ACTIVITY_STATE_PALETTE = [OPS_DARK, OPS_MID, OPS_LIGHT, "#D4EBD9"]
ACTIVITY_OUTCOME_COLORS = {"ناجحة": OPS_DARK, "غير ناجحة": OPS_LIGHT}
ACTIVITY_TIME_CHART_HEIGHT = 460
ACTIVITY_PAIR_CHART_HEIGHT = 420


def _activity_pair_height(n_agents, base=ACTIVITY_PAIR_CHART_HEIGHT, row_px=36, pad=160):
    """ارتفاع موحّد لزوج الشارتات جنب بعض حسب عدد المحصلين."""
    n = max(int(n_agents or 0), 1)
    return max(base, row_px * n + pad)

# الجدولة: 3 درجات من نفس اللوحة فقط
SCHEDULE_PALETTE = list(OPS_SCALE)
SCHEDULE_STATUS_COLORS = {
    "جدولة منتظمة": OPS_POSITIVE,
    "جدولة متعثرة": OPS_NEGATIVE,
    "بدون سداد": OPS_NEUTRAL,
}
SCHEDULE_AGENT_SCALE = list(OPS_SCALE)


def _activity_agent_color_map(values):
    names = sorted({str(value) for value in values if pd.notna(value)})
    return {name: ACTIVITY_AGENT_PALETTE[index % len(ACTIVITY_AGENT_PALETTE)] for index, name in enumerate(names)}


PLOTLY_TEMPLATE = "plotly_dark" if THEME_NAME == "dark" else "plotly_white"
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color=THEME["text_dim"],
    font_family="Tajawal, sans-serif",
    font_size=13,
    margin=dict(t=60, b=50, l=50, r=20),
    title_font_size=18,
    legend_font_size=12,
    hovermode="closest",
    hoverlabel=dict(
        bgcolor=THEME["surface"],
        bordercolor=THEME["border"],
        font=dict(family="Tajawal, sans-serif", size=13, color=THEME["text"]),
    ),
)
PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "responsive": True,
    "scrollZoom": True,
    "doubleClick": "reset+autosize",
    "toImageButtonOptions": {
        "format": "png",
        "filename": "classification_chart",
        "height": 900,
        "width": 1500,
        "scale": 2,
    },
}


def _apply_ops_chart_style(
    fig,
    title,
    *,
    height=430,
    xaxis_title="",
    yaxis_title="",
    show_legend=True,
    margin=None,
    extra=None,
):
    """تنسيق شارتات الإهمال/الجدولة/أخطاء الحالات بنفس أسلوب تويب التصنيف."""
    layout = {
        **PLOTLY_LAYOUT,
        "title": {
            "text": title,
            "x": 0.5,
            "xanchor": "center",
            "font": {"size": 17, "color": THEME["text"]},
        },
        "height": height,
        "xaxis_title": xaxis_title,
        "yaxis_title": yaxis_title,
        "margin": margin or dict(t=70, b=60, l=60, r=36),
        "coloraxis_showscale": False,
        "uniformtext_minsize": 11,
        "uniformtext_mode": "hide",
    }
    if show_legend:
        layout["legend"] = {
            "orientation": "h",
            "yanchor": "bottom",
            "y": -0.28,
            "x": 0.5,
            "xanchor": "center",
            "title_text": "",
            "bgcolor": "rgba(0,0,0,0)",
        }
    else:
        layout["showlegend"] = False
    if extra:
        layout.update(extra)
    fig.update_layout(**layout)
    fig.update_traces(marker_line_width=0)
    return fig


CLASSIFICATION_AGENT_FILTER_KEY = "classification_selected_agent"
PROMISES_AGENT_FILTER_KEY = "promises_selected_agent"
SCHEDULE_AGENT_FILTER_KEY = "schedule_selected_agent"
NEGLECT_AGENT_FILTER_KEY = "neglect_selected_agent"
NEGLECT_FOLLOWUP_AGENT_FILTER_KEY = "neglect_followup_selected_agent"


def _event_value(item, key, default=None):
    if isinstance(item, dict):
        return item.get(key, default)
    return getattr(item, key, default)


def _extract_selected_agent(event, fig):
    selection = _event_value(event, "selection")
    points = _event_value(selection, "points", []) if selection is not None else []
    if not points:
        return None
    point = points[0]
    selected = _event_value(point, "customdata")
    if isinstance(selected, (list, tuple)):
        selected = selected[0] if selected else None
    if selected is None:
        curve_number = _event_value(point, "curve_number", _event_value(point, "curveNumber", 0))
        point_index = _event_value(point, "point_index", _event_value(point, "pointNumber"))
        if point_index is None or curve_number >= len(fig.data):
            return None
        trace = fig.data[curve_number]
        orientation = getattr(trace, "orientation", None)
        values = trace.y if orientation == "h" else trace.x
        if values is not None and point_index < len(values):
            selected = values[point_index]
    return str(selected).strip() if selected is not None else None


def render_selectable_chart(fig, key, filter_key=CLASSIFICATION_AGENT_FILTER_KEY):
    """عرض رسم Plotly مع التقاط اختيار محصّل وإعادة تشغيل الصفحة لتطبيق الفلتر."""
    try:
        event = st.plotly_chart(
            fig,
            use_container_width=True,
            config=PLOTLY_CONFIG,
            key=key,
            on_select="rerun",
            selection_mode=("points",),
        )
    except TypeError:
        # توافق مع إصدارات Streamlit القديمة التي لا تدعم on_select.
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key=key)
        return
    selected = _extract_selected_agent(event, fig)
    if selected:
        current = st.session_state.get(filter_key)
        if selected != current:
            st.session_state[filter_key] = selected
            st.rerun()


def get_promises_view(df, sales_col):
    if not sales_col or sales_col not in df.columns:
        return df
    selected = st.session_state.get(PROMISES_AGENT_FILTER_KEY)
    if not selected:
        return df
    mask = df[sales_col].astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(PROMISES_AGENT_FILTER_KEY, None)
        return df
    return df.loc[mask].copy()


def render_promises_filter_notice():
    selected = st.session_state.get(PROMISES_AGENT_FILTER_KEY)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر النشط: عرض كل ملخصات الوعود للمحصّل «{selected}»")
    with c2:
        if st.button("إظهار الكل", key="clear_promises_agent_filter", use_container_width=True):
            st.session_state.pop(PROMISES_AGENT_FILTER_KEY, None)
            st.rerun()


def get_classification_view(df, sales_col):
    """تطبيق المحصّل المختار على بيانات التصنيف مع إبقاء العرض كاملًا افتراضيًا."""
    if not sales_col or sales_col not in df.columns:
        return df
    selected = st.session_state.get(CLASSIFICATION_AGENT_FILTER_KEY)
    if not selected:
        return df
    mask = df[sales_col].astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(CLASSIFICATION_AGENT_FILTER_KEY, None)
        return df
    return df.loc[mask].copy()


def render_classification_filter_notice(df, sales_col):
    selected = st.session_state.get(CLASSIFICATION_AGENT_FILTER_KEY)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر النشط: عرض كل مؤشرات ورسوم المحصّل «{selected}»")
    with c2:
        if st.button("إظهار الكل", key="clear_classification_agent_filter", use_container_width=True):
            st.session_state.pop(CLASSIFICATION_AGENT_FILTER_KEY, None)
            st.rerun()


def get_schedule_view(df, sales_col):
    """تطبيق فلتر المحصّل المختار من الشارت على بيانات الجدولة."""
    if not sales_col or sales_col not in df.columns:
        return df
    selected = st.session_state.get(SCHEDULE_AGENT_FILTER_KEY)
    if not selected:
        return df
    mask = df[sales_col].astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(SCHEDULE_AGENT_FILTER_KEY, None)
        return df
    return df.loc[mask].copy()


def render_schedule_filter_notice():
    selected = st.session_state.get(SCHEDULE_AGENT_FILTER_KEY)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر النشط: عرض كل مؤشرات ورسوم الجدولة للمحصّل «{selected}»")
    with c2:
        if st.button("إظهار الكل", key="clear_schedule_agent_filter", use_container_width=True):
            st.session_state.pop(SCHEDULE_AGENT_FILTER_KEY, None)
            st.rerun()


def get_neglect_view(df, sales_col, filter_key=NEGLECT_AGENT_FILTER_KEY):
    """تطبيق فلتر المحصّل المختار من الشارت على بيانات الإهمال / متابعة الإهمال."""
    if not sales_col or sales_col not in df.columns:
        return df
    selected = st.session_state.get(filter_key)
    if not selected:
        return df
    mask = df[sales_col].astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(filter_key, None)
        return df
    return df.loc[mask].copy()


def render_neglect_filter_notice(filter_key=NEGLECT_AGENT_FILTER_KEY, clear_key="clear_neglect_agent_filter"):
    selected = st.session_state.get(filter_key)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر النشط: عرض بيانات المحصّل «{selected}» فقط")
    with c2:
        if st.button("إظهار الكل", key=clear_key, use_container_width=True):
            st.session_state.pop(filter_key, None)
            st.rerun()


DASHBOARD_AGENT_FILTER_KEY = "dashboard_selected_agent"
DASHBOARD_DAY_FILTER_KEY = "dashboard_selected_day"
DASHBOARD_OUTCOME_FILTER_KEY = "dashboard_selected_outcome"  # "ناجحة" | "غير ناجحة"

ACTIVITY_NO_ANSWER_STATES = [
    "لا يرد",
    "لا يرد مع التكرار",
    "مغلق",
    "مغلق مع التكرار",
]
ACTIVITY_PROMISE_STATES = ["واعد بالسداد"]
ACTIVITY_PAYMENT_STATES = [
    "سدد كامل المديونية",
    "جدولة",
    "جدولة مقفلة",
    "سدد كامل المديونية بخصم",
]
ACTIVITY_POSITIVE_STATES = ACTIVITY_PROMISE_STATES + ACTIVITY_PAYMENT_STATES


def _state_key(value):
    """توحيد Sub State للتعامل مع اختلاف المسافات والهمزات والصياغة."""
    text = "" if pd.isna(value) else str(value).strip().casefold()
    text = (
        text.replace("أ", "ا")
        .replace("إ", "ا")
        .replace("آ", "ا")
        .replace("ى", "ي")
        .replace("ة", "ه")
    )
    return re.sub(r"[\s_-]+", "", text)


def _classify_activity_sub_state(value):
    """إرجاع مجموعة موحدة للحالة الفرعية المطلوبة في Dashboard النشاط."""
    key = _state_key(value)
    if not key:
        return "غير محدد"
    if "لايرد" in key and "تكرار" in key:
        return "لا يرد مع التكرار"
    if "مغلق" in key and "تكرار" in key:
        return "مغلق مع التكرار"
    if (
        "لايرد" in key
        or "مايرد" in key
        or "مشغول" in key
        or "بريدصوتي" in key
        or "رسالهصوتي" in key
        or "voicemail" in key
    ):
        return "لا يرد"
    if (
        key == "مغلق"
        or ("مغلق" in key and "جدوله" not in key)
        or "مفصول" in key
        or "خارجالخدمه" in key
    ):
        return "مغلق"
    if "واعد" in key and "سداد" in key:
        return "واعد بالسداد"
    if "خصم" in key and "كامل" in key and "مديون" in key:
        return "سدد كامل المديونية بخصم"
    if ("سدد" in key or "سداد" in key) and "كامل" in key and "مديون" in key:
        return "سدد كامل المديونية"
    if "جدوله" in key and ("مغلق" in key or "مقفل" in key):
        return "جدولة مقفلة"
    if "جدوله" in key:
        return "جدولة"
    return "أخرى"


def _classify_unreachable_from_notes(value):
    """تصنيف الإفادة (Notes) إذا كانت تدل على عدم الوصول للعميل.

    ترجع واحدة من ACTIVITY_NO_ANSWER_STATES أو None.
    أمثلة: لا يرد، لايرد مع التكرار، مغلق، مغلق مع التكرار، ما يرد، العميل لا يرد،
    مشغول، مفصول من الخدمة، خارج الخدمة، بريد صوتي، عدم التواصل...
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    raw = str(value).strip()
    if not raw or raw.lower() in {"nan", "none", "null", "-"}:
        return None
    key = _state_key(raw)
    if not key:
        return None

    has_repeat = ("تكرار" in key) or ("متكرر" in key) or ("again" in key) or ("repeat" in key)
    no_answer = (
        "لايرد" in key
        or "مايرد" in key
        or "مابيرد" in key
        or "لميرد" in key
        or "مشبيرد" in key
        or "مبيرد" in key
        or "مشراد" in key
        or "عميللايرد" in key
        or "لايوجدرد" in key
        or "بدونرد" in key
        or "مفيشرد" in key
        or "noreply" in key
        or "noanswer" in key
        or "notanswering" in key
        or "doesnotanswer" in key
        or "unreachable" in key
        or "busy" in key
        or "مشغول" in key
        or "الخطمشغول" in key
        or "رقمشغول" in key
        # بريد صوتي / رسالة صوتية = لم يتم التواصل مع العميل
        or "بريدصوتي" in key
        or "البريدالصوتي" in key
        or "رسالهصوتيه" in key
        or "رسالهصوتي" in key
        or "voicemail" in key
        or "voicemail" in key
        or "answeringmachine" in key
        or "machine" in key and "answer" in key
        # لم يتم التواصل / لم يتواصل العميل
        or "لميتواصل" in key
        or "ماتواصل" in key
        or "لماتواصل" in key
        or "ماتمالتواصل" in key
        or "لمتمتواصل" in key
        or "عدمتواصل" in key
        or "عدمالوصول" in key
        or "عدمالرد" in key
        or "لمالوصول" in key
        or "مفيشتواصل" in key
        or "لاتواصل" in key
        or "nocontact" in key
        or "notreachable" in key
        or "couldnotreach" in key
        or "unabletoreach" in key
        or "clientnotavailable" in key
        or "customernotavailable" in key
    )
    closed = (
        "مغلق" in key
        or "مقفول" in key
        or "الخطمغلق" in key
        or "الرقممغلق" in key
        or "switchedoff" in key
        or "phoneoff" in key
        or "poweredoff" in key
        or "outofservice" in key
        or "خارجالخدمه" in key
        or "خارجالتغطيه" in key
        or "غيرمتاح" in key
        or "الرقمغيرمتاح" in key
        or "مفصول" in key
        or "مفصولمنالخدمه" in key
        or "فصلالخدمه" in key
        or "الخطمفصول" in key
        or "الرقممفصول" in key
        or "disconnected" in key
        or "outofreach" in key
    )

    if no_answer and has_repeat:
        return "لا يرد مع التكرار"
    if closed and has_repeat:
        return "مغلق مع التكرار"
    if no_answer:
        return "لا يرد"
    if closed:
        return "مغلق"
    return None


def _activity_success_mask(df, class_col):
    if not class_col or class_col not in df.columns:
        return pd.Series(False, index=df.index)
    numeric = pd.to_numeric(df[class_col], errors="coerce")
    text = df[class_col].astype(str).str.strip().str.casefold()
    text_mask = text.isin({"1", "true", "yes", "ناجحة", "ناجحه", "successful"})
    return numeric.eq(1) | text_mask


def _calculate_wasted_time_by_day(df, sales_col, time_col, break_start=None, break_end=None):
    """حساب الفجوات بين مكالمات المحصل داخل كل يوم فقط، مع خصم البريك."""
    if not sales_col or sales_col not in df.columns or not time_col or time_col not in df.columns:
        return pd.Series(0.0, index=df.index)
    work = df[[sales_col, time_col]].copy()
    work["_calc_time"] = pd.to_datetime(work[time_col], errors="coerce")
    work["_calc_day"] = work["_calc_time"].dt.date
    work["_orig_idx"] = work.index
    work = work.sort_values([sales_col, "_calc_day", "_calc_time"])
    work["_prev_time"] = work.groupby([sales_col, "_calc_day"])["_calc_time"].shift(1)

    def gap_for_row(row):
        if pd.isna(row["_prev_time"]) or pd.isna(row["_calc_time"]):
            return 0.0
        gap = max((row["_calc_time"] - row["_prev_time"]).total_seconds() / 60, 0.0)
        return round(subtract_break_overlap(row["_prev_time"], row["_calc_time"], break_start, break_end, gap), 1)

    work[WASTED_TIME_COL] = work.apply(gap_for_row, axis=1)
    return work.set_index("_orig_idx")[WASTED_TIME_COL].reindex(df.index).fillna(0.0)


def _calculate_daily_work_hours(df, sales_col, time_col, break_start=None, break_end=None):
    """حساب فترة نشاط كل محصل في كل يوم ثم خصم الجزء المتداخل مع البريك."""
    empty = pd.DataFrame(columns=["المحصّل", "أيام النشاط", "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل"])
    if not sales_col or sales_col not in df.columns or not time_col or time_col not in df.columns:
        return empty
    work = pd.DataFrame({
        "المحصّل": df[sales_col].fillna("غير محدد").astype(str).str.strip(),
        "_activity_time": pd.to_datetime(df[time_col], errors="coerce"),
    }).dropna(subset=["_activity_time"])
    if work.empty:
        return empty
    work["_activity_day"] = work["_activity_time"].dt.date
    daily = work.groupby(["المحصّل", "_activity_day"], as_index=False)["_activity_time"].agg(
        بداية="min", نهاية="max"
    )

    def net_minutes(row):
        minutes = max((row["نهاية"] - row["بداية"]).total_seconds() / 60, 0.0)
        if break_start is not None and break_end is not None:
            minutes = subtract_break_overlap(row["بداية"], row["نهاية"], break_start, break_end, minutes)
        return max(minutes, 0.0)

    daily["دقائق العمل"] = daily.apply(net_minutes, axis=1)
    summary = daily.groupby("المحصّل")["دقائق العمل"].agg(["count", "mean", "sum"]).reset_index()
    summary = summary.rename(columns={"count": "أيام النشاط"})
    summary["متوسط ساعات العمل/اليوم"] = (summary["mean"] / 60).round(2)
    summary["إجمالي ساعات العمل"] = (summary["sum"] / 60).round(2)
    return summary[["المحصّل", "أيام النشاط", "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل"]]


def _build_activity_summary(df, class_col, sales_col, time_col, break_start=None, break_end=None):
    """بناء جدول مؤشرات المحصلين ومجموعة بيانات الرسوم من ملف النشاط."""
    if not sales_col or sales_col not in df.columns:
        return pd.DataFrame(), df.copy(), None
    work = df.copy()
    work["_agent_display"] = work[sales_col].fillna("غير محدد").astype(str).str.strip()
    work.loc[work["_agent_display"] == "", "_agent_display"] = "غير محدد"
    work["_success_bool"] = _activity_success_mask(work, class_col)
    sub_col = find_column(work, PROMISE_SUB_STATE_CANDIDATES)
    if sub_col:
        work["_activity_state"] = work[sub_col].map(_classify_activity_sub_state)
    else:
        work["_activity_state"] = "غير محدد"

    # عدم الوصول للعميل: من نص الإفادة أولاً، ولو مفيش نرجع لـ Sub State
    notes_col = find_column(work, NOTES_CANDIDATES + [ORIGINAL_TEXT_COL, MODEL_TEXT_COL])
    if notes_col:
        work["_unreachable_from_notes"] = work[notes_col].map(_classify_unreachable_from_notes)
    else:
        work["_unreachable_from_notes"] = None

    def _resolve_unreachable(row):
        note_state = row.get("_unreachable_from_notes")
        if isinstance(note_state, str) and note_state in ACTIVITY_NO_ANSWER_STATES:
            return note_state
        sub_state = row.get("_activity_state")
        if isinstance(sub_state, str) and sub_state in ACTIVITY_NO_ANSWER_STATES:
            return sub_state
        return None

    work["_unreachable_state"] = work.apply(_resolve_unreachable, axis=1)

    if time_col and time_col in work.columns and WASTED_TIME_COL not in work.columns:
        work[WASTED_TIME_COL] = _calculate_wasted_time_by_day(
            work, "_agent_display", time_col, break_start, break_end
        )
    if WASTED_TIME_COL in work.columns:
        work[WASTED_TIME_COL] = pd.to_numeric(work[WASTED_TIME_COL], errors="coerce").fillna(0)
        # الوقت المهدر يُحسب فقط لحالات عدم الوصول للعميل
        # (لا يرد / مغلق / مشغول / مفصول من الخدمة ... من الإفادة أو Sub State)
        unreachable_mask = work["_unreachable_state"].notna() & (
            work["_unreachable_state"].astype(str).str.strip() != ""
        )
        work.loc[~unreachable_mask, WASTED_TIME_COL] = 0.0

    agent = work.groupby("_agent_display", dropna=False).size().rename("إجمالي المكالمات").to_frame()
    agent["المكالمات الناجحة"] = work[work["_success_bool"]].groupby("_agent_display").size()
    agent["المكالمات الناجحة"] = agent["المكالمات الناجحة"].fillna(0).astype(int)
    agent["المكالمات غير الناجحة"] = agent["إجمالي المكالمات"] - agent["المكالمات الناجحة"]
    agent["نسبة النجاح (%)"] = (
        agent["المكالمات الناجحة"] / agent["إجمالي المكالمات"].replace(0, pd.NA) * 100
    ).fillna(0).round(1)

    # الحالات الإيجابية من Sub State
    positive_columns = ACTIVITY_PROMISE_STATES + ACTIVITY_PAYMENT_STATES
    positive_table = pd.crosstab(work["_agent_display"], work["_activity_state"])
    for state in positive_columns:
        if state not in positive_table.columns:
            positive_table[state] = 0
    positive_table = positive_table.reindex(columns=positive_columns, fill_value=0)

    # حالات عدم الوصول من الإفادة (Notes) مع احتياطي Sub State
    no_answer_src = work.dropna(subset=["_unreachable_state"])
    if not no_answer_src.empty:
        no_answer_table = pd.crosstab(no_answer_src["_agent_display"], no_answer_src["_unreachable_state"])
    else:
        no_answer_table = pd.DataFrame(index=agent.index)
    for state in ACTIVITY_NO_ANSWER_STATES:
        if state not in no_answer_table.columns:
            no_answer_table[state] = 0
    no_answer_table = no_answer_table.reindex(columns=ACTIVITY_NO_ANSWER_STATES, fill_value=0)

    state_table = positive_table.join(no_answer_table, how="outer").fillna(0)
    agent = agent.join(state_table, how="left").fillna(0)
    agent["إجمالي لا يرد"] = agent[ACTIVITY_NO_ANSWER_STATES].sum(axis=1).astype(int)
    agent["نسبة من إجمالي المكالمات (%)"] = (
        agent["إجمالي المكالمات"] / max(len(work), 1) * 100
    ).round(1)

    if WASTED_TIME_COL in work.columns:
        agent["إجمالي الوقت المهدر (دقيقة)"] = work.groupby("_agent_display")[WASTED_TIME_COL].sum().round(1)
    else:
        agent["إجمالي الوقت المهدر (دقيقة)"] = 0.0

    hours = _calculate_daily_work_hours(work, "_agent_display", time_col, break_start, break_end)
    if not hours.empty:
        agent = agent.reset_index().rename(columns={"_agent_display": "المحصّل"}).merge(hours, on="المحصّل", how="left").set_index("المحصّل")
    else:
        agent["أيام النشاط"] = 0
        agent["متوسط ساعات العمل/اليوم"] = 0.0
        agent["إجمالي ساعات العمل"] = 0.0

    agent = agent.reset_index().rename(columns={"_agent_display": "المحصّل"})
    return agent.fillna(0), work, sub_col


def _activity_layout(**overrides):
    base = {
        **PLOTLY_LAYOUT,
        "title": {"x": 0.5, "xanchor": "center", "font": {"size": 17, "color": THEME["text"]}},
        "margin": dict(t=70, b=60, l=60, r=28),
        "legend": {
            "orientation": "h",
            "yanchor": "top",
            "y": -0.18,
            "x": 0.5,
            "xanchor": "center",
            "bgcolor": "rgba(0,0,0,0)",
        },
        "uniformtext_minsize": 11,
        "uniformtext_mode": "hide",
    }
    overrides = dict(overrides)
    title_x = overrides.pop("title_x", None)
    if "title" in overrides and isinstance(overrides["title"], str):
        overrides["title"] = {**base["title"], "text": overrides["title"]}
    elif "title" in overrides and isinstance(overrides["title"], dict):
        overrides["title"] = {**base["title"], **overrides["title"]}
    if title_x is not None and isinstance(overrides.get("title"), dict):
        overrides["title"]["x"] = title_x
    return {**base, **overrides}


def render_activity_kpi_cards(total, success, agent_count, success_rate, wasted_minutes):
    cards = [
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("📞<br>إجمالي المكالمات", total, {"valueformat": ",d"}, THEME["text"]),
        ("✅<br>المكالمات الناجحة", success, {"valueformat": ",d"}, ACTIVITY_PRIMARY),
        ("📈<br>نسبة النجاح", success_rate, {"valueformat": ".1f", "suffix": "%"}, ACTIVITY_SECONDARY),
        ("⏱️<br>إجمالي الوقت المهدر", wasted_minutes, {"valueformat": ".1f", "suffix": " دقيقة"}, ACTIVITY_MUTED),
    ]
    figure = go.Figure()
    gap = 0.014
    width = (1 - gap * (len(cards) + 1)) / len(cards)
    for index, (label, value, number_format, color) in enumerate(cards):
        x0 = gap + index * (width + gap)
        x1 = x0 + width
        figure.add_shape(
            type="path",
            path=_rounded_rect_path(x0, x1, 0.04, 0.96, radius=0.022),
            xref="paper", yref="paper", layer="below",
            fillcolor=THEME["surface"], line={"color": THEME["border"], "width": 1},
        )
        figure.add_trace(go.Indicator(
            mode="number", value=float(value or 0),
            domain={"x": [x0 + 0.008, x1 - 0.008], "y": [0.13, 0.87]},
            title={"text": label, "font": {"size": 16, "color": THEME["text_dim"]}, "align": "center"},
            number={"font": {"size": 28, "color": color}, **number_format},
        ))
    figure.update_layout(
        height=205, template=PLOTLY_TEMPLATE, paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", font={"family": "Tajawal, sans-serif", "color": THEME["text"]},
        margin={"t": 8, "b": 8, "l": 8, "r": 8},
    )
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG, key="activity_kpi_cards")


def _render_agent_legend_chips(agent_color_map):
    """دليل ألوان مضغوط للمحصلين — سطر واحد قابل للتمرير أفقيًا بدل ليجند ضخم داخل الشارت."""
    if not agent_color_map:
        return
    from html import escape as _esc
    chips = "".join(
        f'<span style="display:inline-flex;align-items:center;gap:5px;padding:3px 9px;margin:0 4px;'
        f'border-radius:999px;background:{THEME["surface"]};border:1px solid {THEME["border"]};'
        f'font-size:11px;color:{THEME["text_dim"]};white-space:nowrap;flex:0 0 auto">'
        f'<span style="width:8px;height:8px;border-radius:50%;background:{color};display:inline-block;flex:0 0 auto"></span>'
        f'{_esc(str(name))}</span>'
        for name, color in agent_color_map.items()
    )
    st.markdown(
        f'<div style="overflow-x:auto;overflow-y:hidden;padding:8px 6px;margin:0 0 12px;'
        f'border:1px solid {THEME["border"]};border-radius:12px;background:{THEME["surface"]};'
        f'white-space:nowrap;scrollbar-width:thin">'
        f'<div style="display:inline-flex;align-items:center;gap:2px;min-width:max-content">{chips}</div></div>',
        unsafe_allow_html=True,
    )


def _render_dashboard_agent_filter_notice():
    agent = st.session_state.get(DASHBOARD_AGENT_FILTER_KEY)
    day = st.session_state.get(DASHBOARD_DAY_FILTER_KEY)
    outcome = st.session_state.get(DASHBOARD_OUTCOME_FILTER_KEY)
    if not agent and not day and not outcome:
        return
    parts = []
    if agent:
        parts.append(f"محصّل: «{agent}»")
    if day:
        parts.append(f"يوم: «{day}»")
    if outcome:
        parts.append(f"نتيجة: «{outcome}»")
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info("🎯 الفلتر التفاعلي النشط: " + " · ".join(parts))
    with c2:
        if st.button("إظهار الكل", key="clear_dashboard_agent_filter", use_container_width=True):
            _clear_dashboard_chart_filter()
            st.rerun()


def _dashboard_activity_view(df, sales_col, class_col=None, time_col=None):
    """تطبيق فلاتر الشارت التفاعلية: محصّل / يوم / نتيجة."""
    view = df.copy()
    agent = st.session_state.get(DASHBOARD_AGENT_FILTER_KEY)
    day = st.session_state.get(DASHBOARD_DAY_FILTER_KEY)
    outcome = st.session_state.get(DASHBOARD_OUTCOME_FILTER_KEY)

    if agent and sales_col and sales_col in view.columns:
        mask = view[sales_col].fillna("غير محدد").astype(str).str.strip().eq(str(agent).strip())
        if mask.any():
            view = view.loc[mask].copy()
        else:
            st.session_state.pop(DASHBOARD_AGENT_FILTER_KEY, None)

    if day and time_col and time_col in view.columns:
        ts = pd.to_datetime(view[time_col], errors="coerce")
        day_mask = ts.dt.strftime("%Y-%m-%d").eq(str(day).strip())
        if day_mask.any():
            view = view.loc[day_mask].copy()
        else:
            st.session_state.pop(DASHBOARD_DAY_FILTER_KEY, None)

    if outcome:
        success_mask = _activity_success_mask(view, class_col)
        if outcome == "ناجحة":
            view = view.loc[success_mask].copy()
        elif outcome == "غير ناجحة":
            view = view.loc[~success_mask].copy()

    return view


def _render_activity_hourly_chart(work, time_col):
    if not time_col or time_col not in work.columns:
        st.info("يلزم وجود عمود Created On لعرض النشاط حسب الساعة.")
        return
    trend = work.copy()
    trend["_activity_time"] = pd.to_datetime(trend[time_col], errors="coerce")
    trend = trend.dropna(subset=["_activity_time"])
    if trend.empty:
        st.info("لا توجد أوقات صالحة لعرض النشاط الساعي.")
        return
    trend["الساعة"] = trend["_activity_time"].dt.hour
    hour_min = int(trend["الساعة"].min())
    hour_max = int(trend["الساعة"].max())
    hourly = trend.groupby("الساعة", as_index=False).size().rename(columns={"size": "عدد المكالمات"})
    # ملء الساعات الفارغة بين min/max بصفر عشان المحور يبقى متصل وواضح
    full_hours = pd.DataFrame({"الساعة": list(range(hour_min, hour_max + 1))})
    hourly = full_hours.merge(hourly, on="الساعة", how="left").fillna(0)
    hourly["عدد المكالمات"] = hourly["عدد المكالمات"].astype(int)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=hourly["الساعة"],
        y=hourly["عدد المكالمات"],
        marker_color=ACTIVITY_PRIMARY,
        text=hourly["عدد المكالمات"],
        texttemplate="%{text:,}",
        textposition="outside",
        cliponaxis=False,
        hovertemplate="<b>الساعة %{x}:00</b><br>عدد المكالمات: %{y:,}<extra></extra>",
    ))
    fig.update_layout(**_activity_layout(
        title="النشاط حسب ساعة اليوم",
        height=420,
        bargap=0.18,
        showlegend=False,
        margin={"t": 56, "b": 70, "l": 60, "r": 28},
        xaxis={
            "title": {"text": "ساعة اليوم", "font": {"size": 13}},
            "dtick": 1,
            "tickvals": list(range(hour_min, hour_max + 1)),
            "ticktext": [f"{h}:00" for h in range(hour_min, hour_max + 1)],
            "range": [hour_min - 0.5, hour_max + 0.5],
            "automargin": True,
            "showgrid": False,
        },
        yaxis={
            "title": {"text": "عدد المكالمات", "font": {"size": 13}},
            "rangemode": "tozero",
            "automargin": True,
            "gridcolor": "rgba(128,145,170,0.18)",
        },
    ))
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="dashboard_activity_hourly")


def _render_activity_outcome_donut(work, class_col):
    if not class_col or class_col not in work.columns:
        st.info("يلزم وجود عمود التصنيف لعرض الناجحة مقابل غير الناجحة.")
        return
    success = int(work["_success_bool"].sum())
    failed = int(len(work) - success)
    donut_df = pd.DataFrame({"النتيجة": ["ناجحة", "غير ناجحة"], "العدد": [success, failed]})
    rate = success / len(work) * 100 if len(work) else 0
    fig = px.pie(
        donut_df, names="النتيجة", values="العدد", hole=0.62,
        color="النتيجة", color_discrete_map=ACTIVITY_OUTCOME_COLORS, template=PLOTLY_TEMPLATE,
    )
    fig.update_layout(**_activity_layout(
        title="نتيجة المكالمات",
        height=ACTIVITY_PAIR_CHART_HEIGHT,
        margin={"t": 70, "b": 40, "l": 20, "r": 20},
        legend_title_text="",
        annotations=[{
            "text": f"<b>{rate:.1f}%</b><br>نجاح",
            "x": 0.5, "y": 0.5, "font": {"size": 20, "color": ACTIVITY_PRIMARY},
            "showarrow": False,
        }],
    ))
    fig.update_traces(
        texttemplate="%{label}<br>%{value:,}<br>%{percent:.1%}",
        textfont=dict(size=14, color=THEME["text"]),
        textinfo="text",
        customdata=donut_df["النتيجة"],
        hovertemplate="<b>%{label}</b><br>العدد: %{value:,}<br>النسبة: %{percent:.1%}<extra></extra>",
    )
    try:
        event = st.plotly_chart(
            fig, use_container_width=True, config=PLOTLY_CONFIG,
            key="dashboard_outcome_donut", on_select="rerun", selection_mode=("points",),
        )
        selection = _event_value(event, "selection") if event is not None else None
        points = _event_value(selection, "points", []) if selection is not None else []
        if points:
            point = points[0]
            selected = _event_value(point, "customdata")
            if isinstance(selected, (list, tuple)):
                selected = selected[0] if selected else None
            if selected is None:
                selected = _event_value(point, "label")
            if selected in ("ناجحة", "غير ناجحة"):
                if st.session_state.get(DASHBOARD_OUTCOME_FILTER_KEY) != selected:
                    st.session_state[DASHBOARD_OUTCOME_FILTER_KEY] = selected
                    st.rerun()
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="dashboard_outcome_donut")


def _render_activity_no_answer_chart(agent):
    available = [state for state in ACTIVITY_NO_ANSWER_STATES if state in agent.columns]
    if not available:
        st.info("لا توجد حالات عدم وصول (من الإفادة أو Sub State) للعرض.")
        return
    plot = agent[["المحصّل"] + available].copy()
    plot["إجمالي لا يرد"] = plot[available].sum(axis=1)
    plot = plot.sort_values("إجمالي لا يرد", ascending=True)
    long = plot.melt(id_vars=["المحصّل"], value_vars=available, var_name="الحالة", value_name="العدد")
    fig = px.bar(
        long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack",
        template=PLOTLY_TEMPLATE, category_orders={"الحالة": ACTIVITY_NO_ANSWER_STATES},
        color_discrete_sequence=ACTIVITY_STATE_PALETTE,
    )
    fig.update_traces(
        marker_line_width=0,
        texttemplate="%{x}",
        textposition="inside",
        insidetextanchor="middle",
        textfont_size=11,
        customdata=long["المحصّل"],
        hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
    )
    row_totals = long.groupby("المحصّل", sort=False)["العدد"].sum()
    fig.add_trace(go.Scatter(
        x=row_totals.values,
        y=row_totals.index.astype(str),
        mode="text",
        text=[f"{int(v)}" for v in row_totals.values],
        textposition="middle right",
        textfont={"size": 12, "color": THEME["text"]},
        showlegend=False,
        hoverinfo="skip",
        cliponaxis=False,
    ))
    fig.update_layout(**_activity_layout(
        title="عدم الوصول للعميل من الإفادة (لا يرد / مغلق)",
        height=ACTIVITY_PAIR_CHART_HEIGHT,
        legend_title_text="",
        legend={"orientation": "h", "yanchor": "top", "y": -0.2, "x": 0.5, "xanchor": "center", "font": {"size": 11}},
        margin={"t": 56, "b": 90, "l": 160, "r": 55},
        xaxis={
            "title": {"text": "عدد الحالات", "font": {"size": 13}},
            "automargin": True,
            "rangemode": "tozero",
            "gridcolor": "rgba(128,145,170,0.18)",
        },
        yaxis={"title": "", "categoryorder": "total ascending", "automargin": True},
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    ))
    render_selectable_chart(fig, "dashboard_no_answer_states", filter_key=DASHBOARD_AGENT_FILTER_KEY)


def _render_activity_leaderboard_chart(agent):
    if agent.empty:
        st.info("لا توجد بيانات محصلين للعرض.")
        return
    plot = agent[["المحصّل", "إجمالي المكالمات", "نسبة النجاح (%)"]].copy()
    plot = plot.sort_values("إجمالي المكالمات", ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=plot["المحصّل"],
        x=plot["إجمالي المكالمات"],
        orientation="h",
        marker_color=ACTIVITY_SECONDARY,
        text=[f"{int(c):,}  |  {float(r):.0f}%" for c, r in zip(plot["إجمالي المكالمات"], plot["نسبة النجاح (%)"])],
        textposition="outside",
        cliponaxis=False,
        customdata=plot["المحصّل"],
        hovertemplate="<b>%{y}</b><br>المكالمات: %{x:,}<br>نسبة النجاح ضمن النص<extra></extra>",
    ))
    fig.update_layout(**_activity_layout(
        title="ترتيب المحصلين (المكالمات | نسبة النجاح %)",
        height=max(400, 34 * len(plot) + 140),
        showlegend=False,
        margin={"t": 56, "b": 56, "l": 170, "r": 90},
        xaxis={
            "title": {"text": "عدد المكالمات", "font": {"size": 13}},
            "rangemode": "tozero",
            "automargin": True,
            "gridcolor": "rgba(128,145,170,0.18)",
        },
        yaxis={
            "title": "",
            "automargin": True,
        },
    ))
    render_selectable_chart(fig, "dashboard_leaderboard", filter_key=DASHBOARD_AGENT_FILTER_KEY)


def _render_activity_positive_states_chart(agent):
    """حالات الوعد بالسداد والسداد الفعلي/الجدولة لكل محصّل."""
    available = [state for state in ACTIVITY_POSITIVE_STATES if state in agent.columns]
    if not available:
        st.info("يلزم وجود عمود Sub State لعرض حالات الوعد والسداد.")
        return
    plot = agent[["المحصّل"] + available].copy()
    plot["إجمالي الحالات الإيجابية"] = plot[available].sum(axis=1)
    plot = plot.sort_values("إجمالي الحالات الإيجابية", ascending=True)
    long = plot.melt(id_vars=["المحصّل"], value_vars=available, var_name="الحالة", value_name="العدد")
    palette = (ACTIVITY_AGENT_PALETTE * 2)[: len(available)]
    fig = px.bar(
        long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack",
        template=PLOTLY_TEMPLATE, category_orders={"الحالة": ACTIVITY_POSITIVE_STATES},
        color_discrete_sequence=palette,
    )
    fig.update_traces(
        marker_line_width=0,
        texttemplate="%{x}",
        textposition="inside",
        insidetextanchor="middle",
        textfont_size=11,
        customdata=long["المحصّل"],
        hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
    )
    row_totals = long.groupby("المحصّل", sort=False)["العدد"].sum()
    fig.add_trace(go.Scatter(
        x=row_totals.values,
        y=row_totals.index.astype(str),
        mode="text",
        text=[f"{int(v)}" for v in row_totals.values],
        textposition="middle right",
        textfont={"size": 12, "color": THEME["text"]},
        showlegend=False,
        hoverinfo="skip",
        cliponaxis=False,
    ))
    pair_h = _activity_pair_height(len(plot))
    fig.update_layout(**_activity_layout(
        title="حالات الوعد والسداد لكل محصل",
        height=pair_h,
        legend_title_text="",
        legend={"orientation": "h", "yanchor": "top", "y": -0.22, "x": 0.5, "xanchor": "center", "font": {"size": 11}},
        margin={"t": 56, "b": 95, "l": 160, "r": 55},
        xaxis={
            "title": {"text": "عدد الحالات", "font": {"size": 13}},
            "automargin": True,
            "rangemode": "tozero",
            "gridcolor": "rgba(128,145,170,0.18)",
        },
        yaxis={"title": "", "categoryorder": "total ascending", "automargin": True},
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    ))
    render_selectable_chart(fig, "dashboard_positive_states", filter_key=DASHBOARD_AGENT_FILTER_KEY)


def _render_activity_hours_efficiency_chart(agent):
    """ساعات العمل (عمود) + الوقت المهدر (عمود أرقام يمين) بدون تداخل."""
    if agent.empty or "إجمالي ساعات العمل" not in agent.columns:
        st.info("لا تتوفر بيانات ساعات عمل كافية لعرض الإنتاجية.")
        return
    plot = agent[["المحصّل", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)"]].copy()
    plot["إجمالي ساعات العمل"] = pd.to_numeric(plot["إجمالي ساعات العمل"], errors="coerce").fillna(0)
    plot["إجمالي الوقت المهدر (دقيقة)"] = pd.to_numeric(plot["إجمالي الوقت المهدر (دقيقة)"], errors="coerce").fillna(0)
    plot = plot.sort_values("إجمالي ساعات العمل", ascending=True)

    max_h = float(plot["إجمالي ساعات العمل"].max()) if len(plot) else 1.0
    # عمود ثابت يمين كل الأعمدة لعرض الماس + رقم الوقت المهدر
    waste_x = max(max_h * 1.22, max_h + 1.2, 1.0)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=plot["المحصّل"],
        x=plot["إجمالي ساعات العمل"],
        orientation="h",
        marker_color=ACTIVITY_PRIMARY,
        name="ساعات العمل",
        text=plot["إجمالي ساعات العمل"].map(lambda v: f"{float(v):.1f}"),
        textposition="inside",
        insidetextanchor="middle",
        textfont={"size": 13, "color": "#FFFFFF"},
        cliponaxis=False,
        customdata=plot["المحصّل"],
        hovertemplate="<b>%{y}</b><br>ساعات العمل: %{x:.1f}<extra></extra>",
    ))
    # الوقت المهدر في صف واحد يمين الأعمدة — مش على محور متراكب
    fig.add_trace(go.Scatter(
        y=plot["المحصّل"],
        x=[waste_x] * len(plot),
        mode="markers+text",
        marker={
            "color": ACTIVITY_MUTED,
            "size": 11,
            "symbol": "diamond",
            "line": {"width": 1, "color": "#FFFFFF"},
        },
        name="الوقت المهدر (دقيقة)",
        text=plot["إجمالي الوقت المهدر (دقيقة)"].map(lambda v: f"{float(v):.0f}"),
        textposition="middle right",
        textfont={"size": 12, "color": THEME["text"]},
        cliponaxis=False,
        customdata=plot["إجمالي الوقت المهدر (دقيقة)"],
        hovertemplate="<b>%{y}</b><br>الوقت المهدر: %{customdata:.0f} دقيقة<extra></extra>",
    ))
    pair_h = _activity_pair_height(len(plot))
    fig.update_layout(**_activity_layout(
        title="ساعات العمل مقابل الوقت المهدر",
        height=pair_h,
        legend={"orientation": "h", "yanchor": "top", "y": -0.18, "x": 0.5, "xanchor": "center", "title_text": ""},
        margin={"t": 56, "b": 95, "l": 170, "r": 110},
        xaxis={
            "title": {"text": "ساعات العمل", "font": {"size": 13}},
            "rangemode": "tozero",
            "range": [0, waste_x * 1.18],
            "automargin": True,
            "gridcolor": "rgba(128,145,170,0.18)",
        },
        yaxis={"title": "", "automargin": True},
        bargap=0.28,
    ))
    render_selectable_chart(fig, "dashboard_hours_efficiency", filter_key=DASHBOARD_AGENT_FILTER_KEY)


def _render_activity_table(agent):
    table_columns = [
        "المحصّل", "إجمالي المكالمات", "المكالمات الناجحة", "نسبة النجاح (%)", "نسبة من إجمالي المكالمات (%)",
        "واعد بالسداد", "سدد كامل المديونية", "جدولة", "جدولة مقفلة", "سدد كامل المديونية بخصم",
        "لا يرد", "لا يرد مع التكرار", "مغلق", "مغلق مع التكرار", "إجمالي لا يرد",
        "أيام النشاط", "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)",
    ]
    table = agent[[c for c in table_columns if c in agent.columns]].copy()
    for col in ["نسبة النجاح (%)", "نسبة من إجمالي المكالمات (%)", "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)"]:
        if col in table.columns:
            table[col] = pd.to_numeric(table[col], errors="coerce").fillna(0).round(2)
    st.dataframe(table.sort_values("إجمالي المكالمات", ascending=False), use_container_width=True, hide_index=True)


def render_activity_dashboard(df, class_col=None, sales_col=None, time_col=None, break_start=None, break_end=None):
    """Dashboard تحليل نشاط المحصلين: KPI، اتجاهات زمنية، حالات Sub State، ساعات العمل، وجدول تفصيلي."""
    if not sales_col or sales_col not in df.columns:
        st.error("لا يوجد عمود واضح للمحصّل (Create By / Sales Person) في الملف.")
        return
    _render_dashboard_agent_filter_notice()
    view = _dashboard_activity_view(df, sales_col, class_col=class_col, time_col=time_col)
    agent, work, sub_col = _build_activity_summary(view, class_col, sales_col, time_col, break_start, break_end)
    if agent.empty:
        st.info("لا توجد مكالمات قابلة للعرض بعد تطبيق الفلاتر.")
        return
    total = len(work)
    success = int(work["_success_bool"].sum())
    success_rate = success / total * 100 if total else 0
    wasted = float(pd.to_numeric(work.get(WASTED_TIME_COL, pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    st.subheader("📌 مؤشرات الأداء الرئيسية")
    render_activity_kpi_cards(total, success, int(agent["المحصّل"].nunique()), success_rate, wasted)
    st.caption("اضغط على أي عنصر في الشارتات (محصّل / يوم / نتيجة) للفلترة · «إظهار الكل» للإلغاء · الألوان موحّدة على 3 درجات.")

    st.markdown("#### 🕒 النشاط حسب الساعة")
    with st.container(border=True):
        _render_activity_hourly_chart(work, time_col)

    st.markdown("#### 🏆 مقارنة أداء المحصلين")
    with st.container(border=True):
        _render_activity_leaderboard_chart(agent)

    st.markdown("#### 📊 نتائج المكالمات وحالات المتابعة")
    outcome_col, no_answer_col = st.columns(2)
    with outcome_col:
        with st.container(border=True):
            _render_activity_outcome_donut(work, class_col)
    with no_answer_col:
        with st.container(border=True):
            _render_activity_no_answer_chart(agent)

    positive_col, hours_col = st.columns(2)
    with positive_col:
        with st.container(border=True):
            _render_activity_positive_states_chart(agent)
    with hours_col:
        with st.container(border=True):
            _render_activity_hours_efficiency_chart(agent)

    st.markdown("#### 📋 الجدول التفصيلي")
    with st.container(border=True):
        st.subheader("📋 جدول أداء كل محصل وحالات Sub State وساعات العمل")
        if not sub_col:
            st.warning("لم يتم العثور على عمود Sub State؛ ستظهر أعمدة الحالات بصفر حتى يتم رفع ملف يحتوي على العمود.")
        _render_activity_table(agent)


def render_full_dashboard(df, class_col=None, sales_col=None, time_col=None, break_start=None, break_end=None):
    render_activity_dashboard(df, class_col, sales_col, time_col, break_start, break_end)


# ==========================================================
# تصدير الداشبورد كصفحة ويب (HTML) مستقلة — نفس الشكل والألوان
# ==========================================================


def build_dashboard_html(df, class_col, sales_col, time_col, source_name="", filter_hint="", filter_summary=None,
                          wallet_df=None, payments_df=None) -> str:
    """إنشاء نسخة HTML مستقلة من Dashboard النشاط — تنسيق واضح وشارتات بـ IDs ثابتة."""
    from html import escape
    import json

    # لوحة HTML موحّدة: لونان تيل + درجة ثالثة أقرب فقط
    background = "#F3F7F7"
    surface = "#FFFFFF"
    border = "#D5E3E4"
    text = "#1F2A2B"
    text_dim = "#5A6F71"
    export_dark = "#2F6F73"      # أساسي غامق
    export_mid = "#5A9093"       # متوسط
    export_light = "#8FB4B7"     # فاتح
    export_success = export_dark
    export_fail = export_light
    export_accent = export_mid
    export_warn = export_mid
    export_scale = [export_dark, export_mid, export_light]
    export_template = "plotly_white"

    def export_layout(**overrides):
        base = {
            "template": export_template,
            "paper_bgcolor": "rgba(0,0,0,0)",
            "plot_bgcolor": "rgba(0,0,0,0)",
            "font": {"family": "Tahoma, Segoe UI, Arial, sans-serif", "color": text, "size": 13},
            "title": {"x": 0.5, "xanchor": "center", "font": {"size": 16, "color": text}},
            "margin": dict(t=64, b=72, l=60, r=36),
            "legend": {
                "orientation": "h", "yanchor": "top", "y": -0.2,
                "x": 0.5, "xanchor": "center", "bgcolor": "rgba(0,0,0,0)",
            },
            "hoverlabel": {"bgcolor": surface, "font": {"color": text, "family": "Tahoma, Arial"}},
            "colorway": [export_dark, export_mid, export_light],
        }
        overrides = dict(overrides)
        title_x = overrides.pop("title_x", None)
        if "title" in overrides and isinstance(overrides["title"], str):
            overrides["title"] = {**base["title"], "text": overrides["title"]}
        elif "title" in overrides and isinstance(overrides["title"], dict):
            overrides["title"] = {**base["title"], **overrides["title"]}
        if title_x is not None and isinstance(overrides.get("title"), dict):
            overrides["title"]["x"] = title_x
        return {**base, **overrides}

    work = df.copy()
    if sales_col and sales_col in work.columns:
        work["_agent_display"] = work[sales_col].fillna("غير محدد").astype(str).str.strip()
    else:
        work["_agent_display"] = "غير محدد"
    work["_success_bool"] = _activity_success_mask(work, class_col)
    sub_col_for_export = find_column(work, PROMISE_SUB_STATE_CANDIDATES)
    if sub_col_for_export:
        work["_activity_state"] = work[sub_col_for_export].map(_classify_activity_sub_state)
    else:
        work["_activity_state"] = ""

    notes_col_export = find_column(work, NOTES_CANDIDATES + [ORIGINAL_TEXT_COL, MODEL_TEXT_COL])
    if notes_col_export:
        work["_unreachable_from_notes"] = work[notes_col_export].map(_classify_unreachable_from_notes)
    else:
        work["_unreachable_from_notes"] = None

    def _resolve_unreachable_export(row):
        note_state = row.get("_unreachable_from_notes")
        if isinstance(note_state, str) and note_state in ACTIVITY_NO_ANSWER_STATES:
            return note_state
        sub_state = row.get("_activity_state")
        if isinstance(sub_state, str) and sub_state in ACTIVITY_NO_ANSWER_STATES:
            return sub_state
        return None

    work["_unreachable_state"] = work.apply(_resolve_unreachable_export, axis=1)

    # الوقت المهدر في HTML فقط لحالات عدم الوصول للعميل
    if WASTED_TIME_COL in work.columns:
        work[WASTED_TIME_COL] = pd.to_numeric(work[WASTED_TIME_COL], errors="coerce").fillna(0)
        unreachable_mask = work["_unreachable_state"].notna() & (
            work["_unreachable_state"].astype(str).str.strip() != ""
        )
        work.loc[~unreachable_mask, WASTED_TIME_COL] = 0.0

    total = len(work)
    success = int(work["_success_bool"].sum())
    rate = success / total * 100 if total else 0
    wasted = float(pd.to_numeric(work.get(WASTED_TIME_COL, pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    agent_count = int(work["_agent_display"].nunique()) if total else 0

    timed_source = (
        pd.to_datetime(work[time_col], errors="coerce")
        if time_col and time_col in work.columns
        else pd.Series(pd.NaT, index=work.index)
    )
    def _clean_state_value(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        text_val = str(val).strip()
        if not text_val or text_val.lower() in {"nan", "none", "null"}:
            return ""
        return text_val

    raw_records = []
    for row_index, row in work.iterrows():
        timestamp = timed_source.loc[row_index] if row_index in timed_source.index else pd.NaT
        raw_records.append({
            "agent": str(row.get("_agent_display", "غير محدد")),
            "time": timestamp.isoformat() if pd.notna(timestamp) else "",
            "success": bool(row.get("_success_bool", False)),
            "wasted": float(pd.to_numeric(row.get(WASTED_TIME_COL, 0), errors="coerce") or 0),
            "state": _clean_state_value(row.get("_activity_state")),
            "unreachable": _clean_state_value(row.get("_unreachable_state")),
        })

    _agent_names = sorted({str(v) for v in work["_agent_display"].dropna().astype(str) if str(v).strip()})
    agent_color_map = {name: export_scale[i % len(export_scale)] for i, name in enumerate(_agent_names)}
    agent_color_json = json.dumps(agent_color_map, ensure_ascii=False)
    # ألوان حالات اللا يرد من نفس 3 درجات فقط
    _state_cols = [export_scale[i % len(export_scale)] for i in range(len(ACTIVITY_NO_ANSWER_STATES))]
    state_color_json = json.dumps(dict(zip(ACTIVITY_NO_ANSWER_STATES, _state_cols)), ensure_ascii=False)
    positive_state_colors = list(export_scale)
    positive_state_color_json = json.dumps(dict(zip(ACTIVITY_POSITIVE_STATES, positive_state_colors)), ensure_ascii=False)
    raw_records_json = json.dumps(raw_records, ensure_ascii=False)

    agent_table = pd.DataFrame()
    if sales_col and sales_col in work.columns:
        agent_table, _, _ = _build_activity_summary(work, class_col, sales_col, time_col)

    # خريطة ساعات العمل لكل محصل — لتحديث شارت الكفاءة في صفحة HTML
    agent_hours_map = {}
    if (
        not agent_table.empty
        and "المحصّل" in agent_table.columns
        and "إجمالي ساعات العمل" in agent_table.columns
    ):
        for _, row in agent_table.iterrows():
            name = str(row.get("المحصّل", "")).strip()
            if not name:
                continue
            agent_hours_map[name] = float(pd.to_numeric(row.get("إجمالي ساعات العمل"), errors="coerce") or 0)
    agent_hours_json = json.dumps(agent_hours_map, ensure_ascii=False)

    export_date_min = export_date_max = ""
    if timed_source.notna().any():
        export_date_min = timed_source.min().strftime("%Y-%m-%d")
        export_date_max = timed_source.max().strftime("%Y-%m-%d")

    def metric_card(card_id, label, value, color):
        return (
            f'<div class="kpi" id="{card_id}">'
            f'<div class="label">{label}</div>'
            f'<div class="value" data-role="value" style="color:{color}">{value}</div></div>'
        )

    # ---- بناء الشارتات بـ IDs ثابتة وتنسيق نظيف (بدون ليجند مزدحم) ----
    chart_specs = []  # (section, title, fig, plot_id)

    if class_col and class_col in work.columns:
        donut_df = pd.DataFrame({"النتيجة": ["ناجحة", "غير ناجحة"], "العدد": [success, max(total - success, 0)]})
        fig = px.pie(
            donut_df, names="النتيجة", values="العدد", hole=0.64,
            color="النتيجة",
            color_discrete_map={"ناجحة": export_success, "غير ناجحة": export_fail},
            template=export_template,
        )
        fig.update_traces(
            textinfo="percent",
            textfont_size=14,
            textposition="inside",
            marker={"line": {"color": surface, "width": 3}},
            hovertemplate="<b>%{label}</b><br>العدد: %{value:,}<br>النسبة: %{percent}<extra></extra>",
        )
        fig.update_layout(**export_layout(
            title="توزيع نتائج المكالمات", height=380,
            margin=dict(t=56, b=48, l=20, r=20),
            showlegend=True,
            legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center", font=dict(size=12)),
            annotations=[{"text": f"<b>{rate:.1f}%</b><br>نجاح", "x": 0.5, "y": 0.5,
                          "font": {"size": 18, "color": export_success}, "showarrow": False}],
        ))
        chart_specs.append(("main", "توزيع نتائج المكالمات", fig, "plot_donut"))

    if time_col and time_col in work.columns:
        work["_activity_time"] = pd.to_datetime(work[time_col], errors="coerce")
        timed = work.dropna(subset=["_activity_time"]).copy()
        if not timed.empty:
            timed["الساعة"] = timed["_activity_time"].dt.hour
            hourly = timed.groupby("الساعة", as_index=False).size().rename(columns={"size": "عدد المكالمات"})
            hour_min = int(hourly["الساعة"].min())
            hour_max = int(hourly["الساعة"].max())
            full_hours = pd.DataFrame({"الساعة": list(range(hour_min, hour_max + 1))})
            hourly = full_hours.merge(hourly, on="الساعة", how="left").fillna(0)
            hourly["عدد المكالمات"] = hourly["عدد المكالمات"].astype(int)
            hour_fig = go.Figure()
            hour_fig.add_trace(go.Bar(
                x=hourly["الساعة"], y=hourly["عدد المكالمات"],
                marker_color=export_success,
                text=hourly["عدد المكالمات"],
                texttemplate="%{text:,}",
                textposition="outside",
                cliponaxis=False,
                hovertemplate="<b>الساعة %{x}:00</b><br>عدد المكالمات: %{y:,}<extra></extra>",
            ))
            hour_fig.update_layout(**export_layout(
                title="النشاط حسب ساعة اليوم", height=420, bargap=0.18,
                xaxis={
                    "dtick": 1,
                    "tickvals": list(range(hour_min, hour_max + 1)),
                    "ticktext": [f"{h}:00" for h in range(hour_min, hour_max + 1)],
                    "range": [hour_min - 0.5, hour_max + 0.5],
                    "title": {"text": "ساعة اليوم", "font": {"size": 13}},
                    "automargin": True,
                    "showgrid": False,
                },
                yaxis={
                    "title": {"text": "عدد المكالمات", "font": {"size": 13}},
                    "rangemode": "tozero",
                    "automargin": True,
                },
                margin=dict(t=60, b=80, l=60, r=28),
                showlegend=False,
            ))
            chart_specs.append(("main", "النشاط الساعي", hour_fig, "plot_hourly"))

    if not agent_table.empty:
        board = agent_table[["المحصّل", "إجمالي المكالمات", "نسبة النجاح (%)"]].copy()
        board = board.sort_values("إجمالي المكالمات", ascending=True)
        # بار واحد واضح + نسبة النجاح كنص على العمود (من غير محور مزدوج مربك)
        board_fig = go.Figure()
        board_fig.add_trace(go.Bar(
            y=board["المحصّل"], x=board["إجمالي المكالمات"], orientation="h",
            marker_color=export_accent,
            text=[f"{int(c):,}  |  {float(r):.0f}%" for c, r in zip(board["إجمالي المكالمات"], board["نسبة النجاح (%)"])],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>المكالمات: %{x:,}<extra></extra>",
        ))
        board_fig.update_layout(**export_layout(
            title="ترتيب المحصلين (المكالمات | نسبة النجاح)",
            height=max(400, 36 * len(board) + 140),
            xaxis_title="عدد المكالمات", yaxis_title="",
            xaxis={"automargin": True, "rangemode": "tozero"},
            yaxis={"automargin": True},
            margin=dict(t=60, b=50, l=170, r=90),
            showlegend=False,
        ))
        chart_specs.append(("rank", "ترتيب أداء المحصلين", board_fig, "plot_leaderboard"))

    if work["_unreachable_state"].notna().any():
        state_counts = work.dropna(subset=["_unreachable_state"]).pivot_table(
            index="_agent_display", columns="_unreachable_state", aggfunc="size", fill_value=0
        )
        for state_name in ACTIVITY_NO_ANSWER_STATES:
            if state_name not in state_counts.columns:
                state_counts[state_name] = 0
        state_counts = state_counts.reindex(columns=ACTIVITY_NO_ANSWER_STATES, fill_value=0).reset_index()
        state_counts = state_counts.rename(columns={"_agent_display": "المحصّل"})
        state_long = state_counts.melt(id_vars=["المحصّل"], var_name="الحالة", value_name="العدد")
        if not state_long.empty and state_long["العدد"].sum() > 0:
            no_fig = px.bar(
                state_long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack",
                template=export_template,
                category_orders={"الحالة": ACTIVITY_NO_ANSWER_STATES},
                color_discrete_sequence=ACTIVITY_STATE_PALETTE,
            )
            no_fig.update_layout(**export_layout(
                title="عدم الوصول للعميل من الإفادة", height=max(400, 34 * state_counts.shape[0] + 160),
                xaxis_title="عدد الحالات", yaxis_title="",
                xaxis={"automargin": True}, yaxis={"automargin": True, "categoryorder": "total ascending"},
                margin=dict(t=60, b=90, l=170, r=28),
                showlegend=True,
                legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center", font=dict(size=11), title_text=""),
            ))
            no_fig.update_traces(marker_line_width=0, hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>")
            chart_specs.append(("states", "حالات لا يرد", no_fig, "plot_no_answer"))

    # شارت الوعد/السداد من Sub State مباشرة (مش من agent_table بس)
    pos_src = (
        work[work["_activity_state"].isin(ACTIVITY_POSITIVE_STATES)].copy()
        if "_activity_state" in work.columns else pd.DataFrame()
    )
    if not pos_src.empty:
        pos_counts = pos_src.pivot_table(
            index="_agent_display", columns="_activity_state", aggfunc="size", fill_value=0
        )
        for state_name in ACTIVITY_POSITIVE_STATES:
            if state_name not in pos_counts.columns:
                pos_counts[state_name] = 0
        pos_counts = pos_counts.reindex(columns=ACTIVITY_POSITIVE_STATES, fill_value=0).reset_index()
        pos_counts = pos_counts.rename(columns={"_agent_display": "المحصّل"})
        available_positive = list(ACTIVITY_POSITIVE_STATES)
        pos_plot = pos_counts.copy()
        pos_plot["_ترتيب"] = pos_plot[available_positive].sum(axis=1)
        pos_plot = pos_plot[pos_plot["_ترتيب"] > 0].sort_values("_ترتيب", ascending=True)
        pos_long = pos_plot.melt(id_vars=["المحصّل"], value_vars=available_positive, var_name="الحالة", value_name="العدد")
        if not pos_long.empty and pos_long["العدد"].sum() > 0:
            pos_fig = px.bar(
                pos_long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack",
                template=export_template,
                category_orders={"الحالة": ACTIVITY_POSITIVE_STATES},
                color_discrete_sequence=positive_state_colors,
            )
            pos_fig.update_layout(**export_layout(
                title="حالات الوعد والسداد", height=max(420, 36 * len(pos_plot) + 170),
                xaxis_title="عدد الحالات", yaxis_title="",
                xaxis={"automargin": True, "rangemode": "tozero", "title": {"text": "عدد الحالات", "font": {"size": 13}}},
                yaxis={"automargin": True, "categoryorder": "total ascending"},
                margin=dict(t=60, b=100, l=170, r=55),
                showlegend=True,
                legend=dict(orientation="h", y=-0.24, x=0.5, xanchor="center", font=dict(size=11), title_text=""),
                uniformtext_minsize=10,
                uniformtext_mode="hide",
            ))
            pos_fig.update_traces(
                marker_line_width=0,
                texttemplate="%{x}",
                textposition="inside",
                insidetextanchor="middle",
                textfont_size=11,
                hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
            )
            pos_totals = pos_long.groupby("المحصّل", sort=False)["العدد"].sum()
            pos_fig.add_trace(go.Scatter(
                x=pos_totals.values,
                y=pos_totals.index.astype(str),
                mode="text",
                text=[f"{int(v)}" for v in pos_totals.values],
                textposition="middle right",
                textfont={"size": 12, "color": text, "family": "Tahoma, Arial"},
                showlegend=False,
                hoverinfo="skip",
                cliponaxis=False,
            ))
            chart_specs.append(("states", "الوعد والسداد", pos_fig, "plot_positive"))

    if (
        not agent_table.empty
        and "إجمالي ساعات العمل" in agent_table.columns
        and "إجمالي الوقت المهدر (دقيقة)" in agent_table.columns
    ):
        eff = agent_table[["المحصّل", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)"]].copy()
        eff["إجمالي ساعات العمل"] = pd.to_numeric(eff["إجمالي ساعات العمل"], errors="coerce").fillna(0)
        eff["إجمالي الوقت المهدر (دقيقة)"] = pd.to_numeric(eff["إجمالي الوقت المهدر (دقيقة)"], errors="coerce").fillna(0)
        eff = eff.sort_values("إجمالي ساعات العمل", ascending=True)
        max_h = float(eff["إجمالي ساعات العمل"].max()) if len(eff) else 1.0
        max_w = float(eff["إجمالي الوقت المهدر (دقيقة)"].max()) if len(eff) else 1.0
        waste_x = max(max_h * 1.22, max_h + 1.2, 1.0)
        eff_fig = go.Figure()
        eff_fig.add_trace(go.Bar(
            y=eff["المحصّل"],
            x=eff["إجمالي ساعات العمل"],
            orientation="h",
            marker_color=export_success,
            name="ساعات العمل",
            text=eff["إجمالي ساعات العمل"].map(lambda v: f"{float(v):.1f}"),
            textposition="inside",
            insidetextanchor="middle",
            textfont={"size": 13, "color": "#FFFFFF"},
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>ساعات العمل: %{x:.1f}<extra></extra>",
        ))
        eff_fig.add_trace(go.Scatter(
            y=eff["المحصّل"],
            x=[waste_x] * len(eff),
            mode="markers+text",
            marker={"color": export_warn, "size": 11, "symbol": "diamond", "line": {"width": 1, "color": "#FFFFFF"}},
            name="الوقت المهدر (دقيقة)",
            text=eff["إجمالي الوقت المهدر (دقيقة)"].map(lambda v: f"{float(v):.0f}"),
            textposition="middle right",
            textfont={"size": 12, "color": text},
            cliponaxis=False,
            customdata=eff["إجمالي الوقت المهدر (دقيقة)"],
            hovertemplate="<b>%{y}</b><br>الوقت المهدر: %{customdata:.0f} دقيقة<extra></extra>",
        ))
        eff_fig.update_layout(**export_layout(
            title="ساعات العمل مقابل الوقت المهدر",
            height=max(420, 36 * len(eff) + 160),
            xaxis_title="ساعات العمل",
            yaxis_title="",
            xaxis={
                "automargin": True,
                "rangemode": "tozero",
                "range": [0, waste_x * 1.18],
                "title": {"text": "ساعات العمل", "font": {"size": 13}},
            },
            yaxis={"automargin": True},
            margin=dict(t=56, b=95, l=170, r=110),
            showlegend=True,
            legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", title_text=""),
            bargap=0.28,
        ))
        chart_specs.append(("ops", "ساعات العمل مقابل الوقت المهدر", eff_fig, "plot_hours_efficiency"))

    agent_options = sorted(work["_agent_display"].dropna().astype(str).unique().tolist())
    state_options = sorted({
        str(s).strip()
        for s in list(work.get("_activity_state", pd.Series(dtype=object)).dropna().astype(str))
        + list(work.get("_unreachable_state", pd.Series(dtype=object)).dropna().astype(str))
        if str(s).strip() and str(s).strip().lower() not in {"nan", "none", "null", "غير محدد", "أخرى"}
    })

    has_extra_pages = wallet_df is not None and payments_df is not None
    nav_html = ""
    if has_extra_pages:
        nav_html = (
            "<nav id='dash-page-nav'>"
            "<button type='button' class='active-page' onclick=\"showDashPage(this,'page-activity')\">📊 نشاط المحصلين</button>"
            "<button type='button' onclick=\"showDashPage(this,'page-wallet')\">💼 المحفظة</button>"
            "<button type='button' onclick=\"showDashPage(this,'page-payments')\">💰 السداد</button>"
            "<button type='button' onclick=\"showDashPage(this,'page-link')\">🔗 نشاط × سداد</button>"
            "</nav>"
            "<script>function showDashPage(btn,id){"
            "document.querySelectorAll('.dash-page').forEach(function(p){p.classList.remove('active');});"
            "document.getElementById(id).classList.add('active');"
            "document.querySelectorAll('#dash-page-nav button').forEach(function(b){b.classList.remove('active-page');});"
            "btn.classList.add('active-page');"
            "setTimeout(function(){document.querySelectorAll('.js-plotly-plot').forEach(function(div){try{Plotly.Plots.resize(div);}catch(e){}});},80);"
            "}</script>"
        )

    parts = [
        "<!doctype html><html lang='ar' dir='rtl'><head><meta charset='utf-8'>",
        "<meta name='viewport' content='width=device-width, initial-scale=1'>",
        "<title>داشبورد تحليل نشاط المحصلين</title>",
        "<style>",
        f"body{{margin:0;background:{background};color:{text};font-family:Tahoma,'Segoe UI',Arial,sans-serif;line-height:1.65}}",
        "main{max-width:1400px;margin:0 auto;padding:24px 18px 48px}",
        f"header.hero{{background:{surface};border:1px solid {border};border-radius:18px;padding:26px 28px;margin-bottom:18px;text-align:center;box-shadow:0 6px 18px rgba(15,23,42,.04)}}",
        f".eyebrow{{font-size:12px;letter-spacing:1.5px;color:{export_accent};margin-bottom:6px}}",
        f"header.hero h1{{margin:0 0 8px;font-size:26px;color:{text}}}",
        f".meta{{color:{text_dim};font-size:13px}}",
        f".panel{{background:{surface};border:1px solid {border};border-radius:16px;padding:16px 16px 12px;margin-bottom:16px}}",
        f"h2.section-title{{margin:4px 0 12px;text-align:center;font-size:17px;color:{text}}}",
        "#interactive-filters,.filters-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;align-items:end}",
        ".kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin:10px 0 12px}",
        ".charts-grid-2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:28px;margin:0 0 28px;align-items:stretch}",".charts-grid-2 > .panel,.charts-grid-2 > .chart-card{display:flex;flex-direction:column;min-height:100%;margin:0;box-shadow:0 6px 18px rgba(15,23,42,.07)}",".charts-grid-2 .js-plotly-plot,.charts-grid-2 .plotly-graph-div{width:100% !important}",".wallet-charts-row{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:28px;margin:0 0 28px;align-items:stretch}",".wallet-charts-row > .panel{margin:0 !important;min-width:0;box-shadow:0 6px 18px rgba(15,23,42,.07)}","@media (max-width:900px){.wallet-charts-row{grid-template-columns:1fr}}",
        "@media (max-width:900px){.charts-grid-2{grid-template-columns:1fr}}",
        f".filter-field{{display:flex;flex-direction:column;gap:6px;color:{text_dim};font-size:12px}}",
        f".filter-field input,.filter-field select,.filter-field button.multi-trigger{{background:{background};color:{text};border:1px solid {border};border-radius:10px;padding:9px 10px;font-size:13px;text-align:right}}",
        f".multi-menu{{display:none;position:absolute;z-index:200;top:calc(100% + 4px);right:0;left:0;background:#fff;border:1px solid {border};border-radius:10px;padding:8px;box-shadow:0 10px 24px rgba(15,23,42,.14);max-height:240px;overflow:auto}}",
        f".filter-field{{position:relative;z-index:1}}",
        f".panel{{overflow:visible}}",
        f".filters-grid,.charts-grid-2,.charts-grid{{overflow:visible}}",
        f"button.multi-trigger{{cursor:pointer;width:100%;display:flex;justify-content:space-between;align-items:center}}",
        f"#kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:16px 0}}",
        f".kpi{{background:{surface};border:1px solid {border};border-radius:14px;padding:16px 10px;text-align:center;min-height:100px}}",
        f".kpi .label{{color:{text_dim};font-size:13px;margin-bottom:8px}}",
        f".kpi .value{{font-size:24px;font-weight:700}}",
        f"#filter-status{{text-align:center;color:{text_dim};font-size:12px;margin-top:10px}}",
        ".charts-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(480px,1fr));gap:20px;margin:14px 0 24px}",
        f".chart-card{{background:{surface};border:1px solid {border};border-radius:16px;padding:18px 16px 14px;min-width:0;overflow:visible;box-shadow:0 4px 14px rgba(15,23,42,.05)}}",
        ".dash-page{display:none}",
        ".dash-page.active{display:block}",
        f"#dash-page-nav{{display:flex;gap:10px;justify-content:center;flex-wrap:wrap;max-width:1400px;margin:16px auto 0;padding:0 18px}}",
        f"#dash-page-nav button{{background:{surface};color:{text};border:1px solid {border};border-radius:12px;padding:10px 18px;font-size:14px;cursor:pointer;font-family:inherit}}",
        f"#dash-page-nav button.active-page{{background:{export_dark};color:#fff;border-color:{export_dark}}}",
        f".chart-card h3{{margin:6px 8px 4px;text-align:center;font-size:15px;color:{text}}}",
        f".table-wrap{{overflow-x:auto;border:1px solid {border};border-radius:14px;background:{surface}}}",
        f"table.data-table{{width:100%;border-collapse:separate;border-spacing:0;font-size:13px}}",
        f"table.data-table thead th{{position:sticky;top:0;background:{export_dark};color:#FFFFFF;font-weight:600;padding:12px 10px;text-align:center;border-bottom:2px solid {export_mid};white-space:nowrap}}",
        f"table.data-table tbody td{{padding:11px 10px;border-bottom:1px solid {border};text-align:center;white-space:nowrap;color:{text}}}",
        f"table.data-table tbody td:first-child{{text-align:right;font-weight:600;color:{export_dark}}}",
        f"table.data-table tbody tr:nth-child(even){{background:#F3F8F8}}",
        f"table.data-table tbody tr:hover{{background:#E7F1F1}}",
        f"table.data-table tbody tr:last-child td{{border-bottom:none}}",
        f"footer{{color:{text_dim};font-size:12px;text-align:center;margin-top:22px}}",
        f".btn-reset{{background:{export_dark};color:#fff;border:0;border-radius:10px;padding:10px 12px;font-size:13px;cursor:pointer}}",
        f".btn-reset:hover{{background:{export_mid}}}",
        "@media (max-width:900px){.charts-grid{grid-template-columns:1fr}}",
        "</style></head><body>",
        nav_html,
        "<main id='page-activity' class='dash-page active'>",
        "<header class='hero'>",
        "<div class='eyebrow'>ACTIVITY DASHBOARD</div>",
        "<h1>📊 تحليل نشاط المحصلين</h1>",
        f"<div class='meta'>مصدر البيانات: {escape(source_name or 'ملف النشاط')}</div>",
    ]
    if filter_hint:
        parts.append(f"<div class='meta' style='margin-top:8px'>الفلاتر عند التصدير: {escape(filter_hint)}</div>")
    parts.append("</header>")

    # فلاتر
    parts.append("<section class='panel'><h2 class='section-title'>🎚️ فلاتر التقرير</h2>")
    parts.append("<div id='interactive-filters'>")
    # agents
    parts.append("<div class='filter-field' style='position:relative'><span>👤 المحصلون</span>")
    parts.append("<button type='button' class='multi-trigger' data-target='agent-menu'><span id='agent-label'>كل المحصلين</span> ⌄</button>")
    parts.append("<div id='agent-menu' class='multi-menu'><label style='display:block;padding:6px;font-weight:700'><input type='checkbox' class='select-all-agent'> كل المحصلين</label>")
    for value in agent_options:
        parts.append(f"<label style='display:block;padding:6px'><input type='checkbox' class='agent-option' value='{escape(value, quote=True)}'> {escape(value)}</label>")
    parts.append("</div></div>")
    # states
    parts.append("<div class='filter-field' style='position:relative'><span>📊 الحالات</span>")
    parts.append("<button type='button' class='multi-trigger' data-target='state-menu'><span id='state-label'>كل الحالات</span> ⌄</button>")
    parts.append("<div id='state-menu' class='multi-menu'><label style='display:block;padding:6px;font-weight:700'><input type='checkbox' class='select-all-state'> كل الحالات</label>")
    for value in state_options:
        parts.append(f"<label style='display:block;padding:6px'><input type='checkbox' class='state-option' value='{escape(value, quote=True)}'> {escape(value)}</label>")
    parts.append("</div></div>")
    parts.append(
        "<div class='filter-field'><span>🏷️ التصنيف</span>"
        "<select id='filter-class'><option value=''>الكل</option>"
        "<option value='success'>ناجحة</option><option value='failure'>غير ناجحة</option></select></div>"
    )
    parts.append(
        f"<div class='filter-field'><span>📅 من</span>"
        f"<input id='filter-date-from' type='date' value='{export_date_min}' min='{export_date_min}' max='{export_date_max}'></div>"
    )
    parts.append(
        f"<div class='filter-field'><span>📅 إلى</span>"
        f"<input id='filter-date-to' type='date' value='{export_date_max}' min='{export_date_min}' max='{export_date_max}'></div>"
    )
    parts.append("<div class='filter-field'><span>&nbsp;</span><button id='reset-filters' class='btn-reset' type='button'>↺ إعادة ضبط</button></div>")
    parts.append("</div><div id='filter-status'>عرض كل البيانات</div></section>")

    # KPI
    parts.append("<section id='kpi-grid'>")
    parts.append(metric_card("kpi-agents", "👥 عدد المحصلين", f"{agent_count:,}", text))
    parts.append(metric_card("kpi-total", "📞 إجمالي المكالمات", f"{total:,}", text))
    parts.append(metric_card("kpi-success", "✅ المكالمات الناجحة", f"{success:,}", export_success))
    parts.append(metric_card("kpi-rate", "📈 نسبة النجاح", f"{rate:.1f}%", export_accent))
    parts.append(metric_card("kpi-wasted", "⏱️ الوقت المهدر", f"{wasted:,.1f} د", export_warn))
    parts.append("</section>")

    section_titles = {
        "main": "النتائج والنشاط",
        "rank": "مقارنة أداء المحصلين",
        "states": "تفاصيل حالات المتابعة",
        "ops": "كفاءة التشغيل (ساعات العمل والوقت المهدر)",
    }
    include_js = True
    current_section = None
    for section, heading, fig, plot_id in chart_specs:
        if section != current_section:
            if current_section is not None:
                parts.append("</div>")  # close charts-grid
            current_section = section
            parts.append(f"<h2 class='section-title'>{section_titles.get(section, section)}</h2>")
            parts.append("<div class='charts-grid'>")
        parts.append(f"<article class='chart-card'><h3>{escape(heading)}</h3>")
        parts.append(pio.to_html(
            fig, full_html=False, include_plotlyjs=("cdn" if include_js else False),
            config={"displayModeBar": False, "responsive": True},
            div_id=plot_id, default_width="100%",
            default_height=f"{int(fig.layout.height or 420)}px",
        ))
        parts.append("</article>")
        include_js = False
    if current_section is not None:
        parts.append("</div>")

    # table
    if not agent_table.empty:
        columns = [
            "المحصّل", "إجمالي المكالمات", "المكالمات الناجحة", "نسبة النجاح (%)",
            "واعد بالسداد", "إجمالي لا يرد", "أيام النشاط",
            "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)",
        ]
        columns = [c for c in columns if c in agent_table.columns]
        parts.append("<section class='panel'><h2 class='section-title'>📋 ملخص أداء كل محصل</h2><div class='table-wrap'><table class='data-table'><thead><tr>")
        for column in columns:
            parts.append(f"<th>{escape(column)}</th>")
        parts.append("</tr></thead><tbody>")
        for _, row in agent_table.sort_values("إجمالي المكالمات", ascending=False).iterrows():
            parts.append("<tr>")
            for column in columns:
                value = row[column]
                if isinstance(value, float):
                    value = f"{value:,.2f}"
                parts.append(f"<td>{escape(str(value))}</td>")
            parts.append("</tr>")
        parts.append("</tbody></table></div></section>")

    interactive_js = """
<script>
const activityData = __ACTIVITY_DATA__;
const agentColors = __AGENT_COLORS__;
const stateColors = __STATE_COLORS__;
const positiveStateColors = __POSITIVE_STATE_COLORS__;
const agentHoursMap = __AGENT_HOURS__;
const donutPlot = document.getElementById('plot_donut');
const hourlyPlot = document.getElementById('plot_hourly');
const leaderboardPlot = document.getElementById('plot_leaderboard');
const statePlot = document.getElementById('plot_no_answer');
const positivePlot = document.getElementById('plot_positive');
const hoursEffPlot = document.getElementById('plot_hours_efficiency');
const fmt = n => Number(n || 0).toLocaleString('en-US');
function setKpi(id, value) {
  const el = document.querySelector('#' + id + ' [data-role="value"]');
  if (el) el.textContent = value;
}
function updateMultiLabels() {
  const agent = [...document.querySelectorAll('.agent-option:checked')].map(o => o.value);
  const state = [...document.querySelectorAll('.state-option:checked')].map(o => o.value);
  const al = document.getElementById('agent-label');
  const sl = document.getElementById('state-label');
  if (al) al.textContent = agent.length ? `${agent.length} محصل محدد` : 'كل المحصلين';
  if (sl) sl.textContent = state.length ? `${state.length} حالة محددة` : 'كل الحالات';
}
function selectedRows() {
  const values = cls => [...document.querySelectorAll('.' + cls + ':checked')].map(o => o.value).filter(Boolean);
  const agent = values('agent-option');
  const state = values('state-option');
  const cls = (document.getElementById('filter-class') || {}).value || '';
  const from = (document.getElementById('filter-date-from') || {}).value || '';
  const to = (document.getElementById('filter-date-to') || {}).value || '';
  return activityData.filter(row => {
    if (agent.length && !agent.includes(row.agent)) return false;
    if (state.length && !(state.includes(row.state) || state.includes(row.unreachable))) return false;
    if (cls === 'success' && !row.success) return false;
    if (cls === 'failure' && row.success) return false;
    if (from && row.time && row.time.slice(0,10) < from) return false;
    if (to && row.time && row.time.slice(0,10) > to) return false;
    return true;
  });
}
function refreshDashboard() {
  const rows = selectedRows();
  const agents = [...new Set(rows.map(r => r.agent))].sort();
  const success = rows.filter(r => r.success).length;
  const rate = rows.length ? success / rows.length * 100 : 0;
  const wasted = rows.reduce((s,r) => s + (r.wasted || 0), 0);
  setKpi('kpi-agents', fmt(agents.length));
  setKpi('kpi-total', fmt(rows.length));
  setKpi('kpi-success', fmt(success));
  setKpi('kpi-rate', rate.toFixed(1) + '%');
  setKpi('kpi-wasted', Number(wasted).toLocaleString('en-US', {maximumFractionDigits:1}) + ' د');
  const status = document.getElementById('filter-status');
  if (status) status.textContent = `عرض ${fmt(rows.length)} مكالمة من أصل ${fmt(activityData.length)} — ${fmt(agents.length)} محصل`;

  if (hourlyPlot) {
    const present = rows.filter(r => r.time).map(r => new Date(r.time).getHours());
    let hourMin = 0, hourMax = 23;
    if (present.length) { hourMin = Math.min(...present); hourMax = Math.max(...present); }
    const hours = [];
    for (let h = hourMin; h <= hourMax; h++) hours.push(h);
    const hourCounts = hours.map(h => rows.filter(r => r.time && new Date(r.time).getHours()===h).length);
    Plotly.react(hourlyPlot, [{
      type:'bar', x:hours, y:hourCounts, marker:{color:'__SUCCESS__'},
      text:hourCounts, texttemplate:'%{text:,}', textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>الساعة %{x}:00</b><br>عدد المكالمات: %{y:,}<extra></extra>'
    }], {
      ...(hourlyPlot.layout || {}),
      xaxis:{
        ...(hourlyPlot.layout && hourlyPlot.layout.xaxis || {}),
        dtick:1, tickvals:hours, ticktext:hours.map(h => h + ':00'),
        range:[hourMin-0.5, hourMax+0.5], title:'ساعة اليوم', automargin:true, showgrid:false
      },
      yaxis:{...(hourlyPlot.layout && hourlyPlot.layout.yaxis || {}), title:'عدد المكالمات', rangemode:'tozero', automargin:true},
      showlegend:false
    });
  }

  if (donutPlot) {
    Plotly.react(donutPlot, [{
      type:'pie', labels:['ناجحة','غير ناجحة'], values:[success, Math.max(rows.length-success,0)],
      hole:0.62, marker:{colors:['__SUCCESS__','__FAIL__']},
      textinfo:'label+value+percent'
    }], donutPlot.layout || {});
  }

  if (leaderboardPlot) {
    const totals = agents.map(a => rows.filter(r => r.agent===a).length);
    const rates = agents.map(a => {
      const rs = rows.filter(r => r.agent===a);
      return rs.length ? rs.filter(r => r.success).length / rs.length * 100 : 0;
    });
    const order = agents.map((a,i)=>({a,t:totals[i],r:rates[i]})).sort((x,y)=>x.t-y.t);
    Plotly.react(leaderboardPlot, [{
      type:'bar', orientation:'h', y:order.map(o=>o.a), x:order.map(o=>o.t),
      marker:{color:'__ACCENT__'},
      text:order.map(o => `${o.t.toLocaleString('en-US')}  |  ${o.r.toFixed(0)}%`),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{y}</b><br>المكالمات: %{x:,}<extra></extra>'
    }], {
      ...(leaderboardPlot.layout || {}),
      showlegend:false,
      margin:{...(leaderboardPlot.layout && leaderboardPlot.layout.margin || {}), l:170, r:90}
    });
  }

  if (statePlot) {
    const stateNames = Object.keys(stateColors);
    // عدم الوصول من حقل unreachable (الإفادة)
    const stateTraces = stateNames.map(state => ({
      type:'bar', orientation:'h', name:state,
      y:agents,
      x:agents.map(a => rows.filter(r => r.agent===a && r.unreachable===state).length),
      marker:{color: stateColors[state] || '#A8B8BC'},
      text:agents.map(a => { const v = rows.filter(r => r.agent===a && r.unreachable===state).length; return v > 0 ? v : ''; }),
      textposition:'inside', insidetextanchor:'middle', textfont:{size:11},
      hovertemplate:'<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>'
    }));
    const stateTotals = agents.map(a => rows.filter(r => r.agent===a && stateNames.includes(r.unreachable)).length);
    const agentsWithState = agents.filter((a,i) => stateTotals[i] > 0);
    const totalsFiltered = stateTotals.filter(v => v > 0);
    const tracesFiltered = stateNames.map(state => ({
      type:'bar', orientation:'h', name:state,
      y:agentsWithState,
      x:agentsWithState.map(a => rows.filter(r => r.agent===a && r.unreachable===state).length),
      marker:{color: stateColors[state] || '#A8B8BC'},
      text:agentsWithState.map(a => { const v = rows.filter(r => r.agent===a && r.unreachable===state).length; return v > 0 ? v : ''; }),
      textposition:'inside', insidetextanchor:'middle', textfont:{size:11},
      hovertemplate:'<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>'
    }));
    tracesFiltered.push({
      type:'scatter', mode:'text', y:agentsWithState, x:totalsFiltered,
      text:totalsFiltered.map(v => String(v)),
      textposition:'middle right', textfont:{size:12, color:'#1F2937'},
      showlegend:false, hoverinfo:'skip', cliponaxis:false
    });
    Plotly.react(statePlot, tracesFiltered, {...(statePlot.layout || {}), barmode:'stack', margin:{...(statePlot.layout && statePlot.layout.margin || {}), r:55}});
  }

  if (positivePlot) {
    const posNames = Object.keys(positiveStateColors);
    // الوعد والسداد من حقل state (Sub State)
    const posTotals = agents.map(a => rows.filter(r => r.agent===a && posNames.includes(r.state)).length);
    const agentsWithPos = agents.filter((a,i) => posTotals[i] > 0);
    const totalsFiltered = posTotals.filter(v => v > 0);
    const positiveTraces = posNames.map(state => ({
      type:'bar', orientation:'h', name:state,
      y:agentsWithPos,
      x:agentsWithPos.map(a => rows.filter(r => r.agent===a && r.state===state).length),
      marker:{color: positiveStateColors[state] || '#6A9A9D'},
      text:agentsWithPos.map(a => { const v = rows.filter(r => r.agent===a && r.state===state).length; return v > 0 ? v : ''; }),
      textposition:'inside', insidetextanchor:'middle', textfont:{size:11},
      hovertemplate:'<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>'
    }));
    positiveTraces.push({
      type:'scatter', mode:'text', y:agentsWithPos, x:totalsFiltered,
      text:totalsFiltered.map(v => String(v)),
      textposition:'middle right', textfont:{size:12, color:'#1F2937'},
      showlegend:false, hoverinfo:'skip', cliponaxis:false
    });
    Plotly.react(positivePlot, positiveTraces, {...(positivePlot.layout || {}), barmode:'stack', margin:{...(positivePlot.layout && positivePlot.layout.margin || {}), r:55}});
  }

  if (hoursEffPlot) {
    const agentWaste = agents.map(a => rows.filter(r => r.agent===a).reduce((s,r)=>s+(r.wasted||0),0));
    const agentHours = agents.map(a => Number((agentHoursMap && agentHoursMap[a]) || 0));
    const order = agents.map((a,i)=>({a, h:agentHours[i], w:agentWaste[i]}))
      .filter(o => o.h > 0 || o.w > 0)
      .sort((x,y)=>x.h-y.h);
    const maxH = Math.max(...order.map(o=>o.h), 1);
    const wasteX = Math.max(maxH * 1.22, maxH + 1.2, 1);
    Plotly.react(hoursEffPlot, [
      {
        type:'bar', orientation:'h', name:'ساعات العمل',
        y:order.map(o=>o.a), x:order.map(o=>o.h),
        marker:{color:'__SUCCESS__'},
        text:order.map(o => o.h.toFixed(1)),
        textposition:'inside', insidetextanchor:'middle',
        textfont:{size:13, color:'#FFFFFF'},
        cliponaxis:false,
        hovertemplate:'<b>%{y}</b><br>ساعات العمل: %{x:.1f}<extra></extra>'
      },
      {
        type:'scatter', mode:'markers+text', name:'الوقت المهدر (دقيقة)',
        y:order.map(o=>o.a), x:order.map(() => wasteX),
        marker:{color:'__WARN__', size:11, symbol:'diamond', line:{width:1, color:'#FFFFFF'}},
        text:order.map(o => String(Math.round(o.w))),
        textposition:'middle right',
        textfont:{size:12, color:'#1F2937'},
        cliponaxis:false,
        customdata:order.map(o => o.w),
        hovertemplate:'<b>%{y}</b><br>الوقت المهدر: %{customdata:.0f} دقيقة<extra></extra>'
      }
    ], {
      ...(hoursEffPlot.layout || {}),
      title:{...(hoursEffPlot.layout && hoursEffPlot.layout.title || {}), text:'ساعات العمل مقابل الوقت المهدر'},
      xaxis:{
        ...(hoursEffPlot.layout && hoursEffPlot.layout.xaxis || {}),
        title:'ساعات العمل', rangemode:'tozero', range:[0, wasteX*1.18], automargin:true
      },
      showlegend:true,
      legend:{orientation:'h', y:-0.18, x:0.5, xanchor:'center'},
      margin:{...(hoursEffPlot.layout && hoursEffPlot.layout.margin || {}), r:110, b:95}
    });
  }
}

document.addEventListener('click', event => {
  const trigger = event.target.closest && event.target.closest('.multi-trigger');
  if (trigger) {
    event.preventDefault();
    event.stopPropagation();
    const menu = document.getElementById(trigger.dataset.target);
    document.querySelectorAll('.multi-menu').forEach(other => {
      if (other !== menu) other.style.display = 'none';
    });
    if (menu) {
      const open = menu.style.display === 'block';
      menu.style.display = open ? 'none' : 'block';
    }
    return;
  }
  if (event.target.closest && event.target.closest('.multi-menu')) {
    event.stopPropagation();
    return;
  }
  document.querySelectorAll('.multi-menu').forEach(menu => { menu.style.display = 'none'; });
});
document.querySelectorAll('.agent-option,.state-option').forEach(option => option.addEventListener('change', () => { updateMultiLabels(); refreshDashboard(); }));
document.querySelector('.select-all-agent')?.addEventListener('change', event => {
  document.querySelectorAll('.agent-option').forEach(option => option.checked = event.target.checked);
  updateMultiLabels(); refreshDashboard();
});
document.querySelector('.select-all-state')?.addEventListener('change', event => {
  document.querySelectorAll('.state-option').forEach(option => option.checked = event.target.checked);
  updateMultiLabels(); refreshDashboard();
});
['filter-class','filter-date-from','filter-date-to'].forEach(id => document.getElementById(id)?.addEventListener('change', refreshDashboard));
document.getElementById('reset-filters')?.addEventListener('click', () => {
  document.querySelectorAll('.agent-option,.state-option,.select-all-agent,.select-all-state').forEach(option => option.checked = false);
  const cls = document.getElementById('filter-class'); if (cls) cls.value = '';
  const from = document.getElementById('filter-date-from'); if (from) from.value = '__DATE_MIN__';
  const to = document.getElementById('filter-date-to'); if (to) to.value = '__DATE_MAX__';
  updateMultiLabels(); refreshDashboard();
});
updateMultiLabels();
refreshDashboard();
function _resizeAllPlots() {
  document.querySelectorAll('.js-plotly-plot').forEach(div => { try { Plotly.Plots.resize(div); } catch (e) {} });
}
window.addEventListener('load', () => { _resizeAllPlots(); setTimeout(_resizeAllPlots, 200); setTimeout(_resizeAllPlots, 600); });
window.addEventListener('resize', _resizeAllPlots);
setTimeout(_resizeAllPlots, 100);
</script>
"""
    interactive_js = (
        interactive_js
        .replace("__ACTIVITY_DATA__", raw_records_json)
        .replace("__AGENT_COLORS__", agent_color_json)
        .replace("__STATE_COLORS__", state_color_json)
        .replace("__POSITIVE_STATE_COLORS__", positive_state_color_json)
        .replace("__AGENT_HOURS__", agent_hours_json)
        .replace("__ACCENT__", json.dumps(export_accent))
        .replace("__SUCCESS__", json.dumps(export_success))
        .replace("__FAIL__", json.dumps(export_fail))
        .replace("__WARN__", json.dumps(export_warn))
        .replace("__DATE_MIN__", export_date_min)
        .replace("__DATE_MAX__", export_date_max)
    )
    parts.append(interactive_js)
    parts.append("<footer>تم إنشاء التقرير من لوحة تحليل نشاط المحصلين</footer></main>")

    if has_extra_pages:
        parts.append(_build_wallet_page_html(wallet_df))
        parts.append(_build_payments_page_html(payments_df))
        parts.append(_build_link_page_html(df, sales_col, class_col, payments_df))

    parts.append("</body></html>")
    return "".join(parts)


# ==========================================================
# تويب 1: التصنيف
# ==========================================================

# ==========================================================
# أدوات اختيار الشركة / الفترة / التجميع اليومي
# ==========================================================

COMPANIES = ["الوطنية للتأمين", "تري للتأمين"]
STATUS_CANDIDATES = [
    "Main State", "Final State", "Status", "State", "Call Status",
    "الحالة الرئيسية", "الحالة النهائية", "الحالة", "حالة المكالمة",
]

CLOSED_WORDS = ["مغلق", "مغلقه", "مغلقه", "closed", "غير متاح"]
def init_activity_state():
    defaults = {
        "selected_company": None,
        "selected_period": None,
        "period_1_start": dt_time(9, 0),
        "period_1_end": dt_time(12, 30),
        "period_2_start": dt_time(13, 0),
        "period_2_end": dt_time(17, 0),
        "period_1_has_break": False,
        "period_2_has_break": False,
        "period_1_break_start": dt_time(11, 0),
        "period_1_break_end": dt_time(11, 15),
        "period_2_break_start": dt_time(15, 0),
        "period_2_break_end": dt_time(15, 15),
        "period_1_break_duration": 15,
        "period_2_break_duration": 15,
        "daily_has_break": False,
        "daily_break_start": dt_time(13, 0),
        "daily_break_end": dt_time(13, 15),
        "daily_break_duration": 15,
        "monthly_has_break": False,
        "monthly_break_start": dt_time(13, 0),
        "monthly_break_end": dt_time(13, 15),
        "monthly_break_duration": 15,
        "selected_week": "week_1",
        "period_results": {},
        "weekly_week_1_has_break": False,
        "weekly_week_1_break_start": dt_time(13, 0),
        "weekly_week_1_break_end": dt_time(13, 15),
        "weekly_week_1_break_duration": 15,
        "weekly_week_1_result": None,
        "weekly_week_2_has_break": False,
        "weekly_week_2_break_start": dt_time(13, 0),
        "weekly_week_2_break_end": dt_time(13, 15),
        "weekly_week_2_break_duration": 15,
        "weekly_week_2_result": None,
        "weekly_week_3_has_break": False,
        "weekly_week_3_break_start": dt_time(13, 0),
        "weekly_week_3_break_end": dt_time(13, 15),
        "weekly_week_3_break_duration": 15,
        "weekly_week_3_result": None,
        "weekly_week_4_has_break": False,
        "weekly_week_4_break_start": dt_time(13, 0),
        "weekly_week_4_break_end": dt_time(13, 15),
        "weekly_week_4_break_duration": 15,
        "weekly_week_4_result": None,
        "daily_result": None,
        "monthly_result": None,
        "dashboard_result": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def render_company_selector():
    st.subheader("🏢 اختر شركة التصنيف")
    c1, c2 = st.columns(2)
    for col, company_name in zip((c1, c2), COMPANIES):
        with col:
            selected = st.session_state.get("selected_company") == company_name
            button_label = f"✓ {company_name}" if selected else company_name
            if st.button(
                button_label,
                key=f"company_{company_name}",
                use_container_width=True,
                type="primary" if selected else "secondary",
            ):
                st.session_state["selected_company"] = company_name
                st.session_state["selected_period"] = None
                st.rerun()
    company = st.session_state.get("selected_company")
    if company:
        st.success(f"الشركة المختارة: {company}")


def render_period_selector():
    if not st.session_state.get("selected_company"):
        return
    st.subheader("🕒 اختر فترة النشاط")
    cols = st.columns(5)
    buttons = [("الفترة الأولى", "🟢", "period_1"), ("الفترة الثانية", "🔵", "period_2"), ("التجميع اليومي", "📊", "daily"), ("التجميع الأسبوعي", "📅", "weekly"), ("التجميع الشهري", "🗓️", "monthly")]
    for col, (title, icon, key) in zip(cols, buttons):
        with col:
            selected = st.session_state.get("selected_period") == key
            if st.button(f"{icon} {title}", key=f"period_btn_{key}", use_container_width=True, type="primary" if selected else "secondary"):
                st.session_state["selected_period"] = key
                st.rerun()
    selected_period = st.session_state.get("selected_period")
    if not selected_period:
        return
    company = st.session_state["selected_company"]
    if selected_period == "period_1":
        render_period_settings("period_1", "الفترة الأولى")
    elif selected_period == "period_2":
        render_period_settings("period_2", "الفترة الثانية")
    elif selected_period == "daily":
        st.subheader("📊 التجميع اليومي")
        st.caption(f"{company} · من بداية اليوم إلى نهايته")
        render_aggregate_tab("daily", "التجميع اليومي")
    elif selected_period == "weekly":
        st.subheader("📅 التجميع الأسبوعي")
        st.caption(f"{company} · اختر الأسبوع لرفع الملف")
        weeks = [("الأسبوع الأول", "week_1"), ("الأسبوع الثاني", "week_2"), ("الأسبوع الثالث", "week_3"), ("الأسبوع الرابع", "week_4")]
        week_cols = st.columns(4)
        for col, (title, key) in zip(week_cols, weeks):
            with col:
                if st.button(title, key=f"btn_{key}", use_container_width=True, type="primary" if st.session_state.get("selected_week") == key else "secondary"):
                    st.session_state["selected_week"] = key
                    st.rerun()
        current = st.session_state.get("selected_week", "week_1")
        title = next(label for label, key in weeks if key == current)
        render_aggregate_tab(f"weekly_{current}", f"التجميع الأسبوعي ({title})")
    elif selected_period == "monthly":
        st.subheader("🗓️ التجميع الشهري")
        st.caption(f"{company} · نشاط الشهر بالكامل")
        render_aggregate_tab("monthly", "التجميع الشهري")


def render_period_upload_and_classify(period_key: str, period_title: str):
    uploaded_file = st.file_uploader(
        f"📂 ارفع ملف {period_title} (CSV أو Excel)",
        type=["csv", "xlsx", "xls"],
        key=f"upload_{period_key}",
        on_change=sync_file_cache,
        args=(f"upload_{period_key}", f"period_upload:{period_key}", (f"period_results:{period_key}",)),
    )
    if uploaded_file is not None:
        st.caption(f"الملف المختار: {uploaded_file.name}")
        classify_period_file(uploaded_file, period_key)
    else:
        _show_period_results_from_cache(period_key)


def _render_break_switch(label: str, has_break_key: str):
    st.session_state[has_break_key] = st.toggle(label, value=st.session_state[has_break_key], key=f"{has_break_key}_switch")
    st.caption("✅ يوجد استراحة" if st.session_state[has_break_key] else "لا يوجد استراحة")


def render_period_settings(period_key: str, period_title: str):
    company = st.session_state["selected_company"]
    start_key, end_key = f"{period_key}_start", f"{period_key}_end"
    has_break_key = f"{period_key}_has_break"
    break_start_key, break_end_key = f"{period_key}_break_start", f"{period_key}_break_end"
    st.subheader(f"⏱️ {period_title}")
    st.caption(f"نشاط {company} · حدّد الميعاد قبل رفع ملف الفترة")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state[start_key] = st.time_input(f"بداية {period_title}", value=st.session_state[start_key], key=f"{start_key}_input")
    with c2:
        st.session_state[end_key] = st.time_input(f"نهاية {period_title}", value=st.session_state[end_key], key=f"{end_key}_input")
    if st.session_state[start_key] >= st.session_state[end_key]:
        st.warning("يجب أن تسبق بداية الفترة نهايتها.")
    _render_break_switch(f"هل يوجد استراحة في {period_title}؟", has_break_key)
    if st.session_state[has_break_key]:
        b1, b2 = st.columns(2)
        with b1:
            st.session_state[break_start_key] = st.time_input("🕐 بداية الاستراحة", value=st.session_state[break_start_key], key=f"{break_start_key}_input")
        with b2:
            duration_key = f"{period_key}_break_duration"
            current_duration = max(5, int(st.session_state[duration_key]))
            st.session_state[duration_key] = st.slider("⏳ مدة الاستراحة (دقيقة)", 5, 120, current_duration, 5, key=f"{duration_key}_slider")
            start_minutes = st.session_state[break_start_key].hour * 60 + st.session_state[break_start_key].minute
            end_minutes = start_minutes + st.session_state[duration_key]
            st.session_state[break_end_key] = dt_time(end_minutes // 60, end_minutes % 60)
        st.info(f"☕ الاستراحة من {st.session_state[break_start_key]:%H:%M} إلى {st.session_state[break_end_key]:%H:%M} ({st.session_state[duration_key]} دقيقة)")
    st.info(f"النشاط: {period_title} {company} من {st.session_state[start_key]:%H:%M} إلى {st.session_state[end_key]:%H:%M}")
    render_period_upload_and_classify(period_key, period_title)


@st.cache_data(show_spinner=False)
def read_uploaded_dataframe(uploaded_file):
    """قراءة الملف مع كاش على محتوى البايتات — مش بيتعاد إلا لو محتوى الملف اتغير.

    Streamlit بيعمل rerun كامل مع أي أكشن، والدالة دي كانت بتتدعي مع كل rerun
    وكانت بتقرأ الإكسيل من الصفر كل مرة. الـ @st.cache_data بيخزن النتيجة
    ويعيدها فوراً طالما محتوى الملف (البايتات + الاسم) نفسهما.
    """
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    return pd.read_excel(io.BytesIO(data))


def _xml_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _sanitize_table_name(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not cleaned or not cleaned[0].isalpha():
        cleaned = "T_" + cleaned
    return cleaned[:60]


def _build_base_workbook_bytes(df: pd.DataFrame, data_sheet_name: str, table_name: str, pivot_sheet_name: str = "Pivot Table") -> bytes:
    """بيبني ملف إكسيل بـ openpyxl فيه شيت البيانات كـ Excel Table (ListObject) رسمي + شيت فاضي للـ Pivot."""
    wb = Workbook()
    ws = wb.active
    ws.title = data_sheet_name
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    last_row = max(ws.max_row, 2)
    last_col_letter = get_column_letter(len(df.columns))
    tab = Table(displayName=table_name, ref=f"A1:{last_col_letter}{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)
    wb.create_sheet(pivot_sheet_name)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _inject_native_pivot_table(
    xlsx_bytes,
    df_columns,
    table_name,
    pivot_sheet_name,
    row_field,
    data_field,
    data_field_label,
    col_field=None,
    subtotal="count",
    data_field2=None,
    data_field2_label=None,
    subtotal2="count",
):
    """بيحقن Pivot Table حقيقي (native، قابل للتحديث والسحب والإفلات) جوه ملف الإكسيل — نفس اللي بتعمله
    يدوي في إكسيل بـ Insert > PivotTable. بيتحدث تلقائي من الـ Excel Table لما تفتح الملف."""
    df_columns = list(df_columns)
    row_idx = df_columns.index(row_field)
    col_idx = df_columns.index(col_field) if col_field else None
    data_idx = df_columns.index(data_field)
    data_idx2 = df_columns.index(data_field2) if data_field2 else None
    n_fields = len(df_columns)
    data_indices = {data_idx}
    if data_idx2 is not None:
        data_indices.add(data_idx2)

    zin = zipfile.ZipFile(BytesIO(xlsx_bytes), "r")
    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    wb_xml = data["xl/workbook.xml"].decode("utf-8")
    wb_rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")

    name_to_rid = {}
    for tag in re.findall(r"<sheet [^>]*/>", wb_xml):
        m_name = re.search(r'name="([^"]+)"', tag)
        m_rid = re.search(r'r:id="([^"]+)"', tag)
        if m_name and m_rid:
            name_to_rid[m_name.group(1)] = m_rid.group(1)
    pivot_rid = name_to_rid[pivot_sheet_name]

    rid_to_target = {}
    for tag in re.findall(r"<Relationship [^>]*/>", wb_rels):
        m_id = re.search(r'Id="([^"]+)"', tag)
        m_target = re.search(r'Target="([^"]+)"', tag)
        if m_id and m_target:
            rid_to_target[m_id.group(1)] = m_target.group(1)
    pivot_sheet_target = rid_to_target[pivot_rid]
    pivot_sheet_path = pivot_sheet_target.lstrip("/")
    if not pivot_sheet_path.startswith("xl/"):
        pivot_sheet_path = "xl/" + pivot_sheet_path
    sheet_file = pivot_sheet_path.split("/")[-1]

    existing_ids = [int(re.sub(r"\D", "", rid)) for rid in re.findall(r'Id="(rId\d+)"', wb_rels)]
    cache_rid = f"rId{max(existing_ids) + 1}"

    wb_xml = wb_xml.replace(
        "</workbook>",
        f'<pivotCaches><pivotCache cacheId="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{cache_rid}"/></pivotCaches></workbook>',
    )
    data["xl/workbook.xml"] = wb_xml.encode("utf-8")

    new_rel = f'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml" Id="{cache_rid}"/>'
    wb_rels = wb_rels.replace("</Relationships>", new_rel + "</Relationships>")
    data["xl/_rels/workbook.xml.rels"] = wb_rels.encode("utf-8")

    data[f"xl/worksheets/_rels/{sheet_file}.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    ct = data["[Content_Types].xml"].decode("utf-8")
    overrides = (
        '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
        '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
        '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'
    )
    data["[Content_Types].xml"] = ct.replace("</Types>", overrides + "</Types>").encode("utf-8")

    cache_fields_xml = "".join(
        f'<cacheField name="{_xml_escape(col)}" numFmtId="0"><sharedItems/></cacheField>' for col in df_columns
    )
    data["xl/pivotCache/pivotCacheDefinition1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" '
        'refreshOnLoad="1" refreshedBy="Claude" refreshedDate="45900" createdVersion="6" refreshedVersion="6" '
        f'minRefreshableVersion="3" recordCount="0">'
        f'<cacheSource type="worksheet"><worksheetSource name="{table_name}"/></cacheSource>'
        f'<cacheFields count="{n_fields}">{cache_fields_xml}</cacheFields>'
        "</pivotCacheDefinition>"
    ).encode("utf-8")
    data["xl/pivotCache/pivotCacheRecords1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="0"/>'
    ).encode("utf-8")
    data["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    pivot_fields_parts = []
    for i in range(n_fields):
        if i == row_idx:
            pivot_fields_parts.append(
                '<pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
            )
        elif col_idx is not None and i == col_idx:
            pivot_fields_parts.append(
                '<pivotField axis="axisCol" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
            )
        elif i in data_indices:
            pivot_fields_parts.append('<pivotField dataField="1" showAll="0"/>')
        else:
            pivot_fields_parts.append('<pivotField showAll="0"/>')
    pivot_fields_xml = "".join(pivot_fields_parts)

    data_fields_parts = [
        f'<dataField name="{_xml_escape(data_field_label)}" fld="{data_idx}" subtotal="{subtotal}" baseField="0" baseItem="0"/>'
    ]
    if data_idx2 is not None:
        label2 = data_field2_label or f"عدد {data_field2}"
        data_fields_parts.append(
            f'<dataField name="{_xml_escape(label2)}" fld="{data_idx2}" subtotal="{subtotal2}" baseField="0" baseItem="0"/>'
        )
    n_data_fields = len(data_fields_parts)
    data_fields_xml = (
        f'<dataFields count="{n_data_fields}">'
        + "".join(data_fields_parts)
        + "</dataFields>"
    )

    # مع أكتر من data field لازم نضيف محور Values (x="-2") على الأعمدة
    # عشان Excel ما يشيلش الـ PivotTable أثناء الإصلاح.
    if col_idx is not None:
        col_fields_xml = (
            f'<colFields count="1"><field x="{col_idx}"/></colFields>'
            '<colItems count="1"><i><x/></i></colItems>'
        )
    elif n_data_fields > 1:
        col_items = "".join(f'<i><x v="{i}"/></i>' if i else "<i><x/></i>" for i in range(n_data_fields))
        # الصيغة الشائعة في OOXML لـ Values على الأعمدة:
        # <colFields count="1"><field x="-2"/></colFields>
        col_items = "".join(
            ("<i><x/></i>" if i == 0 else f'<i><x i="{i}"/></i>') for i in range(n_data_fields)
        )
        col_fields_xml = (
            '<colFields count="1"><field x="-2"/></colFields>'
            f'<colItems count="{n_data_fields}">{col_items}</colItems>'
        )
    else:
        col_fields_xml = ""

    data["xl/pivotTables/pivotTable1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
        'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" '
        'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" '
        'indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
        '<location ref="A3:D30" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'
        f'<pivotFields count="{n_fields}">{pivot_fields_xml}</pivotFields>'
        f'<rowFields count="1"><field x="{row_idx}"/></rowFields>'
        '<rowItems count="1"><i><x/></i></rowItems>'
        f"{col_fields_xml}"
        f"{data_fields_xml}"
        '<pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        "</pivotTableDefinition>"
    ).encode("utf-8")
    data["xl/pivotTables/_rels/pivotTable1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    out_buf = BytesIO()
    zout = zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED)
    for n, content in data.items():
        zout.writestr(n, content)
    zout.close()
    return out_buf.getvalue()


def build_excel_with_native_pivot(result_df: pd.DataFrame, data_sheet_name: str, key_prefix: str):
    """بيبني ملف إكسيل فيه شيت البيانات (كـ Excel Table) + شيت Pivot Table حقيقي (native) —
    المحصل في الصفوف، ومجموع عمود التصنيف (SUM) + عدد Customer Account Number (COUNT) في القيم.
    بيرجع (bytes, تم إضافة بيفوت ولا لأ)."""
    collected_col = find_column(result_df, COLLECTED_BY_CANDIDATES)
    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in result_df.columns else None
    account_col = find_column(result_df, ACCOUNT_NUMBER_CANDIDATES)

    table_name = _sanitize_table_name(f"tbl_{key_prefix}")
    base_bytes = _build_base_workbook_bytes(result_df, data_sheet_name, table_name)

    if not (collected_col and class_col):
        return base_bytes, False

    kwargs = dict(
        xlsx_bytes=base_bytes,
        df_columns=result_df.columns,
        table_name=table_name,
        pivot_sheet_name="Pivot Table",
        row_field=collected_col,
        data_field=class_col,
        data_field_label="مجموع التصنيف",
        subtotal="sum",
    )
    if account_col:
        kwargs.update(
            data_field2=account_col,
            data_field2_label="عدد Customer Account Number",
            subtotal2="count",
        )

    final_bytes = _inject_native_pivot_table(**kwargs)
    return final_bytes, True


def render_pivot_section(df: pd.DataFrame, key_prefix: str):
    """معاينة سريعة جوه السيستم بس (نفس منطق الـ Pivot اللي هيتضاف حقيقي في ملف الإكسيل):
    الصفوف = المحصل، القيم = مجموع (SUM) عمود التصنيف + عدد Customer Account Number."""
    if df is None or df.empty:
        return

    collected_col = find_column(df, COLLECTED_BY_CANDIDATES)
    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    account_col = find_column(df, ACCOUNT_NUMBER_CANDIDATES)

    with st.expander(
        "📊 معاينة الـ Pivot Table (هتلاقي النسخة الحقيقية القابلة للتعديل جوه ملف الإكسيل بعد التحميل)",
        expanded=False,
    ):
        missing = [
            label
            for label, col in [
                ("المحصل (Created by)", collected_col),
                ("التصنيف", class_col),
            ]
            if col is None
        ]
        if missing:
            st.warning("تعذر إنشاء الـ Pivot — الأعمدة دي مش موجودة في الملف: " + "، ".join(missing))
            return

        agg = {class_col: "sum"}
        if account_col:
            agg[account_col] = "count"

        pivot_df = pd.pivot_table(
            df,
            index=collected_col,
            values=list(agg.keys()),
            aggfunc=agg,
            fill_value=0,
            margins=True,
            margins_name="الإجمالي",
        )
        if isinstance(pivot_df.columns, pd.MultiIndex):
            pivot_df.columns = [
                "مجموع التصنيف" if c[0] == class_col else "عدد Customer Account Number"
                for c in pivot_df.columns
            ]
        else:
            rename_map = {class_col: "مجموع التصنيف"}
            if account_col:
                rename_map[account_col] = "عدد Customer Account Number"
            pivot_df = pivot_df.rename(columns=rename_map)

        pivot_df = pivot_df.rename_axis(index="المحصل")
        st.dataframe(pivot_df, use_container_width=True)


def _render_period_results(stored, period_key):
    """عرض نتائج الفترة المحفوظة (جدول + كروت + شارتات + تنزيل) — تُستخدم بعد الضغط على زر التصنيف وبعد شيل الملف."""
    period_title = stored["period_title"]
    result_df = stored["df"]
    sales_col = stored["sales_col"]
    time_col = stored["time_col"]

    render_duplicate_summary(stored.get("duplicate_stats"))
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    render_period_charts(result_df, sales_col, time_col, period_title)

    render_pivot_section(result_df, f"period_{period_key}")

    if result_df is not None:
        excel_bytes, pivot_added = build_excel_with_native_pivot(result_df, "النتائج", f"period_{period_key}")
        if not pivot_added:
            st.caption("⚠️ اتنزل الملف من غير Pivot Table لأن عمود المحصل (Collected by) أو رقم حساب العميل (Customer Account number) مش موجود في الملف.")
        st.download_button(
            f"⬇️ تحميل نتائج {period_title}",
            data=excel_bytes,
            file_name=f"نتائج_{period_title}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"download_{period_key}",
        )


def _show_period_results_from_cache(period_key):
    """عرض النتائج المحفوظة بصمت لما يكون لا يوجد ملف مرفوع — بدون أي رسائل أو أزرار."""
    stored = st.session_state["period_results"].get(period_key)
    if stored:
        _render_period_results(stored, period_key)




def classify_period_file(uploaded_file, period_key):
    period_title = {"period_1": "الفترة الأولى", "period_2": "الفترة الثانية"}[period_key]

    # 💾 لو الفترة دي اتصنّفت قبل كده والنتيجة لسه في الذاكرة — نعرضها من الكاش من غير إعادة قراءة أو تصنيف
    stored = st.session_state["period_results"].get(period_key)
    current_file_hash = uploaded_file_hash(uploaded_file)
    if stored and stored.get("uploaded_hash") == current_file_hash:
        st.success(f"تم تصنيف {period_title} بنجاح ✅ — {len(stored['df']):,} مكالمة")
        _render_period_results(stored, period_key)
        return

    try:
        df = read_uploaded_dataframe(uploaded_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return

    # نفس قاعدة الملف الحالية: حذف أول صف بعد العناوين.
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    if ORIGINAL_TEXT_COL in df.columns:
        df = df.rename(columns={ORIGINAL_TEXT_COL: MODEL_TEXT_COL})

    if MODEL_TEXT_COL not in df.columns:
        st.error(
            f"عمود النص ('{ORIGINAL_TEXT_COL}' أو '{MODEL_TEXT_COL}') غير موجود. "
            f"الأعمدة الموجودة: {', '.join(df.columns.astype(str))}"
        )
        return

    if st.button(
        f"🚀 بدء تصنيف {period_title}",
        key=f"classify_{period_key}",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("جارٍ تجهيز النموذج وتصنيف الملف..."):
            tokenizer, model, device = load_model()
            texts = df[MODEL_TEXT_COL].tolist()
            preds, confidences = predict_batch(texts, tokenizer, model, device)

        result_df = df.copy()
        result_df[CLASSIFICATION_COL] = preds
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]

        sales_col = find_column(result_df, SALES_PERSON_CANDIDATES)
        time_col = find_column(result_df, CREATED_ON_CANDIDATES)
        claim_col = find_column(result_df, CLAIM_CANDIDATES)
        result_df, duplicate_stats = remove_claim_duplicates(result_df, claim_col, time_col)

        break_start = (
            st.session_state[f"{period_key}_break_start"]
            if st.session_state[f"{period_key}_has_break"]
            else None
        )
        break_end = (
            st.session_state[f"{period_key}_break_end"]
            if st.session_state[f"{period_key}_has_break"]
            else None
        )

        if sales_col and time_col:
            result_df = calculate_wasted_time(
                result_df, sales_col, time_col, break_start, break_end
            )
        else:
            st.warning(
                "تعذر العثور على عمود المحصّل أو عمود التاريخ والوقت، فلن يتم حساب الوقت المهدر. "
                "تأكد من أسماء الأعمدة."
            )

        result_df = result_df.rename(columns={MODEL_TEXT_COL: ORIGINAL_TEXT_COL})

        # معلومات الفترة والشركة تبقى مع النتيجة بدون التأثير على الموديل.
        result_df["الشركة"] = st.session_state["selected_company"]
        result_df["الفترة"] = period_title
        result_df["بداية الفترة"] = st.session_state[f"{period_key}_start"].strftime("%H:%M")
        result_df["نهاية الفترة"] = st.session_state[f"{period_key}_end"].strftime("%H:%M")

        st.session_state["period_results"][period_key] = {
            "df": result_df,
            "sales_col": sales_col,
            "time_col": time_col,
            "claim_col": claim_col,
            "duplicate_stats": duplicate_stats,
            "company": st.session_state["selected_company"],
            "period_title": period_title,
            "uploaded_filename": uploaded_file.name,
            "uploaded_hash": current_file_hash,
        }
        st.session_state["last_result_df"] = result_df
        st.session_state["last_sales_col"] = sales_col
        st.session_state["last_time_col"] = time_col
        st.rerun()

    stored = st.session_state["period_results"].get(period_key)
    if stored:
        st.success(f"تم تصنيف {period_title} بنجاح ✅ — {len(stored['df']):,} مكالمة")
        _render_period_results(stored, period_key)


def normalize_status(value):
    text = str(value).strip().lower()
    if any(word in text for word in CLOSED_WORDS):
        return "مغلق"
    return None


def get_status_series(df):
    status_col = find_column(df, STATUS_CANDIDATES)
    if not status_col:
        return None, None
    return status_col, df[status_col].fillna("").map(normalize_status)


def build_agent_activity(df, sales_col):
    if not sales_col or sales_col not in df.columns:
        return pd.DataFrame()

    work = df.copy()
    work["_status_norm"] = None
    status_col, status_series = get_status_series(work)
    if status_col:
        work["_status_norm"] = status_series

    grouped = work.groupby(sales_col).size().rename("إجمالي المكالمات").to_frame()
    grouped["ناجحة"] = work[work[CLASSIFICATION_COL] == 1].groupby(sales_col).size()
    grouped["ناجحة"] = grouped["ناجحة"].fillna(0).astype(int)

    counts = work[work["_status_norm"] == "مغلق"].groupby(sales_col).size()
    grouped["مغلق"] = counts.fillna(0).astype(int)

    # ===== عدد إفادات "لا يرد" في عمود Notes لكل محصّل =====
    notes_col = ORIGINAL_TEXT_COL if ORIGINAL_TEXT_COL in work.columns else None
    lrd_counts = None
    if notes_col:
        lrd = work[notes_col].astype(str).str.contains("لا يرد|لايرد|لا\\s*يرد", regex=True, na=False)
        lrd_counts = work[lrd].groupby(sales_col).size()
    grouped["إفادات لا يرد"] = lrd_counts.reindex(grouped.index, fill_value=0).astype(int) if lrd_counts is not None else 0
    total_calls = work.groupby(sales_col).size()
    lrd_pct = []
    for _, r in grouped.iterrows():
        lrd_n = int(r["إفادات لا يرد"])
        total_n = int(r["إجمالي المكالمات"])
        lrd_pct.append(round(lrd_n / total_n * 100, 1) if total_n else 0.0)
    grouped["نسبة لا يرد (%)"] = lrd_pct

    # ===== متوسط الفجوة/المدة لكل محصّل (دقيقة) =====
    # لو الملف فيه عمود "مدة المكالمة" نستخدمه، وإلا نستخدم متوسط الوقت المهدر بين المكالمات.
    dur_col = find_column(work, DURATION_CANDIDATES)
    avg_key = "متوسط مدة المكالمة (دقيقة)"
    if dur_col and dur_col in work.columns:
        dur = pd.to_numeric(work[dur_col], errors="coerce")
        avg = dur.groupby(work[sales_col]).mean()
        grouped[avg_key] = avg.round(1)
        grouped[avg_key] = grouped[avg_key].fillna(0)
    else:
        wt_col = WASTED_TIME_COL if WASTED_TIME_COL in work.columns else None
        if wt_col:
            wt = pd.to_numeric(work[wt_col], errors="coerce")
            avg = wt.groupby(work[sales_col]).mean()
            grouped[avg_key] = avg.round(1)
            grouped[avg_key] = grouped[avg_key].fillna(0)
        else:
            grouped[avg_key] = 0.0

    return grouped.fillna(0).reset_index().rename(columns={sales_col: "المحصّل"})


def _rounded_rect_path(x0, x1, y0, y1, radius=0.018):
    """إنشاء مسار SVG مستدير الزوايا داخل إحداثيات Plotly الورقية."""
    radius = min(radius, (x1 - x0) / 3, (y1 - y0) / 3)
    return (
        f"M {x0 + radius},{y0} "
        f"L {x1 - radius},{y0} Q {x1},{y0} {x1},{y0 + radius} "
        f"L {x1},{y1 - radius} Q {x1},{y1} {x1 - radius},{y1} "
        f"L {x0 + radius},{y1} Q {x0},{y1} {x0},{y1 - radius} "
        f"L {x0},{y0 + radius} Q {x0},{y0} {x0 + radius},{y0} Z"
    )


def render_kpi_dashboard(total, success, agent_count, success_rate, avg_wasted=None):
    """لوحة KPI مركزية مبنية بـ Plotly لضمان محاذاة موحدة داخل كل كارت."""
    cards = [
        ("📞<br>إجمالي المكالمات", total, {"valueformat": ",d"}, THEME["text"]),
        ("✅<br>المكالمات الناجحة", success, {"valueformat": ",d"}, COLOR_SUCCESS),
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("📈<br>نسبة النجاح", success_rate, {"valueformat": ".1f", "suffix": "%"}, COLOR_ACCENT),
    ]
    if avg_wasted is not None:
        cards.append(("⏱️<br>متوسط الوقت المهدر", avg_wasted, {"valueformat": ".1f", "suffix": " دقيقة"}, COLOR_WARN))

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
                value=value,
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
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG)


def render_period_charts(df, sales_col, time_col, period_title):
    if not sales_col or sales_col not in df.columns:
        st.info("لا يوجد عمود واضح للمحصّل لعرض نشاط المحصّلين.")
        return
    render_classification_filter_notice(df, sales_col)
    df = get_classification_view(df, sales_col)
    agent = build_agent_activity(df, sales_col)
    if agent.empty:
        return
    total = len(df)
    success = int((df[CLASSIFICATION_COL] == 1).sum()) if CLASSIFICATION_COL in df.columns else 0
    success_rate = round(success / total * 100, 1) if total else 0
    avg_wasted = None
    if WASTED_TIME_COL in df.columns:
        avg_wasted = float(pd.to_numeric(df[WASTED_TIME_COL], errors="coerce").mean())
        if pd.isna(avg_wasted):
            avg_wasted = 0.0
    st.subheader("📌 ملخص نتائج التصنيف")
    render_kpi_dashboard(total, success, len(agent), success_rate, avg_wasted)
    st.divider()
    render_agent_activity_charts(agent, df, sales_col, period_title)
    with st.expander("📋 عرض جدول تفاصيل كل محصّل"):
        st.dataframe(agent, use_container_width=True, hide_index=True)


def render_agent_activity_charts(agent, df, sales_col, period_title):
    agent_sorted = agent.sort_values("ناجحة", ascending=False).reset_index(drop=True)
    names = [str(n) for n in agent_sorted["المحصّل"]]

    def total_vs_success_chart():
        fig = go.Figure([
            go.Bar(name="إجمالي المكالمات", x=names, y=agent_sorted["إجمالي المكالمات"], marker_color=CLASSIFICATION_SECONDARY, customdata=names),
            go.Bar(name="المكالمات الناجحة", x=names, y=agent_sorted["ناجحة"], marker_color=CLASSIFICATION_HIGHLIGHT, customdata=names),
        ])
        fig.update_layout(
            **PLOTLY_LAYOUT,
            title=f"📞 إجمالي المكالمات والناجحة — {period_title}",
            barmode="group",
            xaxis_title="",
            yaxis_title="عدد المكالمات",
            height=430,
            legend=dict(orientation="h", yanchor="bottom", y=-0.28, x=0.5, xanchor="center"),
        )
        fig.update_traces(
            marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>",
        )
        render_selectable_chart(fig, "classification_total_success_chart")

    rates = [
        (int(row["ناجحة"]) / int(row["إجمالي المكالمات"]) * 100)
        if int(row["إجمالي المكالمات"]) else 0
        for _, row in agent_sorted.iterrows()
    ]

    def success_rate_chart():
        rate_df = agent_sorted.assign(**{"نسبة النجاح": rates}).sort_values("نسبة النجاح")
        fig = px.bar(
            rate_df,
            x="نسبة النجاح",
            y="المحصّل",
            orientation="h",
            text="نسبة النجاح",
            color="نسبة النجاح",
            color_continuous_scale=CLASSIFICATION_SCALE,
            template=PLOTLY_TEMPLATE,
        )
        fig.update_layout(
            **PLOTLY_LAYOUT,
            title="📈 نسبة نجاح كل محصّل",
            xaxis_range=[0, 100],
            xaxis_title="نسبة النجاح (%)",
            yaxis_title="",
            height=430,
            coloraxis_showscale=False,
        )
        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
            cliponaxis=False,
            marker_line_width=0,
            customdata=rate_df["المحصّل"],
            hovertemplate="<b>%{y}</b><br>نسبة النجاح: %{x:.1f}%<extra></extra>",
        )
        render_selectable_chart(fig, "classification_success_rate_chart")

    st.subheader(f"📈 تحليلات الأداء — {period_title}")
    first, second = st.columns(2)
    with first:
        with st.container(border=True):
            total_vs_success_chart()
    with second:
        with st.container(border=True):
            success_rate_chart()

    with st.container(border=True):
        render_success_fail_chart(agent, period_title)

    duration_col, no_answer_col = st.columns(2)
    with duration_col:
        with st.container(border=True):
            render_avg_duration_chart(agent, period_title)
    with no_answer_col:
        with st.container(border=True):
            render_no_answer_chart(df, sales_col, period_title)


def render_success_fail_chart(agent, period_title):
    ordered = agent.sort_values("ناجحة", ascending=False).reset_index(drop=True)
    names = [str(n) for n in ordered["المحصّل"]]
    failed = ordered["إجمالي المكالمات"] - ordered["ناجحة"]
    fig = go.Figure([
        go.Bar(name="المكالمات الناجحة", x=names, y=ordered["ناجحة"], marker_color=CLASSIFICATION_HIGHLIGHT, customdata=names),
        go.Bar(name="المكالمات غير الناجحة", x=names, y=failed, marker_color=CLASSIFICATION_SECONDARY, customdata=names),
    ])
    fig.update_layout(**PLOTLY_LAYOUT, title=f"✅ الناجحة مقابل غير الناجحة ({period_title})", barmode="group", xaxis_title="", yaxis_title="عدد المكالمات")
    fig.update_traces(hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>")
    render_selectable_chart(fig, f"classification_success_fail_{period_title}")


def render_avg_duration_chart(agent, period_title):
    key = "متوسط مدة المكالمة (دقيقة)"
    if key not in agent.columns:
        return
    ordered = agent.sort_values(key, ascending=True)
    fig = px.bar(ordered, x=key, y="المحصّل", orientation="h", text=key, color=key, color_continuous_scale=CLASSIFICATION_SCALE, template=PLOTLY_TEMPLATE)
    fig.update_layout(**PLOTLY_LAYOUT, title=f"⏱️ متوسط مدة المكالمات ({period_title})", xaxis_title="دقيقة", yaxis_title="")
    fig.update_traces(customdata=ordered["المحصّل"], hovertemplate="<b>%{y}</b><br>متوسط المدة: %{x:.1f} دقيقة<extra></extra>")
    render_selectable_chart(fig, f"classification_avg_duration_{period_title}")


def render_no_answer_chart(df, sales_col, period_title):
    notes_col = ORIGINAL_TEXT_COL if ORIGINAL_TEXT_COL in df.columns else None
    if not notes_col or sales_col not in df.columns:
        return
    work = df.copy()
    work["_lrd"] = work[notes_col].astype(str).str.contains(r"لا يرد|لايرد|لا\s*يرد", regex=True, na=False)
    counts = work[work["_lrd"]].groupby(sales_col).size().reset_index(name="عدد إفادات لا يرد")
    if counts.empty:
        return
    fig = px.bar(counts, x=sales_col, y="عدد إفادات لا يرد", color_discrete_sequence=[CLASSIFICATION_PRIMARY], template=PLOTLY_TEMPLATE)
    fig.update_layout(**PLOTLY_LAYOUT, title=f"📝 إفادات لا يرد لكل محصّل ({period_title})", xaxis_title="", yaxis_title="العدد")
    fig.update_traces(customdata=counts[sales_col], hovertemplate="<b>%{x}</b><br>عدد إفادات لا يرد: %{y}<extra></extra>")
    render_selectable_chart(fig, f"classification_no_answer_{period_title}")


def _render_aggregate_break_settings(period_key, period_title):
    has_break_key = f"{period_key}_has_break"
    break_start_key = f"{period_key}_break_start"
    break_end_key = f"{period_key}_break_end"
    _render_break_switch(f"🕐 هل يوجد استراحة في {period_title}؟", has_break_key)
    if st.session_state[has_break_key]:
        b1, b2 = st.columns(2)
        with b1:
            st.session_state[break_start_key] = st.time_input("🕐 بداية الاستراحة", value=st.session_state[break_start_key], key=f"{break_start_key}_input")
        with b2:
            duration_key = f"{period_key}_break_duration"
            st.session_state[duration_key] = st.slider("⏳ مدة الاستراحة (دقيقة)", 5, 120, int(st.session_state[duration_key]), 5, key=f"{duration_key}_slider")
            total = st.session_state[break_start_key].hour * 60 + st.session_state[break_start_key].minute + st.session_state[duration_key]
            st.session_state[break_end_key] = dt_time(total // 60, total % 60)
        st.info(f"☕ الاستراحة من {st.session_state[break_start_key]:%H:%M} إلى {st.session_state[break_end_key]:%H:%M}")


def render_aggregate_tab(period_key, period_title):
    company = st.session_state["selected_company"]
    _render_aggregate_break_settings(period_key, period_title)
    uploaded_file = st.file_uploader(
        f"📂 ارفع ملف {period_title} (CSV أو Excel)",
        type=["csv", "xlsx", "xls"],
        key=f"upload_{period_key}",
        on_change=sync_file_cache,
        args=(f"upload_{period_key}", f"period_upload:{period_key}", (f"{period_key}_result",)),
    )
    if uploaded_file is not None:
        st.caption(f"الملف المختار: {uploaded_file.name}")
        _classify_aggregate_file(uploaded_file, company, period_key, period_title)
    else:
        _show_aggregate_results_from_cache(period_key, period_title)


def _classify_aggregate_file(uploaded_file, company, period_key, period_title):
    result_key = f"{period_key}_result"
    stored = st.session_state.get(result_key)
    current_file_hash = uploaded_file_hash(uploaded_file)
    if stored and stored.get("uploaded_hash") == current_file_hash:
        _render_aggregate_results(stored, period_title, period_key)
        return
    try:
        df = read_uploaded_dataframe(uploaded_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)
    if ORIGINAL_TEXT_COL in df.columns:
        df = df.rename(columns={ORIGINAL_TEXT_COL: MODEL_TEXT_COL})
    if MODEL_TEXT_COL not in df.columns:
        st.error(f"عمود النص غير موجود. الأعمدة الموجودة: {', '.join(df.columns.astype(str))}")
        return
    if st.button(f"🚀 بدء تصنيف {period_title}", key=f"btn_{period_key}", type="primary", use_container_width=True):
        with st.spinner(f"جارٍ تصنيف {period_title}..."):
            tokenizer, model, device = load_model()
            texts = df[MODEL_TEXT_COL].tolist()
            preds, confidences = predict_batch(texts, tokenizer, model, device)
        result_df = df.copy()
        result_df[CLASSIFICATION_COL] = preds
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]
        sales_col = find_column(result_df, SALES_PERSON_CANDIDATES)
        time_col = find_column(result_df, CREATED_ON_CANDIDATES)
        claim_col = find_column(result_df, CLAIM_CANDIDATES)
        result_df, duplicate_stats = remove_claim_duplicates(result_df, claim_col, time_col)
        has_break = st.session_state[f"{period_key}_has_break"]
        break_start = st.session_state[f"{period_key}_break_start"] if has_break else None
        break_end = st.session_state[f"{period_key}_break_end"] if has_break else None
        if sales_col and time_col:
            result_df = calculate_wasted_time(result_df, sales_col, time_col, break_start, break_end)
        result_df = result_df.rename(columns={MODEL_TEXT_COL: ORIGINAL_TEXT_COL})
        result_df["الشركة"] = company
        result_df["الفترة"] = period_title
        if has_break:
            result_df["بداية الاستراحة"] = break_start.strftime("%H:%M")
            result_df["نهاية الاستراحة"] = break_end.strftime("%H:%M")
            result_df["مدة الاستراحة_دقيقة"] = st.session_state[f"{period_key}_break_duration"]
        st.session_state[result_key] = {
            "df": result_df, "sales_col": sales_col, "time_col": time_col,
            "claim_col": claim_col, "duplicate_stats": duplicate_stats,
            "company": company, "uploaded_filename": uploaded_file.name,
            "uploaded_hash": current_file_hash,
        }
        st.rerun()

def _render_aggregate_results(stored, period_title, period_key):
    result_df = stored["df"]
    sales_col = stored["sales_col"]
    time_col = stored["time_col"]
    render_duplicate_summary(stored.get("duplicate_stats"))
    st.dataframe(result_df, use_container_width=True, hide_index=True)
    render_period_charts(result_df, sales_col, time_col, period_title)
    render_pivot_section(result_df, f"agg_{period_key}")
    excel_bytes, pivot_added = build_excel_with_native_pivot(result_df, period_title, f"agg_{period_key}")
    if not pivot_added:
        st.caption("⚠️ اتنزل الملف من غير Pivot Table لأن عمود المحصل (Collected by) أو رقم حساب العميل (Customer Account number) مش موجود في الملف.")
    st.download_button(
        f"⬇️ تحميل نتائج {period_title}",
        data=excel_bytes,
        file_name=f"نتائج_{period_title}_{stored['company']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"dl_{period_title}",
        type="primary",
    )

def _show_aggregate_results_from_cache(period_key, period_title):
    stored = st.session_state.get(f"{period_key}_result")
    if stored:
        st.success(f"تم تصنيف {period_title} بنجاح ✅ — {len(stored['df']):,} مكالمة")
        _render_aggregate_results(stored, period_title, period_key)
def page_classification():
    init_activity_state()
    page_header("CALL QUALITY CLASSIFIER", "🎯 تصنيف المكالمات", "اختر الشركة → اختر الفترة → حدّد الميعاد والاستراحة → ارفع الملف → ابدأ التصنيف")
    render_company_selector()
    render_period_selector()


# ==========================================================
# تويبات 2-5
# ==========================================================

def page_placeholder(eyebrow, title, subtitle, icon):
    page_header(eyebrow, f"{icon} {title}", subtitle)
    st.info("هذا القسم ما زال قيد التجهيز، وسيتم تحديد مصدر بياناته ومنطقه في مرحلة لاحقة.")


# ==========================================================
# تويب 6: الداشبورد
# ==========================================================

DASHBOARD_SOURCE_KEY = "dashboard_uploaded_source"


def _normalize_match_id(val):
    """توحيد شكل المعرف للمطابقة (يزيل .0 القادمة من إكسل ويشيل المسافات)."""
    if pd.isna(val):
        return ""
    text = str(val).strip()
    if text.lower() in {"nan", "none", "nat", ""}:
        return ""
    if text.endswith(".0"):
        try:
            float(text)
            text = text[:-2]
        except ValueError:
            pass
    return text.strip()


def _run_neglect_followup_pipeline(new_file, old_file):
    """متابعة الإهمال — XLOOKUP على Account Number:

    1) من المحفظة الحديثة: نجيب قيمة ``Follow up Last Date``
    2) على شيت الإهمال القديم:
       - لو فيه عمود بالاسم الحرفي «تاريخ اخر متابعة» → نحدّثه
       - لو مش موجود → ننشئه
       - باقي أعمدة الشيت (بما فيها أي Follow up Last Date قديم) تفضل زي ما هي
    3) فرق الأيام = تاريخ اليوم − تاريخ اخر متابعة
    4) أقل من 8 أيام = تم التغطية، غير كده = لم يتم التغطية
    """
    TARGET_LAST_COL = "تاريخ اخر متابعة"
    try:
        _init_promises_today()
        df_new = read_uploaded_dataframe(new_file)
        df_old = read_uploaded_dataframe(old_file)

        if len(df_new) > 0:
            df_new = df_new.iloc[1:].reset_index(drop=True)

        if len(df_old) > 0:
            first_row = df_old.iloc[0]
            if first_row.isna().all() or all(
                str(v).strip() in {"", "nan", "None"} for v in first_row.values
            ):
                df_old = df_old.iloc[1:].reset_index(drop=True)

        account_candidates = ACCOUNT_NUMBER_CANDIDATES + CLAIM_CANDIDATES + ID_CANDIDATES
        acc_col_new = find_column(df_new, account_candidates)
        acc_col_old = find_column(df_old, account_candidates)
        # المصدر من المحفظة الحديثة فقط: Follow up Last Date
        last_date_new = find_column(df_new, NEGLECT_LAST_DATE_CANDIDATES)

        if not acc_col_new or not acc_col_old:
            st.error(
                "تعذر العثور على عمود Account Number (رقم المطالبة) للمطابقة.\n\n"
                f"أعمدة المحفظة الحديثة: {', '.join(map(str, df_new.columns))}\n\n"
                f"أعمدة تقرير الإهمال: {', '.join(map(str, df_old.columns))}"
            )
            return
        if not last_date_new:
            st.error(
                "تعذر العثور على عمود Follow up Last Date في المحفظة الحديثة.\n\n"
                f"الأعمدة الموجودة: {', '.join(map(str, df_new.columns))}"
            )
            return

        work_new = pd.DataFrame({
            "_key": df_new[acc_col_new].map(_normalize_match_id),
            "_last": [parse_date_cell(v) for v in df_new[last_date_new]],
        })
        work_new = work_new[work_new["_key"] != ""]
        lookup_map = (
            work_new.dropna(subset=["_last"])
            .sort_values("_last")
            .groupby("_key", sort=False)["_last"]
            .max()
            .to_dict()
        )

        result_df = df_old.copy()
        keys_old = result_df[acc_col_old].map(_normalize_match_id)
        looked_up = keys_old.map(lookup_map)

        # عمود بالاسم الحرفي فقط — لو موجود نحدّثه، لو لأ ننشئه. من غير لمس باقي الأعمدة.
        updated_existing = TARGET_LAST_COL in result_df.columns
        result_df[TARGET_LAST_COL] = looked_up

        target_date = st.session_state.get(TODAY_KEY, datetime.now().date())

        def _days_since(val):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return None
            d = val if hasattr(val, "year") else parse_date_cell(val)
            if d is None:
                return None
            try:
                return (target_date - d).days
            except Exception:
                return None

        result_df["فرق_الأيام"] = result_df[TARGET_LAST_COL].apply(_days_since)

        def _coverage(days):
            if days is None or (isinstance(days, float) and pd.isna(days)):
                return "لم يتم التغطية"
            return "تم التغطية" if int(days) < 8 else "لم يتم التغطية"

        result_df["الملاحظات"] = result_df["فرق_الأيام"].apply(_coverage)

        # تنسيق «تاريخ اخر متابعة» للعرض/التصدير
        result_df[TARGET_LAST_COL] = result_df[TARGET_LAST_COL].apply(
            lambda d: d.strftime("%Y-%m-%d")
            if d is not None and pd.notna(d) and hasattr(d, "strftime")
            else ("" if d is None or (isinstance(d, float) and pd.isna(d)) else str(d))
        )

        matched = int(result_df["فرق_الأيام"].notna().sum())
        st.session_state["neglect_followup_result"] = {
            "df": result_df,
            "meta": {
                "acc_col_new": acc_col_new,
                "acc_col_old": acc_col_old,
                "last_date_new": last_date_new,
                "target_last_col": TARGET_LAST_COL,
                "updated_existing": updated_existing,
                "matched": matched,
                "total_old": len(result_df),
                "mapping_size": len(lookup_map),
                "target_date": target_date,
            },
        }
        st.rerun()
    except Exception as e:
        st.error(f"خطأ في معالجة الملفات: {e}")


def render_neglect_kpi_dashboard(cards, chart_key="neglect_kpi"):
    """كروت KPI بنفس ستايل التصنيف (Plotly rounded cards)."""
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
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG, key=chart_key)


def _show_neglect_followup_results(df, meta=None):
    st.subheader("📊 نتائج متابعة الإهمال")
    meta = meta or {}
    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    target_last_col = meta.get("target_last_col") or (
        find_column(df, NEGLECT_LAST_DATE_CANDIDATES) or "تاريخ اخر متابعة"
    )

    # فلتر تفاعلي من الشارت
    render_neglect_filter_notice(
        filter_key=NEGLECT_FOLLOWUP_AGENT_FILTER_KEY,
        clear_key="clear_neglect_followup_agent_filter",
    )
    df = get_neglect_view(df, sales_col, filter_key=NEGLECT_FOLLOWUP_AGENT_FILTER_KEY)

    covered = int((df["الملاحظات"] == "تم التغطية").sum()) if "الملاحظات" in df.columns else 0
    not_covered = int((df["الملاحظات"] == "لم يتم التغطية").sum()) if "الملاحظات" in df.columns else 0
    total = len(df)
    pct = covered / total * 100 if total else 0
    matched = int(df["فرق_الأيام"].notna().sum()) if "فرق_الأيام" in df.columns else 0

    cards = [
        ("📋<br>إجمالي الحالات", total, {"valueformat": ",d"}, THEME["text"]),
        ("✅<br>تم التغطية", covered, {"valueformat": ",d"}, OPS_POSITIVE),
        ("❌<br>لم يتم التغطية", not_covered, {"valueformat": ",d"}, OPS_NEGATIVE),
        ("📈<br>نسبة التغطية", pct, {"valueformat": ".1f", "suffix": "%"}, OPS_MID),
        ("🔗<br>حالات لها تاريخ", matched, {"valueformat": ",d"}, OPS_LIGHT),
    ]
    render_neglect_kpi_dashboard(cards, chart_key="neglect_followup_kpi")

    if meta:
        action = "تحديث العمود الموجود" if meta.get("updated_existing") else "إنشاء عمود جديد"
        st.caption(
            f"{action}: «{target_last_col}» · "
            f"XLOOKUP: المحفظة[{meta.get('acc_col_new')}] ↔ الإهمال[{meta.get('acc_col_old')}] · "
            f"المصدر: {meta.get('last_date_new')} · "
            f"حجم الخريطة: {meta.get('mapping_size', 0):,} · "
            f"تاريخ اليوم: {meta.get('target_date')}"
        )

    if total and "الملاحظات" in df.columns:
        st.markdown("#### 📈 تحليلات التغطية")
        left, right = st.columns(2)
        with left:
            type_counts = (
                df["الملاحظات"].value_counts()
                .rename_axis("الحالة")
                .reset_index(name="العدد")
            )
            fig = px.pie(
                type_counts,
                values="العدد",
                names="الحالة",
                hole=0.55,
                color="الحالة",
                color_discrete_map={
                    "تم التغطية": OPS_POSITIVE,
                    "لم يتم التغطية": OPS_NEGATIVE,
                },
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig, "توزيع التغطية", height=430,
                margin=dict(t=70, b=50, l=30, r=30),
            )
            fig.update_traces(
                texttemplate="%{label}<br>%{value:,.0f} (%{percent:.1%})",
                textfont=dict(size=14, color=THEME["text"]),
                textinfo="text",
                hovertemplate="<b>%{label}</b><br>العدد: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
            )
            with st.container(border=True):
                st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="neglect_followup_pie")
        with right:
            if sales_col and sales_col in df.columns:
                agent_type = (
                    df.groupby([sales_col, "الملاحظات"]).size()
                    .reset_index(name="العدد")
                )
                agent_order = (
                    agent_type.groupby(sales_col)["العدد"].sum()
                    .sort_values(ascending=False).head(15).index.tolist()
                )
                agent_type = agent_type[agent_type[sales_col].isin(agent_order)]
                # ترتيب العرض من الأقل للأعلى عشان الشارت الأفقي يبقى أوضح
                agent_type["_sort"] = agent_type[sales_col].map({n: i for i, n in enumerate(agent_order)})
                agent_type = agent_type.sort_values(["_sort", "الملاحظات"], ascending=[False, True])
                fig = px.bar(
                    agent_type,
                    x="العدد",
                    y=sales_col,
                    color="الملاحظات",
                    barmode="group",
                    orientation="h",
                    text="العدد",
                    color_discrete_map={
                        "تم التغطية": OPS_POSITIVE,
                        "لم يتم التغطية": OPS_NEGATIVE,
                    },
                    template=PLOTLY_TEMPLATE,
                    category_orders={sales_col: list(reversed(agent_order))},
                )
                _apply_ops_chart_style(
                    fig, "التغطية حسب المحصّل — اضغط للاختيار",
                    height=430, xaxis_title="عدد الحالات",
                    margin=dict(t=70, b=70, l=160, r=50),
                )
                fig.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                    cliponaxis=False,
                    customdata=agent_type[sales_col],
                    hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f}<extra></extra>",
                )
                with st.container(border=True):
                    render_selectable_chart(
                        fig,
                        "neglect_followup_by_agent",
                        filter_key=NEGLECT_FOLLOWUP_AGENT_FILTER_KEY,
                    )
            else:
                st.info("لا يوجد عمود محصّل لعرض التوزيع.")

    preferred = [
        c for c in [
            "الملاحظات",
            target_last_col,
            "فرق_الأيام",
            sales_col,
            meta.get("acc_col_old") if meta else None,
        ]
        if c and c in df.columns
    ]
    other_cols = [c for c in df.columns if c not in preferred]
    ordered = preferred + other_cols
    st.subheader("📋 تفاصيل متابعة الإهمال")
    st.dataframe(df[ordered], use_container_width=True, hide_index=True)

    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
        df[ordered].to_excel(writer, index=False, sheet_name="متابعة الإهمال")
    st.download_button(
        "⬇️ تحميل تقرير متابعة الإهمال المحدث (Excel)",
        data=out_excel.getvalue(),
        file_name=f"متابعة_الإهمال_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


def page_neglect():
    """تويب الإهمال: فلترة الحالات المتأخرة في المتابعة + حساب فرق الأيام + داشبورد."""
    init_neglect_state()
    _init_promises_today()
    
    page_header(
        "NEGLECT TRACKING",
        "⚠️ الإهمال ومتابعة الإهمال",
        "ارفع المحفظة وسيتم استخراج الحالات التي مرّ عليها أكثر من 7 أيام دون متابعة وتحتاج إلى تدخل سريع",
    )
    
    # اختيار الوضع
    m1, m2 = st.columns(2)
    with m1:
        if st.button("🚨 الإهمال", use_container_width=True, type="primary" if st.session_state["neglect_mode"] == "neglect" else "secondary"):
            st.session_state["neglect_mode"] = "neglect"
            st.rerun()
    with m2:
        if st.button("🔍 متابعة الإهمال", use_container_width=True, type="primary" if st.session_state["neglect_mode"] == "followup" else "secondary"):
            st.session_state["neglect_mode"] = "followup"
            st.rerun()

    st.divider()
    
    if st.session_state["neglect_mode"] == "neglect":
        # إدارة حالات Sub State
        with st.expander("⚙️ إدارة حالات Sub State المستهدفة (الإهمال)"):
            available = st.session_state.get("neglect_available_states", [])
            if available:
                st.subheader("🔍 اختر حالات إضافية من الملف المرفوع")
                to_add_options = [s for s in available if s not in st.session_state["neglect_sub_states"]]
                selected_to_add = st.multiselect("اختر الحالات لإضافتها لقائمة الإهمال:", to_add_options)
                if st.button("➕ إضافة الحالات المختارة"):
                    if selected_to_add:
                        st.session_state["neglect_sub_states"].extend(selected_to_add)
                        st.success(f"تمت إضافة {len(selected_to_add)} حالة بنجاح!")
                        st.rerun()
            else:
                st.info("💡 ارفع ملف أولاً أستطيع استخراج كل الحالات المتاحة تختار منها.")
            
            st.divider()
            st.write("📋 الحالات المشمولة حالياً في الإهمال:")
            cols = st.columns(3)
            for i, state in enumerate(st.session_state["neglect_sub_states"]):
                with cols[i % 3]:
                    if st.button(f"❌ {state}", key=f"del_{i}", use_container_width=True):
                        st.session_state["neglect_sub_states"].remove(state)
                        st.rerun()

        uploaded = st.file_uploader(
            "📂 ارفع ملف المحفظة (Excel أو CSV) لفلترة الإهمال",
            type=["xlsx", "xls", "csv"],
            key="neglect_upload",
            on_change=sync_file_cache,
            args=("neglect_upload", "neglect", (NEGLECT_RESULT_KEY,)),
        )
        if uploaded is not None:
            _run_neglect_pipeline(uploaded)
        else:
            cached = st.session_state.get(NEGLECT_RESULT_KEY)
            if cached:
                _show_neglect_results(cached["df"], cached)
    else:
        # وضع متابعة الإهمال - هذا الجزء كان مفقوداً في النسخة السابقة
        st.info("💡 يطابق هذا الوضع تقرير إهمال قديمًا مع محفظة اليوم الحديثة لتحديد الحالات التي تمت متابعتها.")
        c1, c2 = st.columns(2)
        with c1:
            new_portfolio = st.file_uploader(
                "📂 ارفع محفظة اليوم الحديثة",
                type=["xlsx", "xls", "csv"],
                key="new_portfolio_up",
                on_change=sync_file_cache,
                args=("new_portfolio_up", "neglect_followup_new", ("neglect_followup_result",)),
            )
        with c2:
            old_neglect = st.file_uploader(
                "📂 ارفع تقرير الإهمال القديم",
                type=["xlsx", "xls", "csv"],
                key="old_neglect_up",
                on_change=sync_file_cache,
                args=("old_neglect_up", "neglect_followup_old", ("neglect_followup_result",)),
            )
        
        if new_portfolio and old_neglect:
            if st.button("🚀 بدء المطابقة ومتابعة الإهمال", use_container_width=True, type="primary"):
                _run_neglect_followup_pipeline(new_portfolio, old_neglect)
        
        cached_followup = st.session_state.get("neglect_followup_result")
        if cached_followup:
            _show_neglect_followup_results(cached_followup["df"], cached_followup.get("meta"))

def _run_neglect_pipeline(uploaded):

    _file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    cached = st.session_state.get(NEGLECT_RESULT_KEY)
    if cached and cached.get("file_hash") == _file_hash:
        _show_neglect_results(cached["df"], cached)
        return

    try:
        df = read_uploaded_dataframe(uploaded)
        if len(df) > 0:
            df = df.iloc[1:].reset_index(drop=True)
            
        sales_col = find_column(df, SALES_PERSON_CANDIDATES)
        substate_col = find_column(df, PROMISE_SUB_STATE_CANDIDATES)
        duedate_col = find_column(df, PROMISE_DUE_DATE_CANDIDATES)
        lastdate_col = find_column(df, NEGLECT_LAST_DATE_CANDIDATES)
        net_col = find_column(df, PROMISE_NET_AMOUNT_CANDIDATES)
        
        if not all([sales_col, substate_col, duedate_col]):
            st.error("الملف يفتقد أعمدة أساسية (المحصّل، الحالة، أو تاريخ المتابعة).")
            return
            
        # استخراج كافة الحالات المتاحة في الملف للكاش
        all_states = sorted(df[substate_col].dropna().unique().tolist())
        st.session_state["neglect_available_states"] = all_states

        # 1. فلترة المحصّلين
        df = df[~df[sales_col].astype(str).str.strip().isin(PROMISE_EXCLUDED_SALES)].copy()
        
        # 2. فلترة Sub State
        df = df[df[substate_col].astype(str).str.strip().isin(st.session_state["neglect_sub_states"])].copy()
        
        # 3. فلترة التاريخ (ما عدا اليوم)
        target_date = st.session_state[TODAY_KEY]
        df['temp_due'] = [parse_date_cell(v) for v in df[duedate_col]]
        df = df[df['temp_due'] != target_date].copy()
        
        # 4. حساب فرق الأيام والفلترة (> 7 أيام)
        if lastdate_col:
            df['temp_last'] = [parse_date_cell(v) for v in df[lastdate_col]]
            df['فرق_الأيام'] = [(target_date - d).days if d else 0 for d in df['temp_last']]
            # فلترة الحالات اللي بقالها أكتر من 7 أيام
            df = df[df['فرق_الأيام'] > 7].copy()
        
        st.session_state[NEGLECT_RESULT_KEY] = {
            "df": df, "file_hash": _file_hash, "sales_col": sales_col,
            "substate_col": substate_col, "duedate_col": duedate_col,
            "lastdate_col": lastdate_col, "net_col": net_col
        }
        st.rerun()
    except Exception as e:
        st.error(f"خطأ في معالجة الملف: {e}")

def _show_neglect_results(df, meta):
    sales_col = meta.get("sales_col")
    net_col = meta.get("net_col")

    render_neglect_filter_notice(
        filter_key=NEGLECT_AGENT_FILTER_KEY,
        clear_key="clear_neglect_agent_filter",
    )
    df = get_neglect_view(df, sales_col, filter_key=NEGLECT_AGENT_FILTER_KEY)

    total = len(df)
    total_amount = (
        pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum()
        if net_col and net_col in df.columns else 0
    )
    agent_count = int(df[sales_col].nunique()) if sales_col and sales_col in df.columns else 0
    avg_days = (
        float(pd.to_numeric(df["فرق_الأيام"], errors="coerce").mean())
        if "فرق_الأيام" in df.columns else None
    )
    if avg_days is not None and pd.isna(avg_days):
        avg_days = 0.0

    st.subheader("📊 ملخص الإهمال")
    cards = [
        ("⚠️<br>إجمالي الحالات", total, {"valueformat": ",d"}, OPS_MID),
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("💰<br>إجمالي المديونية", total_amount, {"valueformat": ",.0f"}, OPS_DARK),
        ("📅<br>متوسط فرق الأيام", avg_days if avg_days is not None else 0, {"valueformat": ".1f", "suffix": " يوم"}, OPS_MID),
    ]
    render_neglect_kpi_dashboard(cards, chart_key="neglect_main_kpi")

    if total and sales_col and sales_col in df.columns:
        st.markdown("#### 📈 تحليلات الإهمال")
        left, right = st.columns(2)
        with left:
            agent_counts = (
                df[sales_col].value_counts().head(15).sort_values().reset_index()
            )
            agent_counts.columns = ["المحصّل", "عدد الحالات"]
            fig = px.bar(
                agent_counts,
                x="عدد الحالات",
                y="المحصّل",
                orientation="h",
                text="عدد الحالات",
                color="عدد الحالات",
                color_continuous_scale=OPS_SCALE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig, "أعلى المحصّلين — اضغط للاختيار",
                height=430, xaxis_title="عدد الحالات", show_legend=False,
                margin=dict(t=70, b=55, l=160, r=50),
            )
            fig.update_traces(
                texttemplate="%{x:,.0f}",
                textposition="outside",
                cliponaxis=False,
                customdata=agent_counts["المحصّل"],
                hovertemplate="<b>%{y}</b><br>عدد الحالات: %{x:,}<extra></extra>",
            )
            with st.container(border=True):
                render_selectable_chart(
                    fig,
                    "neglect_by_agent",
                    filter_key=NEGLECT_AGENT_FILTER_KEY,
                )
        with right:
            sub_col = meta.get("substate_col")
            if sub_col and sub_col in df.columns:
                state_counts = (
                    df[sub_col].astype(str).str.strip().value_counts()
                    .head(12)
                    .sort_values(ascending=True)
                    .reset_index()
                )
                state_counts.columns = ["الحالة", "العدد"]
                total_states = max(int(state_counts["العدد"].sum()), 1)
                state_counts["النسبة"] = (state_counts["العدد"] / total_states * 100).round(1)
                # شارت أفقي أوضح من الدائرة لما الحالات كتير والأسماء طويلة
                fig = px.bar(
                    state_counts,
                    x="العدد",
                    y="الحالة",
                    orientation="h",
                    text="العدد",
                    color="العدد",
                    color_continuous_scale=OPS_SCALE,
                    template=PLOTLY_TEMPLATE,
                )
                _apply_ops_chart_style(
                    fig,
                    "توزيع حالات Sub State",
                    height=max(430, 36 * len(state_counts) + 120),
                    xaxis_title="عدد الحالات",
                    show_legend=False,
                    margin=dict(t=70, b=55, l=180, r=70),
                )
                fig.update_traces(
                    texttemplate="%{x:,.0f}  (%{customdata:.1f}%)",
                    textposition="outside",
                    cliponaxis=False,
                    customdata=state_counts["النسبة"],
                    hovertemplate="<b>%{y}</b><br>العدد: %{x:,.0f}<br>النسبة: %{customdata:.1f}%<extra></extra>",
                    marker_line_width=0,
                )
                with st.container(border=True):
                    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="neglect_by_state")
            else:
                st.info("لا يوجد عمود Sub State لعرض التوزيع.")

    st.subheader("📋 جدول حالات الإهمال التفصيلي")
    display_cols = [
        c for c in [
            sales_col,
            meta.get("substate_col"),
            meta.get("duedate_col"),
            meta.get("lastdate_col"),
            "فرق_الأيام",
            net_col,
        ]
        if c and c in df.columns
    ]
    st.dataframe(df[display_cols] if display_cols else df, use_container_width=True, hide_index=True)

    export_cols = [c for c in df.columns if c not in {"temp_due", "temp_last"}]
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
        df[export_cols].to_excel(writer, index=False, sheet_name="حالات الإهمال")
    st.download_button(
        "⬇️ تحميل تقرير الإهمال (Excel)",
        data=out_excel.getvalue(),
        file_name=f"تقرير_الإهمال_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


def _render_dashboard_work_settings():
    """إعداد استراحة Dashboard النشاط؛ تُستخدم لخصم البريك من ساعات العمل اليومية."""
    st.subheader("⏱️ إعدادات حساب ساعات العمل")
    st.caption("سيتم حساب زمن النشاط من أول مكالمة لآخر مكالمة لكل محصل في كل يوم، مع خصم وقت الاستراحة المحدد.")
    st.session_state.setdefault("dashboard_has_break", False)
    st.session_state.setdefault("dashboard_break_start", dt_time(13, 0))
    st.session_state.setdefault("dashboard_break_end", dt_time(13, 15))
    has_break = st.checkbox("☕ يوجد وقت استراحة يتم خصمه", key="dashboard_has_break")
    if not has_break:
        return None, None
    c1, c2 = st.columns(2)
    with c1:
        break_start = st.time_input("بداية الاستراحة", key="dashboard_break_start")
    with c2:
        break_end = st.time_input("نهاية الاستراحة", key="dashboard_break_end")
    if break_start >= break_end:
        st.warning("يجب أن تسبق بداية الاستراحة نهايتها؛ لذلك لن يتم الخصم حتى يتم تصحيح الوقت.")
        return None, None
    st.info(f"سيتم خصم الاستراحة من {break_start:%H:%M} إلى {break_end:%H:%M} من ساعات العمل اليومية.")
    return break_start, break_end


def page_dashboard():
    """تويب داشبورد مستقلة تمامًا عن التصنيف — ترفع فيها ملف النشاط بعد التصنيف
    (فيه عمود التصنيف جاهز) وتعرض لك داشبورد كاملة بالكروت والشارتات،
    وممكن تتنزّل كصفحة ويب HTML مستقلة تفتحها في أي متصفح."""
    page_header(
        "ACTIVITY DASHBOARD",
        "📊 تحليل نشاط المحصّلين",
        "ارفع ملف النشاط المصنّف بعد التصنيف لبناء لوحة تحكم متكاملة",
        centered=True,
    )

    init_activity_state()
    break_start, break_end = _render_dashboard_work_settings()

    # 💾 الكاش الشفاف: لو فيه داشبورد محفوظة لآخر ملف مرفوع — نعرضها من غير إعادة معالجة
    cached = st.session_state.get("dashboard_result")
    current_source_hash = st.session_state.get(DASHBOARD_SOURCE_HASH_KEY)

    dash_file = st.file_uploader(
        "📂 ارفع ملف النشاط المصنّف (بعد التصنيف) — CSV أو Excel",
        type=["csv", "xlsx", "xls"],
        key="dash_upload_v3",
        on_change=sync_file_cache,
        args=("dash_upload_v3", "dashboard", ("dashboard_result", "dashboard_source")),
    )

    if dash_file is None and cached is None:
        st.info("ارفع ملف النشاط بعد تصنيفه، وسيتم بناء لوحة التحكم فورًا.")
        return

    # 🆕 رفع اختياري للمحفظة + السداد لبناء صفحات تحليل إضافية والربط بينهم وبين النشاط
    _render_dashboard_extra_uploaders()

    # لو اتشال الملف والنتيجة لسه في الذاكرة — نعرضها من الكاش
    if dash_file is None:
        _show_dashboard_from_cache(break_start, break_end)
        return

    # 💾 لو الملف ده اتعرج قبل كده — نعرض الكاش من غير إعادة معالجة
    current_file_hash = uploaded_file_hash(dash_file)
    if current_source_hash == current_file_hash and cached is not None:
        _render_dashboard(cached["df"], cached["class_col"], cached["sales_col"],
                          cached["time_col"], dash_file.name, filter_hint="",
                          break_start=break_start, break_end=break_end)
        return

    try:
        df = read_uploaded_dataframe(dash_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return

    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    time_col = find_column(df, CREATED_ON_CANDIDATES)

    if class_col is None:
        st.warning(
            f"⚠️ لا يوجد عمود '{CLASSIFICATION_COL}' في الملف؛ ستُعرض المكالمات دون تفاصيل النجاح أو الفشل. "
            "تأكد إنك رفعت الملف بعد التصنيف."
        )

    # 💾 نحفظ النتيجة في الكاش (شفاف — من غير أي كروت أو أزرار إضافية)
    st.session_state[DASHBOARD_SOURCE_KEY] = dash_file.name
    st.session_state[DASHBOARD_SOURCE_HASH_KEY] = current_file_hash
    st.session_state["dashboard_result"] = {
        "df": df, "class_col": class_col, "sales_col": sales_col,
        "time_col": time_col, "source_name": dash_file.name,
        "source_hash": current_file_hash,
    }

    # 🎚️ السلايسرز: فلتر المحصّلين + فلتر التواريخ (للعرض فقط — الكاش محفوظ)
    _render_dashboard(df, class_col, sales_col, time_col, dash_file.name, filter_hint="",
                      break_start=break_start, break_end=break_end)


def _render_dashboard_extra_uploaders():
    """رفع اختياري للمحفظة الحديثة وملف السداد — يفعّل 3 صفحات تحليل إضافية (المحفظة/السداد/الربط)
    بجانب صفحة نشاط المحصلين الأساسية، في التطبيق وفي ملف HTML المُصدَّر معًا."""
    with st.expander("➕ إضافة المحفظة والسداد (اختياري) — لعرض تحليل موسّع بأربع صفحات", expanded=False):
        st.caption(
            "الرفع هنا اختياري بالكامل. لو سبت الخانتين فاضيين هتفضل شايف صفحة نشاط المحصلين بس. "
            "ولو رفعت المحفظة والسداد مع بعض، هتتضاف 3 صفحات: تحليل المحفظة، تحليل السداد، وربط نشاط المحصلين بالسداد — "
            "وكلهم هيظهروا كمان جوه ملف الـ HTML اللي بتنزّله، مع أزرار للتنقل بين الصفحات."
        )
        c1, c2 = st.columns(2)
        with c1:
            wallet_file = st.file_uploader(
                "📂 ارفع المحفظة الحديثة (Excel أو CSV)",
                type=["csv", "xlsx", "xls"], key="dash_wallet_upload",
            )
        with c2:
            payments_file = st.file_uploader(
                "📂 ارفع ملف السداد (Excel أو CSV)",
                type=["csv", "xlsx", "xls"], key="dash_payments_upload",
            )
        _sync_dashboard_side_file(wallet_file, DASH_WALLET_CACHE_KEY)
        _sync_dashboard_side_file(payments_file, DASH_PAYMENTS_CACHE_KEY)

        wallet_info = st.session_state.get(DASH_WALLET_CACHE_KEY)
        payments_info = st.session_state.get(DASH_PAYMENTS_CACHE_KEY)
        if wallet_info and payments_info:
            st.success(
                f"✅ هيتم بناء 4 صفحات: نشاط المحصلين + المحفظة ({wallet_info['name']}) "
                f"+ السداد ({payments_info['name']}) + الربط بينهم."
            )
        elif wallet_info or payments_info:
            st.info("محتاج ترفع الملفين (المحفظة والسداد) مع بعض عشان تظهر صفحات التحليل الإضافية وصفحة الربط.")


def _sync_dashboard_side_file(uploaded_file, cache_key):
    """كاش بسيط بالهاش لملف جانبي (محفظة/سداد) في تويب النشاط."""
    if uploaded_file is None:
        st.session_state.pop(cache_key, None)
        return
    file_hash = uploaded_file_hash(uploaded_file)
    cached = st.session_state.get(cache_key)
    if cached and cached.get("hash") == file_hash:
        return
    try:
        df = read_uploaded_dataframe(uploaded_file)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        st.session_state.pop(cache_key, None)
        return
    st.session_state[cache_key] = {"df": df, "hash": file_hash, "name": uploaded_file.name}


def _render_dashboard(df, class_col, sales_col, time_col, source_name, filter_hint="", break_start=None, break_end=None):
    """يعرض صفحة واحدة فقط مع سلايسرزها: النشاط / المحفظة / السداد / الربط."""
    wallet_info = st.session_state.get(DASH_WALLET_CACHE_KEY)
    payments_info = st.session_state.get(DASH_PAYMENTS_CACHE_KEY)
    has_extra_pages = bool(wallet_info) and bool(payments_info)

    active_page = "📊 نشاط المحصلين"
    if has_extra_pages:
        pages = ["📊 نشاط المحصلين", "💼 المحفظة", "💰 السداد", "🔗 نشاط × سداد"]
        chosen = st.segmented_control(
            "اختر الصفحة", pages, default=pages[0], key=DASH_ACTIVE_PAGE_KEY,
            label_visibility="collapsed",
        )
        active_page = chosen or pages[0]
        st.divider()

    activity_df = df
    activity_hint = filter_hint or ""
    if active_page == "📊 نشاط المحصلين":
        activity_df, activity_hint = _render_slicers(df, sales_col, time_col)
        if activity_hint:
            st.caption(f"الفلاتر المطبقة: {activity_hint}")
        render_full_dashboard(
            activity_df, class_col=class_col, sales_col=sales_col, time_col=time_col,
            break_start=break_start, break_end=break_end,
        )
    elif active_page == "💼 المحفظة" and wallet_info:
        render_wallet_page(wallet_info["df"])
    elif active_page == "💰 السداد" and payments_info:
        render_payments_page(payments_info["df"])
    elif active_page == "🔗 نشاط × سداد" and payments_info:
        render_activity_payments_link_page(df, sales_col, class_col, payments_info["df"])

    st.divider()
    # التصدير من البيانات الأصلية؛ كل صفحة HTML فيها سلايسرزها التفاعلية
    dashboard_html = build_dashboard_html(
        df, class_col=class_col, sales_col=sales_col, time_col=time_col,
        source_name=source_name, filter_hint=activity_hint if active_page == "📊 نشاط المحصلين" else "",
        filter_summary=st.session_state.get("dashboard_filter_summary", {}),
        wallet_df=(wallet_info["df"] if wallet_info else None),
        payments_df=(payments_info["df"] if payments_info else None),
    )
    st.download_button("🌐 تحميل لوحة التحكم كصفحة ويب HTML", data=dashboard_html.encode("utf-8"), file_name="داشبورد_النشاط.html", mime="text/html", use_container_width=True, key="dash_html_download_v3", type="primary")
    st.download_button("⬇️ تحميل البيانات كـ CSV", data=df.to_csv(index=False).encode("utf-8-sig"), file_name="بيانات_النشاط.csv", mime="text/csv", use_container_width=True, key="dash_csv_download_v3")


def _clear_dashboard_chart_filter():
    st.session_state.pop(DASHBOARD_AGENT_FILTER_KEY, None)
    st.session_state.pop(DASHBOARD_DAY_FILTER_KEY, None)
    st.session_state.pop(DASHBOARD_OUTCOME_FILTER_KEY, None)


def _render_native_multi_slicer(label, options, state_key, empty_label):
    """Multi-select نظيف بلا Tags: الاختيارات تظهر داخل Popover فقط، والزر يعرض ملخصًا مختصرًا."""
    options = [str(option) for option in options]
    widget_keys = [f"{state_key}__{index}" for index in range(len(options))]
    selected = [option for option, widget_key in zip(options, widget_keys) if st.session_state.get(widget_key, False)]
    trigger_text = empty_label if not selected else f"تم اختيار {len(selected)}"
    st.caption(label)
    with st.popover(trigger_text, use_container_width=True):
        action_all, action_clear = st.columns(2)
        with action_all:
            if st.button("تحديد الكل", key=f"{state_key}__select_all", use_container_width=True):
                for widget_key in widget_keys:
                    st.session_state[widget_key] = True
                st.rerun()
        with action_clear:
            if st.button("إلغاء الكل", key=f"{state_key}__clear_all", use_container_width=True):
                for widget_key in widget_keys:
                    st.session_state[widget_key] = False
                st.rerun()
        st.caption("اختار أكثر من قيمة؛ لن تظهر الاختيارات كوسوم خارج القائمة.")
        for option, widget_key in zip(options, widget_keys):
            st.session_state.setdefault(widget_key, False)
            st.checkbox(option, key=widget_key, label_visibility="visible")
    return [option for option, widget_key in zip(options, widget_keys) if st.session_state.get(widget_key, False)]


def _render_slicers(df, sales_col, time_col):
    agents = sorted([str(a) for a in df[sales_col].dropna().unique()]) if sales_col and sales_col in df.columns else []
    date_min = date_max = None
    if time_col and time_col in df.columns:
        ts = pd.to_datetime(df[time_col], errors="coerce")
        if ts.notna().any():
            date_min, date_max = ts.min().date(), ts.max().date()
    sub_col = find_column(df, PROMISE_SUB_STATE_CANDIDATES)
    substates = sorted([str(s) for s in df[sub_col].dropna().unique()]) if sub_col and sub_col in df.columns else []
    class_col = CLASSIFICATION_COL if CLASSIFICATION_COL in df.columns else None
    all_agent_label = "كل المحصلين"
    all_state_label = "كل الحالات"
    class_labels = ["الكل", "ناجحة", "غير ناجحة"]
    class_values = {"الكل": None, "ناجحة": 1, "غير ناجحة": 0}

    # عنوان مستقل للفلاتر، وكل Slicer في خانة مستقلة بذاتها مثل لوحات Power BI.
    header_col, action_col = st.columns([5, 1])
    with header_col:
        st.subheader("🎚️ فلاتر التحليل")
        st.caption("كل فلتر مستقل؛ اترك الاختيار على «الكل» لعرض كل البيانات.")
    with action_col:
        if st.button("↺ إعادة ضبط", key="clear_dashboard_slicers_v5", use_container_width=True):
            for key in ("dash_agent_slicer_v5", "dash_state_slicer_v5", "dash_date_slicer_v5", "dash_class_slicer_v5"):
                st.session_state.pop(key, None)
            for prefix in ("dash_agent_slicer_v6", "dash_state_slicer_v6"):
                for session_key in list(st.session_state.keys()):
                    if session_key.startswith(prefix + "__"):
                        st.session_state.pop(session_key, None)
            _clear_dashboard_chart_filter()
            st.rerun()

    slicer_cols = st.columns(4, gap="small")
    with slicer_cols[0]:
        selected_agents = _render_native_multi_slicer(
            "👤 المحصلون", agents, "dash_agent_slicer_v6", "كل المحصلين"
        )
    with slicer_cols[1]:
        selected_substates = _render_native_multi_slicer(
            "📊 الحالات الفرعية", substates, "dash_state_slicer_v6", "كل الحالات"
        ) if substates else []
    with slicer_cols[2]:
        date_range = st.date_input(
            "📅 التاريخ",
            value=(date_min, date_max) if date_min is not None else None,
            min_value=date_min,
            max_value=date_max,
            key="dash_date_slicer_v5",
        ) if date_min is not None else None
    with slicer_cols[3]:
        selected_class = st.selectbox(
            "🏷️ التصنيف",
            class_labels,
            index=0,
            key="dash_class_slicer_v5",
        )

    selected_agents = [str(agent) for agent in selected_agents]
    selected_substates = [str(state) for state in selected_substates]
    date_summary = "كل التواريخ"
    if isinstance(date_range, tuple) and len(date_range) == 2:
        date_summary = f"{date_range[0]} إلى {date_range[1]}"
    st.session_state["dashboard_filter_summary"] = {
        "المحصل": "كل المحصلين" if not selected_agents else f"تم اختيار {len(selected_agents)} محصل",
        "الحالة الفرعية": "كل الحالات" if not selected_substates else f"تم اختيار {len(selected_substates)} حالة",
        "التاريخ": date_summary,
        "التصنيف": selected_class,
    }
    filtered = df.copy()
    hint_parts = []
    if sales_col and selected_agents:
        filtered = filtered[filtered[sales_col].astype(str).isin(selected_agents)]
        hint_parts.append(f"المحصلون: {', '.join(selected_agents)}")
    if time_col and isinstance(date_range, tuple) and len(date_range) == 2:
        d0, d1 = date_range
        if d0 != date_min or d1 != date_max:
            ts = pd.to_datetime(filtered[time_col], errors="coerce")
            filtered = filtered[ts.notna() & (ts.dt.date >= d0) & (ts.dt.date <= d1)]
            hint_parts.append(f"التاريخ: {d0} إلى {d1}")
    if class_col and class_values[selected_class] is not None:
        filtered = filtered[filtered[class_col] == class_values[selected_class]]
        hint_parts.append(f"التصنيف: {selected_class}")
    if sub_col and selected_substates:
        filtered = filtered[filtered[sub_col].astype(str).isin(selected_substates)]
        hint_parts.append(f"الحالات: {', '.join(selected_substates)}")
    return filtered, " · ".join(hint_parts)


def _show_dashboard_from_cache(break_start=None, break_end=None):
    """عرض الداشبورد المحفوظة بعد شيل الملف — من الكاش بدون إعادة معالجة."""
    cached = st.session_state.get("dashboard_result")
    if cached is None:
        return
    st.info(f"📌 لوحة التحكم محفوظة في الذاكرة — آخر ملف مرفوع: {cached['source_name']}")
    _render_dashboard(cached["df"], cached["class_col"], cached["sales_col"],
                      cached["time_col"], cached["source_name"], filter_hint="",
                      break_start=break_start, break_end=break_end)


# ==========================================================
# صفحات إضافية اختيارية داخل تحليل نشاط المحصلين: المحفظة / السداد / الربط
# ==========================================================

def _build_kpi_figure(cards, height=200):
    """كروت KPI عامة (نفس أسلوب كروت الجدولة/الإهمال) — قابلة لإعادة الاستخدام
    في أي صفحة تحليل جديدة، وتصلح للتصدير كـ HTML لأنها Plotly figure عادية."""
    figure = go.Figure()
    count = len(cards)
    gap = 0.018
    width = (1 - gap * (count + 1)) / count
    for index, (label, value, number_format, number_color) in enumerate(cards):
        x0 = gap + index * (width + gap)
        x1 = x0 + width
        figure.add_shape(
            type="path", xref="paper", yref="paper",
            path=_rounded_rect_path(x0, x1, 0.06, 0.94, radius=0.022),
            line={"color": THEME["border"], "width": 1}, fillcolor=THEME["surface"], layer="below",
        )
        figure.add_trace(go.Indicator(
            mode="number", value=float(value or 0),
            domain={"x": [x0 + 0.012, x1 - 0.012], "y": [0.12, 0.88]},
            title={"text": label, "font": {"size": 17, "color": THEME["text_dim"]}, "align": "center"},
            number={"font": {"size": 30, "color": number_color}, **number_format},
        ))
    figure.update_layout(
        height=height, template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Tajawal, sans-serif", "color": THEME["text"]},
        margin={"t": 8, "b": 8, "l": 8, "r": 8},
    )
    return figure


def _wallet_aging_bucket(days):
    """تصنيف أيام عمر الإسناد إلى بوكيتات ثابتة."""
    if days is None or (isinstance(days, float) and pd.isna(days)):
        return "غير محدد"
    try:
        d = int(days)
    except (TypeError, ValueError):
        return "غير محدد"
    if d < 0:
        d = 0
    for low, high, label in WALLET_AGING_BUCKETS:
        if high is None:
            if d >= low:
                return label
        elif low <= d <= high:
            return label
    return "غير محدد"


def _wallet_prepare_frame(df):
    """تجهيز إطار المحفظة: أعمدة رقمية + عمر الإسناد + التحصيل والمتبقي."""
    work = df.copy()
    sales_col = find_column(work, SALES_PERSON_CANDIDATES)
    sub_col = find_column(work, PROMISE_SUB_STATE_CANDIDATES)
    net_col = find_column(work, PROMISE_NET_AMOUNT_CANDIDATES)
    original_col = find_column(work, CASE_ORIGINAL_DEBT_CANDIDATES)
    payment_col = find_column(work, CASE_PAYMENT_CANDIDATES)
    assign_col = find_column(work, WALLET_ASSIGN_DATE_CANDIDATES)
    debit_col = find_column(work, WALLET_DEBIT_DATE_CANDIDATES)
    nationality_col = find_column(work, WALLET_NATIONALITY_CANDIDATES)
    customer_state_col = find_column(work, WALLET_CUSTOMER_STATE_CANDIDATES)
    customer_id_col = (
        find_column(work, WALLET_CUSTOMER_ID_CANDIDATES)
        or find_column(work, ACCOUNT_NUMBER_CANDIDATES)
        or find_column(work, ID_CANDIDATES)
    )

    if net_col and net_col in work.columns:
        work[net_col] = pd.to_numeric(work[net_col], errors="coerce").fillna(0)
    if original_col and original_col in work.columns:
        work[original_col] = pd.to_numeric(work[original_col], errors="coerce").fillna(0)
    if payment_col and payment_col in work.columns:
        work[payment_col] = pd.to_numeric(work[payment_col], errors="coerce").fillna(0)

    if net_col:
        work[WALLET_REMAINING_COL] = work[net_col]
    else:
        work[WALLET_REMAINING_COL] = 0.0

    # تم التحصيل: نعتمد أولاً على عمود Payment
    if payment_col:
        work[WALLET_COLLECTED_COL] = work[payment_col].clip(lower=0)
    elif original_col and net_col:
        work[WALLET_COLLECTED_COL] = (work[original_col] - work[net_col]).clip(lower=0)
    elif original_col:
        work[WALLET_COLLECTED_COL] = work[original_col].clip(lower=0)
    else:
        work[WALLET_COLLECTED_COL] = 0.0

    # تواريخ للسلايسرز فقط — مش بنحسب عمر الإسناد منها
    if assign_col and assign_col in work.columns:
        work["_assign_date"] = work[assign_col].map(parse_date_cell)
    else:
        work["_assign_date"] = None

    if debit_col and debit_col in work.columns:
        work["_debit_date"] = work[debit_col].map(parse_date_cell)
    else:
        work["_debit_date"] = None

    # عمر الإسناد من العمود الجاهز في ملف المحفظة
    aging_src_col = find_column(work, WALLET_AGING_SOURCE_CANDIDATES)
    if aging_src_col and aging_src_col in work.columns:
        work[WALLET_AGING_COL] = (
            work[aging_src_col].astype(str).str.strip()
            .replace({"": "غير محدد", "nan": "غير محدد", "None": "غير محدد", "none": "غير محدد", "null": "غير محدد"})
        )
    else:
        work[WALLET_AGING_COL] = "غير محدد"

    cols = {
        "sales_col": sales_col,
        "sub_col": sub_col,
        "net_col": net_col,
        "original_col": original_col,
        "payment_col": payment_col,
        "assign_col": assign_col,
        "debit_col": debit_col,
        "nationality_col": nationality_col,
        "customer_state_col": customer_state_col,
        "customer_id_col": customer_id_col,
        "aging_src_col": aging_src_col,
    }
    return work, cols


def _wallet_apply_slicers(work, cols):
    """سلايسرز صفحة المحفظة: محصل / تاريخ إسناد / تاريخ حادث / جنسية / حالة العميل."""
    filtered = work
    hint_parts = []

    sales_col = cols.get("sales_col")
    if sales_col and sales_col in filtered.columns:
        agents = sorted({
            v for v in filtered[sales_col].astype(str).str.strip().tolist()
            if v and v.lower() not in {"nan", "none", "null", ""}
        })
        selected_agents = st.multiselect(
            "👤 المحصل", options=agents, default=[], key="wallet_slicer_agent"
        )
        if selected_agents:
            filtered = filtered[filtered[sales_col].astype(str).str.strip().isin(selected_agents)]
            hint_parts.append(f"المحصل: {len(selected_agents)}")

    assign_col = cols.get("assign_col")
    if assign_col and "_assign_date" in filtered.columns:
        valid_dates = [d for d in filtered["_assign_date"].tolist() if d is not None]
        if valid_dates:
            dmin, dmax = min(valid_dates), max(valid_dates)
            c1, c2 = st.columns(2)
            with c1:
                start_assign = st.date_input(
                    "📅 تاريخ الإسناد من", value=dmin, min_value=dmin, max_value=dmax,
                    key="wallet_slicer_assign_from",
                )
            with c2:
                end_assign = st.date_input(
                    "📅 تاريخ الإسناد إلى", value=dmax, min_value=dmin, max_value=dmax,
                    key="wallet_slicer_assign_to",
                )
            if start_assign and end_assign:
                mask = filtered["_assign_date"].map(
                    lambda d: d is not None and start_assign <= d <= end_assign
                )
                filtered = filtered[mask]
                hint_parts.append(f"إسناد: {start_assign} → {end_assign}")

    debit_col = cols.get("debit_col")
    if debit_col and "_debit_date" in filtered.columns:
        valid_dates = [d for d in filtered["_debit_date"].tolist() if d is not None]
        if valid_dates:
            dmin, dmax = min(valid_dates), max(valid_dates)
            c1, c2 = st.columns(2)
            with c1:
                start_debit = st.date_input(
                    "📅 تاريخ الحادث من", value=dmin, min_value=dmin, max_value=dmax,
                    key="wallet_slicer_debit_from",
                )
            with c2:
                end_debit = st.date_input(
                    "📅 تاريخ الحادث إلى", value=dmax, min_value=dmin, max_value=dmax,
                    key="wallet_slicer_debit_to",
                )
            if start_debit and end_debit:
                mask = filtered["_debit_date"].map(
                    lambda d: d is not None and start_debit <= d <= end_debit
                )
                filtered = filtered[mask]
                hint_parts.append(f"حادث: {start_debit} → {end_debit}")

    nationality_col = cols.get("nationality_col")
    if nationality_col and nationality_col in filtered.columns:
        nations = sorted({
            v for v in filtered[nationality_col].astype(str).str.strip().tolist()
            if v and v.lower() not in {"nan", "none", "null", ""}
        })
        selected_nations = st.multiselect(
            "🌍 جنسية العميل", options=nations, default=[], key="wallet_slicer_nationality"
        )
        if selected_nations:
            filtered = filtered[filtered[nationality_col].astype(str).str.strip().isin(selected_nations)]
            hint_parts.append(f"الجنسية: {len(selected_nations)}")

    customer_state_col = cols.get("customer_state_col")
    if customer_state_col and customer_state_col in filtered.columns:
        states = sorted({
            v for v in filtered[customer_state_col].astype(str).str.strip().tolist()
            if v and v.lower() not in {"nan", "none", "null", ""}
        })
        selected_states = st.multiselect(
            "🏷️ حالة العميل", options=states, default=[], key="wallet_slicer_customer_state"
        )
        if selected_states:
            filtered = filtered[filtered[customer_state_col].astype(str).str.strip().isin(selected_states)]
            hint_parts.append(f"حالة العميل: {len(selected_states)}")

    return filtered, " · ".join(hint_parts)


def _build_wallet_aging_table(work, sub_col):
    """ماتريكس Power BI style: صفوف = عمر الإسناد، أعمدة = الحالات، قيم = عدد الحسابات + صف إجمالي."""
    if work is None or work.empty:
        return pd.DataFrame()

    aging = work[WALLET_AGING_COL].astype(str).str.strip().replace({"": "غير محدد"})
    if sub_col and sub_col in work.columns:
        state = work[sub_col].astype(str).str.strip().replace({"": "غير محدد"})
    else:
        state = pd.Series(["—"] * len(work), index=work.index)

    pivot = (
        pd.crosstab(aging, state, margins=True, margins_name="الإجمالي")
        .rename_axis("عمر الإسناد")
        .reset_index()
    )

    # ترتيب الصفوف: حاول البوكيتات المعروفة أولاً ثم الباقي ثم الإجمالي
    bucket_order = [label for _, _, label in WALLET_AGING_BUCKETS] + ["غير محدد"]
    order_map = {lab: i for i, lab in enumerate(bucket_order)}
    order_map["الإجمالي"] = 9999

    def _row_key(val):
        return order_map.get(str(val), 500)

    pivot["_ord"] = pivot["عمر الإسناد"].map(_row_key)
    pivot = pivot.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

    # أعمدة مساعدة مالية لكل عمر (اختياري للعرض بجانب المصفوفة)
    money = (
        work.assign(_aging=aging)
        .groupby("_aging")
        .agg(
            **{
                "تم التحصيل": (WALLET_COLLECTED_COL, "sum"),
                "باقي المديونية": (WALLET_REMAINING_COL, "sum"),
            }
        )
        .reset_index()
        .rename(columns={"_aging": "عمر الإسناد"})
    )
    total_money = pd.DataFrame([{
        "عمر الإسناد": "الإجمالي",
        "تم التحصيل": float(work[WALLET_COLLECTED_COL].sum()),
        "باقي المديونية": float(work[WALLET_REMAINING_COL].sum()),
    }])
    money = pd.concat([money, total_money], ignore_index=True)
    pivot = pivot.merge(money, on="عمر الإسناد", how="left")

    # ترتيب الأعمدة: عمر الإسناد | الحالات... | الإجمالي | تم التحصيل | باقي المديونية
    state_cols = [c for c in pivot.columns if c not in {"عمر الإسناد", "الإجمالي", "تم التحصيل", "باقي المديونية"}]
    # رتب الحالات حسب إجمالي العمود تنازلي
    if "الإجمالي" in pivot.columns:
        # صف الإجمالي موجود
        total_row = pivot[pivot["عمر الإسناد"] == "الإجمالي"]
        if not total_row.empty:
            state_cols = sorted(state_cols, key=lambda c: float(total_row.iloc[0].get(c, 0) or 0), reverse=True)
    ordered = ["عمر الإسناد"] + state_cols
    if "الإجمالي" in pivot.columns:
        ordered.append("الإجمالي")
    ordered += [c for c in ("تم التحصيل", "باقي المديونية") if c in pivot.columns]
    pivot = pivot[ordered]
    return pivot


def _build_wallet_analysis(df):
    """تحليل كامل لملف المحفظة: KPIs + الحالات + عمر الإسناد + الجنسية + حالة العميل."""
    work, cols = _wallet_prepare_frame(df)
    sales_col = cols["sales_col"]
    sub_col = cols["sub_col"]
    nationality_col = cols["nationality_col"]
    customer_state_col = cols["customer_state_col"]
    customer_id_col = cols["customer_id_col"]

    total_accounts = len(work)
    total_amount = float(work[WALLET_REMAINING_COL].sum())
    total_collected = float(work[WALLET_COLLECTED_COL].sum())
    agent_count = (
        int(work[sales_col].astype(str).str.strip().nunique())
        if sales_col and sales_col in work.columns else 0
    )
    if customer_id_col and customer_id_col in work.columns:
        customer_count = int(
            work[customer_id_col].astype(str).str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "none": pd.NA, "null": pd.NA, "None": pd.NA})
            .nunique(dropna=True)
        )
    else:
        customer_count = total_accounts
    avg_amount = (total_amount / total_accounts) if total_accounts else 0.0

    figs = {}
    figs["kpi"] = _build_kpi_figure([
        ("💰<br>إجمالي المديونية", total_amount, {"valueformat": ",.0f"}, OPS_DARK),
        ("👥<br>عدد المحصلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("🧑<br>عدد العملاء", customer_count, {"valueformat": ",d"}, THEME["text"]),
        ("📋<br>عدد الحسابات", total_accounts, {"valueformat": ",d"}, THEME["text"]),
        ("📊<br>متوسط المديونية", avg_amount, {"valueformat": ",.0f"}, OPS_MID),
    ], height=160)

    def _wallet_bar_style(fig, title, *, height, xaxis_title="", yaxis_title="", show_legend=False, margin=None):
        """تنسيق شارتات المحفظة: قيم ظاهرة · جريد خفيف · زي نشاط المحصلين."""
        _apply_ops_chart_style(
            fig, title, height=height, xaxis_title=xaxis_title, yaxis_title=yaxis_title,
            show_legend=show_legend, margin=margin or dict(t=56, b=48, l=120, r=56),
        )
        fig.update_layout(
            uniformtext_minsize=10,
            uniformtext_mode=False,
            xaxis=dict(
                showgrid=True, gridcolor="rgba(128,128,128,0.12)", gridwidth=1,
                zeroline=False, automargin=True, title_font_size=12,
            ),
            yaxis=dict(
                showgrid=False, zeroline=False, automargin=True, title_font_size=12,
            ),
            bargap=0.28,
        )
        fig.update_traces(
            texttemplate="%{x:,.0f}" if fig.data and getattr(fig.data[0], "orientation", None) == "h" else "%{y:,.0f}",
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=12, color=THEME["text"], family="Tahoma, Segoe UI, Arial, sans-serif"),
            marker_line_width=0,
            constraintext="none",
        )
        return fig

    if sub_col and sub_col in work.columns:
        state_counts = (
            work[sub_col].astype(str).str.strip().value_counts().head(12)
            .sort_values(ascending=True).reset_index()
        )
        state_counts.columns = ["الحالة", "عدد الحسابات"]
        fig = px.bar(
            state_counts, x="عدد الحسابات", y="الحالة", orientation="h", text="عدد الحسابات",
            color="عدد الحسابات", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE,
        )
        h = 400
        _wallet_bar_style(
            fig, "توزيع حسابات المحفظة حسب الحالة",
            height=h, xaxis_title="عدد الحسابات",
            margin=dict(t=56, b=40, l=140, r=64),
        )
        fig.update_traces(
            hovertemplate="<b>%{y}</b><br>عدد الحسابات: %{x:,.0f}<extra></extra>",
            texttemplate="%{x:,.0f}",
        )
        figs["states"] = fig

        by_state_amt = (
            work.groupby(work[sub_col].astype(str).str.strip(), dropna=False)[WALLET_REMAINING_COL]
            .sum().sort_values(ascending=True).tail(12).reset_index()
        )
        by_state_amt.columns = ["الحالة", "المبلغ"]
        fig2 = px.bar(
            by_state_amt, x="المبلغ", y="الحالة", orientation="h", text="المبلغ",
            color="المبلغ", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE,
        )
        h2 = 400
        _wallet_bar_style(
            fig2, "المديونية حسب الحالة",
            height=h2, xaxis_title="المبلغ",
            margin=dict(t=56, b=40, l=140, r=72),
        )
        fig2.update_traces(
            hovertemplate="<b>%{y}</b><br>المبلغ: %{x:,.0f}<extra></extra>",
            texttemplate="%{x:,.0f}",
        )
        figs["state_amount"] = fig2

    if sales_col and sales_col in work.columns:
        agent_counts = (
            work[sales_col].astype(str).str.strip().value_counts().head(12)
            .sort_values(ascending=True).reset_index()
        )
        agent_counts.columns = ["المحصل", "عدد الحسابات"]
        fig3 = px.bar(
            agent_counts, x="عدد الحسابات", y="المحصل", orientation="h", text="عدد الحسابات",
            color="عدد الحسابات", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE,
        )
        h3 = 400
        _wallet_bar_style(
            fig3, "توزيع حسابات المحفظة حسب المحصل",
            height=h3, xaxis_title="عدد الحسابات",
            margin=dict(t=56, b=40, l=150, r=64),
        )
        fig3.update_traces(
            hovertemplate="<b>%{y}</b><br>عدد الحسابات: %{x:,.0f}<extra></extra>",
            texttemplate="%{x:,.0f}",
        )
        figs["by_agent_count"] = fig3

    if float(work[WALLET_COLLECTED_COL].sum()) > 0 or float(work[WALLET_REMAINING_COL].sum()) > 0:
        bucket_order = [label for _, _, label in WALLET_AGING_BUCKETS] + ["غير محدد"]
        aging_summary = (
            work.groupby(WALLET_AGING_COL, dropna=False)
            .agg(
                تم_التحصيل=(WALLET_COLLECTED_COL, "sum"),
                باقي_المديونية=(WALLET_REMAINING_COL, "sum"),
                عدد_الحسابات=(WALLET_AGING_COL, "size"),
            )
            .reset_index()
        )
        aging_summary["_ord"] = aging_summary[WALLET_AGING_COL].map(
            {lab: i for i, lab in enumerate(bucket_order)}
        ).fillna(99)
        aging_summary = aging_summary.sort_values("_ord")
        aging_long = aging_summary.melt(
            id_vars=[WALLET_AGING_COL],
            value_vars=["تم_التحصيل", "باقي_المديونية"],
            var_name="النوع",
            value_name="المبلغ",
        )
        aging_long["النوع"] = aging_long["النوع"].map({
            "تم_التحصيل": "تم التحصيل",
            "باقي_المديونية": "باقي المديونية",
        })
        fig_aging = px.bar(
            aging_long,
            x=WALLET_AGING_COL,
            y="المبلغ",
            color="النوع",
            barmode="group",
            text="المبلغ",
            color_discrete_map={"تم التحصيل": OPS_POSITIVE, "باقي المديونية": OPS_NEGATIVE},
            template=PLOTLY_TEMPLATE,
            category_orders={WALLET_AGING_COL: bucket_order},
        )
        _wallet_bar_style(
            fig_aging, "عمر الإسناد — التحصيل والمتبقي حسب البوكيت",
            height=400, xaxis_title="عمر الإسناد", yaxis_title="",
            show_legend=True,
            margin=dict(t=56, b=88, l=72, r=40),
        )
        fig_aging.update_layout(
            xaxis=dict(showgrid=False, automargin=True, title_standoff=12),
            yaxis=dict(
                showgrid=True, gridcolor="rgba(128,128,128,0.12)", gridwidth=1,
                zeroline=False, automargin=True,
                title=dict(text="المبلغ", standoff=18),
                tickfont=dict(size=11),
                tickformat="~s",
            ),
            legend=dict(orientation="h", yanchor="top", y=-0.24, x=0.5, xanchor="center", bgcolor="rgba(0,0,0,0)"),
        )
        fig_aging.update_traces(
            texttemplate="%{y:,.0f}",
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=12, color=THEME["text"], family="Tahoma, Segoe UI, Arial, sans-serif"),
            marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,.0f}<extra></extra>",
        )
        figs["aging"] = fig_aging

    state_src_col = customer_state_col or sub_col
    if state_src_col and state_src_col in work.columns:
        cust_state = (
            work[state_src_col].astype(str).str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "none": pd.NA, "null": pd.NA})
            .dropna().value_counts().head(10).reset_index()
        )
        cust_state.columns = ["حالة العميل", "العدد"]
        if not cust_state.empty:
            fig_pie = px.pie(
                cust_state, values="العدد", names="حالة العميل",
                color="حالة العميل", color_discrete_sequence=OPS_SCALE + ACTIVITY_AGENT_PALETTE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig_pie, "توزيع حالة العميل", height=360, margin=dict(t=56, b=24, l=24, r=24),
                show_legend=False,
            )
            fig_pie.update_traces(
                texttemplate="%{label}<br>%{value:,.0f}<br>%{percent:.1%}",
                textposition="outside",
                textfont=dict(size=11, color=THEME["text"], family="Tahoma, Segoe UI, Arial, sans-serif"),
                hovertemplate="<b>%{label}</b><br>العدد: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
            )
            fig_pie.update_layout(showlegend=False, uniformtext_minsize=9, uniformtext_mode=False)
            figs["customer_state_pie"] = fig_pie

    if nationality_col and nationality_col in work.columns:
        nation = (
            work[nationality_col].astype(str).str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "none": pd.NA, "null": pd.NA})
            .dropna().value_counts().head(10).reset_index()
        )
        nation.columns = ["الجنسية", "العدد"]
        if not nation.empty:
            fig_donut = px.pie(
                nation, values="العدد", names="الجنسية", hole=0.52,
                color="الجنسية", color_discrete_sequence=OPS_SCALE + ACTIVITY_AGENT_PALETTE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig_donut, "توزيع جنسية العميل", height=360, margin=dict(t=56, b=24, l=24, r=24),
                show_legend=False,
            )
            fig_donut.update_traces(
                texttemplate="%{label}<br>%{value:,.0f}<br>%{percent:.1%}",
                textposition="outside",
                textfont=dict(size=11, color=THEME["text"], family="Tahoma, Segoe UI, Arial, sans-serif"),
                hovertemplate="<b>%{label}</b><br>العدد: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
            )
            fig_donut.update_layout(showlegend=False, uniformtext_minsize=9, uniformtext_mode=False)
            figs["nationality_donut"] = fig_donut

    aging_table = _build_wallet_aging_table(work, sub_col)
    meta = {
        **cols,
        "total_accounts": total_accounts,
        "total_amount": total_amount,
        "total_collected": total_collected,
        "agent_count": agent_count,
        "customer_count": customer_count,
        "avg_amount": avg_amount,
    }
    return figs, work, aging_table, meta


def render_wallet_page(df):
    st.subheader("💼 تحليل المحفظة الكاملة")
    prepared, cols = _wallet_prepare_frame(df)

    with st.expander("🔎 فلاتر المحفظة", expanded=True):
        filtered, hint = _wallet_apply_slicers(prepared, cols)
    if hint:
        st.caption(f"الفلاتر: {hint}")
    st.caption(f"المعروض: {len(filtered):,} / {len(prepared):,} حساب")

    figs, work, aging_table, meta = _build_wallet_analysis(filtered)

    # لا نفرض ارتفاع ثابت يقص القيم — كل شارت بيحافظ على ارتفاعه المناسب
    st.plotly_chart(figs["kpi"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_kpi")

    col1, col2 = st.columns(2)
    with col1:
        if "states" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["states"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_states")
    with col2:
        if "state_amount" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["state_amount"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_state_amount")

    col3, col4 = st.columns(2)
    with col3:
        if "by_agent_count" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["by_agent_count"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_agent_count")
        else:
            st.info("لا يوجد عمود محصل.")
    with col4:
        if "aging" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["aging"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_aging")
        else:
            st.info("لا يتوفر عمود عمر الاسناد في ملف المحفظة، أو لا توجد مبالغ للعرض.")

    c5, c6 = st.columns(2)
    with c5:
        if "customer_state_pie" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["customer_state_pie"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_customer_state")
        else:
            st.info("لا يوجد عمود لحالة العميل.")
    with c6:
        if "nationality_donut" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["nationality_donut"], use_container_width=True, config=PLOTLY_CONFIG, key="wallet_nationality")
        else:
            st.info("لا يوجد عمود لجنسية العميل.")

    st.subheader("📋 ماتريكس عمر الإسناد")
    st.caption("صفوف: عمر الإسناد · أعمدة: الحالات · القيم: عدد الحسابات (+ إجمالي التحصيل والمتبقي)")
    if aging_table is not None and not aging_table.empty:
        display_table = aging_table.copy()
        for col_name in display_table.columns:
            if col_name == "عمر الإسناد":
                continue
            display_table[col_name] = pd.to_numeric(display_table[col_name], errors="coerce").fillna(0).map(
                lambda v: f"{v:,.0f}"
            )
        st.dataframe(display_table, use_container_width=True, hide_index=True)

        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
            aging_table.to_excel(writer, index=False, sheet_name="ماتريكس_عمر_الاسناد")
        st.download_button(
            "⬇️ تحميل الماتريكس",
            data=out_excel.getvalue(),
            file_name="wallet_aging_matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="wallet_aging_download",
        )
    else:
        st.info("لا توجد بيانات لبناء ماتريكس عمر الإسناد.")

    with st.expander("📋 عرض بيانات المحفظة"):
        st.dataframe(work, use_container_width=True, hide_index=True)



def _build_payments_analysis(df):
    """تحليل كامل لملف السداد: KPIs + اتجاه زمني + أعلى المحصلين تحصيلاً + توزيع المبالغ."""
    collector_col = find_column(df, COLLECTED_BY_CANDIDATES) or find_column(df, SALES_PERSON_CANDIDATES)
    date_col = find_column(df, SCHEDULE_PAYMENT_DATE_CANDIDATES) or find_column(df, CREATED_ON_CANDIDATES)
    amount_col = find_column(df, CASE_PAYMENT_CANDIDATES) or find_column(df, PROMISE_NET_AMOUNT_CANDIDATES)
    work = df.copy()
    if amount_col and amount_col in work.columns:
        work[amount_col] = pd.to_numeric(work[amount_col], errors="coerce").fillna(0)

    total_payments = len(work)
    total_amount = float(work[amount_col].sum()) if amount_col else 0.0
    agent_count = int(work[collector_col].astype(str).str.strip().nunique()) if collector_col and collector_col in work.columns else 0
    avg_payment = (total_amount / total_payments) if total_payments and amount_col else 0.0

    figs = {}
    figs["kpi"] = _build_kpi_figure([
        ("🧾<br>إجمالي عدد السدادات", total_payments, {"valueformat": ",d"}, THEME["text"]),
        ("👥<br>عدد المحصلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("💰<br>إجمالي المبلغ المحصل", total_amount, {"valueformat": ",.0f"}, OPS_DARK),
        ("📊<br>متوسط السداد", avg_payment, {"valueformat": ",.0f"}, OPS_MID),
    ])

    if date_col and date_col in work.columns:
        ts = pd.to_datetime(work[date_col], errors="coerce")
        trend = work.assign(_date=ts.dt.date).dropna(subset=["_date"])
        if not trend.empty:
            if amount_col:
                daily = trend.groupby("_date")[amount_col].sum().reset_index()
                y_col, y_title = amount_col, "إجمالي المبلغ"
            else:
                daily = trend.groupby("_date").size().reset_index(name="عدد السدادات")
                y_col, y_title = "عدد السدادات", "عدد السدادات"
            fig = px.line(daily, x="_date", y=y_col, markers=True, template=PLOTLY_TEMPLATE,
                          color_discrete_sequence=[OPS_DARK])
            _apply_ops_chart_style(fig, "اتجاه السداد اليومي", height=420, xaxis_title="التاريخ",
                                    yaxis_title=y_title, show_legend=False)
            fig.update_traces(hovertemplate="<b>%{x}</b><br>" + y_title + ": %{y:,.0f}<extra></extra>")
            figs["trend"] = fig

    if collector_col and collector_col in work.columns:
        counts = (
            work[collector_col].astype(str).str.strip().value_counts().head(15)
            .sort_values(ascending=True).reset_index()
        )
        counts.columns = ["المحصّل", "عدد السدادات"]
        fig2 = px.bar(counts, x="عدد السدادات", y="المحصّل", orientation="h", text="عدد السدادات",
                      color="عدد السدادات", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE)
        _apply_ops_chart_style(fig2, "توزيع عدد السدادات حسب المحصل", height=max(430, 36 * len(counts) + 120),
                                xaxis_title="عدد السدادات", show_legend=False, margin=dict(t=70, b=55, l=170, r=60))
        fig2.update_traces(texttemplate="%{x:,.0f}", textposition="outside", cliponaxis=False, marker_line_width=0,
                            hovertemplate="<b>%{y}</b><br>عدد السدادات: %{x:,.0f}<extra></extra>")
        figs["by_agent_count"] = fig2

        if amount_col:
            amt = (
                work.groupby(work[collector_col].astype(str).str.strip())[amount_col].sum()
                .sort_values(ascending=True).tail(15).reset_index()
            )
            amt.columns = ["المحصّل", "إجمالي التحصيل"]
            fig3 = px.bar(amt, x="إجمالي التحصيل", y="المحصّل", orientation="h", text="إجمالي التحصيل",
                          color="إجمالي التحصيل", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE)
            _apply_ops_chart_style(fig3, "أعلى المحصلين تحصيلاً", height=max(430, 36 * len(amt) + 120),
                                    xaxis_title="المبلغ", show_legend=False, margin=dict(t=70, b=55, l=170, r=60))
            fig3.update_traces(texttemplate="%{x:,.0f}", textposition="outside", cliponaxis=False, marker_line_width=0,
                                hovertemplate="<b>%{y}</b><br>المبلغ: %{x:,.0f}<extra></extra>")
            figs["by_agent_amount"] = fig3

    if amount_col and amount_col in work.columns and (work[amount_col] > 0).any():
        fig4 = px.histogram(work[work[amount_col] > 0], x=amount_col, nbins=30,
                             color_discrete_sequence=[OPS_MID], template=PLOTLY_TEMPLATE)
        _apply_ops_chart_style(fig4, "توزيع مبالغ السداد", height=380, xaxis_title="المبلغ",
                                yaxis_title="عدد السدادات", show_legend=False)
        figs["amount_hist"] = fig4

    return figs, work, collector_col, amount_col



def render_payments_page(df):
    st.subheader("💰 تحليل السداد الكامل")
    figs, work, _collector_col, _amount_col = _build_payments_analysis(df)
    st.plotly_chart(figs["kpi"], use_container_width=True, config=PLOTLY_CONFIG, key="payments_kpi")
    if "trend" in figs:
        with st.container(border=True):
            st.plotly_chart(figs["trend"], use_container_width=True, config=PLOTLY_CONFIG, key="payments_trend")
    col1, col2 = st.columns(2)
    with col1:
        if "by_agent_count" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["by_agent_count"], use_container_width=True, config=PLOTLY_CONFIG, key="payments_agent_count")
    with col2:
        if "by_agent_amount" in figs:
            with st.container(border=True):
                st.plotly_chart(figs["by_agent_amount"], use_container_width=True, config=PLOTLY_CONFIG, key="payments_agent_amount")
    if "amount_hist" in figs:
        with st.container(border=True):
            st.plotly_chart(figs["amount_hist"], use_container_width=True, config=PLOTLY_CONFIG, key="payments_hist")
    with st.expander("📋 عرض بيانات السداد"):
        st.dataframe(work, use_container_width=True, hide_index=True)


def _build_activity_payments_link(activity_df, sales_col, class_col, payments_df):
    """يربط نشاط كل محصل (مكالمات/نسبة نجاح) بسداده الفعلي (عدد/مبلغ السدادات) — بدون المحفظة."""
    _, payments_work, collector_col, amount_col = _build_payments_analysis(payments_df)
    if not sales_col or sales_col not in activity_df.columns or not collector_col:
        return None

    act = activity_df.copy()
    act["_agent"] = act[sales_col].astype(str).str.strip()
    act["_success"] = _activity_success_mask(act, class_col)
    agent_activity = act.groupby("_agent").agg(
        عدد_المكالمات=("_agent", "count"),
        عدد_الناجحة=("_success", "sum"),
    ).reset_index()
    agent_activity["نسبة_النجاح"] = (
        agent_activity["عدد_الناجحة"] / agent_activity["عدد_المكالمات"] * 100
    ).round(1)

    pay = payments_work.copy()
    pay["_agent"] = pay[collector_col].astype(str).str.strip()
    if amount_col:
        agent_payments = pay.groupby("_agent").agg(
            عدد_السدادات=("_agent", "count"),
            إجمالي_التحصيل=(amount_col, "sum"),
        ).reset_index()
    else:
        agent_payments = pay.groupby("_agent").agg(عدد_السدادات=("_agent", "count")).reset_index()
        agent_payments["إجمالي_التحصيل"] = 0.0

    merged = pd.merge(agent_activity, agent_payments, on="_agent", how="inner")
    if merged.empty:
        return None
    merged = merged.rename(columns={"_agent": "المحصّل"})
    merged["الفعالية"] = merged.apply(
        lambda r: (r["إجمالي_التحصيل"] / r["عدد_المكالمات"]) if r["عدد_المكالمات"] else 0.0, axis=1
    )

    total_agents = len(merged)
    total_collected = float(merged["إجمالي_التحصيل"].sum())
    total_calls = float(merged["عدد_المكالمات"].sum())
    overall_success = (merged["عدد_الناجحة"].sum() / total_calls * 100) if total_calls else 0.0
    avg_effectiveness = float(merged["الفعالية"].mean()) if total_agents else 0.0

    figs = {}
    figs["kpi"] = _build_kpi_figure([
        ("👥<br>محصلون مشتركون", total_agents, {"valueformat": ",d"}, THEME["text"]),
        ("📞<br>نسبة النجاح الإجمالية", overall_success, {"valueformat": ".1f", "suffix": "%"}, OPS_DARK),
        ("💰<br>إجمالي التحصيل", total_collected, {"valueformat": ",.0f"}, OPS_MID),
        ("⚡<br>متوسط التحصيل للمكالمة", avg_effectiveness, {"valueformat": ",.1f"}, OPS_LIGHT),
    ])

    top = merged.sort_values("إجمالي_التحصيل", ascending=True).tail(15)
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(x=top["عدد_الناجحة"], y=top["المحصّل"], name="مكالمات ناجحة",
                              orientation="h", marker_color=OPS_DARK))
    fig_bar.add_trace(go.Bar(x=top["عدد_السدادات"], y=top["المحصّل"], name="سدادات فعلية",
                              orientation="h", marker_color=OPS_LIGHT))
    _apply_ops_chart_style(fig_bar, "مكالمات ناجحة مقابل سدادات فعلية لكل محصل",
                            height=max(430, 36 * len(top) + 140), xaxis_title="العدد",
                            show_legend=True, margin=dict(t=70, b=90, l=170, r=60))
    fig_bar.update_layout(barmode="group")
    figs["compare_bar"] = fig_bar

    fig_scatter = px.scatter(
        merged, x="نسبة_النجاح", y="إجمالي_التحصيل", size="عدد_المكالمات",
        color="عدد_المكالمات", color_continuous_scale=OPS_SCALE, template=PLOTLY_TEMPLATE,
        hover_name="المحصّل",
    )
    _apply_ops_chart_style(fig_scatter, "نسبة نجاح المكالمات مقابل إجمالي التحصيل",
                            height=460, xaxis_title="نسبة النجاح %", yaxis_title="إجمالي التحصيل",
                            show_legend=False)
    fig_scatter.update_traces(
        hovertemplate="<b>%{hovertext}</b><br>نسبة النجاح: %{x:.1f}%<br>التحصيل: %{y:,.0f}<extra></extra>"
    )
    figs["scatter"] = fig_scatter

    return figs, merged


def render_activity_payments_link_page(activity_df, sales_col, class_col, payments_df):
    st.subheader("🔗 ربط نشاط المحصلين بالسداد")
    result = _build_activity_payments_link(activity_df, sales_col, class_col, payments_df)
    if result is None:
        st.info("تعذّر الربط — تأكد إن عمود المحصل موجود في ملفي النشاط والسداد.")
        return
    figs, merged = result
    st.plotly_chart(figs["kpi"], use_container_width=True, config=PLOTLY_CONFIG, key="link_kpi")
    with st.container(border=True):
        st.plotly_chart(figs["compare_bar"], use_container_width=True, config=PLOTLY_CONFIG, key="link_compare")
    with st.container(border=True):
        st.plotly_chart(figs["scatter"], use_container_width=True, config=PLOTLY_CONFIG, key="link_scatter")
    st.subheader("📋 جدول تفصيلي: النشاط مقابل السداد لكل محصل")
    st.dataframe(merged, use_container_width=True, hide_index=True)


def _fig_html_card(fig, key_prefix, index):
    """يحوّل Plotly figure لبطاقة HTML بنفس أسلوب chart-card، بارتفاع صريح لتفادي مشاكل العرض."""
    height = int(fig.layout.height or 420)
    return (
        "<article class='chart-card'>"
        + pio.to_html(
            fig, full_html=False, include_plotlyjs=False,
            config={"displayModeBar": False, "responsive": True},
            div_id=f"{key_prefix}_{index}", default_width="100%", default_height=f"{height}px",
        )
        + "</article>"
    )


def _build_wallet_page_html(wallet_df):
    """صفحة HTML للمحفظة بنفس أسلوب نشاط المحصلين: سلايسرز + كروت + شارتات + جدول."""
    from html import escape
    import json as _json
    import base64 as _b64

    figs, work, aging_table, meta = _build_wallet_analysis(wallet_df)
    sales_col = meta.get("sales_col")
    sub_col = meta.get("sub_col")
    nationality_col = meta.get("nationality_col")
    customer_state_col = meta.get("customer_state_col") or sub_col

    def _clean(v):
        s = str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else ""
        if s.lower() in {"", "nan", "none", "null"}:
            return ""
        return s

    records = []
    for _, row in work.iterrows():
        ad = row.get("_assign_date")
        dd = row.get("_debit_date")
        records.append({
            "agent": _clean(row[sales_col]) if sales_col and sales_col in work.columns else "",
            "state": _clean(row[sub_col]) if sub_col and sub_col in work.columns else "",
            "customer_state": _clean(row[customer_state_col]) if customer_state_col and customer_state_col in work.columns else "",
            "nationality": _clean(row[nationality_col]) if nationality_col and nationality_col in work.columns else "",
            "aging": _clean(row.get(WALLET_AGING_COL, "غير محدد")) or "غير محدد",
            "collected": float(row.get(WALLET_COLLECTED_COL, 0) or 0),
            "remaining": float(row.get(WALLET_REMAINING_COL, 0) or 0),
            "assign_date": ad.isoformat() if hasattr(ad, "isoformat") else "",
            "debit_date": dd.isoformat() if hasattr(dd, "isoformat") else "",
            "customer": _clean(row[meta["customer_id_col"]]) if meta.get("customer_id_col") and meta["customer_id_col"] in work.columns else "",
        })

    agents = sorted({r["agent"] for r in records if r["agent"]})
    nations = sorted({r["nationality"] for r in records if r["nationality"]})
    cust_states = sorted({r["customer_state"] for r in records if r["customer_state"]})
    aging_vals = sorted({r["aging"] for r in records if r["aging"]})
    assign_dates = sorted(d for d in (r["assign_date"] for r in records) if d)
    debit_dates = sorted(d for d in (r["debit_date"] for r in records) if d)
    assign_min = assign_dates[0] if assign_dates else ""
    assign_max = assign_dates[-1] if assign_dates else ""
    debit_min = debit_dates[0] if debit_dates else ""
    debit_max = debit_dates[-1] if debit_dates else ""

    ops_dark, ops_mid = OPS_DARK, OPS_MID

    def _multi(label, menu_id, label_id, all_cls, opt_cls, values, empty_label):
        html = [
            f"<div class='filter-field' style='position:relative'><span>{label}</span>",
            f"<button type='button' class='multi-trigger' data-target='{menu_id}'><span id='{label_id}'>{empty_label}</span> ⌄</button>",
            f"<div id='{menu_id}' class='multi-menu'>",
            f"<label style='display:block;padding:6px;font-weight:700'><input type='checkbox' class='{all_cls}'> {empty_label}</label>",
        ]
        for value in values:
            html.append(
                f"<label style='display:block;padding:6px'>"
                f"<input type='checkbox' class='{opt_cls}' value='{escape(value, quote=True)}'> {escape(value)}</label>"
            )
        html.append("</div></div>")
        return "".join(html)

    def _chart_card(fig, plot_id, height=None):
        # ارتفاع موحّد + قيم بولد أسود زي نشاط المحصلين (HTML ثيم فاتح)
        h = int(height or (fig.layout.height or 400))
        fig = fig.update_layout(
            height=h,
            title=None,
            font=dict(family="Tahoma, Segoe UI, Arial, sans-serif", size=12, color="#111827"),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        # cliponaxis مدعوم في bar فقط — pie/donut بيرفضوه
        text_style = dict(size=12, color="#111827", family="Tahoma, Segoe UI, Arial, sans-serif")
        for trace in fig.data:
            ttype = getattr(trace, "type", None)
            if ttype in ("bar", "scatter", "histogram"):
                trace.update(textfont=text_style, cliponaxis=False)
            else:
                try:
                    trace.update(textfont=text_style)
                except Exception:
                    pass
        return pio.to_html(
            fig, full_html=False, include_plotlyjs=False,
            config={"displayModeBar": False, "responsive": True},
            div_id=plot_id, default_width="100%", default_height=f"{h}px",
        )

    parts = []
    parts.append("<div id='page-wallet' class='dash-page'>")
    parts.append("<header class='hero' style='margin-bottom:18px'>")
    parts.append("<div class='eyebrow'>WALLET DASHBOARD</div>")
    parts.append("<h1 style='margin:0 0 8px;font-size:24px'>💼 تحليل المحفظة الكاملة</h1>")
    parts.append("<div class='meta'>سلايسرز تفاعلية · كروت وشارتات بنفس تنسيق نشاط المحصلين</div>")
    parts.append("</header>")

    # Filters
    parts.append("<section class='panel'>")
    parts.append("<h2 class='section-title'>🎚️ فلاتر المحفظة</h2>")
    parts.append("<div class='filters-grid' id='wallet-interactive-filters'>")
    parts.append(_multi("👤 المحصل", "wallet-agent-menu", "wallet-agent-label", "wallet-select-all-agent", "wallet-agent-option", agents, "كل المحصلين"))
    parts.append(_multi("🌍 الجنسية", "wallet-nation-menu", "wallet-nation-label", "wallet-select-all-nation", "wallet-nation-option", nations, "كل الجنسيات"))
    parts.append(_multi("🏷️ حالة العميل", "wallet-cstate-menu", "wallet-cstate-label", "wallet-select-all-cstate", "wallet-cstate-option", cust_states, "كل الحالات"))
    parts.append(_multi("⏳ عمر الإسناد", "wallet-aging-menu", "wallet-aging-label", "wallet-select-all-aging", "wallet-aging-option", aging_vals, "كل الأعمار"))
    if assign_min:
        parts.append(f"<div class='filter-field'><span>📅 تاريخ الإسناد من</span><input id='wallet-assign-from' type='date' value='{assign_min}' min='{assign_min}' max='{assign_max}'></div>")
        parts.append(f"<div class='filter-field'><span>📅 تاريخ الإسناد إلى</span><input id='wallet-assign-to' type='date' value='{assign_max}' min='{assign_min}' max='{assign_max}'></div>")
    if debit_min:
        parts.append(f"<div class='filter-field'><span>📅 تاريخ الحادث من</span><input id='wallet-debit-from' type='date' value='{debit_min}' min='{debit_min}' max='{debit_max}'></div>")
        parts.append(f"<div class='filter-field'><span>📅 تاريخ الحادث إلى</span><input id='wallet-debit-to' type='date' value='{debit_max}' min='{debit_min}' max='{debit_max}'></div>")
    parts.append("<div class='filter-field'><span>&nbsp;</span><button id='wallet-reset-filters' class='btn-reset' type='button'>↺ إعادة ضبط</button></div>")
    parts.append("</div>")
    parts.append("<div id='wallet-filter-status' class='meta' style='margin-top:12px;text-align:center'></div>")
    parts.append("</section>")

    # KPIs
    parts.append("<section class='kpi-grid'>")
    for kid, label, val, color in [
        ("wallet-kpi-amount", "💰 إجمالي المديونية", f"{meta.get('total_amount', 0):,.0f}", ops_dark),
        ("wallet-kpi-agents", "👥 عدد المحصلين", f"{meta.get('agent_count', 0):,}", None),
        ("wallet-kpi-customers", "🧑 عدد العملاء", f"{meta.get('customer_count', 0):,}", None),
        ("wallet-kpi-accounts", "📋 عدد الحسابات", f"{meta.get('total_accounts', 0):,}", None),
        ("wallet-kpi-avg", "📊 متوسط المديونية", f"{meta.get('avg_amount', 0):,.0f}", ops_mid),
    ]:
        style = f" style='color:{color}'" if color else ""
        parts.append(f"<div class='kpi'><div class='label'>{label}</div><div class='value' id='{kid}'{style}>{val}</div></div>")
    parts.append("</section>")

    plot_id_map = {}
    chart_idx = 1

    def take(key, title, subtitle):
        nonlocal chart_idx
        if key not in figs:
            return ""
        pid = f"wallet_{chart_idx}"
        plot_id_map[key] = pid
        chart_idx += 1
        return (
            "<section class='panel' style='margin:0'>"
            f"<h2 class='section-title' style='margin:4px 0 6px'>{title}</h2>"
            f"<div class='meta' style='text-align:center;margin:0 0 10px;font-size:12px'>{escape(subtitle)}</div>"
            f"<div class='chart-card' style='border:none;box-shadow:none;padding:4px 0 0;margin:0'>"
            + _chart_card(figs[key], pid)
            + "</div></section>"
        )

    def _append_chart_row(items):
        items = [x for x in items if x]
        if not items:
            return
        parts.append("<div class='wallet-charts-row'>" + "".join(items) + "</div>")

    _append_chart_row([
        take("states", "📊 توزيع الحسابات حسب الحالة", "عدد الحسابات لكل Sub State"),
        take("state_amount", "💰 المديونية حسب الحالة", "إجمالي Net Amount لكل حالة"),
    ])
    _append_chart_row([
        take("by_agent_count", "👤 توزيع الحسابات حسب المحصل", "عدد الحسابات المسندة لكل محصل"),
        take("aging", "⏳ عمر الإسناد — التحصيل والمتبقي", "Payment مقابل الباقي حسب عمود عمر الإسناد"),
    ])
    _append_chart_row([
        take("customer_state_pie", "🏷️ توزيع حالة العميل", "نسب حالات العميل"),
        take("nationality_donut", "🌍 توزيع جنسية العميل", "نسب جنسيات العملاء"),
    ])

    # Matrix table (filled by JS on load/filter — Power BI style)
    parts.append("<section class='panel'>")
    parts.append("<h2 class='section-title'>📋 ماتريكس عمر الإسناد</h2>")
    parts.append("<div class='meta' style='text-align:center;margin-bottom:8px'>صفوف: عمر الإسناد · أعمدة: الحالات · القيم: عدد الحسابات · مع إجمالي التحصيل والمتبقي</div>")
    parts.append("<div class='table-wrap'><table class='data-table' id='wallet-matrix-table'>")
    parts.append("<thead id='wallet-matrix-thead'></thead>")
    parts.append("<tbody id='wallet-matrix-tbody'></tbody>")
    parts.append("</table></div></section>")
    parts.append("</div>")  # page-wallet

    js_tpl = _b64.b64decode("CihmdW5jdGlvbigpewpjb25zdCB3YWxsZXREYXRhID0gX19XQUxMRVRfREFUQV9fOwpjb25zdCB3YWxsZXRQbG90SWRzID0gX19XQUxMRVRfUExPVF9JRFNfXzsKY29uc3QgV19EQVJLID0gIl9fT1BTX0RBUktfXyI7CmNvbnN0IFdfTUlEID0gIl9fT1BTX01JRF9fIjsKY29uc3QgV19MSUdIVCA9ICJfX09QU19MSUdIVF9fIjsKY29uc3QgV19QT1MgPSAiX19PUFNfUE9TX18iOwpjb25zdCBXX05FRyA9ICJfX09QU19ORUdfXyI7CmNvbnN0IFdfQVNTSUdOX01JTiA9ICJfX0FTU0lHTl9NSU5fXyI7CmNvbnN0IFdfQVNTSUdOX01BWCA9ICJfX0FTU0lHTl9NQVhfXyI7CmNvbnN0IFdfREVCSVRfTUlOID0gIl9fREVCSVRfTUlOX18iOwpjb25zdCBXX0RFQklUX01BWCA9ICJfX0RFQklUX01BWF9fIjsKCmZ1bmN0aW9uIHdGbXQobil7IHJldHVybiBOdW1iZXIobnx8MCkudG9Mb2NhbGVTdHJpbmcoImVuLVVTIik7IH0KZnVuY3Rpb24gd0ZtdDAobil7IHJldHVybiBOdW1iZXIobnx8MCkudG9Mb2NhbGVTdHJpbmcoImVuLVVTIiwge21heGltdW1GcmFjdGlvbkRpZ2l0czowfSk7IH0KZnVuY3Rpb24gd1NldEtwaShpZCwgdmFsKXsgY29uc3QgZWw9ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpOyBpZihlbCkgZWwudGV4dENvbnRlbnQgPSB2YWw7IH0KZnVuY3Rpb24gd0NoZWNrZWQoY2xzKXsKICByZXR1cm4gQXJyYXkucHJvdG90eXBlLnNsaWNlLmNhbGwoZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgiLiIrY2xzKyI6Y2hlY2tlZCIpKS5tYXAoZnVuY3Rpb24obyl7IHJldHVybiBvLnZhbHVlOyB9KTsKfQoKZnVuY3Rpb24gd1NlbGVjdGVkUm93cygpewogIHZhciByb3dzID0gd2FsbGV0RGF0YS5zbGljZSgpOwogIHZhciBhZ2VudHMgPSB3Q2hlY2tlZCgid2FsbGV0LWFnZW50LW9wdGlvbiIpOwogIGlmIChhZ2VudHMubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIoZnVuY3Rpb24ocil7IHJldHVybiBhZ2VudHMuaW5kZXhPZihyLmFnZW50KSA+PSAwOyB9KTsKICB2YXIgbmF0aW9ucyA9IHdDaGVja2VkKCJ3YWxsZXQtbmF0aW9uLW9wdGlvbiIpOwogIGlmIChuYXRpb25zLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKGZ1bmN0aW9uKHIpeyByZXR1cm4gbmF0aW9ucy5pbmRleE9mKHIubmF0aW9uYWxpdHkpID49IDA7IH0pOwogIHZhciBjc3RhdGVzID0gd0NoZWNrZWQoIndhbGxldC1jc3RhdGUtb3B0aW9uIik7CiAgaWYgKGNzdGF0ZXMubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIoZnVuY3Rpb24ocil7IHJldHVybiBjc3RhdGVzLmluZGV4T2Yoci5jdXN0b21lcl9zdGF0ZSkgPj0gMDsgfSk7CiAgdmFyIGFnaW5nID0gd0NoZWNrZWQoIndhbGxldC1hZ2luZy1vcHRpb24iKTsKICBpZiAoYWdpbmcubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIoZnVuY3Rpb24ocil7IHJldHVybiBhZ2luZy5pbmRleE9mKHIuYWdpbmcpID49IDA7IH0pOwogIHZhciBhZiA9IChkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LWFzc2lnbi1mcm9tIikgfHwge30pLnZhbHVlIHx8ICIiOwogIHZhciBhdCA9IChkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LWFzc2lnbi10byIpIHx8IHt9KS52YWx1ZSB8fCAiIjsKICBpZiAoYWYpIHJvd3MgPSByb3dzLmZpbHRlcihmdW5jdGlvbihyKXsgcmV0dXJuICFyLmFzc2lnbl9kYXRlIHx8IHIuYXNzaWduX2RhdGUgPj0gYWY7IH0pOwogIGlmIChhdCkgcm93cyA9IHJvd3MuZmlsdGVyKGZ1bmN0aW9uKHIpeyByZXR1cm4gIXIuYXNzaWduX2RhdGUgfHwgci5hc3NpZ25fZGF0ZSA8PSBhdDsgfSk7CiAgdmFyIGRmID0gKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtZGViaXQtZnJvbSIpIHx8IHt9KS52YWx1ZSB8fCAiIjsKICB2YXIgZHQgPSAoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoIndhbGxldC1kZWJpdC10byIpIHx8IHt9KS52YWx1ZSB8fCAiIjsKICBpZiAoZGYpIHJvd3MgPSByb3dzLmZpbHRlcihmdW5jdGlvbihyKXsgcmV0dXJuICFyLmRlYml0X2RhdGUgfHwgci5kZWJpdF9kYXRlID49IGRmOyB9KTsKICBpZiAoZHQpIHJvd3MgPSByb3dzLmZpbHRlcihmdW5jdGlvbihyKXsgcmV0dXJuICFyLmRlYml0X2RhdGUgfHwgci5kZWJpdF9kYXRlIDw9IGR0OyB9KTsKICByZXR1cm4gcm93czsKfQoKZnVuY3Rpb24gd0NvdW50TWFwKHJvd3MsIGtleSl7CiAgdmFyIG0gPSB7fTsKICByb3dzLmZvckVhY2goZnVuY3Rpb24ocil7IHZhciBrID0gcltrZXldIHx8ICLYutmK2LEg2YXYrdiv2K8iOyBtW2tdID0gKG1ba118fDApICsgMTsgfSk7CiAgcmV0dXJuIG07Cn0KZnVuY3Rpb24gd1N1bU1hcChyb3dzLCBrZXksIHZhbEtleSl7CiAgdmFyIG0gPSB7fTsKICByb3dzLmZvckVhY2goZnVuY3Rpb24ocil7IHZhciBrID0gcltrZXldIHx8ICLYutmK2LEg2YXYrdiv2K8iOyBtW2tdID0gKG1ba118fDApICsgKHJbdmFsS2V5XXx8MCk7IH0pOwogIHJldHVybiBtOwp9CmZ1bmN0aW9uIHdUb3BFbnRyaWVzKG1hcCwgbil7CiAgcmV0dXJuIE9iamVjdC5rZXlzKG1hcCkubWFwKGZ1bmN0aW9uKGspeyByZXR1cm4gW2ssIG1hcFtrXV07IH0pLnNvcnQoZnVuY3Rpb24oYSxiKXsgcmV0dXJuIGFbMV0tYlsxXTsgfSkuc2xpY2UoLW4pOwp9CgpmdW5jdGlvbiB3VXBkYXRlTGFiZWxzKCl7CiAgdmFyIGEgPSB3Q2hlY2tlZCgid2FsbGV0LWFnZW50LW9wdGlvbiIpOwogIHZhciBhbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtYWdlbnQtbGFiZWwiKTsKICBpZiAoYWwpIGFsLnRleHRDb250ZW50ID0gYS5sZW5ndGggPyAoYS5sZW5ndGggKyAiINmF2K3YtdmEINmF2K3Yr9ivIikgOiAi2YPZhCDYp9mE2YXYrdi12YTZitmGIjsKICB2YXIgbiA9IHdDaGVja2VkKCJ3YWxsZXQtbmF0aW9uLW9wdGlvbiIpOwogIHZhciBubCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtbmF0aW9uLWxhYmVsIik7CiAgaWYgKG5sKSBubC50ZXh0Q29udGVudCA9IG4ubGVuZ3RoID8gKG4ubGVuZ3RoICsgIiDYrNmG2LPZitipIikgOiAi2YPZhCDYp9mE2KzZhtiz2YrYp9iqIjsKICB2YXIgYyA9IHdDaGVja2VkKCJ3YWxsZXQtY3N0YXRlLW9wdGlvbiIpOwogIHZhciBjbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtY3N0YXRlLWxhYmVsIik7CiAgaWYgKGNsKSBjbC50ZXh0Q29udGVudCA9IGMubGVuZ3RoID8gKGMubGVuZ3RoICsgIiDYrdin2YTYqSIpIDogItmD2YQg2KfZhNit2KfZhNin2KoiOwogIHZhciBnID0gd0NoZWNrZWQoIndhbGxldC1hZ2luZy1vcHRpb24iKTsKICB2YXIgZ2wgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LWFnaW5nLWxhYmVsIik7CiAgaWYgKGdsKSBnbC50ZXh0Q29udGVudCA9IGcubGVuZ3RoID8gKGcubGVuZ3RoICsgIiDYudmF2LEiKSA6ICLZg9mEINin2YTYo9i52YXYp9ixIjsKfQoKZnVuY3Rpb24gd0JhckgocGxvdCwgZW50cmllcywgaG92ZXJTdWZmaXgpewogIGlmICghcGxvdCkgcmV0dXJuOwogIHZhciBoID0gNDAwOwogIFBsb3RseS5yZWFjdChwbG90LCBbewogICAgdHlwZTogImJhciIsIG9yaWVudGF0aW9uOiAiaCIsCiAgICB5OiBlbnRyaWVzLm1hcChmdW5jdGlvbihlKXsgcmV0dXJuIGVbMF07IH0pLAogICAgeDogZW50cmllcy5tYXAoZnVuY3Rpb24oZSl7IHJldHVybiBlWzFdOyB9KSwKICAgIHRleHQ6IGVudHJpZXMubWFwKGZ1bmN0aW9uKGUpeyByZXR1cm4gZVsxXTsgfSksCiAgICB0ZXh0dGVtcGxhdGU6ICIle3g6LC4wZn0iLCB0ZXh0cG9zaXRpb246ICJvdXRzaWRlIiwgY2xpcG9uYXhpczogZmFsc2UsCiAgICB0ZXh0Zm9udDoge3NpemU6IDEyLCBjb2xvcjogIiMxMTE4MjciLCBmYW1pbHk6ICJUYWhvbWEsIFNlZ29lIFVJLCBBcmlhbCwgc2Fucy1zZXJpZiJ9LAogICAgbWFya2VyOiB7Y29sb3I6IGVudHJpZXMubWFwKGZ1bmN0aW9uKF8saSl7IHJldHVybiBbV19EQVJLLFdfTUlELFdfTElHSFRdW2klM107IH0pLCBsaW5lOiB7d2lkdGg6IDB9fSwKICAgIGhvdmVydGVtcGxhdGU6ICI8Yj4le3l9PC9iPjxicj4iICsgaG92ZXJTdWZmaXggKyAiOiAle3g6LC4wZn08ZXh0cmE+PC9leHRyYT4iCiAgfV0sIE9iamVjdC5hc3NpZ24oe30sIHBsb3QubGF5b3V0IHx8IHt9LCB7CiAgICBzaG93bGVnZW5kOiBmYWxzZSwgaGVpZ2h0OiBoLAogICAgZm9udDoge2ZhbWlseTogIlRhaG9tYSwgU2Vnb2UgVUksIEFyaWFsLCBzYW5zLXNlcmlmIiwgc2l6ZTogMTIsIGNvbG9yOiAiIzExMTgyNyJ9LAogICAgbWFyZ2luOiBPYmplY3QuYXNzaWduKHt9LCAocGxvdC5sYXlvdXQgJiYgcGxvdC5sYXlvdXQubWFyZ2luKSB8fCB7fSwge3Q6IDQwLCBiOiA0MCwgbDogMTQwLCByOiA3Mn0pLAogICAgeGF4aXM6IE9iamVjdC5hc3NpZ24oe30sIChwbG90LmxheW91dCAmJiBwbG90LmxheW91dC54YXhpcykgfHwge30sIHsKICAgICAgc2hvd2dyaWQ6IHRydWUsIGdyaWRjb2xvcjogInJnYmEoMTI4LDEyOCwxMjgsMC4xMikiLCBncmlkd2lkdGg6IDEsIHplcm9saW5lOiBmYWxzZSwgYXV0b21hcmdpbjogdHJ1ZQogICAgfSksCiAgICB5YXhpczogT2JqZWN0LmFzc2lnbih7fSwgKHBsb3QubGF5b3V0ICYmIHBsb3QubGF5b3V0LnlheGlzKSB8fCB7fSwgewogICAgICBzaG93Z3JpZDogZmFsc2UsIHplcm9saW5lOiBmYWxzZSwgYXV0b21hcmdpbjogdHJ1ZQogICAgfSkKICB9KSk7Cn0KCmZ1bmN0aW9uIHdCdWlsZE1hdHJpeChyb3dzKXsKICB2YXIgc3RhdGVzID0ge307CiAgdmFyIGFnaW5ncyA9IHt9OwogIHZhciBjZWxsID0ge307CiAgdmFyIGNvbGxlY3RlZCA9IHt9OwogIHZhciByZW1haW5pbmcgPSB7fTsKICByb3dzLmZvckVhY2goZnVuY3Rpb24ocil7CiAgICB2YXIgYSA9IHIuYWdpbmcgfHwgIti62YrYsSDZhdit2K/YryI7CiAgICB2YXIgcyA9IHIuc3RhdGUgfHwgIuKAlCI7CiAgICBzdGF0ZXNbc10gPSAxOwogICAgYWdpbmdzW2FdID0gMTsKICAgIHZhciBrID0gYSArICJ8fCIgKyBzOwogICAgY2VsbFtrXSA9IChjZWxsW2tdfHwwKSArIDE7CiAgICBjb2xsZWN0ZWRbYV0gPSAoY29sbGVjdGVkW2FdfHwwKSArIChyLmNvbGxlY3RlZHx8MCk7CiAgICByZW1haW5pbmdbYV0gPSAocmVtYWluaW5nW2FdfHwwKSArIChyLnJlbWFpbmluZ3x8MCk7CiAgfSk7CiAgdmFyIHN0YXRlTGlzdCA9IE9iamVjdC5rZXlzKHN0YXRlcykuc29ydChmdW5jdGlvbihhLGIpewogICAgdmFyIHRhPTAsIHRiPTA7CiAgICBPYmplY3Qua2V5cyhhZ2luZ3MpLmZvckVhY2goZnVuY3Rpb24oYWcpeyB0YSArPSBjZWxsW2FnKyJ8fCIrYV18fDA7IHRiICs9IGNlbGxbYWcrInx8IitiXXx8MDsgfSk7CiAgICByZXR1cm4gdGIgLSB0YTsKICB9KTsKICB2YXIgYWdpbmdMaXN0ID0gT2JqZWN0LmtleXMoYWdpbmdzKS5zb3J0KCk7CiAgdmFyIGhlYWQgPSAiPHRyPjx0aD7YudmF2LEg2KfZhNil2LPZhtin2K88L3RoPiIgKyBzdGF0ZUxpc3QubWFwKGZ1bmN0aW9uKHMpeyByZXR1cm4gIjx0aD4iK3MrIjwvdGg+IjsgfSkuam9pbigiIikgKwogICAgICAgICAgICAgIjx0aD7Yp9mE2KXYrNmF2KfZhNmKPC90aD48dGg+2KrZhSDYp9mE2KrYrdi12YrZhDwvdGg+PHRoPtio2KfZgtmKINin2YTZhdiv2YrZiNmG2YrYqTwvdGg+PC90cj4iOwogIHZhciBib2R5ID0gYWdpbmdMaXN0Lm1hcChmdW5jdGlvbihhKXsKICAgIHZhciByb3dUb3RhbCA9IDA7CiAgICB2YXIgdGRzID0gc3RhdGVMaXN0Lm1hcChmdW5jdGlvbihzKXsKICAgICAgdmFyIHYgPSBjZWxsW2ErInx8IitzXXx8MDsKICAgICAgcm93VG90YWwgKz0gdjsKICAgICAgcmV0dXJuICI8dGQ+IiArICh2ID8gd0ZtdCh2KSA6ICLigJQiKSArICI8L3RkPiI7CiAgICB9KS5qb2luKCIiKTsKICAgIHJldHVybiAiPHRyPjx0ZD4iK2ErIjwvdGQ+Iit0ZHMrIjx0ZD48Yj4iK3dGbXQocm93VG90YWwpKyI8L2I+PC90ZD48dGQ+Iit3Rm10MChjb2xsZWN0ZWRbYV18fDApKyI8L3RkPjx0ZD4iK3dGbXQwKHJlbWFpbmluZ1thXXx8MCkrIjwvdGQ+PC90cj4iOwogIH0pLmpvaW4oIiIpOwogIHZhciBjb2xUb3RhbHMgPSBzdGF0ZUxpc3QubWFwKGZ1bmN0aW9uKHMpewogICAgdmFyIHQgPSAwOwogICAgYWdpbmdMaXN0LmZvckVhY2goZnVuY3Rpb24oYSl7IHQgKz0gY2VsbFthKyJ8fCIrc118fDA7IH0pOwogICAgcmV0dXJuIHQ7CiAgfSk7CiAgdmFyIGdyYW5kID0gY29sVG90YWxzLnJlZHVjZShmdW5jdGlvbihzLHYpeyByZXR1cm4gcyt2OyB9LCAwKTsKICB2YXIgc3VtQyA9IGFnaW5nTGlzdC5yZWR1Y2UoZnVuY3Rpb24ocyxhKXsgcmV0dXJuIHMrKGNvbGxlY3RlZFthXXx8MCk7IH0sIDApOwogIHZhciBzdW1SID0gYWdpbmdMaXN0LnJlZHVjZShmdW5jdGlvbihzLGEpeyByZXR1cm4gcysocmVtYWluaW5nW2FdfHwwKTsgfSwgMCk7CiAgdmFyIGZvb3QgPSAiPHRyIHN0eWxlPSdmb250LXdlaWdodDo3MDA7YmFja2dyb3VuZDojRTdGMUYxJz48dGQ+2KfZhNil2KzZhdin2YTZijwvdGQ+IiArCiAgICAgICAgICAgICBjb2xUb3RhbHMubWFwKGZ1bmN0aW9uKHYpeyByZXR1cm4gIjx0ZD4iK3dGbXQodikrIjwvdGQ+IjsgfSkuam9pbigiIikgKwogICAgICAgICAgICAgIjx0ZD4iK3dGbXQoZ3JhbmQpKyI8L3RkPjx0ZD4iK3dGbXQwKHN1bUMpKyI8L3RkPjx0ZD4iK3dGbXQwKHN1bVIpKyI8L3RkPjwvdHI+IjsKICByZXR1cm4geyBoZWFkOiBoZWFkLCBib2R5OiBib2R5ICsgZm9vdCB9Owp9CgpmdW5jdGlvbiB3UmVmcmVzaCgpewogIHZhciByb3dzID0gd1NlbGVjdGVkUm93cygpOwogIHZhciBhZ2VudHMgPSBbXSwgc2VlbkEgPSB7fTsKICByb3dzLmZvckVhY2goZnVuY3Rpb24ocil7IGlmIChyLmFnZW50ICYmICFzZWVuQVtyLmFnZW50XSkgeyBzZWVuQVtyLmFnZW50XT0xOyBhZ2VudHMucHVzaChyLmFnZW50KTsgfSB9KTsKICB2YXIgY3VzdG9tZXJzID0gW10sIHNlZW5DID0ge307CiAgcm93cy5mb3JFYWNoKGZ1bmN0aW9uKHIpeyBpZiAoci5jdXN0b21lciAmJiAhc2VlbkNbci5jdXN0b21lcl0pIHsgc2VlbkNbci5jdXN0b21lcl09MTsgY3VzdG9tZXJzLnB1c2goci5jdXN0b21lcik7IH0gfSk7CiAgdmFyIHJlbWFpbmluZyA9IHJvd3MucmVkdWNlKGZ1bmN0aW9uKHMscil7IHJldHVybiBzICsgKHIucmVtYWluaW5nfHwwKTsgfSwgMCk7CiAgdmFyIGF2ZyA9IHJvd3MubGVuZ3RoID8gcmVtYWluaW5nIC8gcm93cy5sZW5ndGggOiAwOwogIHdTZXRLcGkoIndhbGxldC1rcGktYW1vdW50Iiwgd0ZtdDAocmVtYWluaW5nKSk7CiAgd1NldEtwaSgid2FsbGV0LWtwaS1hZ2VudHMiLCB3Rm10KGFnZW50cy5sZW5ndGgpKTsKICB3U2V0S3BpKCJ3YWxsZXQta3BpLWN1c3RvbWVycyIsIHdGbXQoY3VzdG9tZXJzLmxlbmd0aCB8fCByb3dzLmxlbmd0aCkpOwogIHdTZXRLcGkoIndhbGxldC1rcGktYWNjb3VudHMiLCB3Rm10KHJvd3MubGVuZ3RoKSk7CiAgd1NldEtwaSgid2FsbGV0LWtwaS1hdmciLCB3Rm10MChhdmcpKTsKICB2YXIgc3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LWZpbHRlci1zdGF0dXMiKTsKICBpZiAoc3QpIHN0LnRleHRDb250ZW50ID0gIti52LHYtiAiICsgd0ZtdChyb3dzLmxlbmd0aCkgKyAiINit2LPYp9ioINmF2YYg2KPYtdmEICIgKyB3Rm10KHdhbGxldERhdGEubGVuZ3RoKSArICIgIHwgICIgKyB3Rm10KGFnZW50cy5sZW5ndGgpICsgIiDZhdit2LXZhCI7CgogIHdCYXJIKHdhbGxldFBsb3RJZHMuc3RhdGVzICYmIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKHdhbGxldFBsb3RJZHMuc3RhdGVzKSwgd1RvcEVudHJpZXMod0NvdW50TWFwKHJvd3MsInN0YXRlIiksIDEyKSwgIti52K/YryDYp9mE2K3Ys9in2KjYp9iqIik7CiAgd0Jhckgod2FsbGV0UGxvdElkcy5zdGF0ZV9hbW91bnQgJiYgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod2FsbGV0UGxvdElkcy5zdGF0ZV9hbW91bnQpLCB3VG9wRW50cmllcyh3U3VtTWFwKHJvd3MsInN0YXRlIiwicmVtYWluaW5nIiksIDEyKSwgItin2YTZhdio2YTYuiIpOwogIHdCYXJIKHdhbGxldFBsb3RJZHMuYnlfYWdlbnRfY291bnQgJiYgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod2FsbGV0UGxvdElkcy5ieV9hZ2VudF9jb3VudCksIHdUb3BFbnRyaWVzKHdDb3VudE1hcChyb3dzLCJhZ2VudCIpLCAxMiksICLYudiv2K8g2KfZhNit2LPYp9io2KfYqiIpOwoKICB2YXIgYWdpbmdQbG90ID0gd2FsbGV0UGxvdElkcy5hZ2luZyAmJiBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCh3YWxsZXRQbG90SWRzLmFnaW5nKTsKICBpZiAoYWdpbmdQbG90KSB7CiAgICB2YXIgY29sbCA9IHdTdW1NYXAocm93cywiYWdpbmciLCJjb2xsZWN0ZWQiKTsKICAgIHZhciByZW0gPSB3U3VtTWFwKHJvd3MsImFnaW5nIiwicmVtYWluaW5nIik7CiAgICB2YXIgbGFiZWxzID0gW10sIHNlZW5MID0ge307CiAgICBPYmplY3Qua2V5cyhjb2xsKS5jb25jYXQoT2JqZWN0LmtleXMocmVtKSkuZm9yRWFjaChmdW5jdGlvbihrKXsgaWYgKCFzZWVuTFtrXSkgeyBzZWVuTFtrXT0xOyBsYWJlbHMucHVzaChrKTsgfSB9KTsKICAgIGxhYmVscy5zb3J0KCk7CiAgICBQbG90bHkucmVhY3QoYWdpbmdQbG90LCBbCiAgICAgIHt0eXBlOiJiYXIiLCBuYW1lOiLYqtmFINin2YTYqtit2LXZitmEIiwgeDpsYWJlbHMsIHk6bGFiZWxzLm1hcChmdW5jdGlvbihrKXtyZXR1cm4gY29sbFtrXXx8MDt9KSwgbWFya2VyOntjb2xvcjpXX1BPU30sCiAgICAgICB0ZXh0OmxhYmVscy5tYXAoZnVuY3Rpb24oayl7cmV0dXJuIGNvbGxba118fDA7fSksIHRleHR0ZW1wbGF0ZToiJXt5OiwuMGZ9IiwgdGV4dHBvc2l0aW9uOiJvdXRzaWRlIiwgY2xpcG9uYXhpczpmYWxzZSwgdGV4dGZvbnQ6e3NpemU6MTIsY29sb3I6IiMxMTE4MjciLGZhbWlseToiVGFob21hLCBTZWdvZSBVSSwgQXJpYWwsIHNhbnMtc2VyaWYifSwKICAgICAgIGhvdmVydGVtcGxhdGU6IjxiPiV7eH08L2I+PGJyPtiq2YUg2KfZhNiq2K3YtdmK2YQ6ICV7eTosLjBmfTxleHRyYT48L2V4dHJhPiJ9LAogICAgICB7dHlwZToiYmFyIiwgbmFtZToi2KjYp9mC2Yog2KfZhNmF2K/ZitmI2YbZitipIiwgeDpsYWJlbHMsIHk6bGFiZWxzLm1hcChmdW5jdGlvbihrKXtyZXR1cm4gcmVtW2tdfHwwO30pLCBtYXJrZXI6e2NvbG9yOldfTkVHfSwKICAgICAgIHRleHQ6bGFiZWxzLm1hcChmdW5jdGlvbihrKXtyZXR1cm4gcmVtW2tdfHwwO30pLCB0ZXh0dGVtcGxhdGU6IiV7eTosLjBmfSIsIHRleHRwb3NpdGlvbjoib3V0c2lkZSIsIGNsaXBvbmF4aXM6ZmFsc2UsIHRleHRmb250OntzaXplOjEyLGNvbG9yOiIjMTExODI3IixmYW1pbHk6IlRhaG9tYSwgU2Vnb2UgVUksIEFyaWFsLCBzYW5zLXNlcmlmIn0sCiAgICAgICBob3ZlcnRlbXBsYXRlOiI8Yj4le3h9PC9iPjxicj7YqNin2YLZiiDYp9mE2YXYr9mK2YjZhtmK2Kk6ICV7eTosLjBmfTxleHRyYT48L2V4dHJhPiJ9CiAgICBdLCBPYmplY3QuYXNzaWduKHt9LCBhZ2luZ1Bsb3QubGF5b3V0fHx7fSwge2Jhcm1vZGU6Imdyb3VwIiwgaGVpZ2h0OjQwMCwgbWFyZ2luOk9iamVjdC5hc3NpZ24oe30sIChhZ2luZ1Bsb3QubGF5b3V0JiZhZ2luZ1Bsb3QubGF5b3V0Lm1hcmdpbil8fHt9LCB7dDo0OCxiOjg4LGw6ODAscjo0MH0pLCB4YXhpczpPYmplY3QuYXNzaWduKHt9LCAoYWdpbmdQbG90LmxheW91dCYmYWdpbmdQbG90LmxheW91dC54YXhpcyl8fHt9LCB7c2hvd2dyaWQ6ZmFsc2UsYXV0b21hcmdpbjp0cnVlfSksIHlheGlzOk9iamVjdC5hc3NpZ24oe30sIChhZ2luZ1Bsb3QubGF5b3V0JiZhZ2luZ1Bsb3QubGF5b3V0LnlheGlzKXx8e30sIHtzaG93Z3JpZDp0cnVlLGdyaWRjb2xvcjoicmdiYSgxMjgsMTI4LDEyOCwwLjEyKSIsemVyb2xpbmU6ZmFsc2UsYXV0b21hcmdpbjp0cnVlLHRpdGxlOk9iamVjdC5hc3NpZ24oe30sICgoYWdpbmdQbG90LmxheW91dCYmYWdpbmdQbG90LmxheW91dC55YXhpcyl8fHt9KS50aXRsZXx8e30sIHt0ZXh0OiLYp9mE2YXYqNmE2LoiLHN0YW5kb2ZmOjE4fSksdGlja2Zvcm1hdDoifnMifSl9KSk7CiAgfQoKICB2YXIgcGllUGxvdCA9IHdhbGxldFBsb3RJZHMuY3VzdG9tZXJfc3RhdGVfcGllICYmIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKHdhbGxldFBsb3RJZHMuY3VzdG9tZXJfc3RhdGVfcGllKTsKICBpZiAocGllUGxvdCkgewogICAgdmFyIGVudHJpZXMgPSB3VG9wRW50cmllcyh3Q291bnRNYXAocm93cywiY3VzdG9tZXJfc3RhdGUiKSwgMTApLnJldmVyc2UoKTsKICAgIFBsb3RseS5yZWFjdChwaWVQbG90LCBbewogICAgICB0eXBlOiJwaWUiLCBsYWJlbHM6ZW50cmllcy5tYXAoZnVuY3Rpb24oZSl7cmV0dXJuIGVbMF07fSksIHZhbHVlczplbnRyaWVzLm1hcChmdW5jdGlvbihlKXtyZXR1cm4gZVsxXTt9KSwKICAgICAgbWFya2VyOntjb2xvcnM6W1dfREFSSyxXX01JRCxXX0xJR0hULCIjQjhDNUM4IiwiIzNEN0U4MiIsIiM3RUFCQUUiXX0sCiAgICAgIHRleHRpbmZvOiJsYWJlbCtwZXJjZW50IiwKICAgICAgaG92ZXJ0ZW1wbGF0ZToiPGI+JXtsYWJlbH08L2I+PGJyPtin2YTYudiv2K86ICV7dmFsdWU6LC4wZn08YnI+2KfZhNmG2LPYqNipOiAle3BlcmNlbnQ6LjElfTxleHRyYT48L2V4dHJhPiIKICAgIH1dLCBPYmplY3QuYXNzaWduKHt9LCBwaWVQbG90LmxheW91dHx8e30sIHtoZWlnaHQ6MzYwLCBzaG93bGVnZW5kOmZhbHNlfSkpOwogIH0KCiAgdmFyIGRvbnV0UGxvdCA9IHdhbGxldFBsb3RJZHMubmF0aW9uYWxpdHlfZG9udXQgJiYgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQod2FsbGV0UGxvdElkcy5uYXRpb25hbGl0eV9kb251dCk7CiAgaWYgKGRvbnV0UGxvdCkgewogICAgdmFyIG5FbnRyaWVzID0gd1RvcEVudHJpZXMod0NvdW50TWFwKHJvd3MsIm5hdGlvbmFsaXR5IiksIDEwKS5yZXZlcnNlKCk7CiAgICBQbG90bHkucmVhY3QoZG9udXRQbG90LCBbewogICAgICB0eXBlOiJwaWUiLCBsYWJlbHM6bkVudHJpZXMubWFwKGZ1bmN0aW9uKGUpe3JldHVybiBlWzBdO30pLCB2YWx1ZXM6bkVudHJpZXMubWFwKGZ1bmN0aW9uKGUpe3JldHVybiBlWzFdO30pLCBob2xlOjAuNTUsCiAgICAgIG1hcmtlcjp7Y29sb3JzOltXX0RBUkssV19NSUQsV19MSUdIVCwiI0I4QzVDOCIsIiMzRDdFODIiLCIjN0VBQkFFIl19LAogICAgICB0ZXh0aW5mbzoibGFiZWwrcGVyY2VudCIsCiAgICAgIGhvdmVydGVtcGxhdGU6IjxiPiV7bGFiZWx9PC9iPjxicj7Yp9mE2LnYr9ivOiAle3ZhbHVlOiwuMGZ9PGJyPtin2YTZhtiz2KjYqTogJXtwZXJjZW50Oi4xJX08ZXh0cmE+PC9leHRyYT4iCiAgICB9XSwgT2JqZWN0LmFzc2lnbih7fSwgZG9udXRQbG90LmxheW91dHx8e30sIHtoZWlnaHQ6MzYwLCBzaG93bGVnZW5kOmZhbHNlfSkpOwogIH0KCiAgdmFyIHRoZWFkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoIndhbGxldC1tYXRyaXgtdGhlYWQiKTsKICB2YXIgdGJvZHkgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LW1hdHJpeC10Ym9keSIpOwogIGlmICh0aGVhZCAmJiB0Ym9keSkgewogICAgdmFyIG14ID0gd0J1aWxkTWF0cml4KHJvd3MpOwogICAgdGhlYWQuaW5uZXJIVE1MID0gbXguaGVhZDsKICAgIHRib2R5LmlubmVySFRNTCA9IG14LmJvZHk7CiAgfQp9CgpmdW5jdGlvbiB3QmluZCgpewogIFsid2FsbGV0LWFnZW50LW9wdGlvbiIsIndhbGxldC1uYXRpb24tb3B0aW9uIiwid2FsbGV0LWNzdGF0ZS1vcHRpb24iLCJ3YWxsZXQtYWdpbmctb3B0aW9uIl0uZm9yRWFjaChmdW5jdGlvbihjbHMpewogICAgQXJyYXkucHJvdG90eXBlLnNsaWNlLmNhbGwoZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgiLiIrY2xzKSkuZm9yRWFjaChmdW5jdGlvbihvKXsKICAgICAgby5hZGRFdmVudExpc3RlbmVyKCJjaGFuZ2UiLCBmdW5jdGlvbigpeyB3VXBkYXRlTGFiZWxzKCk7IHdSZWZyZXNoKCk7IH0pOwogICAgfSk7CiAgfSk7CiAgZnVuY3Rpb24gYmluZFNlbGVjdEFsbChhbGxDbHMsIG9wdENscyl7CiAgICB2YXIgZWwgPSBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCIuIithbGxDbHMpOwogICAgaWYgKCFlbCkgcmV0dXJuOwogICAgZWwuYWRkRXZlbnRMaXN0ZW5lcigiY2hhbmdlIiwgZnVuY3Rpb24oZSl7CiAgICAgIEFycmF5LnByb3RvdHlwZS5zbGljZS5jYWxsKGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoIi4iK29wdENscykpLmZvckVhY2goZnVuY3Rpb24obyl7IG8uY2hlY2tlZCA9IGUudGFyZ2V0LmNoZWNrZWQ7IH0pOwogICAgICB3VXBkYXRlTGFiZWxzKCk7IHdSZWZyZXNoKCk7CiAgICB9KTsKICB9CiAgYmluZFNlbGVjdEFsbCgid2FsbGV0LXNlbGVjdC1hbGwtYWdlbnQiLCAid2FsbGV0LWFnZW50LW9wdGlvbiIpOwogIGJpbmRTZWxlY3RBbGwoIndhbGxldC1zZWxlY3QtYWxsLW5hdGlvbiIsICJ3YWxsZXQtbmF0aW9uLW9wdGlvbiIpOwogIGJpbmRTZWxlY3RBbGwoIndhbGxldC1zZWxlY3QtYWxsLWNzdGF0ZSIsICJ3YWxsZXQtY3N0YXRlLW9wdGlvbiIpOwogIGJpbmRTZWxlY3RBbGwoIndhbGxldC1zZWxlY3QtYWxsLWFnaW5nIiwgIndhbGxldC1hZ2luZy1vcHRpb24iKTsKICBbIndhbGxldC1hc3NpZ24tZnJvbSIsIndhbGxldC1hc3NpZ24tdG8iLCJ3YWxsZXQtZGViaXQtZnJvbSIsIndhbGxldC1kZWJpdC10byJdLmZvckVhY2goZnVuY3Rpb24oaWQpewogICAgdmFyIGVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpOwogICAgaWYgKGVsKSBlbC5hZGRFdmVudExpc3RlbmVyKCJjaGFuZ2UiLCB3UmVmcmVzaCk7CiAgfSk7CiAgdmFyIHJlc2V0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoIndhbGxldC1yZXNldC1maWx0ZXJzIik7CiAgaWYgKHJlc2V0KSByZXNldC5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsIGZ1bmN0aW9uKCl7CiAgICBBcnJheS5wcm90b3R5cGUuc2xpY2UuY2FsbChkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCIud2FsbGV0LWFnZW50LW9wdGlvbiwud2FsbGV0LW5hdGlvbi1vcHRpb24sLndhbGxldC1jc3RhdGUtb3B0aW9uLC53YWxsZXQtYWdpbmctb3B0aW9uLC53YWxsZXQtc2VsZWN0LWFsbC1hZ2VudCwud2FsbGV0LXNlbGVjdC1hbGwtbmF0aW9uLC53YWxsZXQtc2VsZWN0LWFsbC1jc3RhdGUsLndhbGxldC1zZWxlY3QtYWxsLWFnaW5nIikpLmZvckVhY2goZnVuY3Rpb24obyl7IG8uY2hlY2tlZCA9IGZhbHNlOyB9KTsKICAgIHZhciBhZiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtYXNzaWduLWZyb20iKTsgaWYgKGFmKSBhZi52YWx1ZSA9IFdfQVNTSUdOX01JTjsKICAgIHZhciBhdCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ3YWxsZXQtYXNzaWduLXRvIik7IGlmIChhdCkgYXQudmFsdWUgPSBXX0FTU0lHTl9NQVg7CiAgICB2YXIgZGYgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgid2FsbGV0LWRlYml0LWZyb20iKTsgaWYgKGRmKSBkZi52YWx1ZSA9IFdfREVCSVRfTUlOOwogICAgdmFyIGR0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoIndhbGxldC1kZWJpdC10byIpOyBpZiAoZHQpIGR0LnZhbHVlID0gV19ERUJJVF9NQVg7CiAgICB3VXBkYXRlTGFiZWxzKCk7IHdSZWZyZXNoKCk7CiAgfSk7CiAgd1VwZGF0ZUxhYmVscygpOwogIHdSZWZyZXNoKCk7Cn0KaWYgKGRvY3VtZW50LnJlYWR5U3RhdGUgPT09ICJsb2FkaW5nIikgZG9jdW1lbnQuYWRkRXZlbnRMaXN0ZW5lcigiRE9NQ29udGVudExvYWRlZCIsIHdCaW5kKTsKZWxzZSB3QmluZCgpOwp9KSgpOwo=").decode("utf-8")
    js_out = (
        js_tpl
        .replace("__WALLET_DATA__", _json.dumps(records, ensure_ascii=False))
        .replace("__WALLET_PLOT_IDS__", _json.dumps(plot_id_map, ensure_ascii=False))
        .replace("__OPS_DARK__", OPS_DARK)
        .replace("__OPS_MID__", OPS_MID)
        .replace("__OPS_LIGHT__", OPS_LIGHT)
        .replace("__OPS_POS__", OPS_POSITIVE)
        .replace("__OPS_NEG__", OPS_NEGATIVE)
        .replace("__ASSIGN_MIN__", assign_min)
        .replace("__ASSIGN_MAX__", assign_max)
        .replace("__DEBIT_MIN__", debit_min)
        .replace("__DEBIT_MAX__", debit_max)
    )
    parts.append("<script>\n" + js_out + "\n</script>")
    return "".join(parts)



def _build_payments_page_html(payments_df):
    figs, _work, _c, _a = _build_payments_analysis(payments_df)
    parts = ["<div id='page-payments' class='dash-page'>",
             "<h2 class='section-title'>💰 تحليل السداد الكامل</h2>"]
    parts.append(_fig_html_card(figs["kpi"], "payments", 0))
    if "trend" in figs:
        parts.append("<div class='charts-grid'>" + _fig_html_card(figs["trend"], "payments", 1) + "</div>")
    parts.append("<div class='charts-grid'>")
    for i, key in enumerate(("by_agent_count", "by_agent_amount", "amount_hist"), start=2):
        if key in figs:
            parts.append(_fig_html_card(figs[key], "payments", i))
    parts.append("</div></div>")
    return "".join(parts)


def _build_link_page_html(activity_df, sales_col, class_col, payments_df):
    parts = ["<div id='page-link' class='dash-page'>",
             "<h2 class='section-title'>🔗 ربط نشاط المحصلين بالسداد</h2>"]
    result = _build_activity_payments_link(activity_df, sales_col, class_col, payments_df)
    if result is None:
        parts.append("<p style='text-align:center'>تعذّر الربط — تأكد إن عمود المحصل موجود في ملفي النشاط والسداد.</p></div>")
        return "".join(parts)
    figs, merged = result
    parts.append(_fig_html_card(figs["kpi"], "link", 0))
    parts.append("<div class='charts-grid'>")
    parts.append(_fig_html_card(figs["compare_bar"], "link", 1))
    parts.append(_fig_html_card(figs["scatter"], "link", 2))
    parts.append("</div>")
    columns = ["المحصّل", "عدد_المكالمات", "عدد_الناجحة", "نسبة_النجاح", "عدد_السدادات", "إجمالي_التحصيل", "الفعالية"]
    columns = [c for c in columns if c in merged.columns]
    parts.append("<section class='panel'><h2 class='section-title'>📋 جدول النشاط مقابل السداد</h2>"
                  "<div class='table-wrap'><table class='data-table'><thead><tr>")
    for column in columns:
        parts.append(f"<th>{column}</th>")
    parts.append("</tr></thead><tbody>")
    for _, row in merged.sort_values("إجمالي_التحصيل", ascending=False).iterrows():
        parts.append("<tr>")
        for column in columns:
            value = row[column]
            if isinstance(value, float):
                value = f"{value:,.1f}"
            parts.append(f"<td>{value}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div></section></div>")
    return "".join(parts)


# ==========================================================
# ==========================================================
# تويب الجدولة المتعثرة
# ==========================================================

def _build_last_payment_map(payments_df, debitor_col, date_col):
    """يرجع dict: Debitor -> أحدث تاريخ سداد (date)."""
    work = payments_df[[debitor_col, date_col]].copy()
    work["_debitor"] = work[debitor_col].astype(str).str.strip()
    work["_pay_date"] = [parse_date_cell(v) for v in work[date_col]]
    work = work[work["_debitor"].ne("") & work["_debitor"].ne("nan") & work["_pay_date"].notna()]
    if work.empty:
        return {}
    # أحدث تاريخ لكل Debitor
    idx = work.groupby("_debitor")["_pay_date"].idxmax()
    latest = work.loc[idx]
    return dict(zip(latest["_debitor"], latest["_pay_date"]))


def _run_schedule_stalled_pipeline(portfolio_file, payments_file):
    """محفظة حديثة + ملف السدادات → حالات جدولة منتظمة / متعثرة."""
    portfolio_hash = hashlib.sha256(portfolio_file.getvalue()).hexdigest()
    payments_hash = hashlib.sha256(payments_file.getvalue()).hexdigest()
    combined_hash = hashlib.sha256(f"{portfolio_hash}:{payments_hash}".encode()).hexdigest()

    cached = st.session_state.get(SCHEDULE_RESULT_KEY)
    if cached and cached.get("file_hash") == combined_hash:
        return False  # موجود في الكاش

    try:
        portfolio_df = read_uploaded_dataframe(portfolio_file)
        payments_df = read_uploaded_dataframe(payments_file)
    except Exception as e:
        st.error(f"تعذر قراءة أحد الملفات: {e}")
        return False

    total_in_portfolio = len(portfolio_df)
    # 1) حذف أول صف بعد العناوين من المحفظة
    if len(portfolio_df) > 0:
        portfolio_df = portfolio_df.iloc[1:].reset_index(drop=True)

    # ملف السدادات: نحذف أول صف إن وُجد نمط مشابه (صف فارغ/عنوان مكرر)
    # نتركه كما هو إن لم يكن ضرورياً — غالباً السدادات بدون صف زائد، لكن لو عدد الأعمدة غريب نتركه.
    # المستخدم طلب الحذف صراحة للمحفظة فقط.

    sales_col = find_column(portfolio_df, SALES_PERSON_CANDIDATES)
    substate_col = find_column(portfolio_df, PROMISE_SUB_STATE_CANDIDATES)
    portfolio_debitor_col = find_column(portfolio_df, SCHEDULE_DEBITOR_CANDIDATES)
    payments_debitor_col = find_column(payments_df, SCHEDULE_DEBITOR_CANDIDATES)
    payment_date_col = find_column(payments_df, SCHEDULE_PAYMENT_DATE_CANDIDATES)
    net_col = find_column(portfolio_df, PROMISE_NET_AMOUNT_CANDIDATES)

    missing = []
    if not sales_col:
        missing.append("المحصّل (Salesperson)")
    if not substate_col:
        missing.append("الحالة الفرعية (Sub State)")
    if not portfolio_debitor_col:
        missing.append("Debitor في المحفظة")
    if not payments_debitor_col:
        missing.append("Debitor في السدادات")
    if not payment_date_col:
        missing.append("Date of Creation في السدادات")
    if missing:
        st.error(
            "تعذر العثور على أعمدة مهمة.\n"
            f"الأعمدة الناقصة: {', '.join(missing)}\n\n"
            f"أعمدة المحفظة: {', '.join(map(str, portfolio_df.columns))}\n"
            f"أعمدة السدادات: {', '.join(map(str, payments_df.columns))}"
        )
        return False

    # 2) فلترة Salesperson — استبعاد المحصّلين المحددين
    sales_vals = portfolio_df[sales_col].astype(str).str.strip()
    keep_sales = ~sales_vals.isin(PROMISE_EXCLUDED_SALES)
    dropped_sales = int((~keep_sales).sum())
    df = portfolio_df[keep_sales].copy()

    # 3) فلترة Sub State = جدولة فقط
    sub_vals = df[substate_col].astype(str).str.strip()
    # توحيد بسيط: قبول "جدولة" و"جدوله"
    keep_sub = sub_vals.apply(lambda s: _state_key(s) == _state_key(SCHEDULE_SUB_STATE_VALUE))
    dropped_sub = int((~keep_sub).sum())
    df = df[keep_sub].copy()

    # 4) VLOOKUP: تاريخ آخر سداد من ملف السدادات عبر Debitor
    last_pay_map = _build_last_payment_map(payments_df, payments_debitor_col, payment_date_col)
    debitor_keys = df[portfolio_debitor_col].astype(str).str.strip()
    df[SCHEDULE_LAST_PAYMENT_COL] = debitor_keys.map(last_pay_map)

    # 5) حالة الجدولة: ≤ شهرين (60 يوم) منتظمة، أكثر متعثرة
    _init_promises_today()
    target_date = st.session_state[TODAY_KEY]

    def _schedule_status(last_pay):
        if last_pay is None or (isinstance(last_pay, float) and pd.isna(last_pay)):
            return "بدون سداد", None
        try:
            days = (target_date - last_pay).days
        except Exception:
            return "بدون سداد", None
        if days < 0:
            days = 0
        if days <= SCHEDULE_REGULAR_MAX_DAYS:
            return "جدولة منتظمة", days
        return "جدولة متعثرة", days

    status_days = df[SCHEDULE_LAST_PAYMENT_COL].apply(_schedule_status)
    df[SCHEDULE_STATUS_COL] = status_days.apply(lambda x: x[0])
    df[SCHEDULE_DAYS_SINCE_COL] = status_days.apply(lambda x: x[1])

    matched = int(df[SCHEDULE_LAST_PAYMENT_COL].notna().sum())
    unmatched = len(df) - matched
    regular_count = int((df[SCHEDULE_STATUS_COL] == "جدولة منتظمة").sum())
    stalled_count = int((df[SCHEDULE_STATUS_COL] == "جدولة متعثرة").sum())
    no_pay_count = int((df[SCHEDULE_STATUS_COL] == "بدون سداد").sum())

    st.session_state[SCHEDULE_RESULT_KEY] = {
        "df": df,
        "file_hash": combined_hash,
        "portfolio_name": portfolio_file.name,
        "payments_name": payments_file.name,
        "sales_col": sales_col,
        "substate_col": substate_col,
        "debitor_col": portfolio_debitor_col,
        "net_col": net_col,
        "target_date": target_date,
        "total_in_portfolio": total_in_portfolio,
        "after_filters": len(df),
        "dropped_sales": dropped_sales,
        "dropped_sub": dropped_sub,
        "matched": matched,
        "unmatched": unmatched,
        "regular_count": regular_count,
        "stalled_count": stalled_count,
        "no_pay_count": no_pay_count,
    }
    return True


def render_schedule_kpi_dashboard(total, regular, stalled, no_pay, agent_count=None):
    """كروت KPI للجدولة بنفس أسلوب كروت التصنيف (Plotly Indicators + زوايا دائرية)."""
    cards = [
        ("📋<br>إجمالي الجدولة", total, {"valueformat": ",d"}, THEME["text"]),
        ("📗<br>جدولة منتظمة", regular, {"valueformat": ",d"}, SCHEDULE_STATUS_COLORS["جدولة منتظمة"]),
        ("📕<br>جدولة متعثرة", stalled, {"valueformat": ",d"}, SCHEDULE_STATUS_COLORS["جدولة متعثرة"]),
        ("⚪<br>بدون سداد", no_pay, {"valueformat": ",d"}, SCHEDULE_STATUS_COLORS["بدون سداد"]),
    ]
    if agent_count is not None:
        cards.insert(1, ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]))

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
                title={"text": label, "font": {"size": 17, "color": THEME["text_dim"]}, "align": "center"},
                number={"font": {"size": 30, "color": number_color}, **number_format},
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
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG, key="schedule_kpi_dashboard")


def _show_schedule_stalled_results(df, meta):
    sales_col = meta.get("sales_col")
    net_col = meta.get("net_col")
    debitor_col = meta.get("debitor_col")
    matched = int(meta.get("matched", 0) or 0)
    no_pay_meta = int(meta.get("no_pay_count", 0) or 0)

    # فلتر المحصّل من الضغط على الشارت
    render_schedule_filter_notice()
    df_all = df
    df = get_schedule_view(df, sales_col)

    total = len(df)
    regular = int((df[SCHEDULE_STATUS_COL] == "جدولة منتظمة").sum()) if SCHEDULE_STATUS_COL in df.columns else 0
    stalled = int((df[SCHEDULE_STATUS_COL] == "جدولة متعثرة").sum()) if SCHEDULE_STATUS_COL in df.columns else 0
    no_pay = int((df[SCHEDULE_STATUS_COL] == "بدون سداد").sum()) if SCHEDULE_STATUS_COL in df.columns else 0
    agent_count = int(df[sales_col].nunique()) if sales_col and sales_col in df.columns else 0

    selected_agent = st.session_state.get(SCHEDULE_AGENT_FILTER_KEY)
    if not selected_agent:
        st.success(
            f"✅ تم التحليل: {len(df_all):,} حالة جدولة بعد الفلترة · "
            f"مطابقة سدادات: {matched:,} · بدون سداد: {no_pay_meta:,}"
        )
        st.caption(
            f"المحفظة: {meta.get('portfolio_name', '—')} · "
            f"السدادات: {meta.get('payments_name', '—')} · "
            f"تاريخ المرجع: {meta.get('target_date')}"
        )
        st.caption("💡 اضغط على أي محصّل في الشارت لتصفية كل المؤشرات والرسوم عليه.")

    st.subheader("📊 ملخص الجدولة" + (f" — {selected_agent}" if selected_agent else ""))
    render_schedule_kpi_dashboard(total, regular, stalled, no_pay, agent_count=agent_count)

    # فلتر عرض الحالة
    status_filter = st.multiselect(
        "تصفية حسب حالة الجدولة",
        options=["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"],
        default=["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"],
        key="schedule_status_filter",
    )
    view = df[df[SCHEDULE_STATUS_COL].isin(status_filter)].copy() if status_filter else df.copy()

    if total and sales_col and sales_col in df.columns:
        st.markdown("#### 📈 تحليلات الجدولة التفاعلية")
        agent_status = (
            df.groupby([sales_col, SCHEDULE_STATUS_COL])
            .size()
            .reset_index(name="العدد")
        )
        agent_order = (
            agent_status.groupby(sales_col)["العدد"]
            .sum()
            .sort_values(ascending=False)
            .head(15)
            .index.tolist()
        )
        agent_status = agent_status[agent_status[sales_col].isin(agent_order)].copy()
        agent_status["_sort"] = agent_status[sales_col].map({name: i for i, name in enumerate(agent_order)})
        agent_status = agent_status.sort_values(["_sort", SCHEDULE_STATUS_COL], ascending=[True, True])

        left, right = st.columns(2)
        with left:
            fig = px.bar(
                agent_status,
                x="العدد",
                y=sales_col,
                color=SCHEDULE_STATUS_COL,
                barmode="group",
                orientation="h",
                text="العدد",
                color_discrete_map=SCHEDULE_STATUS_COLORS,
                template=PLOTLY_TEMPLATE,
                category_orders={
                    sales_col: agent_order,
                    SCHEDULE_STATUS_COL: ["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"],
                },
            )
            _apply_ops_chart_style(
                fig, "حالات الجدولة حسب المحصّل",
                height=430, xaxis_title="العدد",
                margin=dict(t=70, b=70, l=170, r=55),
                extra={"xaxis": dict(tickformat=",.0f", automargin=True)},
            )
            fig.update_traces(
                texttemplate="%{x:,.0f}",
                textposition="outside",
                textfont=dict(size=13, color=THEME["text"]),
                cliponaxis=False,
                customdata=agent_status[sales_col],
                hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f}<extra></extra>",
            )
            with st.container(border=True):
                render_selectable_chart(fig, "schedule_by_agent", filter_key=SCHEDULE_AGENT_FILTER_KEY)

        with right:
            pie_counts = (
                df[SCHEDULE_STATUS_COL]
                .value_counts()
                .reindex(["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"], fill_value=0)
                .rename_axis("الحالة")
                .reset_index(name="العدد")
            )
            pie_counts = pie_counts[pie_counts["العدد"] > 0]
            pie = px.pie(
                pie_counts,
                values="العدد",
                names="الحالة",
                hole=0.55,
                color="الحالة",
                color_discrete_map=SCHEDULE_STATUS_COLORS,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                pie, "توزيع حالات الجدولة", height=430,
                margin=dict(t=70, b=50, l=30, r=30),
            )
            pie.update_traces(
                texttemplate="%{label}<br>%{value:,.0f} (%{percent:.1%})",
                textfont=dict(size=14, color=THEME["text"]),
                textinfo="text",
                hovertemplate="<b>%{label}</b><br>العدد: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
            )
            with st.container(border=True):
                st.plotly_chart(pie, use_container_width=True, config=PLOTLY_CONFIG, key="schedule_status_pie")

        stacked = (
            df.groupby([sales_col, SCHEDULE_STATUS_COL])
            .size()
            .unstack(fill_value=0)
        )
        for col in ["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"]:
            if col not in stacked.columns:
                stacked[col] = 0
        stacked = stacked[["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"]]
        stacked["الإجمالي"] = stacked.sum(axis=1)
        stacked = stacked.sort_values("الإجمالي", ascending=True).tail(15)
        stack_fig = go.Figure()
        agent_labels = stacked.index.astype(str).tolist()
        for status_name in ["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"]:
            stack_fig.add_trace(
                go.Bar(
                    name=status_name,
                    y=agent_labels,
                    x=stacked[status_name].tolist(),
                    orientation="h",
                    marker_color=SCHEDULE_STATUS_COLORS[status_name],
                    text=[f"{v:,}" if v else "" for v in stacked[status_name].tolist()],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(size=12, color="#F3F6FA"),
                    customdata=agent_labels,
                    hovertemplate=f"<b>%{{y}}</b><br>{status_name}: %{{x:,}}<extra></extra>",
                )
            )
        _apply_ops_chart_style(
            stack_fig, "توزيع مكدّس حسب المحصّل",
            height=430, xaxis_title="العدد",
            margin=dict(t=70, b=70, l=170, r=50),
            extra={"barmode": "stack", "xaxis": dict(tickformat=",.0f", automargin=True)},
        )
        with st.container(border=True):
            render_selectable_chart(stack_fig, "schedule_stacked_by_agent", filter_key=SCHEDULE_AGENT_FILTER_KEY)

    # ملخص حسب المحصّل
    if sales_col and sales_col in df.columns:
        summary = (
            df.groupby(sales_col)[SCHEDULE_STATUS_COL]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )
        for col in ["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"]:
            if col not in summary.columns:
                summary[col] = 0
        summary["الإجمالي"] = summary[["جدولة منتظمة", "جدولة متعثرة", "بدون سداد"]].sum(axis=1)
        summary = summary.rename(columns={sales_col: "المحصّل"}).sort_values("الإجمالي", ascending=False)
        st.subheader("📊 ملخص حسب المحصّل")
        st.dataframe(summary, use_container_width=True, hide_index=True)

    display_cols = [
        c for c in [
            sales_col,
            debitor_col,
            meta.get("substate_col"),
            SCHEDULE_LAST_PAYMENT_COL,
            SCHEDULE_DAYS_SINCE_COL,
            SCHEDULE_STATUS_COL,
            net_col,
        ]
        if c and c in view.columns
    ]
    st.subheader("📋 تفاصيل حالات الجدولة")
    if display_cols:
        st.dataframe(view[display_cols], use_container_width=True, hide_index=True)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True)

    def _excel_bytes(report_df, sheet_name):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            report_df.to_excel(writer, index=False, sheet_name=sheet_name)
        return buf.getvalue()

    report_date = datetime.now().strftime("%Y-%m-%d")
    # التحميل يعتمد على العرض الحالي (بعد فلتر المحصّل + الحالة)
    stalled_df = view[view[SCHEDULE_STATUS_COL] == "جدولة متعثرة"].copy() if SCHEDULE_STATUS_COL in view.columns else view.iloc[0:0].copy()
    regular_df = view[view[SCHEDULE_STATUS_COL] == "جدولة منتظمة"].copy() if SCHEDULE_STATUS_COL in view.columns else view.iloc[0:0].copy()

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "⬇️ تحميل كل حالات الجدولة",
            data=_excel_bytes(view, "الجدولة"),
            file_name=f"تقرير_الجدولة_كامل_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
            key="schedule_download_all",
            disabled=view.empty,
        )
    with d2:
        st.download_button(
            "⬇️ تحميل الجدولة المتعثرة فقط",
            data=_excel_bytes(stalled_df, "متعثرة"),
            file_name=f"تقرير_الجدولة_المتعثرة_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
            key="schedule_download_stalled",
            disabled=stalled_df.empty,
        )
    with d3:
        st.download_button(
            "⬇️ تحميل الجدولة المنتظمة فقط",
            data=_excel_bytes(regular_df, "منتظمة"),
            file_name=f"تقرير_الجدولة_المنتظمة_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="schedule_download_regular",
            disabled=regular_df.empty,
        )


def page_schedule_stalled():
    """صفحة الجدولة المتعثرة: محفظة حديثة + ملف السدادات → منتظمة / متعثرة."""
    _init_promises_today()
    page_header(
        "STALLED SCHEDULES",
        "📅 الجدولة المتعثرة",
        "ارفع المحفظة الحديثة وملف السدادات لتحديد حالات الجدولة المنتظمة والمتعثرة (أكثر من شهرين بلا سداد)",
    )

    portfolio_key = "schedule_portfolio_upload"
    payments_key = "schedule_payments_upload"
    result_keys = (SCHEDULE_RESULT_KEY,)

    c1, c2 = st.columns(2)
    with c1:
        portfolio_file = st.file_uploader(
            "📂 ارفع المحفظة الحديثة (Excel أو CSV)",
            type=["xlsx", "xls", "csv"],
            key=portfolio_key,
            on_change=sync_file_cache,
            args=(portfolio_key, "schedule_portfolio", result_keys),
        )
    with c2:
        payments_file = st.file_uploader(
            "📂 ارفع ملف السدادات (Excel أو CSV)",
            type=["xlsx", "xls", "csv"],
            key=payments_key,
            on_change=sync_file_cache,
            args=(payments_key, "schedule_payments", result_keys),
        )

    if portfolio_file is not None and payments_file is not None:
        st.caption(
            f"المحفظة: {portfolio_file.name} · السدادات: {payments_file.name}"
        )
        with st.spinner("جارٍ تحليل الجدولة وربط السدادات..."):
            _run_schedule_stalled_pipeline(portfolio_file, payments_file)

    cached = st.session_state.get(SCHEDULE_RESULT_KEY)
    if cached and cached.get("df") is not None:
        if portfolio_file is None and payments_file is None:
            st.success(
                f"✅ نتيجة محفوظة من: {cached.get('portfolio_name', '—')} + "
                f"{cached.get('payments_name', '—')}. لن تُحذف عند التنقل بين التبويبات."
            )
        _show_schedule_stalled_results(cached["df"], cached)
    else:
        st.info("📂 ارفع ملف المحفظة الحديثة وملف السدادات معاً لبدء التحليل.")



# التنقل (Sidebar Navigation)
# ==========================================================

# ==========================================================
# تويب أخطاء الحالات
# ==========================================================

def _to_amount(val):
    """تحويل قيمة Payment / Net Amount لرقم."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        try:
            return float(val)
        except (TypeError, ValueError):
            return 0.0
    txt = str(val).strip().replace(",", "").replace(" ", "").replace("جنيه", "").replace("EGP", "")
    if not txt or txt.lower() in {"nan", "none", "-", "—"}:
        return 0.0
    try:
        return float(txt)
    except (TypeError, ValueError):
        try:
            return float(pd.to_numeric(txt, errors="coerce") or 0)
        except Exception:
            return 0.0


def _match_case_payment_state(value):
    """تصنيف Sub State لإحدى حالات السداد/الجدولة المعروفة، أو None."""
    key = _state_key(value)
    if not key:
        return None
    # الأخص أولاً
    if "جدوله" in key and ("مقفل" in key or "مغلق" in key) and "خصم" in key:
        return "جدولة مقفلة بخصم"
    if "جدوله" in key and ("مقفل" in key or "مغلق" in key):
        return "جدولة مقفلة"
    if "خصم" in key and "كامل" in key and "مديون" in key:
        return "سدد كامل المديونية بخصم"
    if ("سدد" in key or "سداد" in key) and "كامل" in key and "مديون" in key:
        return "سدد كامل المديونية"
    if key == "جدوله" or (key.startswith("جدوله") and "مقفل" not in key and "مغلق" not in key and "خصم" not in key):
        return "جدولة"
    return None


def _detect_case_errors_for_row(sub_state, payment_amt, net_amt, original_amt=None):
    """كشف أخطاء الحالات حسب القواعد المتفق عليها.

    قواعد Payment:
      1) Sub State ∈ {جدولة، سدد كامل المديونية} و Payment = 0 → خطأ
      2) Sub State ∉ حالات السداد و Payment > 0 → خطأ
         (Payment = 0 أو Net Amount = 0 مع حالة غير سداد → مش خطأ)

    قواعد Net Amount (حالات السداد فقط):
      3) سدد كامل المديونية + Net Amount ≥ 50 → خطأ
         (متبقي < 50 مقبول)
      4) جدولة مقفلة + Net Amount ≥ 50 → خطأ
         (متبقي < 50 مقبول)
      5) جدولة مقفلة بخصم + Net Amount = 0 → خطأ
      6) سدد كامل المديونية بخصم + Net Amount = 0 → خطأ
      7) جدولة عادية مع متبقي في Net Amount → طبيعي (مش خطأ)

    ملاحظة: حالة غير سداد + (Payment = 0 أو Net Amount = 0 أو الاتنين) → مش خطأ حالة.
    """
    errors = []
    matched = _match_case_payment_state(sub_state)
    is_payment_state = matched is not None

    # —— قواعد Payment ——
    # 1) جدولة / سدد كامل المديونية + Payment = 0 → خطأ
    if matched in CASE_PAYMENT_MUST_BE_POSITIVE and payment_amt == 0:
        errors.append((
            "Payment صفر مع حالة سداد",
            f"الحالة «{matched}» تدل على سداد لكن Payment = 0",
        ))
    # 2) حالة غير سداد + Payment > 0 → خطأ
    #    (الصفر في Payment أو Net Amount مع غير سداد → مش خطأ)
    if not is_payment_state and payment_amt > 0:
        errors.append((
            "Payment أكبر من صفر بدون حالة سداد",
            f"الحالة «{sub_state}» لا تدل على سداد لكن Payment = {payment_amt:,.2f}",
        ))

    # —— قواعد Net Amount (حالات السداد فقط) ——
    # غير سداد + Net Amount = 0 → مش خطأ (متعمد)
    if matched == "سدد كامل المديونية" and net_amt >= CASE_NET_AMOUNT_ERROR_MIN:
        errors.append((
            "متبقي بعد سداد كامل",
            f"«سدد كامل المديونية» لكن Net Amount = {net_amt:,.2f} (المقبول أقل من {CASE_NET_AMOUNT_ERROR_MIN:g})",
        ))
    elif matched == "جدولة مقفلة" and net_amt >= CASE_NET_AMOUNT_ERROR_MIN:
        errors.append((
            "متبقي مع جدولة مقفلة",
            f"«جدولة مقفلة» لكن Net Amount = {net_amt:,.2f} (المقبول أقل من {CASE_NET_AMOUNT_ERROR_MIN:g})",
        ))
    elif matched == "جدولة مقفلة بخصم" and net_amt == 0:
        errors.append((
            "Net Amount صفر مع جدولة مقفلة بخصم",
            "«جدولة مقفلة بخصم» لكن Net Amount = 0",
        ))
    elif matched == "سدد كامل المديونية بخصم" and net_amt == 0:
        errors.append((
            "Net Amount صفر مع سداد بخصم",
            "«سدد كامل المديونية بخصم» لكن Net Amount = 0",
        ))
    # matched == "جدولة": متبقي متوقع — مش خطأ

    return errors



def _run_case_errors_pipeline(uploaded):
    file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    cached = st.session_state.get(CASE_ERRORS_RESULT_KEY)
    if cached and cached.get("file_hash") == file_hash:
        return False

    try:
        raw_df = read_uploaded_dataframe(uploaded)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return False

    total_in_file = len(raw_df)
    df = raw_df
    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    sales_col = find_column(df, SALES_PERSON_CANDIDATES)
    substate_col = find_column(df, PROMISE_SUB_STATE_CANDIDATES)
    payment_col = find_column(df, CASE_PAYMENT_CANDIDATES)
    net_col = find_column(df, PROMISE_NET_AMOUNT_CANDIDATES)
    original_col = find_column(df, CASE_ORIGINAL_DEBT_CANDIDATES)

    missing = []
    if not sales_col:
        missing.append("المحصّل (Salesperson)")
    if not substate_col:
        missing.append("الحالة الفرعية (Sub State)")
    if not payment_col:
        missing.append("Payment")
    if not net_col:
        missing.append("Net Amount")
    if missing:
        st.error(
            "تعذر العثور على أعمدة مهمة.\n"
            f"الأعمدة الناقصة: {', '.join(missing)}\n\n"
            f"الأعمدة الموجودة: {', '.join(map(str, df.columns))}"
        )
        return False

    # فلترة المحصّلين
    sales_vals = df[sales_col].astype(str).str.strip()
    keep_sales = ~sales_vals.isin(PROMISE_EXCLUDED_SALES)
    dropped_sales = int((~keep_sales).sum())
    df = df[keep_sales].copy()

    payments = df[payment_col].apply(_to_amount)
    nets = df[net_col].apply(_to_amount)
    originals = df[original_col].apply(_to_amount) if original_col else None
    sub_states = df[substate_col].astype(str).str.strip()

    error_rows = []
    rule_counter = {}
    for idx in df.index:
        sub = sub_states.loc[idx]
        pay = float(payments.loc[idx])
        net = float(nets.loc[idx])
        orig = float(originals.loc[idx]) if originals is not None else None
        found = _detect_case_errors_for_row(sub, pay, net, original_amt=orig)
        if not found:
            continue
        row = df.loc[idx].copy()
        reasons = [r[1] for r in found]
        rules = [r[0] for r in found]
        row[CASE_ERROR_REASON_COL] = " | ".join(reasons)
        row[CASE_ERROR_RULE_COL] = " | ".join(rules)
        row["_payment_num"] = pay
        row["_net_num"] = net
        error_rows.append(row)
        for rule, _ in found:
            rule_counter[rule] = rule_counter.get(rule, 0) + 1

    errors_df = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(columns=list(df.columns) + [CASE_ERROR_REASON_COL, CASE_ERROR_RULE_COL])

    st.session_state[CASE_ERRORS_RESULT_KEY] = {
        "df": errors_df,
        "source_df": df,
        "file_hash": file_hash,
        "filename": uploaded.name,
        "sales_col": sales_col,
        "substate_col": substate_col,
        "payment_col": payment_col,
        "net_col": net_col,
        "original_col": original_col,
        "total_in_file": total_in_file,
        "after_sales_filter": len(df),
        "dropped_sales": dropped_sales,
        "error_count": len(errors_df),
        "rule_counter": rule_counter,
    }
    return True


CASE_ERRORS_AGENT_FILTER_KEY = "case_errors_selected_agent"
CASE_ERROR_PALETTE = list(OPS_SCALE)  # نفس 3 درجات التويبات التشغيلية
CASE_ERROR_RULE_COLORS = {
    "Payment صفر مع حالة سداد": "#2F6F73",
    "Payment أكبر من صفر بدون حالة سداد": "#5A8A8D",
    "Net Amount صفر بدون حالة سداد": "#8FA8AB",
    "متبقي بعد سداد كامل": "#3D7A7D",
    "متبقي مع جدولة مقفلة": "#6B9598",
    "Net Amount صفر مع جدولة مقفلة بخصم": "#9BB0B3",
    "Net Amount صفر مع سداد بخصم": "#B2C1C3",
}


def get_case_errors_view(df, sales_col):
    if not sales_col or sales_col not in df.columns:
        return df
    selected = st.session_state.get(CASE_ERRORS_AGENT_FILTER_KEY)
    if not selected:
        return df
    mask = df[sales_col].astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(CASE_ERRORS_AGENT_FILTER_KEY, None)
        return df
    return df.loc[mask].copy()


def render_case_errors_filter_notice():
    selected = st.session_state.get(CASE_ERRORS_AGENT_FILTER_KEY)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر النشط: عرض أخطاء الحالات للمحصّل «{selected}»")
    with c2:
        if st.button("إظهار الكل", key="clear_case_errors_agent_filter", use_container_width=True):
            st.session_state.pop(CASE_ERRORS_AGENT_FILTER_KEY, None)
            st.rerun()


def render_case_errors_kpi(total_rows, error_count, agent_count, error_rate=0):
    cards = [
        ("📋<br>صفوف بعد الفلترة", total_rows, {"valueformat": ",d"}, THEME["text"]),
        ("🚨<br>أخطاء الحالات", error_count, {"valueformat": ",d"}, CASE_ERROR_PALETTE[0]),
        ("👥<br>محصّلون عليهم أخطاء", agent_count, {"valueformat": ",d"}, CASE_ERROR_PALETTE[1]),
        ("📈<br>نسبة الأخطاء %", error_rate, {"valueformat": ".1f", "suffix": "%"}, CASE_ERROR_PALETTE[2]),
    ]
    figure = go.Figure()
    gap = 0.018
    width = (1 - gap * (len(cards) + 1)) / len(cards)
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
                title={"text": label, "font": {"size": 16, "color": THEME["text_dim"]}, "align": "center"},
                number={"font": {"size": 30, "color": number_color}, **number_format},
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
    st.plotly_chart(figure, use_container_width=True, config=PLOTLY_CONFIG, key="case_errors_kpi")


def _show_case_errors_results(errors_df, meta):
    sales_col = meta.get("sales_col")
    substate_col = meta.get("substate_col")
    payment_col = meta.get("payment_col")
    net_col = meta.get("net_col")

    render_case_errors_filter_notice()
    df = get_case_errors_view(errors_df, sales_col)

    error_count = len(df)
    agent_count = int(df[sales_col].nunique()) if sales_col and sales_col in df.columns and error_count else 0
    selected_agent = st.session_state.get(CASE_ERRORS_AGENT_FILTER_KEY)
    if not selected_agent:
        st.success(
            f"✅ تم الفحص: {meta.get('after_sales_filter', 0):,} صف بعد استبعاد المحصّلين · "
            f"أخطاء: {meta.get('error_count', 0):,}"
        )
        st.caption(
            f"الملف: {meta.get('filename', '—')} · "
            f"مستبعدون: {meta.get('dropped_sales', 0):,} · "
            "💡 اضغط على محصّل في الشارت لتصفية النتائج عليه."
        )

    st.subheader("📊 ملخص أخطاء الحالات" + (f" — {selected_agent}" if selected_agent else ""))
    base_rows = meta.get("after_sales_filter", 0) if not selected_agent else error_count
    error_rate = (error_count / base_rows * 100.0) if base_rows else 0.0
    render_case_errors_kpi(
        base_rows,
        error_count,
        agent_count,
        error_rate,
    )

    if error_count == 0:
        st.info("لا توجد أخطاء حالات مطابقة للقواعد الحالية.")
        return

    if sales_col and sales_col in df.columns:
        st.markdown("#### 📈 تحليلات أخطاء الحالات")
        agent_counts = (
            df.groupby(sales_col)
            .size()
            .reset_index(name="عدد الأخطاء")
            .sort_values("عدد الأخطاء", ascending=True)
            .tail(15)
        )
        left, right = st.columns(2)
        with left:
            fig = px.bar(
                agent_counts,
                x="عدد الأخطاء",
                y=sales_col,
                orientation="h",
                text="عدد الأخطاء",
                color="عدد الأخطاء",
                color_continuous_scale=OPS_SCALE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                fig, "عدد الأخطاء حسب المحصّل",
                height=430, xaxis_title="عدد الأخطاء", show_legend=False,
                margin=dict(t=70, b=55, l=170, r=50),
            )
            fig.update_traces(
                texttemplate="%{x:,}",
                textposition="outside",
                cliponaxis=False,
                customdata=agent_counts[sales_col],
                hovertemplate="<b>%{y}</b><br>الأخطاء: %{x:,}<extra></extra>",
            )
            with st.container(border=True):
                render_selectable_chart(fig, "case_errors_by_agent", filter_key=CASE_ERRORS_AGENT_FILTER_KEY)

        with right:
            # إجمالي المديونية (Net Amount) للأخطاء حسب المحصّل
            if net_col and net_col in df.columns:
                net_work = df.copy()
                if "_net_num" in net_work.columns:
                    net_work["_amt"] = pd.to_numeric(net_work["_net_num"], errors="coerce").fillna(0)
                else:
                    net_work["_amt"] = pd.to_numeric(net_work[net_col], errors="coerce").fillna(0)
                agent_net = (
                    net_work.groupby(sales_col)["_amt"]
                    .sum()
                    .reset_index(name="إجمالي Net Amount")
                    .sort_values("إجمالي Net Amount", ascending=True)
                    .tail(15)
                )
                net_fig = px.bar(
                    agent_net,
                    x="إجمالي Net Amount",
                    y=sales_col,
                    orientation="h",
                    text="إجمالي Net Amount",
                    color="إجمالي Net Amount",
                    color_continuous_scale=OPS_SCALE,
                    template=PLOTLY_TEMPLATE,
                )
                _apply_ops_chart_style(
                    net_fig, "إجمالي Net Amount للأخطاء حسب المحصّل",
                    height=430, xaxis_title="Net Amount", show_legend=False,
                    margin=dict(t=70, b=55, l=170, r=50),
                )
                net_fig.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                    cliponaxis=False,
                    customdata=agent_net[sales_col],
                    hovertemplate="<b>%{y}</b><br>Net Amount: %{x:,.0f}<extra></extra>",
                )
                with st.container(border=True):
                    render_selectable_chart(net_fig, "case_errors_net_by_agent", filter_key=CASE_ERRORS_AGENT_FILTER_KEY)
            else:
                st.info("لا يوجد عمود Net Amount لعرض رسم المديونية.")

        if substate_col and substate_col in df.columns:
            state_counts = (
                df.groupby(substate_col)
                .size()
                .reset_index(name="العدد")
                .sort_values("العدد", ascending=True)
                .tail(12)
            )
            state_fig = px.bar(
                state_counts,
                x="العدد",
                y=substate_col,
                orientation="h",
                text="العدد",
                color="العدد",
                color_continuous_scale=OPS_SCALE,
                template=PLOTLY_TEMPLATE,
            )
            _apply_ops_chart_style(
                state_fig, "الأخطاء حسب Sub State",
                height=430, xaxis_title="العدد", show_legend=False,
                margin=dict(t=70, b=55, l=180, r=50),
            )
            state_fig.update_traces(texttemplate="%{x:,}", textposition="outside", cliponaxis=False)
            with st.container(border=True):
                st.plotly_chart(state_fig, use_container_width=True, config=PLOTLY_CONFIG, key="case_errors_by_state")

    display_cols = [
        c for c in [
            sales_col,
            substate_col,
            payment_col,
            net_col,
            CASE_ERROR_RULE_COL,
            CASE_ERROR_REASON_COL,
        ]
        if c and c in df.columns
    ]
    st.subheader("📋 تفاصيل أخطاء الحالات")
    if display_cols:
        st.dataframe(df[display_cols], use_container_width=True, hide_index=True)
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)

    def _excel_bytes(report_df, sheet_name):
        buf = io.BytesIO()
        export = report_df.drop(columns=[c for c in ["_payment_num", "_net_num"] if c in report_df.columns], errors="ignore")
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            export.to_excel(writer, index=False, sheet_name=sheet_name)
        return buf.getvalue()

    report_date = datetime.now().strftime("%Y-%m-%d")
    st.download_button(
        "⬇️ تحميل تقرير أخطاء الحالات",
        data=_excel_bytes(df, "أخطاء الحالات"),
        file_name=f"تقرير_أخطاء_الحالات_{report_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
        key="case_errors_download",
        disabled=df.empty,
    )


def page_case_errors():
    """صفحة أخطاء الحالات: تعارض Sub State مع Payment / Net Amount."""
    page_header(
        "CASE ERRORS",
        "🧾 أخطاء الحالات",
        "ارفع المحفظة الحديثة لاكتشاف تعارضات الحالة مع Payment و Net Amount",
    )

    upload_key = "case_errors_upload"
    cache_scope = "case_errors_upload"
    result_keys = (CASE_ERRORS_RESULT_KEY,)

    uploaded = st.file_uploader(
        "📂 ارفع المحفظة الحديثة (Excel أو CSV)",
        type=["xlsx", "xls", "csv"],
        key=upload_key,
        on_change=sync_file_cache,
        args=(upload_key, cache_scope, result_keys),
    )

    if uploaded is not None:
        st.caption(f"الملف: {uploaded.name}")
        with st.spinner("جارٍ فحص أخطاء الحالات..."):
            _run_case_errors_pipeline(uploaded)
    else:
        cached_file = st.session_state.get(APP_DATA_CACHE_KEY, {}).get(cache_scope, {})
        cached = st.session_state.get(CASE_ERRORS_RESULT_KEY)
        if cached or cached_file:
            # الملف اتشال → sync_file_cache يمسح النتائج؛ لو لسه فيه كاش قديم نعرض رسالة رفع
            pass

    cached = st.session_state.get(CASE_ERRORS_RESULT_KEY)
    if cached and cached.get("df") is not None:
        if uploaded is None:
            st.success(
                f"✅ نتيجة محفوظة من: {cached.get('filename', '—')}. "
                "لن تُحذف عند التنقل بين التبويبات."
            )
        _show_case_errors_results(cached["df"], cached)
    else:
        st.info("📂 ارفع ملف المحفظة لبدء فحص أخطاء الحالات.")



# ==========================================================
# تويب التوزيع — توزيع عملاء المحصل المستقيل / المحصل الجديد
# ==========================================================

DISTRIBUTION_RESULT_KEY = "distribution_result"
DISTRIBUTION_SCENARIO_KEY = "distribution_scenario"
DISTRIBUTION_UPLOAD_KEY = "distribution_upload"
DISTRIBUTION_CACHE_SCOPE = "distribution_upload"


def _distribution_find_numeric_cols(df):
    """يرجع الأعمدة الرقمية المحتملة للمبالغ."""
    numeric = []
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            numeric.append(c)
            continue
        # جرب تحويل عينة
        sample = pd.to_numeric(df[c].dropna().head(50), errors="coerce")
        if sample.notna().sum() >= max(3, len(sample) // 2):
            numeric.append(c)
    return numeric


def _greedy_assign_customers(
    customers_df,
    target_collectors,
    amount_col,
    max_diff_amount,
    max_diff_clients,
    max_diff_accounts,
    client_weight=1.0,
    amount_weight=1.0,
    account_weight=1.0,
):
    """
    توزيع مطالبات/عملاء المحصل المستقيل فقط على المستهدفين.
    الهدف: أقرب تساوي ممكن في (المبلغ، عدد العملاء، عدد الحسابات/المطالبات)
    بين المحصلين المستهدفين — مش موازنة المحفظة الأصلية بتاعتهم.

    مراحل:
    1) إسناد أولي: كل عميل يروح للأقل حملاً حسب score مطبّع.
    2) بحث محلي مكثّف: نقل + تبادل عملاء كاملين لحد ثبات التحسين.
    """
    if customers_df.empty or not target_collectors:
        return {}, pd.DataFrame(), []

    cust_info = {}
    for _, row in customers_df.iterrows():
        ck = row["client_key"]
        cust_info[ck] = {
            "amount": float(row["amount"] or 0),
            "n_accounts": int(row["n_accounts"] or 0),
            "n_rows": int(row["n_rows"] or 0),
        }

    n_targets = len(target_collectors)
    total_amount = float(customers_df["amount"].sum())
    total_clients = int(len(customers_df))
    total_accounts = int(customers_df["n_accounts"].sum())
    total_rows = int(customers_df["n_rows"].sum())

    avg_amount = max(total_amount / n_targets, 1e-9)
    avg_clients = max(total_clients / n_targets, 1e-9)
    avg_accounts = max(total_accounts / n_targets, 1e-9)
    avg_rows = max(total_rows / n_targets, 1e-9)

    state = {
        name: {"amount": 0.0, "clients": 0, "accounts": 0, "rows": 0, "client_keys": []}
        for name in target_collectors
    }

    ordered = customers_df.sort_values(
        ["amount", "n_accounts", "n_rows"], ascending=[False, False, False]
    ).reset_index(drop=True)

    def _ranges(st):
        amts = [st[n]["amount"] for n in target_collectors]
        clis = [st[n]["clients"] for n in target_collectors]
        accs = [st[n]["accounts"] for n in target_collectors]
        rows = [st[n]["rows"] for n in target_collectors]
        return (
            max(amts) - min(amts),
            max(clis) - min(clis),
            max(accs) - min(accs),
            max(rows) - min(rows),
        )

    def _cost(st):
        """تكلفة: تقليل فروقات المبالغ والعملاء والحسابات والمطالبات للمطالبات الموزّعة فقط."""
        a, c, k, r = _ranges(st)
        # عقوبات لو عدّينا الـ tolerance
        pen = 0.0
        if max_diff_amount is not None and a > max_diff_amount:
            pen += 5000.0 * ((a - max_diff_amount) / avg_amount)
        if max_diff_clients is not None and c > max_diff_clients:
            pen += 2000.0 * (c - max_diff_clients)
        if max_diff_accounts is not None and k > max_diff_accounts:
            pen += 2000.0 * (k - max_diff_accounts)

        base = (
            amount_weight * (a / avg_amount)
            + client_weight * (c / avg_clients)
            + account_weight * (k / avg_accounts)
            + 0.5 * account_weight * (r / avg_rows)  # المطالبات كبُعد إضافي
        )
        # ثانوي: مجموع مربعات الانحراف عن المتوسط (يساعد على التوزيع المتساوي)
        amt_var = sum((st[n]["amount"] - avg_amount) ** 2 for n in target_collectors) / (avg_amount ** 2)
        cli_var = sum((st[n]["clients"] - avg_clients) ** 2 for n in target_collectors) / (avg_clients ** 2)
        acc_var = sum((st[n]["accounts"] - avg_accounts) ** 2 for n in target_collectors) / (avg_accounts ** 2)
        base += 0.15 * (amount_weight * amt_var + client_weight * cli_var + account_weight * acc_var)
        return (pen + base, a, c, k, r)

    def _imbalance_assign(st):
        amts = [st[n]["amount"] / avg_amount for n in target_collectors]
        clis = [st[n]["clients"] / avg_clients for n in target_collectors]
        accs = [st[n]["accounts"] / avg_accounts for n in target_collectors]
        def _r(v):
            return max(v) - min(v)
        return amount_weight * _r(amts) + client_weight * _r(clis) + account_weight * _r(accs)

    # ----- 1) إسناد أولي -----
    for _, row in ordered.iterrows():
        ck = row["client_key"]
        info = cust_info[ck]
        amt, nacc, nrows = info["amount"], info["n_accounts"], info["n_rows"]

        best, best_score = None, None
        for name in target_collectors:
            state[name]["amount"] += amt
            state[name]["clients"] += 1
            state[name]["accounts"] += nacc
            state[name]["rows"] += nrows
            score = _imbalance_assign(state)
            # بعد الإضافة: فضّل الأقل مبلغاً ثم عملاء
            tie = (score, state[name]["amount"], state[name]["clients"], state[name]["accounts"], name)
            if best_score is None or tie < best_score:
                best_score = tie
                best = name
            state[name]["amount"] -= amt
            state[name]["clients"] -= 1
            state[name]["accounts"] -= nacc
            state[name]["rows"] -= nrows

        state[best]["amount"] += amt
        state[best]["clients"] += 1
        state[best]["accounts"] += nacc
        state[best]["rows"] += nrows
        state[best]["client_keys"].append(ck)

    # ----- 2) بحث محلي: نقل + تبادل -----
    moved_count = 0
    swap_count = 0
    max_rounds = max(400, len(customers_df) * 6)

    def _transfer(src, dst, ck):
        info = cust_info[ck]
        state[src]["client_keys"].remove(ck)
        state[src]["amount"] -= info["amount"]
        state[src]["clients"] -= 1
        state[src]["accounts"] -= info["n_accounts"]
        state[src]["rows"] -= info["n_rows"]
        state[dst]["client_keys"].append(ck)
        state[dst]["amount"] += info["amount"]
        state[dst]["clients"] += 1
        state[dst]["accounts"] += info["n_accounts"]
        state[dst]["rows"] += info["n_rows"]

    for _round in range(max_rounds):
        cur = _cost(state)
        cur_cost = cur[0]
        improved = False

        by_amt_desc = sorted(target_collectors, key=lambda n: (state[n]["amount"], state[n]["accounts"], state[n]["clients"]), reverse=True)
        by_amt_asc = list(reversed(by_amt_desc))
        rich_list = by_amt_desc[: min(8, n_targets)]
        poor_list = by_amt_asc[: min(8, n_targets)]

        # (أ) نقل
        best_transfer = None
        for src in rich_list:
            if state[src]["clients"] == 0:
                continue
            src_keys = sorted(state[src]["client_keys"], key=lambda ck: cust_info[ck]["amount"])
            for dst in poor_list:
                if src == dst or state[src]["amount"] <= state[dst]["amount"]:
                    continue
                for ck in src_keys:
                    info = cust_info[ck]
                    # مؤقت
                    state[src]["amount"] -= info["amount"]
                    state[src]["clients"] -= 1
                    state[src]["accounts"] -= info["n_accounts"]
                    state[src]["rows"] -= info["n_rows"]
                    state[dst]["amount"] += info["amount"]
                    state[dst]["clients"] += 1
                    state[dst]["accounts"] += info["n_accounts"]
                    state[dst]["rows"] += info["n_rows"]
                    new = _cost(state)
                    state[src]["amount"] += info["amount"]
                    state[src]["clients"] += 1
                    state[src]["accounts"] += info["n_accounts"]
                    state[src]["rows"] += info["n_rows"]
                    state[dst]["amount"] -= info["amount"]
                    state[dst]["clients"] -= 1
                    state[dst]["accounts"] -= info["n_accounts"]
                    state[dst]["rows"] -= info["n_rows"]
                    if new[0] < cur_cost - 1e-12:
                        if best_transfer is None or new[0] < best_transfer[0][0]:
                            best_transfer = (new, src, dst, ck)

        if best_transfer is not None:
            _, src, dst, ck = best_transfer
            _transfer(src, dst, ck)
            moved_count += 1
            improved = True
            continue

        # (ب) تبادل 1↔1
        best_swap = None
        pairs = []
        for a in rich_list:
            for b in poor_list:
                if a != b:
                    pairs.append((a, b))
        seen = set()
        uniq = []
        for a, b in pairs:
            key = tuple(sorted((a, b)))
            if key not in seen:
                seen.add(key)
                uniq.append((a, b))

        for a_name, b_name in uniq:
            keys_a = sorted(state[a_name]["client_keys"], key=lambda ck: cust_info[ck]["amount"], reverse=True)
            keys_b = sorted(state[b_name]["client_keys"], key=lambda ck: cust_info[ck]["amount"], reverse=True)
            if not keys_a or not keys_b:
                continue
            cand_a = list(dict.fromkeys(keys_a[:15] + keys_a[-15:]))
            cand_b = list(dict.fromkeys(keys_b[:15] + keys_b[-15:]))
            for ck_a in cand_a:
                ia = cust_info[ck_a]
                for ck_b in cand_b:
                    ib = cust_info[ck_b]
                    # فرق المبلغ بعد التبادل على a: +ib -ia
                    da = ib["amount"] - ia["amount"]
                    dk = ib["n_accounts"] - ia["n_accounts"]
                    dr = ib["n_rows"] - ia["n_rows"]
                    state[a_name]["amount"] += da
                    state[a_name]["accounts"] += dk
                    state[a_name]["rows"] += dr
                    state[b_name]["amount"] -= da
                    state[b_name]["accounts"] -= dk
                    state[b_name]["rows"] -= dr
                    new = _cost(state)
                    state[a_name]["amount"] -= da
                    state[a_name]["accounts"] -= dk
                    state[a_name]["rows"] -= dr
                    state[b_name]["amount"] += da
                    state[b_name]["accounts"] += dk
                    state[b_name]["rows"] += dr
                    if new[0] < cur_cost - 1e-12:
                        if best_swap is None or new[0] < best_swap[0][0]:
                            best_swap = (new, a_name, ck_a, b_name, ck_b)

        if best_swap is not None:
            _, a_name, ck_a, b_name, ck_b = best_swap
            _transfer(a_name, b_name, ck_a)
            _transfer(b_name, a_name, ck_b)
            swap_count += 1
            improved = True
            continue

        if not improved:
            break

    # ----- 3) مرحلة أخيرة: ركّز على تقليل فرق المبالغ فقط لو لسه فوق الـ tolerance -----
    # مع السماح بفرق عملاء/حسابات في حدود الـ tolerance
    if max_diff_amount is not None:
        for _round in range(max_rounds // 2):
            a_spread, c_spread, k_spread, _ = _ranges(state)
            if a_spread <= max_diff_amount:
                break
            by_amt_desc = sorted(target_collectors, key=lambda n: state[n]["amount"], reverse=True)
            by_amt_asc = list(reversed(by_amt_desc))
            src = by_amt_desc[0]
            dst = by_amt_asc[0]
            if src == dst or state[src]["amount"] <= state[dst]["amount"]:
                break

            best_ck = None
            best_new_spread = a_spread
            for ck in list(state[src]["client_keys"]):
                info = cust_info[ck]
                # بعد النقل
                new_src = state[src]["amount"] - info["amount"]
                new_dst = state[dst]["amount"] + info["amount"]
                # spread تقريبي
                others_max = max((state[n]["amount"] for n in target_collectors if n not in (src, dst)), default=0.0)
                others_min = min((state[n]["amount"] for n in target_collectors if n not in (src, dst)), default=new_src)
                new_spread = max(others_max, new_src, new_dst) - min(others_min, new_src, new_dst)
                # قيود العملاء/الحسابات
                new_c_src = state[src]["clients"] - 1
                new_c_dst = state[dst]["clients"] + 1
                new_k_src = state[src]["accounts"] - info["n_accounts"]
                new_k_dst = state[dst]["accounts"] + info["n_accounts"]
                all_c = [state[n]["clients"] for n in target_collectors if n not in (src, dst)] + [new_c_src, new_c_dst]
                all_k = [state[n]["accounts"] for n in target_collectors if n not in (src, dst)] + [new_k_src, new_k_dst]
                new_c_sp = max(all_c) - min(all_c)
                new_k_sp = max(all_k) - min(all_k)
                if max_diff_clients is not None and new_c_sp > max(max_diff_clients, c_spread):
                    continue
                if max_diff_accounts is not None and new_k_sp > max(max_diff_accounts, k_spread):
                    continue
                if new_spread < best_new_spread - 1e-6:
                    best_new_spread = new_spread
                    best_ck = ck

            if best_ck is None:
                # جرب تبادل لتقليل فرق المبلغ
                found_swap = False
                for ck_a in list(state[src]["client_keys"]):
                    ia = cust_info[ck_a]
                    for ck_b in list(state[dst]["client_keys"]):
                        ib = cust_info[ck_b]
                        if ia["amount"] <= ib["amount"]:
                            continue  # لازم src يدي أكبر وياخد أصغر
                        new_src = state[src]["amount"] - ia["amount"] + ib["amount"]
                        new_dst = state[dst]["amount"] - ib["amount"] + ia["amount"]
                        others_max = max((state[n]["amount"] for n in target_collectors if n not in (src, dst)), default=0.0)
                        others_min = min((state[n]["amount"] for n in target_collectors if n not in (src, dst)), default=new_src)
                        new_spread = max(others_max, new_src, new_dst) - min(others_min, new_src, new_dst)
                        if new_spread < a_spread - 1e-6:
                            _transfer(src, dst, ck_a)
                            _transfer(dst, src, ck_b)
                            swap_count += 1
                            found_swap = True
                            break
                    if found_swap:
                        break
                if not found_swap:
                    break
            else:
                _transfer(src, dst, best_ck)
                moved_count += 1

    assign_map = {}
    for name in target_collectors:
        for ck in state[name]["client_keys"]:
            assign_map[ck] = name

    summary_rows = []
    for name in target_collectors:
        s = state[name]
        summary_rows.append({
            "المحصّل": name,
            "عدد العملاء": s["clients"],
            "عدد الحسابات": s["accounts"],
            "عدد المطالبات": s["rows"],
            "إجمالي المبلغ": round(s["amount"], 2),
            "الهدف_مبلغ": round(avg_amount, 2),
            "الهدف_عملاء": round(avg_clients, 2),
            "الهدف_حسابات": round(avg_accounts, 2),
        })
    summary = pd.DataFrame(summary_rows)
    # رتب من الأعلى مبلغ للأقل عشان يتقرأ بسهولة
    summary = summary.sort_values("إجمالي المبلغ", ascending=False).reset_index(drop=True)

    warnings = []
    if len(summary) > 1:
        amounts = summary["إجمالي المبلغ"]
        clients = summary["عدد العملاء"]
        accounts = summary["عدد الحسابات"]
        amount_spread = float(amounts.max() - amounts.min())
        client_spread = int(clients.max() - clients.min())
        account_spread = int(accounts.max() - accounts.min())
        total_moves = moved_count + swap_count
        if total_moves:
            warnings.append(f"تحسين: نقل {moved_count} + تبادل {swap_count} — المتوسط المستهدف للمبلغ ≈ {avg_amount:,.0f}")
        if max_diff_amount is not None and amount_spread > max_diff_amount:
            warnings.append(
                f"فرق المبالغ ({amount_spread:,.0f}) أكبر من المسموح ({max_diff_amount:,.0f}) "
                f"— أقرب توازن ممكن بدون تقسيم عميل"
            )
        if max_diff_clients is not None and client_spread > max_diff_clients:
            warnings.append(f"فرق عدد العملاء ({client_spread}) أكبر من المسموح ({max_diff_clients})")
        if max_diff_accounts is not None and account_spread > max_diff_accounts:
            warnings.append(f"فرق عدد الحسابات ({account_spread}) أكبر من المسموح ({max_diff_accounts})")
        summary.attrs["spreads"] = {
            "amount": amount_spread,
            "clients": client_spread,
            "accounts": account_spread,
        }
        summary.attrs["moved_count"] = moved_count
        summary.attrs["swap_count"] = swap_count
        summary.attrs["avg_amount"] = avg_amount

    return assign_map, summary, warnings



def page_distribution():
    """توزيع عملاء المحصل المستقيل على باقي المحصلين (أو إنشاء محفظة لمحصل جديد)."""
    page_header(
        "DISTRIBUTION",
        "⚖️ التوزيع",
        "توزيع عملاء المحصل المستقيل بالتساوي على المحصلين المختارين، مع موازنة المبالغ وعدد العملاء والحسابات",
    )

    # ---- اختيار السيناريو (radio زي الوعود + لوجو جنب كل حالة) ----
    scenario_labels = {
        "resigned": "🚪 موظف مستقيل",
        "new_collector": "🆕 محصّل جديد",
    }
    label_to_key = {v: k for k, v in scenario_labels.items()}
    current = st.session_state.get(DISTRIBUTION_SCENARIO_KEY, "resigned")
    current_label = scenario_labels.get(current, scenario_labels["resigned"])
    chosen_label = st.radio(
        "اختر حالة التوزيع",
        options=list(scenario_labels.values()),
        index=list(scenario_labels.values()).index(current_label),
        horizontal=True,
        key="dist_scenario_radio",
    )
    scenario = label_to_key[chosen_label]
    st.session_state[DISTRIBUTION_SCENARIO_KEY] = scenario

    if scenario == "new_collector":
        st.info("🚧 حالة «محصل جديد — بناء محفظة» هتتضاف في الخطوة الجاية. حالياً ركزنا على حالة الموظف المستقيل.")
        st.caption("الفكرة: تختار المحصلين اللي هياخد منهم عملاء، وتحدد نسب أو أعداد، ويتبنى ملف محفظة جديد للمحصل الجديد.")
        return

    # ---- حالة المحصل المستقيل ----
    st.markdown("---")
    st.subheader("1️⃣ رفع المحفظة واختيار المحصل المستقيل")

    uploaded = st.file_uploader(
        "📂 ارفع ملف المحفظة (Excel أو CSV)",
        type=["xlsx", "xls", "csv"],
        key=DISTRIBUTION_UPLOAD_KEY,
        on_change=sync_file_cache,
        args=(DISTRIBUTION_UPLOAD_KEY, DISTRIBUTION_CACHE_SCOPE, [DISTRIBUTION_RESULT_KEY]),
    )

    if uploaded is None:
        cached = st.session_state.get(DISTRIBUTION_RESULT_KEY)
        if cached:
            st.success(f"✅ نتيجة توزيع محفوظة من: {cached.get('filename', '—')}")
            _render_distribution_results(cached)
        else:
            st.info("📂 ارفع ملف المحفظة لبدء التوزيع.")
        return

    try:
        raw_df = read_uploaded_dataframe(uploaded)
    except Exception as e:
        st.error(f"تعذر قراءة الملف: {e}")
        return

    # حذف أول صف بعد العناوين لو موجود (نمط باقي التطبيق)
    df = raw_df.iloc[1:].copy() if len(raw_df) > 1 else raw_df.copy()
    df = df.reset_index(drop=True)

    # اختيار عمود المحصّل وعمود Sales Team من الملف (بدون تغيير باقي منطق التوزيع)
    col_options = list(df.columns)
    detected_sales = find_column(df, SALES_PERSON_CANDIDATES)
    detected_team = find_column(df, SALES_TEAM_CANDIDATES)

    st.markdown("##### 🔧 تحديد الأعمدة")
    col_pick_1, col_pick_2 = st.columns(2)
    with col_pick_1:
        sales_default_idx = col_options.index(detected_sales) if detected_sales in col_options else 0
        sales_col = st.selectbox(
            "👤 عمود المحصّل (Sales Person)",
            options=col_options,
            index=sales_default_idx,
            key="dist_sales_col_select",
            help="العمود اللي فيه اسم المحصل المستقيل وباقي المحصلين",
        )
    with col_pick_2:
        team_options = ["— بدون فلترة بـ Sales Team —"] + col_options
        team_default_idx = (
            team_options.index(detected_team)
            if detected_team and detected_team in team_options
            else 0
        )
        team_col_choice = st.selectbox(
            "👥 عمود Sales Team (اختياري)",
            options=team_options,
            index=team_default_idx,
            key="dist_team_col_select",
            help="لو اخترت عمود فريق، تقدر تفلتر المحفظة على فريق معيّن قبل التوزيع",
        )
        team_col = None if isinstance(team_col_choice, str) and team_col_choice.startswith("—") else team_col_choice

    if team_col and team_col in df.columns:
        team_vals = df[team_col].astype(str).str.strip()
        team_choices = sorted(
            {
                v for v in team_vals.tolist()
                if v and v.lower() not in {"nan", "none", "null", ""}
            }
        )
        if team_choices:
            selected_teams = st.multiselect(
                "فلترة حسب Sales Team (اختياري — فاضي = كل الفرق)",
                options=team_choices,
                default=[],
                key="dist_team_filter_values",
            )
            if selected_teams:
                df = df[team_vals.isin(selected_teams)].copy()
                df = df.reset_index(drop=True)
                st.caption(f"بعد فلترة الفريق: **{len(df):,}** صف")

    if not sales_col or sales_col not in df.columns:
        st.error(
            "تعذر استخدام عمود المحصّل المختار. "
            f"الأعمدة الموجودة: {', '.join(map(str, df.columns))}"
        )
        return

    # أعمدة المفاتيح
    customer_id_col = (
        find_column(df, WALLET_CUSTOMER_ID_CANDIDATES)
        or find_column(df, ACCOUNT_NUMBER_CANDIDATES)
        or find_column(df, ID_CANDIDATES)
    )
    account_col = find_column(df, ACCOUNT_NUMBER_CANDIDATES) or customer_id_col
    net_col = find_column(df, PROMISE_NET_AMOUNT_CANDIDATES)
    numeric_cols = _distribution_find_numeric_cols(df)

    team_label = team_col if team_col else "—"
    st.caption(
        f"الملف: **{uploaded.name}** · عدد الصفوف: **{len(df):,}** · "
        f"عمود المحصّل: **{sales_col}** · Sales Team: **{team_label}**"
    )

    # قائمة المحصلين
    sales_vals = df[sales_col].astype(str).str.strip()
    all_collectors = sorted({v for v in sales_vals.tolist() if v and v.lower() not in {"nan", "none", "null", ""}})

    if len(all_collectors) < 2:
        st.warning("الملف يحتوي على محصل واحد فقط أو لا يوجد محصلين — لا يمكن التوزيع.")
        return

    resigned = st.selectbox(
        "👤 اختر المحصل المستقيل (هيتوزع عملاؤه)",
        options=all_collectors,
        key="dist_resigned_select",
    )

    resigned_mask = sales_vals == resigned
    resigned_df = df.loc[resigned_mask].copy()
    others = [c for c in all_collectors if c != resigned]

    st.markdown(f"**عملاء/صفوف المحصل المستقيل:** {len(resigned_df):,} صف")

    # ---- اختيار المستهدفين ----
    st.markdown("---")
    st.subheader("2️⃣ المحصلين اللي هيتوزع عليهم")
    st.caption("حدد من القائمة مين هياخد من عملاء المستقيل. افتراضياً كل الباقي محددين.")

    # checkboxes في شبكة
    if "dist_target_collectors" not in st.session_state:
        st.session_state["dist_target_collectors"] = list(others)

    # أزرار تحديد الكل / إلغاء
    b1, b2, _ = st.columns([1, 1, 2])
    with b1:
        if st.button("✅ تحديد الكل", key="dist_select_all", use_container_width=True):
            st.session_state["dist_target_collectors"] = list(others)
            st.rerun()
    with b2:
        if st.button("⬜ إلغاء الكل", key="dist_select_none", use_container_width=True):
            st.session_state["dist_target_collectors"] = []
            st.rerun()

    n_cols = 3
    cols = st.columns(n_cols)
    selected_targets = []
    for i, name in enumerate(others):
        with cols[i % n_cols]:
            checked = name in st.session_state.get("dist_target_collectors", others)
            if st.checkbox(name, value=checked, key=f"dist_cb_{i}_{name}"):
                selected_targets.append(name)

    st.session_state["dist_target_collectors"] = selected_targets

    if not selected_targets:
        st.warning("لازم تختار محصل واحد على الأقل للتوزيع عليه.")
        return

    st.success(f"سيتم التوزيع على **{len(selected_targets)}** محصل.")

    # ---- أعمدة الموازنة ----
    st.markdown("---")
    st.subheader("3️⃣ أعمدة الموازنة والـ Tolerances")

    c1, c2 = st.columns(2)
    with c1:
        # عمود العميل (مفتاح عدم التكرار)
        id_options = [c for c in [customer_id_col, account_col] if c] + [c for c in df.columns if c not in {customer_id_col, account_col, sales_col}]
        id_options = list(dict.fromkeys([c for c in id_options if c]))  # unique preserve order
        client_key_col = st.selectbox(
            "🔑 عمود هوية العميل (عشان ميتوزعش عميل على أكتر من محصل)",
            options=id_options,
            index=0 if customer_id_col in id_options else 0,
            key="dist_client_key",
        )
        account_key_col = st.selectbox(
            "📋 عمود الحساب (لعدّ عدد الحسابات) — ممكن يكون نفس عمود العميل",
            options=id_options,
            index=id_options.index(account_col) if account_col in id_options else 0,
            key="dist_account_key",
        )
    with c2:
        amount_options = [c for c in numeric_cols if c]
        if net_col and net_col not in amount_options:
            amount_options = [net_col] + amount_options
        if not amount_options:
            amount_options = list(df.columns)
        default_amount = [net_col] if net_col and net_col in amount_options else (amount_options[:1] if amount_options else [])
        amount_cols = st.multiselect(
            "💰 أعمدة المبالغ اللي هتتوازن (هيتجمعوا لكل عميل)",
            options=amount_options,
            default=default_amount,
            key="dist_amount_cols",
        )

    st.markdown("**⚖️ أوزان الموازنة** — كلّما زوّدت وزن بُعد، الخوارزمية هتهتم تقلل فرقه أكتر:")
    w1, w2, w3 = st.columns(3)
    with w1:
        amount_weight = st.slider("وزن المبالغ", 0.0, 3.0, 1.0, 0.1, key="dist_w_amt")
    with w2:
        client_weight = st.slider("وزن عدد العملاء", 0.0, 3.0, 1.5, 0.1, key="dist_w_cli",
                                  help="افتراضي أعلى شوية عشان العملاء متتهملش")
    with w3:
        account_weight = st.slider("وزن عدد الحسابات", 0.0, 3.0, 1.5, 0.1, key="dist_w_acc",
                                   help="افتراضي أعلى شوية عشان الحسابات متتهملش")

    st.markdown("**أقصى فرق مسموح بعد التوزيع (Tolerance):**")
    t1, t2, t3 = st.columns(3)
    with t1:
        max_diff_amount = st.number_input(
            "أقصى فرق في المبالغ",
            min_value=0.0,
            value=5000.0,
            step=500.0,
            key="dist_max_diff_amt",
            help="لو الفرق بين أعلى وأقل محصل في المبلغ زاد عن الرقم ده هيظهر تحذير",
        )
    with t2:
        max_diff_clients = st.number_input(
            "أقصى فرق في عدد العملاء",
            min_value=0,
            value=2,
            step=1,
            key="dist_max_diff_cli",
        )
    with t3:
        max_diff_accounts = st.number_input(
            "أقصى فرق في عدد الحسابات",
            min_value=0,
            value=3,
            step=1,
            key="dist_max_diff_acc",
        )

    # ---- تشغيل التوزيع ----
    st.markdown("---")
    run = st.button("🚀 تنفيذ التوزيع", type="primary", use_container_width=True, key="dist_run_btn")

    if not run and DISTRIBUTION_RESULT_KEY not in st.session_state:
        st.info("اضغط «تنفيذ التوزيع» بعد ضبط الاختيارات.")
        return

    if run:
        if not amount_cols:
            st.error("اختار عمود مبلغ واحد على الأقل.")
            return
        if not client_key_col:
            st.error("اختار عمود هوية العميل.")
            return

        work = resigned_df.copy()
        # تنظيف المفاتيح
        work["_client_key"] = work[client_key_col].astype(str).str.strip()
        work["_account_key"] = work[account_key_col].astype(str).str.strip()
        work = work[work["_client_key"].notna() & work["_client_key"].ne("") & work["_client_key"].str.lower().ne("nan")]

        # حساب المبلغ لكل صف
        for col in amount_cols:
            work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0)
        work["_row_amount"] = work[amount_cols].sum(axis=1)

        # تجميع على مستوى العميل
        grouped = work.groupby("_client_key", as_index=False).agg(
            amount=("_row_amount", "sum"),
            n_accounts=("_account_key", "nunique"),
            n_rows=("_row_amount", "count"),
        )
        grouped = grouped.rename(columns={"_client_key": "client_key"})

        if grouped.empty:
            st.error("لا توجد سجلات صالحة للتوزيع بعد تنظيف المفاتيح.")
            return

        assign_map, summary, warnings = _greedy_assign_customers(
            grouped,
            selected_targets,
            amount_col="amount",
            max_diff_amount=max_diff_amount,
            max_diff_clients=max_diff_clients,
            max_diff_accounts=max_diff_accounts,
            amount_weight=amount_weight,
            client_weight=client_weight,
            account_weight=account_weight,
        )

        # --- تطبيق الإسناد على الملف ---
        # تطبيع المفاتيح عشان الـ map يطابق الصفوف فعلاً
        def _norm_key(s):
            return (
                s.astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)  # 123.0 -> 123 من إكسيل
                .str.replace(r"\s+", " ", regex=True)
            )

        assign_map_norm = {_norm_key(pd.Series([k])).iloc[0]: v for k, v in assign_map.items()}

        full_df = df.copy()
        resigned_mask = full_df[sales_col].astype(str).str.strip() == resigned
        resigned_idx = full_df.index[resigned_mask]
        keys_on_rows = _norm_key(full_df.loc[resigned_idx, client_key_col])
        mapped = keys_on_rows.map(assign_map_norm)
        unmapped = int(mapped.isna().sum())
        if unmapped:
            # محاولة ثانية بالمفتاح الخام
            mapped2 = full_df.loc[resigned_idx, client_key_col].astype(str).str.strip().map(assign_map)
            mapped = mapped.fillna(mapped2)
            unmapped = int(mapped.isna().sum())
        full_df.loc[resigned_idx, sales_col] = mapped.fillna(selected_targets[0]).values

        # المطالبات الموزّعة فقط (بعد الإسناد) — ده المصدر الحقيقي للتحقق
        distributed_df = full_df.loc[resigned_idx].copy()
        # أعمدة المبلغ على مستوى الصف
        for col in amount_cols:
            if col in distributed_df.columns:
                distributed_df[col] = pd.to_numeric(distributed_df[col], errors="coerce").fillna(0)
        if amount_cols:
            distributed_df["_amt"] = distributed_df[amount_cols].sum(axis=1)
        else:
            distributed_df["_amt"] = 0.0
        distributed_df["_client"] = _norm_key(distributed_df[client_key_col])
        distributed_df["_account"] = _norm_key(distributed_df[account_key_col]) if account_key_col in distributed_df.columns else distributed_df["_client"]

        # ملخص فعلي من الشيت بعد الإسناد (مش من الخوارزمية بس)
        actual_summary = (
            distributed_df.groupby(sales_col, as_index=False)
            .agg(
                **{
                    "عدد العملاء": ("_client", "nunique"),
                    "عدد الحسابات": ("_account", "nunique"),
                    "عدد المطالبات": ("_amt", "count"),
                    "إجمالي المبلغ": ("_amt", "sum"),
                }
            )
            .sort_values("إجمالي المبلغ", ascending=False)
            .reset_index(drop=True)
        )
        actual_summary = actual_summary.rename(columns={sales_col: "المحصّل"})
        actual_summary["إجمالي المبلغ"] = actual_summary["إجمالي المبلغ"].round(2)

        # استبدال ملخص الخوارزمية بالملخص الفعلي من البيانات المكتوبة
        summary = actual_summary

        if len(summary) > 1:
            amount_spread = float(summary["إجمالي المبلغ"].max() - summary["إجمالي المبلغ"].min())
            client_spread = int(summary["عدد العملاء"].max() - summary["عدد العملاء"].min())
            account_spread = int(summary["عدد الحسابات"].max() - summary["عدد الحسابات"].min())
            warnings = list(warnings or [])
            # حدّث التحذيرات حسب الواقع
            warnings = [w for w in warnings if "فرق المبالغ" not in w and "فرق عدد" not in w]
            if max_diff_amount is not None and amount_spread > max_diff_amount:
                warnings.append(
                    f"فرق المبالغ الفعلي في الشيت ({amount_spread:,.0f}) أكبر من المسموح ({max_diff_amount:,.0f})"
                )
            if max_diff_clients is not None and client_spread > max_diff_clients:
                warnings.append(f"فرق عدد العملاء الفعلي ({client_spread}) أكبر من المسموح ({max_diff_clients})")
            if max_diff_accounts is not None and account_spread > max_diff_accounts:
                warnings.append(f"فرق عدد الحسابات الفعلي ({account_spread}) أكبر من المسموح ({max_diff_accounts})")
            if unmapped:
                warnings.append(f"⚠️ {unmapped} صف لم يُطابق مفتاح العميل وتم إسناده افتراضياً — راجع عمود الهوية")

        # تفاصيل الإسناد: صف واحد لكل عميل
        detail = grouped.copy()
        detail["client_key_norm"] = _norm_key(detail["client_key"])
        detail["المحصّل_الجديد"] = detail["client_key_norm"].map(assign_map_norm).fillna(
            detail["client_key"].astype(str).str.strip().map(assign_map)
        )
        detail = detail.rename(columns={
            "client_key": "هوية العميل",
            "amount": "المبلغ",
            "n_accounts": "عدد الحسابات",
            "n_rows": "عدد المطالبات",
        })
        detail = detail[["هوية العميل", "المحصّل_الجديد", "المبلغ", "عدد الحسابات", "عدد المطالبات"]]
        detail = detail.sort_values(["المحصّل_الجديد", "المبلغ"], ascending=[True, False]).reset_index(drop=True)

        # شيت المطالبات الموزّعة فقط (للتأكد من الأرقام)
        row_level_export = distributed_df.copy()
        # ترتيب أعمدة مقروء
        front = [sales_col, client_key_col]
        if account_key_col and account_key_col not in front:
            front.append(account_key_col)
        for c in amount_cols:
            if c not in front:
                front.append(c)
        rest = [c for c in row_level_export.columns if c not in front and not str(c).startswith("_")]
        row_level_export = row_level_export[front + rest]

        result = {
            "filename": uploaded.name,
            "file_hash": uploaded_file_hash(uploaded),
            "resigned": resigned,
            "targets": selected_targets,
            "client_key_col": client_key_col,
            "account_key_col": account_key_col,
            "amount_cols": amount_cols,
            "summary": summary,
            "detail": detail,
            "row_level_export": row_level_export,
            "distributed_df": distributed_df,
            "full_df": full_df,
            "resigned_df": resigned_df,
            "warnings": warnings,
            "n_clients": len(grouped),
            "total_amount": float(grouped["amount"].sum()),
            "sales_col": sales_col,
            "unmapped_rows": unmapped,
        }
        st.session_state[DISTRIBUTION_RESULT_KEY] = result
        st.success("✅ تم التوزيع بنجاح!")
        st.rerun()

    # عرض النتائج لو موجودة لنفس الملف
    cached = st.session_state.get(DISTRIBUTION_RESULT_KEY)
    if cached and cached.get("file_hash") == uploaded_file_hash(uploaded):
        _render_distribution_results(cached)
    elif cached:
        st.warning("يوجد نتيجة توزيع قديمة لملف مختلف. اضغط «تنفيذ التوزيع» من جديد.")


def _render_distribution_results(result):
    """عرض ملخص ونتائج التوزيع + التحميل."""
    st.markdown("---")
    st.subheader("📊 نتائج التوزيع")

    resigned = result.get("resigned", "—")
    n_clients = result.get("n_clients", 0)
    total_amount = result.get("total_amount", 0)
    targets = result.get("targets", [])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("👤 المحصل المستقيل", resigned)
    k2.metric("👥 عدد العملاء الموزعين", f"{n_clients:,}")
    k3.metric("💰 إجمالي المبالغ", f"{total_amount:,.0f}")
    k4.metric("🎯 عدد المحصلين المستهدفين", f"{len(targets)}")

    summary = result.get("summary")
    # حساب وعرض الفروقات الفعلية (أهم مقياس لجودة التوزيع)
    if summary is not None and not summary.empty and len(summary) > 1:
        amt_spread = float(summary["إجمالي المبلغ"].max() - summary["إجمالي المبلغ"].min())
        cli_spread = int(summary["عدد العملاء"].max() - summary["عدد العملاء"].min())
        acc_spread = int(summary["عدد الحسابات"].max() - summary["عدد الحسابات"].min())
        st.markdown("#### 📏 الفروقات الفعلية بعد التوزيع (كلّما أصغر كلّما أفضل)")
        s1, s2, s3 = st.columns(3)
        s1.metric("فرق المبالغ", f"{amt_spread:,.0f}")
        s2.metric("فرق عدد العملاء", f"{cli_spread}")
        s3.metric("فرق عدد الحسابات", f"{acc_spread}")

    warnings = result.get("warnings") or []
    if warnings:
        for w in warnings:
            st.warning(f"⚠️ {w}")
    else:
        st.success("التوزيع ضمن حدود الـ Tolerance المحددة.")

    if summary is not None and not summary.empty:
        st.markdown("#### ملخص التوزيع حسب المحصّل")
        st.dataframe(summary, use_container_width=True, hide_index=True)

        # رسم بسيط
        if len(summary) > 0:
            fig = px.bar(
                summary,
                x="المحصّل",
                y=["عدد العملاء", "عدد الحسابات"],
                barmode="group",
                template=PLOTLY_TEMPLATE,
                color_discrete_sequence=OPS_SCALE,
            )
            _apply_ops_chart_style(fig, "توزيع العملاء والحسابات", height=380, xaxis_title="", yaxis_title="العدد")
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="dist_summary_chart")

            fig2 = px.bar(
                summary,
                x="المحصّل",
                y="إجمالي المبلغ",
                template=PLOTLY_TEMPLATE,
                color="إجمالي المبلغ",
                color_continuous_scale=OPS_SCALE,
            )
            _apply_ops_chart_style(fig2, "توزيع المبالغ", height=380, xaxis_title="", yaxis_title="المبلغ")
            st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG, key="dist_amount_chart")

    detail = result.get("detail")
    if detail is not None and not detail.empty:
        with st.expander("📋 تفاصيل إسناد كل عميل", expanded=False):
            st.dataframe(detail, use_container_width=True, hide_index=True)

    # تحميل
    full_df = result.get("full_df")
    if full_df is not None:
        st.markdown("#### ⬇️ تحميل النتائج")
        d1, d2 = st.columns(2)
        with d1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                if summary is not None:
                    summary.to_excel(writer, index=False, sheet_name="ملخص_التوزيع")
                row_level = result.get("row_level_export")
                if row_level is not None and not row_level.empty:
                    row_level.to_excel(writer, index=False, sheet_name="المطالبات_الموزعة_فقط")
                if detail is not None:
                    detail.to_excel(writer, index=False, sheet_name="إسناد_العملاء")
                full_df.to_excel(writer, index=False, sheet_name="المحفظة_الكاملة")
            st.download_button(
                "📥 تحميل المحفظة الكاملة بعد التوزيع",
                data=buf.getvalue(),
                file_name=f"محفظة_بعد_توزيع_{resigned}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
                key="dist_dl_full",
            )
        with d2:
            row_level = result.get("row_level_export")
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
                if row_level is not None and not getattr(row_level, "empty", True):
                    row_level.to_excel(writer, index=False, sheet_name="المطالبات_الموزعة_فقط")
                else:
                    pd.DataFrame({"ملاحظة": ["لا توجد مطالبات موزّعة"]}).to_excel(
                        writer, index=False, sheet_name="المطالبات_الموزعة_فقط"
                    )
            st.download_button(
                "📥 تحميل تقرير الإسناد فقط",
                data=buf2.getvalue(),
                file_name=f"تقرير_توزيع_{resigned}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dist_dl_report",
            )


PAGES = {
    "🎯 تصنيف المكالمات": page_classification,
    "📚 الوعود": page_promises,
    "⚠️ الإهمال والمتابعة": page_neglect,
    "📅 الجدولة المتعثرة": page_schedule_stalled,
    "🧾 أخطاء الحالات": page_case_errors,
    "⚖️ التوزيع": page_distribution,
    "📊 تحليل نشاط المحصّلين": page_dashboard,
}

DEFAULT_PAGE = next(iter(PAGES))
with st.sidebar:
    import base64
    from io import BytesIO
    _AHLY_LOGO_B64 = "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"
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.image(BytesIO(base64.b64decode(_AHLY_LOGO_B64)), width=110)
    st.markdown(
        "<h2 style='text-align:center;margin:0.25rem 0 0.1rem;color:#126D3C;font-weight:800'>الوطنية</h2>"
        "<p style='text-align:center;opacity:0.75;margin:0 0 0.5rem;font-size:0.9rem'>اختر القسم من القائمة</p>",
        unsafe_allow_html=True,
    )
    st.divider()
    selected_page = st.session_state.get("selected_page", DEFAULT_PAGE)
    for page_label in PAGES:
        if st.button(
            page_label,
            key=f"sidebar_page_{page_label}",
            use_container_width=True,
            type="primary" if selected_page == page_label else "secondary",
        ):
            st.session_state["selected_page"] = page_label
            st.rerun()

PAGES[selected_page]()
