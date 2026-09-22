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
    page_title="إيجادة — إدارة المحفظة ونشاط المحصلين",
    page_icon="💼",
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
        "accent_gold": "#C9A84C",
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
        # هوية إيجادة Strategic Insights — من واجهة رؤى الأداء
        "bg": "#F4F1E8",
        "bg_glow": "#EDE9DE",
        "surface": "#FFFFFF",
        "surface_2": "#F7F4EC",
        "surface_3": "#FBFAF6",
        "surface_hover": "#E8F0E8",
        "sidebar_bg": "#FFFFFF",
        "accent_surface": "#E6F0EA",
        "accent": "#1B5E45",
        "accent_strong": "#0D3D2E",
        "accent_gold": "#C9A84C",
        "on_accent": "#FFFFFF",
        "success": "#1B5E45",
        "danger": "#C45C5C",
        "warn": "#D4A017",
        "text": "#0D3D2E",
        "text_dim": "#3D5C4E",
        "text_muted": "#6B8578",
        "placeholder": "#8A9E94",
        "border": "rgba(13, 61, 46, 0.18)",
        "border_soft": "rgba(13, 61, 46, 0.10)",
        "input_bg": "#FFFFFF",
        "chart_marker": "#D5DFD4",
        "chart_text": "#0D3D2E",
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
    """حقن هوية إيجادة Strategic Insights: سكايلاين + ألوان النظام."""
    t = THEME
    is_light = THEME_NAME == "light"

    _SKYLINE_B64 = "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"
    skyline = f'url("data:image/png;base64,{_SKYLINE_B64}")'

    if is_light:
        bg_image = skyline
        bg_pos = "center bottom"
        bg_size = "100% auto"
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

/* ===== خلفية إيجادة: كريمي + سكايلاين ===== */
.stApp {{
  background-color: {t["bg"]} !important;
  background-image: {bg_image} !important;
  background-position: {bg_pos} !important;
  background-size: {bg_size} !important;
  background-repeat: {bg_repeat} !important;
  background-attachment: fixed !important;
  image-rendering: high-quality !important;
}}

/* شريط Streamlit الأصلي — شفاف */
header[data-testid="stHeader"] {{
  background: transparent !important;
}}
div[data-testid="stDecoration"] {{
  display: none !important;
}}

/* شريط إيجادة العلوي — بعرض الصفحة كامل مثل الصورة المرجعية */
.ejada-topbar {{
  background: linear-gradient(105deg, #06261C 0%, #0A3326 25%, #0D3D2E 50%, #134D3A 75%, #0A3326 100%);
  border-bottom: 3px solid rgba(201, 168, 76, 0.55);
  padding: 0.95rem 1.75rem 1.05rem;
  margin: 0 -2rem 0 -2rem;
  border-radius: 14px 14px 0 0;
  text-align: center;
  box-shadow: 0 6px 24px rgba(13, 61, 46, 0.28);
  position: relative;
  overflow: hidden;
}}
.ejada-topbar::before {{
  content: "";
  position: absolute;
  inset: 0;
  background:
    repeating-linear-gradient(
      135deg,
      transparent,
      transparent 14px,
      rgba(201, 168, 76, 0.045) 14px,
      rgba(201, 168, 76, 0.045) 28px
    ),
    radial-gradient(ellipse 60% 80% at 10% 50%, rgba(201,168,76,0.08), transparent 55%),
    radial-gradient(ellipse 50% 70% at 90% 50%, rgba(201,168,76,0.07), transparent 50%);
  pointer-events: none;
}}
.ejada-topbar-inner {{
  position: relative;
  z-index: 1;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  max-width: 1100px;
  margin: 0 auto;
}}
.ejada-logo-mark {{
  width: 52px;
  height: 52px;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
}}
.ejada-logo-mark svg {{
  width: 48px;
  height: 48px;
  filter: drop-shadow(0 2px 6px rgba(0,0,0,0.25));
}}
.ejada-topbar-title {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.12rem;
  flex: 1;
}}
.ejada-ar {{
  color: #F5E6C8;
  font-size: 1.28rem;
  font-weight: 800;
  letter-spacing: 0.03em;
  line-height: 1.35;
  text-shadow: 0 1px 2px rgba(0,0,0,0.2);
}}
.ejada-en {{
  color: #C9A84C;
  font-size: 1.02rem;
  font-weight: 700;
  letter-spacing: 0.06em;
  line-height: 1.3;
}}
.ejada-topbar-sub-out {{
  text-align: center;
  color: #0D3D2E;
  font-size: 1.15rem;
  font-weight: 800;
  letter-spacing: 0.02em;
  margin: 1.1rem 0 0.85rem 0;
  padding: 0 0.5rem;
}}
.ejada-user-chip {{
  display: inline-flex;
  align-items: center;
  gap: 0.4rem;
  background: #FFFFFF;
  border: 1px solid rgba(13,61,46,0.14);
  border-radius: 999px;
  padding: 0.35rem 0.85rem 0.35rem 0.55rem;
  font-size: 0.78rem;
  font-weight: 700;
  color: #0D3D2E;
  box-shadow: 0 2px 8px rgba(13,61,46,0.08);
  float: left;
  margin-bottom: 0.5rem;
}}
.ejada-user-chip .dot {{
  width: 22px;
  height: 22px;
  border-radius: 50%;
  background: linear-gradient(135deg, #1B5E45, #0D3D2E);
  color: #F5E6C8;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 0.7rem;
}}

/* ===== الحاوية الرئيسية (بطاقة بيضاء) ===== */
.block-container {{
  background: {t["surface"]} !important;
  border: 1px solid {t["border"]} !important;
  border-top: none !important;
  border-radius: 16px !important;
  box-shadow: 0 12px 40px rgba(13, 61, 46, 0.12) !important;
  padding-top: 0 !important;
  padding-bottom: 2rem !important;
  padding-left: 2rem !important;
  padding-right: 2rem !important;
  max-width: min(1480px, 94vw) !important;
  margin-left: auto !important;
  margin-right: auto !important;
  margin-top: 1rem !important;
  margin-bottom: 3.5rem !important;
  overflow: hidden !important;
}}

/* ===== الشريط الجانبي ===== */
section[data-testid="stSidebar"] {{
  background: {t["sidebar_bg"]} !important;
  border-left: 1px solid {t["border_soft"]} !important;
  box-shadow: -4px 0 24px rgba(13, 61, 46, 0.06) !important;
}}
section[data-testid="stSidebar"] .block-container {{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  margin-top: 0.5rem !important;
  padding-top: 0.75rem !important;
}}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
  color: {t["accent_strong"]} !important;
}}

/* ===== هيدر الصفحة ===== */
.wq-page-badge {{
  display: inline-block;
  background: linear-gradient(135deg, {t["accent"]}18, {t["accent"]}08);
  color: {t["accent_strong"]} !important;
  border: 1px solid {t["accent"]}40;
  border-radius: 999px;
  padding: 0.25rem 0.85rem;
  font-size: 0.72rem;
  font-weight: 700;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  margin-bottom: 0.5rem;
}}
.wq-page-header {{
  background: linear-gradient(135deg, {t["accent"]}10 0%, rgba(201,168,76,0.08) 50%, transparent 100%);
  border: 1px solid {t["border"]};
  border-right: 4px solid {t["accent_strong"]};
  border-radius: 14px;
  padding: 1rem 1.25rem 1.1rem 1.25rem;
  margin-bottom: 1.1rem;
}}
.wq-page-title {{
  color: {t["accent_strong"]} !important;
  font-weight: 800 !important;
  font-size: 1.85rem !important;
  margin: 0.35rem 0 0.3rem 0 !important;
}}
.wq-page-sub {{
  color: {t["text_muted"]} !important;
  font-size: 0.92rem !important;
  opacity: 0.95;
  margin: 0 !important;
}}

.block-container hr {{
  border: none !important;
  border-top: 2px solid {t["accent"]}28 !important;
  margin: 1rem 0 1.25rem 0 !important;
}}

/* ===== أزرار Primary ===== */
div.stButton > button[kind="primary"],
div.stButton > button[data-testid="baseButton-primary"],
button[kind="primary"] {{
  background: {t["accent_strong"]} !important;
  background-image: linear-gradient(180deg, {t["accent"]} 0%, {t["accent_strong"]} 100%) !important;
  color: {t["on_accent"]} !important;
  border: none !important;
  border-radius: 999px !important;
  font-weight: 700 !important;
  box-shadow: 0 2px 10px rgba(13, 61, 46, 0.28) !important;
}}
div.stButton > button[kind="primary"]:hover,
button[kind="primary"]:hover {{
  filter: brightness(1.08);
  box-shadow: 0 4px 16px rgba(13, 61, 46, 0.35) !important;
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

/* تبويبات */
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

/* بطاقات */
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

h1, h2, h3 {{
  color: {t["text"]} !important;
  font-weight: 800 !important;
}}

.stProgress > div > div > div > div {{
  background-color: {t["accent_strong"]} !important;
}}

div.stDownloadButton > button {{
  border-radius: 999px !important;
  font-weight: 700 !important;
}}

/* ===== تنقل الشريط الجانبي — مثل الصورة المرجعية ===== */
section[data-testid="stSidebar"] div.stButton > button {{
  border-radius: 12px !important;
  font-weight: 700 !important;
  font-size: 0.92rem !important;
  padding: 0.65rem 0.9rem !important;
  margin-bottom: 0.35rem !important;
  justify-content: flex-start !important;
  text-align: right !important;
  transition: all 0.18s ease !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="secondary"] {{
  background: transparent !important;
  color: {t["text_dim"]} !important;
  border: 1px solid transparent !important;
  box-shadow: none !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {{
  background: {t["accent_surface"]} !important;
  color: {t["accent_strong"]} !important;
  border-color: {t["border_soft"]} !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="primary"] {{
  background: linear-gradient(135deg, {t["accent"]} 0%, {t["accent_strong"]} 100%) !important;
  color: {t["on_accent"]} !important;
  border: none !important;
  box-shadow: 0 3px 12px rgba(13, 61, 46, 0.22) !important;
}}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {{
  gap: 0.15rem !important;
}}

/* ===== هوية السايدبار — مطابقة للصورة المرجعية ===== */
.ejada-side-brand {{
  text-align: center;
  padding: 0.85rem 0.4rem 1rem;
  margin-bottom: 0.6rem;
  border-bottom: 1px solid rgba(13,61,46,0.10);
}}
.ejada-side-logo {{
  display: flex;
  justify-content: center;
  margin-bottom: 0.4rem;
}}
.ejada-side-name {{
  color: #0D3D2E;
  font-weight: 800;
  font-size: 1.1rem;
  letter-spacing: 0.03em;
  line-height: 1.3;
}}
.ejada-side-tag {{
  color: #6B8578;
  font-size: 0.68rem;
  font-weight: 600;
  margin-top: 0.15rem;
  letter-spacing: 0.02em;
}}
.ejada-side-footer {{
  margin-top: 1.75rem;
  padding-top: 0.85rem;
  border-top: 1px solid rgba(13,61,46,0.10);
  text-align: center;
  color: #8A9E94;
  font-size: 0.66rem;
  font-weight: 500;
  line-height: 1.45;
}}
section[data-testid="stSidebar"] {{
  background: #FAFBF9 !important;
  border-left: none !important;
  border-right: 1px solid rgba(13,61,46,0.08) !important;
  box-shadow: 2px 0 18px rgba(13, 61, 46, 0.05) !important;
}}
section[data-testid="stSidebar"] div.stButton > button {{
  border-radius: 14px !important;
  font-weight: 700 !important;
  font-size: 0.88rem !important;
  padding: 0.72rem 0.85rem !important;
  margin-bottom: 0.28rem !important;
  justify-content: flex-start !important;
  text-align: right !important;
  direction: rtl !important;
  transition: all 0.18s ease !important;
  min-height: 2.65rem !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="secondary"] {{
  background: transparent !important;
  color: #3D5C4E !important;
  border: 1px solid transparent !important;
  box-shadow: none !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {{
  background: rgba(27, 94, 69, 0.08) !important;
  color: #0D3D2E !important;
  border-color: rgba(13,61,46,0.08) !important;
}}
section[data-testid="stSidebar"] div.stButton > button[kind="primary"] {{
  background: linear-gradient(135deg, #1B5E45 0%, #0D3D2E 100%) !important;
  color: #FFFFFF !important;
  border: none !important;
  box-shadow: 0 4px 14px rgba(13, 61, 46, 0.25) !important;
}}

</style>
""",

        unsafe_allow_html=True,
    )

    # شريط علوي بهوية إيجادة — بعرض كامل مثل الصورة المرجعية
    st.markdown(
        """
<div class="ejada-topbar">
  <div class="ejada-topbar-inner">
    <div class="ejada-logo-mark"><svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <defs>
    <linearGradient id="gGold" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#F5E6C8"/>
      <stop offset="45%" stop-color="#C9A84C"/>
      <stop offset="100%" stop-color="#A8842E"/>
    </linearGradient>
  </defs>
  <circle cx="32" cy="32" r="30" fill="none" stroke="url(#gGold)" stroke-width="1.5" opacity="0.7"/>
  <path d="M32 10c6 8 14 12 18 14-4 2-10 8-12 18-1-8-6-14-12-16 6-2 8-8 6-16z" fill="url(#gGold)" opacity="0.95"/>
  <path d="M22 38c4-2 8-1 12 2 2 6-2 12-6 14-6-4-10-10-6-16z" fill="url(#gGold)" opacity="0.75"/>
  <ellipse cx="32" cy="48" rx="10" ry="3" fill="url(#gGold)" opacity="0.55"/>
</svg></div>
    <div class="ejada-topbar-title">
      <span class="ejada-ar">رؤى الأداء الاستراتيجي - إيجادة</span>
      <span class="ejada-en">EJADA Strategic Insights</span>
    </div>
    <div class="ejada-logo-mark"><svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <defs>
    <linearGradient id="gGold" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#F5E6C8"/>
      <stop offset="45%" stop-color="#C9A84C"/>
      <stop offset="100%" stop-color="#A8842E"/>
    </linearGradient>
  </defs>
  <circle cx="32" cy="32" r="30" fill="none" stroke="url(#gGold)" stroke-width="1.5" opacity="0.7"/>
  <path d="M32 10c6 8 14 12 18 14-4 2-10 8-12 18-1-8-6-14-12-16 6-2 8-8 6-16z" fill="url(#gGold)" opacity="0.95"/>
  <path d="M22 38c4-2 8-1 12 2 2 6-2 12-6 14-6-4-10-10-6-16z" fill="url(#gGold)" opacity="0.75"/>
  <ellipse cx="32" cy="48" rx="10" ry="3" fill="url(#gGold)" opacity="0.55"/>
</svg></div>
  </div>
</div>
<div class="ejada-topbar-sub-out">نظام تحليل وادارة متقدم</div>
<div style="clear:both;text-align:left;direction:ltr;">
  <span class="ejada-user-chip"><span class="dot">👤</span> User Profile</span>
</div>
<div style="clear:both;"></div>
""",
        unsafe_allow_html=True,
    )


_inject_wataniya_identity_css()

# Streamlit native theme is the source of truth for the app UI.
# Plotly receives the matching palette below; custom CSS applies Wataniya identity.

def page_header(eyebrow: str, title: str, subtitle: str, centered: bool = False):
    """هيدر مميز بهوية الوطنية + خلفية خفيفة."""
    align = "center" if centered else "right"
    badge = ""
    if eyebrow:
        badge = f"<div class='wq-page-badge'>{eyebrow.upper()}</div>"
    st.markdown(
        f"""
<div class="wq-page-header" style="text-align:{align};">
  {badge}
  <h1 class="wq-page-title" style="text-align:{align};">{title}</h1>
  <p class="wq-page-sub" style="text-align:{align};">{subtitle}</p>
</div>
""",
        unsafe_allow_html=True,
    )


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
    "🌐 التحليل الاستراتيجي": page_dashboard,
    "🎯 تصنيف المكالمات": page_classification,
    "🎙️ تقييم الجودة AI": page_classification,
    "📚 الوعود": page_promises,
    "⚠️ الإهمال والمتابعة": page_neglect,
    "📅 الجدولة المتعثرة": page_schedule_stalled,
    "🧾 أخطاء الحالات": page_case_errors,
    "⚖️ ادارة الموارد": page_distribution,
    "📊 تقارير الأداء": page_dashboard,
}

# ربط أسماء العرض بنفس دوال الصفحات (بدون تكرار غير مقصود في الواجهة)
# ملاحظة: "تقييم الجودة AI" يفتح نفس صفحة التصنيف، و"التحليل الاستراتيجي/تقارير الأداء" يفتحان الداشبورد.
_PAGE_ALIASES = {
    "🌐 التحليل الاستراتيجي": "📊 تقارير الأداء",
}
# إزالة التكرار من قائمة الأزرار المعروضة مع الإبقاء على كل الوظائف
_SIDEBAR_ITEMS = [
    ("🎯", "تصنيف المكالمات", "🎯 تصنيف المكالمات"),
    ("📚", "الوعود", "📚 الوعود"),
    ("⚠️", "الإهمال والمتابعة", "⚠️ الإهمال والمتابعة"),
    ("📅", "الجدولة المتعثرة", "📅 الجدولة المتعثرة"),
    ("🧾", "أخطاء الحالات", "🧾 أخطاء الحالات"),
    ("⚖️", "ادارة الموارد", "⚖️ ادارة الموارد"),
    ("📊", "تقارير الأداء", "📊 تقارير الأداء"),
    ("🌐", "التحليل الاستراتيجي", "🌐 التحليل الاستراتيجي"),
    ("🎙️", "تقييم الجودة AI", "🎙️ تقييم الجودة AI"),
]

DEFAULT_PAGE = "🎯 تصنيف المكالمات"
with st.sidebar:
    # شعار إيجادة أعلى السايدبار — بنفس روح الصورة المرجعية
    st.markdown(
        """
<div class="ejada-side-brand">
  <div class="ejada-side-logo">
    <svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" width="36" height="36">
      <defs>
        <linearGradient id="sg" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#1B5E45"/>
          <stop offset="100%" stop-color="#0D3D2E"/>
        </linearGradient>
      </defs>
      <rect x="2" y="2" width="44" height="44" rx="12" fill="url(#sg)"/>
      <path d="M24 10c4 5 9 8 12 9-3 1-7 5-8 12-1-5-4-9-8-10 4-1 5-5 4-11z" fill="#C9A84C"/>
      <ellipse cx="24" cy="34" rx="8" ry="2.5" fill="#C9A84C" opacity="0.7"/>
    </svg>
  </div>
  <div class="ejada-side-name">إيجادة</div>
  <div class="ejada-side-tag">EJADA Strategic Insights</div>
</div>
""",
        unsafe_allow_html=True,
    )

    selected_page = st.session_state.get("selected_page", DEFAULT_PAGE)
    # ترتيب العرض: الأقرب للصورة أولاً ثم بقية الوظائف
    display_order = [
        ("🌐", "التحليل الاستراتيجي", "🌐 التحليل الاستراتيجي"),
        ("🎙️", "تقييم الجودة AI", "🎙️ تقييم الجودة AI"),
        ("🎯", "تصنيف المكالمات", "🎯 تصنيف المكالمات"),
        ("📚", "الوعود", "📚 الوعود"),
        ("⚠️", "الإهمال والمتابعة", "⚠️ الإهمال والمتابعة"),
        ("📅", "الجدولة المتعثرة", "📅 الجدولة المتعثرة"),
        ("🧾", "أخطاء الحالات", "🧾 أخطاء الحالات"),
        ("⚖️", "ادارة الموارد", "⚖️ ادارة الموارد"),
        ("📊", "تقارير الأداء", "📊 تقارير الأداء"),
    ]

    for icon, label, page_key in display_order:
        is_active = selected_page == page_key
        btn_label = f"{icon}  {label}"
        if st.button(
            btn_label,
            key=f"sidebar_page_{page_key}",
            use_container_width=True,
            type="primary" if is_active else "secondary",
        ):
            st.session_state["selected_page"] = page_key
            st.rerun()

    st.markdown(
        """
<div class="ejada-side-footer">
  2024 - تم التصميم بواسطة<br/>Ejada Strategic Consulting
</div>
""",
        unsafe_allow_html=True,
    )

# تطبيع المفتاح لو كان من أسماء بديلة
_resolved = st.session_state.get("selected_page", DEFAULT_PAGE)
if _resolved not in PAGES:
    _resolved = DEFAULT_PAGE
    st.session_state["selected_page"] = _resolved
PAGES[_resolved]()
