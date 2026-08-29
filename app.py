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
  - تحليل نشاط المحصّلين (Dashboard + تصدير HTML)

أسماء الأعمدة قابلة للتعديل من قسم الإعدادات أعلى الملف.
"""

import io
import hashlib
import re
from datetime import time as dt_time
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification

# ==========================================================
# إعدادات وأسماء الأعمدة
# ==========================================================

MODEL_REPO = "Mahmoud252002/7oudaModel"
MAX_LENGTH = 256

ORIGINAL_TEXT_COL = "Notes"         # اسم العمود الأصلي في الملف
MODEL_TEXT_COL = "الافادة"          # الاسم اللي بيتحول له مؤقتًا لـ الموديل
CLASSIFICATION_COL = "التصنيف"      # عمود النتيجة: 1 = ناجحة / 0 = غير ناجحة
WASTED_TIME_COL = "الوقت_المهدر_دقيقة"

ID_CANDIDATES = ["Account ID", "account id", "AccountID", "ID", "id", "رقم الحساب", "الرقم التعريفي", "Account No", "account no", "Account Number"]
SALES_PERSON_CANDIDATES = ["Create By", "create by", "CreateBy", "Created By", "created by", "Sales Person", "sales person", "المحصّل", "Salesperson", "salesperson", "SalesPerson"]
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

NEGLECT_LAST_DATE_CANDIDATES = ["Follow up Last Date", "follow up last date", "Last Follow Up", "تاريخ آخر متابعة"]
NEGLECT_RESULT_KEY = "neglect_result"
APP_DATA_CACHE_KEY = "app_uploaded_data_cache"
DASHBOARD_SOURCE_HASH_KEY = "dashboard_source_hash"


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
    _file_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
    cached = st.session_state.get(result_key)
    if cached and cached.get("file_hash") == _file_hash:
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

    # 2) فلترة Salesperson — نستبعد المحصّلين المحددين
    sales_vals = df[sales_col].astype(str).str.strip()
    keep_sales = ~sales_vals.isin(PROMISE_EXCLUDED_SALES)
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
            count_fig = px.bar(
                agent_summary.head(15).sort_values("عدد الوعود"),
                x="عدد الوعود",
                y=sales_col,
                orientation="h",
                text="عدد الوعود",
                color="عدد الوعود",
                color_continuous_scale=[THEME["surface_2"], COLOR_ACCENT],
                template=PLOTLY_TEMPLATE,
            )
            count_fig.update_layout(**PLOTLY_LAYOUT, title="عدد الوعود حسب المحصّل", xaxis_title="عدد الوعود", yaxis_title="", coloraxis_showscale=False, height=430)
            count_fig.update_traces(customdata=agent_summary.head(15).sort_values("عدد الوعود")[sales_col], hovertemplate="<b>%{y}</b><br>عدد الوعود: %{x:,}<extra></extra>")
            render_selectable_chart(count_fig, f"promises_count_{mode_label}", filter_key=PROMISES_AGENT_FILTER_KEY)
        with right:
            if net_col and "إجمالي المديونية" in agent_summary.columns:
                amount_fig = px.bar(
                    agent_summary.head(15).sort_values("إجمالي المديونية"),
                    x="إجمالي المديونية",
                    y=sales_col,
                    orientation="h",
                    text="إجمالي المديونية",
                    color="إجمالي المديونية",
                    color_continuous_scale=[COLOR_ACCENT, COLOR_WARN],
                    template=PLOTLY_TEMPLATE,
                )
                amount_fig.update_layout(**PLOTLY_LAYOUT, title="إجمالي المديونية حسب المحصّل", xaxis_title="إجمالي المديونية", yaxis_title="", coloraxis_showscale=False, height=430)
                amount_fig.update_traces(customdata=agent_summary.head(15).sort_values("إجمالي المديونية")[sales_col], hovertemplate="<b>%{y}</b><br>إجمالي المديونية: %{x:,.0f}<extra></extra>")
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
        ("📗<br>الوعود القائمة", standing_count, {"valueformat": ",d"}, COLOR_SUCCESS),
        ("📕<br>الوعود المكسورة", broken_count, {"valueformat": ",d"}, COLOR_FAIL),
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("💰<br>إجمالي المديونية", total_amount, {"valueformat": ",.0f"}, COLOR_WARN),
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
                color_discrete_map={"الوعود القائمة": COLOR_SUCCESS, "الوعود المكسورة": COLOR_FAIL},
                template=PLOTLY_TEMPLATE,
            )
            fig.update_layout(**{
                **PLOTLY_LAYOUT,
                "title": "القائمة والمكسورة حسب المحصّل",
                "xaxis_title": "عدد الوعود",
                "yaxis_title": "",
                "height": 500,
                "legend_title_text": "",
                "margin": dict(t=78, b=62, l=170, r=70),
                "xaxis": dict(tickformat=",.0f", automargin=True),
                "uniformtext_minsize": 12,
                "uniformtext_mode": "hide",
            })
            fig.update_traces(
                texttemplate="%{x:,.0f}",
                textposition="outside",
                textfont=dict(size=15, color=THEME["text"]),
                cliponaxis=False,
                customdata=agent_type[sales_col],
                hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f} وعد<extra></extra>",
            )
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
                    color_discrete_map={"الوعود القائمة": COLOR_SUCCESS, "الوعود المكسورة": COLOR_FAIL},
                    template=PLOTLY_TEMPLATE,
                )
                fig.update_layout(**{
                    **PLOTLY_LAYOUT,
                    "title": "إجمالي المديونية حسب المحصّل",
                    "xaxis_title": "إجمالي المديونية",
                    "yaxis_title": "",
                    "height": 500,
                    "legend_title_text": "",
                    "margin": dict(t=78, b=62, l=170, r=105),
                    "xaxis": dict(tickformat=",.0f", separatethousands=True, automargin=True),
                    "uniformtext_minsize": 11,
                    "uniformtext_mode": "hide",
                })
                fig.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                    textfont=dict(size=14, color=THEME["text"]),
                    cliponaxis=False,
                    customdata=agent_amount[sales_col],
                    hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f} جنيه<extra></extra>",
                )
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
            color_discrete_map={"الوعود القائمة": COLOR_SUCCESS, "الوعود المكسورة": COLOR_FAIL},
            template=PLOTLY_TEMPLATE,
        )
        fig.update_layout(**{
            **PLOTLY_LAYOUT,
            "title": "توزيع الوعود القائمة والمكسورة",
            "height": 420,
            "legend_title_text": "",
            "margin": dict(t=78, b=45, l=35, r=35),
        })
        fig.update_traces(
            texttemplate="%{label}<br>%{value:,.0f} (%{percent:.1%})",
            textfont=dict(size=16, color=THEME["text"]),
            textinfo="text",
            hovertemplate="<b>%{label}</b><br>عدد الوعود: %{value:,.0f}<br>النسبة: %{percent:.1%}<extra></extra>",
        )
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
    else:
        cached_file = st.session_state.get(APP_DATA_CACHE_KEY, {}).get(cache_scope, {})
        cached_result = st.session_state.get(standing_key) or st.session_state.get(broken_key)
        if cached_result or cached_file:
            saved_name = (cached_result or cached_file).get("filename", "المحفظة المحفوظة")
            st.success(f"✅ محفظة {company_label} محفوظة: {saved_name}. لن تُحذف عند التنقل بين التبويبات.")

    combined, meta = _combine_promises_cached_results(company_label, result_keys=result_keys)
    if combined.empty or meta is None:
        st.info(f"📂 ارفع محفظة {company_label} لعرض الوعود القائمة والمكسورة معًا.")
        return
    render_combined_promises_dashboard(combined, meta, company_label)


st.set_page_config(
    page_title="لوحة تحليل المكالمات",
    page_icon="🎙️",
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
    "dark": {
        "bg": "#0E1420",
        "bg_glow": "#10192C",
        "surface": "#151F30",
        "surface_2": "#1B2A42",
        "surface_3": "#111B2C",
        "surface_hover": "#1D304C",
        "sidebar_bg": "#0B111C",
        "accent_surface": "#122B2A",
        "accent": "#4F8C8D",
        "accent_strong": "#3D7375",
        "on_accent": "#FFFFFF",
        "success": "#4F927D",
        "danger": "#B96578",
        "warn": "#B39A5A",
        "text": "#F3F6FA",
        "text_dim": "#B9C6D6",
        "text_muted": "#9FB0C6",
        "placeholder": "#8FA3B8",
        "border": "rgba(15, 157, 138, 0.20)",
        "border_soft": "rgba(148, 163, 184, 0.18)",
        "input_bg": "#1B2A42",
        "chart_marker": "#0E1420",
        "chart_text": "#FFFFFF",
        "danger_soft": "#3B1420", "danger_text": "#FECDD3",
        "warn_soft": "#3A2A08", "warn_text": "#FDE68A",
    },
    "light": {
        "bg": "#F4F7FB",
        "bg_glow": "#EAF1F8",
        "surface": "#FFFFFF",
        "surface_2": "#EEF4F8",
        "surface_3": "#F7FAFC",
        "surface_hover": "#E2ECF3",
        "sidebar_bg": "#FFFFFF",
        "accent_surface": "#E6F7F4",
        "accent": "#397B7D",
        "accent_strong": "#2F6668",
        "on_accent": "#FFFFFF",
        "success": "#397D69",
        "danger": "#A65367",
        "warn": "#977D42",
        "text": "#172033",
        "text_dim": "#526174",
        "text_muted": "#6B7A8C",
        "placeholder": "#718096",
        "border": "rgba(8, 127, 112, 0.22)",
        "border_soft": "rgba(71, 85, 105, 0.18)",
        "input_bg": "#FFFFFF",
        "chart_marker": "#CBD5E1",
        "chart_text": "#172033",
        "danger_soft": "#FDE2E7", "danger_text": "#7F1D35",
        "warn_soft": "#FEF3C7", "warn_text": "#78350F",
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


# Streamlit native theme is the source of truth for the app UI.
# Plotly receives the matching palette below; no custom CSS is injected.

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
# Palette النشاط: ثلاث عائلات لونية هادئة فقط بدرجات متقاربة.
ACTIVITY_AGENT_PALETTE = [
    "#2F6F73", "#477F82", "#628B8E", "#7D9A9D", "#98AEB2", "#B2C1C3",
    "#5F7D8C", "#8095A2", "#A6B4B9",
]
ACTIVITY_STATE_PALETTE = ["#2F6F73", "#628B8E", "#8095A2", "#A6B4B9"]
ACTIVITY_OUTCOME_COLORS = {"ناجحة": "#2F6F73", "غير ناجحة": "#8095A2"}


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


CLASSIFICATION_AGENT_FILTER_KEY = "classification_selected_agent"
PROMISES_AGENT_FILTER_KEY = "promises_selected_agent"


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


DASHBOARD_AGENT_FILTER_KEY = "dashboard_selected_agent"

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
    if "لايرد" in key:
        return "لا يرد"
    if key == "مغلق" or ("مغلق" in key and "جدوله" not in key):
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

    if time_col and time_col in work.columns and WASTED_TIME_COL not in work.columns:
        work[WASTED_TIME_COL] = _calculate_wasted_time_by_day(
            work, "_agent_display", time_col, break_start, break_end
        )
    if WASTED_TIME_COL in work.columns:
        work[WASTED_TIME_COL] = pd.to_numeric(work[WASTED_TIME_COL], errors="coerce").fillna(0)

    agent = work.groupby("_agent_display", dropna=False).size().rename("إجمالي المكالمات").to_frame()
    agent["المكالمات الناجحة"] = work[work["_success_bool"]].groupby("_agent_display").size()
    agent["المكالمات الناجحة"] = agent["المكالمات الناجحة"].fillna(0).astype(int)
    agent["المكالمات غير الناجحة"] = agent["إجمالي المكالمات"] - agent["المكالمات الناجحة"]
    agent["نسبة النجاح (%)"] = (
        agent["المكالمات الناجحة"] / agent["إجمالي المكالمات"].replace(0, pd.NA) * 100
    ).fillna(0).round(1)

    state_columns = ACTIVITY_NO_ANSWER_STATES + ACTIVITY_PROMISE_STATES + ACTIVITY_PAYMENT_STATES
    state_table = pd.crosstab(work["_agent_display"], work["_activity_state"])
    for state in state_columns:
        if state not in state_table.columns:
            state_table[state] = 0
    state_table = state_table.reindex(columns=state_columns, fill_value=0)
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
    return {**PLOTLY_LAYOUT, **overrides}


def render_activity_kpi_cards(total, success, agent_count, success_rate, wasted_minutes):
    cards = [
        ("👥<br>عدد المحصّلين", agent_count, {"valueformat": ",d"}, THEME["text"]),
        ("📞<br>إجمالي المكالمات", total, {"valueformat": ",d"}, THEME["text"]),
        ("✅<br>المكالمات الناجحة", success, {"valueformat": ",d"}, ACTIVITY_OUTCOME_COLORS["ناجحة"]),
        ("📈<br>نسبة النجاح", success_rate, {"valueformat": ".1f", "suffix": "%"}, ACTIVITY_AGENT_PALETTE[1]),
        ("⏱️<br>إجمالي الوقت المهدر", wasted_minutes, {"valueformat": ".1f", "suffix": " دقيقة"}, ACTIVITY_AGENT_PALETTE[3]),
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


def _render_dashboard_agent_filter_notice():
    selected = st.session_state.get(DASHBOARD_AGENT_FILTER_KEY)
    if not selected:
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        st.info(f"🎯 الفلتر التفاعلي النشط: كل المؤشرات للمحصّل «{selected}»")
    with c2:
        if st.button("إظهار الكل", key="clear_dashboard_agent_filter", use_container_width=True):
            st.session_state.pop(DASHBOARD_AGENT_FILTER_KEY, None)
            st.rerun()


def _dashboard_activity_view(df, sales_col):
    selected = st.session_state.get(DASHBOARD_AGENT_FILTER_KEY)
    if not selected or not sales_col or sales_col not in df.columns:
        return df
    mask = df[sales_col].fillna("غير محدد").astype(str).str.strip().eq(str(selected).strip())
    if not mask.any():
        st.session_state.pop(DASHBOARD_AGENT_FILTER_KEY, None)
        return df
    return df.loc[mask].copy()


def _render_activity_daily_chart(work, time_col, class_col=None):
    if not time_col or time_col not in work.columns:
        st.info("يلزم وجود عمود Created On لعرض النشاط على مدار الأيام.")
        return
    trend = work.copy()
    trend["_activity_time"] = pd.to_datetime(trend[time_col], errors="coerce")
    trend = trend.dropna(subset=["_activity_time"])
    if trend.empty:
        st.info("لا توجد تواريخ صالحة لعرض النشاط اليومي.")
        return
    trend["اليوم"] = trend["_activity_time"].dt.strftime("%Y-%m-%d")
    trend["_success_for_day"] = _activity_success_mask(trend, class_col)
    daily = trend.groupby(["اليوم", "_agent_display"], as_index=False).agg(
        **{"عدد المكالمات": ("_agent_display", "size"), "المكالمات الناجحة": ("_success_for_day", "sum")}
    )
    daily_totals = daily.groupby("اليوم", as_index=False).agg(
        **{"إجمالي المكالمات": ("عدد المكالمات", "sum"), "إجمالي الناجحة": ("المكالمات الناجحة", "sum")}
    )
    daily_totals["نسبة النجاح (%)"] = (
        daily_totals["إجمالي الناجحة"] / daily_totals["إجمالي المكالمات"].replace(0, pd.NA) * 100
    ).fillna(0).round(1)
    sort_options = {
        "التاريخ تصاعديًا": "date",
        "إجمالي المكالمات تنازليًا": "calls",
        "نسبة النجاح تنازليًا": "success_rate",
    }
    sort_label = st.selectbox(
        "ترتيب الـ Histogram اليومي",
        list(sort_options.keys()),
        key="activity_daily_sort_v1",
        help="الترتيب يغيّر ترتيب الأيام على المحور الأفقي فقط.",
    )
    sort_mode = sort_options[sort_label]
    if sort_mode == "calls":
        ordered_days = daily_totals.sort_values(["إجمالي المكالمات", "اليوم"], ascending=[False, True])["اليوم"].tolist()
    elif sort_mode == "success_rate":
        ordered_days = daily_totals.sort_values(["نسبة النجاح (%)", "إجمالي المكالمات", "اليوم"], ascending=[False, False, True])["اليوم"].tolist()
    else:
        ordered_days = sorted(daily_totals["اليوم"].tolist())
    daily["اليوم"] = pd.Categorical(daily["اليوم"], categories=ordered_days, ordered=True)
    daily = daily.sort_values(["اليوم", "_agent_display"])
    daily_totals["اليوم"] = pd.Categorical(daily_totals["اليوم"], categories=ordered_days, ordered=True)
    daily_totals = daily_totals.sort_values("اليوم")

    fig = px.bar(
        daily, x="اليوم", y="عدد المكالمات", color="_agent_display", barmode="group",
        text_auto=True, custom_data=["_agent_display"], template=PLOTLY_TEMPLATE,
        labels={"_agent_display": "المحصّل"}, color_discrete_sequence=ACTIVITY_AGENT_PALETTE,
    )
    fig.add_trace(go.Scatter(
        x=daily_totals["اليوم"].astype(str), y=daily_totals["نسبة النجاح (%)"],
        name="نسبة النجاح", mode="lines+markers+text", text=daily_totals["نسبة النجاح (%)"].map(lambda value: f"{value:.1f}%"),
        textposition="top center", line={"color": ACTIVITY_AGENT_PALETTE[0], "width": 3},
        marker={"color": ACTIVITY_AGENT_PALETTE[0], "size": 9, "line": {"color": THEME["surface"], "width": 2}},
        yaxis="y2", customdata=[[""] for _ in range(len(daily_totals))],
        hovertemplate="<b>%{x}</b><br>نسبة النجاح: %{y:.1f}%<extra></extra>",
    ))
    fig.update_layout(**_activity_layout(
        title="📊 Combo Chart يومي: المكالمات ونسبة النجاح", title_x=0.5,
        xaxis_title="اليوم", yaxis_title="عدد المكالمات", height=400, bargap=0.14,
        legend_title_text="", hovermode="x unified",
        legend={"orientation": "h", "yanchor": "top", "y": -0.16, "x": 0.5, "xanchor": "center"},
        margin={"t": 62, "b": 78, "l": 50, "r": 55},
        xaxis={"type": "category", "categoryorder": "array", "categoryarray": ordered_days, "tickangle": -25},
        yaxis={"title": "عدد المكالمات", "rangemode": "tozero"},
        yaxis2={"title": "نسبة النجاح (%)", "overlaying": "y", "side": "right", "range": [0, 100], "ticksuffix": "%", "showgrid": False},
    ))
    fig.update_traces(selector={"type": "bar"}, marker_line_width=0, hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>")
    render_selectable_chart(fig, "dashboard_activity_daily", filter_key=DASHBOARD_AGENT_FILTER_KEY)


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
    hourly = trend.groupby(["الساعة", "_agent_display"], as_index=False).size().rename(columns={"size": "عدد المكالمات"})
    fig = px.bar(
        hourly, x="الساعة", y="عدد المكالمات", color="_agent_display", barmode="stack",
        text_auto=True, custom_data=["_agent_display"], template=PLOTLY_TEMPLATE,
        labels={"_agent_display": "المحصّل"}, color_discrete_sequence=ACTIVITY_AGENT_PALETTE,
    )
    fig.update_layout(**_activity_layout(
        title="🕒 Histogram ساعي لنشاط المحصلين", title_x=0.5, xaxis_title="ساعة اليوم", yaxis_title="عدد المكالمات",
        xaxis={"dtick": 1, "tickvals": list(range(hour_min, hour_max + 1)), "range": [max(-0.5, hour_min - 0.5), min(23.5, hour_max + 0.5)]}, height=400, bargap=0.06, legend_title_text="",
        legend={"orientation": "h", "yanchor": "top", "y": -0.16, "x": 0.5, "xanchor": "center"},
        margin={"t": 62, "b": 78, "l": 50, "r": 16},
    ))
    fig.update_traces(
        marker_line_width=0,
        customdata=trend["_agent_display"],
        hovertemplate="<b>الساعة %{x}:00</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>",
    )
    render_selectable_chart(fig, "dashboard_activity_hourly", filter_key=DASHBOARD_AGENT_FILTER_KEY)


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
    fig.update_traces(
        textinfo="percent", textfont_size=15,
        marker={"line": {"color": THEME["surface"], "width": 3}},
        hovertemplate="<b>%{label}</b><br>العدد: %{value:,}<br>النسبة: %{percent}<extra></extra>",
    )
    fig.update_layout(**_activity_layout(
        title="🎯 الناجحة مقابل غير الناجحة", title_x=0.5, height=400,
        legend={"orientation": "h", "yanchor": "top", "y": -0.12, "x": 0.5, "xanchor": "center"},
        margin={"t": 62, "b": 62, "l": 16, "r": 16},
        annotations=[{"text": f"{rate:.1f}%<br>نجاح", "x": 0.5, "y": 0.5, "font": {"size": 22, "color": COLOR_SUCCESS}, "showarrow": False}],
    ))
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key="dashboard_outcome_donut")


def _render_activity_no_answer_chart(agent):
    available = [state for state in ACTIVITY_NO_ANSWER_STATES if state in agent.columns]
    if not available:
        st.info("يلزم وجود عمود Sub State لعرض حالات لا يرد ومغلق.")
        return
    plot = agent[["المحصّل"] + available].copy()
    plot["إجمالي لا يرد"] = plot[available].sum(axis=1)
    plot = plot.sort_values("إجمالي لا يرد", ascending=True)
    long = plot.melt(id_vars=["المحصّل"], value_vars=available, var_name="الحالة", value_name="العدد")
    fig = px.bar(
        long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack", text_auto=True,
        template=PLOTLY_TEMPLATE, category_orders={"الحالة": ACTIVITY_NO_ANSWER_STATES},
        color_discrete_sequence=ACTIVITY_STATE_PALETTE,
    )
    fig.update_layout(**_activity_layout(
        title="📵 حالات لا يرد لكل محصل (تشمل مغلق والتكرار)", title_x=0.5, xaxis_title="عدد الحالات", yaxis_title="",
        height=400, legend_title_text="", legend={"orientation": "h", "yanchor": "top", "y": -0.16, "x": 0.5, "xanchor": "center"},
        margin={"t": 62, "b": 78, "l": 100, "r": 16}, yaxis={"categoryorder": "total ascending"},
    ))
    fig.update_traces(
        marker_line_width=0,
        customdata=long["المحصّل"],
        hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>",
    )
    render_selectable_chart(fig, "dashboard_no_answer_states", filter_key=DASHBOARD_AGENT_FILTER_KEY)


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
    view = _dashboard_activity_view(df, sales_col)
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
    st.caption("اضغط على اسم أي محصل داخل الرسوم التفاعلية لتطبيق فلتر موحد على الكروت والرسوم والجدول.")

    daily_col, hourly_col = st.columns(2)
    with daily_col:
        with st.container(border=True):
            _render_activity_daily_chart(work, time_col, class_col)
    with hourly_col:
        with st.container(border=True):
            _render_activity_hourly_chart(work, time_col)

    outcome_col, no_answer_col = st.columns(2)
    with outcome_col:
        with st.container(border=True):
            _render_activity_outcome_donut(work, class_col)
    with no_answer_col:
        with st.container(border=True):
            _render_activity_no_answer_chart(agent)

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


def build_dashboard_html(df, class_col, sales_col, time_col, source_name="", filter_hint="", filter_summary=None) -> str:
    """إنشاء نسخة HTML مستقلة من Dashboard النشاط بنفس التسلسل والألوان والرسوم الأساسية."""
    from html import escape
    import json

    # التقرير المصدّر له Light Mode مستقل حتى يظل واضحًا عند فتحه في أي متصفح.
    background = "#F5F7FB"
    surface = "#FFFFFF"
    border = "#D9E2EC"
    text = "#1F2937"
    text_dim = "#526174"
    export_success = "#2F6F73"
    export_fail = "#8095A2"
    export_accent = "#477F82"
    export_warn = "#628B8E"
    export_template = "plotly_white"
    work = df.copy()
    if sales_col and sales_col in work.columns:
        work["_agent_display"] = work[sales_col].fillna("غير محدد").astype(str).str.strip()
    else:
        work["_agent_display"] = "غير محدد"
    work["_success_bool"] = _activity_success_mask(work, class_col)
    sub_col_for_export = find_column(work, PROMISE_SUB_STATE_CANDIDATES)
    if sub_col_for_export:
        work["_activity_state"] = work[sub_col_for_export].map(_classify_activity_sub_state)
    total = len(work)
    success = int(work["_success_bool"].sum())
    rate = success / total * 100 if total else 0
    wasted = pd.to_numeric(work.get(WASTED_TIME_COL, pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    agent_count = int(work["_agent_display"].nunique()) if total else 0
    filter_summary = filter_summary or {}
    filter_items = [
        ("👤 المحصل", filter_summary.get("المحصل", "كل المحصلين")),
        ("📊 الحالة الفرعية", filter_summary.get("الحالة الفرعية", "كل الحالات")),
        ("📅 التاريخ", filter_summary.get("التاريخ", "كل التواريخ")),
        ("🏷️ التصنيف", filter_summary.get("التصنيف", "الكل")),
    ]
    timed_source = pd.to_datetime(work[time_col], errors="coerce") if time_col and time_col in work.columns else pd.Series(pd.NaT, index=work.index)
    raw_records = []
    for row_index, row in work.iterrows():
        timestamp = timed_source.loc[row_index] if row_index in timed_source.index else pd.NaT
        raw_records.append({
            "agent": str(row.get("_agent_display", "غير محدد")),
            "time": timestamp.isoformat() if pd.notna(timestamp) else "",
            "success": bool(row.get("_success_bool", False)),
            "wasted": float(pd.to_numeric(row.get(WASTED_TIME_COL, 0), errors="coerce") or 0),
            "state": str(_classify_activity_sub_state(row.get(find_column(work, PROMISE_SUB_STATE_CANDIDATES), ""))) if find_column(work, PROMISE_SUB_STATE_CANDIDATES) else "",
        })
    raw_records_json = json.dumps(raw_records, ensure_ascii=False).replace("</", "<\\/")
    date_values = sorted({record["time"][:10] for record in raw_records if record.get("time")})
    export_date_min = date_values[0] if date_values else ""
    export_date_max = date_values[-1] if date_values else ""
    agent_color_json = json.dumps(_activity_agent_color_map(work["_agent_display"]), ensure_ascii=False)
    state_color_json = json.dumps(dict(zip(ACTIVITY_NO_ANSWER_STATES, ACTIVITY_STATE_PALETTE)), ensure_ascii=False)

    def metric_card(card_id, label, value, color):
        return (
            f'<div id="{card_id}" style="background:{surface};border:1px solid {border};border-radius:14px;'
            f'padding:22px 14px;text-align:center;min-height:112px;box-sizing:border-box">'
            f'<div style="color:{text_dim};font-size:15px;margin-bottom:12px">{label}</div>'
            f'<div data-role="value" style="color:{color};font-size:28px;font-weight:700;line-height:1.2">{value}</div></div>'
        )

    parts = [
        '<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        '<title>داشبورد تحليل نشاط المحصلين</title></head>',
        f'<body style="margin:0;background:{background};color:{text};font-family:Tahoma,Arial,sans-serif;line-height:1.6">',
        '<main style="max-width:1500px;margin:0 auto;padding:28px 30px">',
        f'<header style="background:{surface};border:1px solid {border};border-radius:16px;padding:24px 28px;margin-bottom:22px">'
        '<div style="font-size:13px;color:' + COLOR_ACCENT + ';letter-spacing:1px">ACTIVITY DASHBOARD</div>'
        '<h1 style="margin:4px 0 2px;font-size:30px">📊 تحليل نشاط المحصلين</h1>'
        f'<div style="color:{text_dim};font-size:14px">مصدر البيانات: {escape(source_name or "ملف النشاط")}</div>',
    ]
    if filter_hint:
        parts.append(f'<div style="margin-top:12px;color:{text_dim};font-size:13px">الفلاتر النشطة: {escape(filter_hint)}</div>')
    agent_options = sorted(work["_agent_display"].dropna().astype(str).unique().tolist())
    state_options = sorted(work["_activity_state"].dropna().astype(str).unique().tolist()) if "_activity_state" in work.columns else []
    parts.extend([
        '</header>',
        f'<section style="background:{surface};border:0;border-radius:12px;padding:8px 0;margin-bottom:14px">',
        f'<h2 style="margin:0 0 8px;text-align:center;font-size:18px;color:{text}">🎚️ فلاتر التقرير التفاعلية</h2>',
        '<section id="interactive-filters" style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;align-items:end">',
        f'<div style="position:relative;display:flex;flex-direction:column;gap:7px;color:{text_dim};font-size:13px"><span>👤 المحصلون</span><button type="button" class="multi-trigger" data-target="agent-menu" style="background:{background};color:{text};border:1px solid {border};border-radius:9px;padding:8px;font-size:13px;text-align:right;cursor:pointer"><span id="agent-label">كل المحصلين</span>⌄</button><div id="agent-menu" class="multi-menu" style="display:none;position:absolute;z-index:20;top:74px;right:0;left:0;background:#FFFFFF;color:{text};border:1px solid {border};border-radius:10px;padding:8px;box-shadow:0 10px 24px rgba(15,23,42,.16);max-height:230px;overflow-y:auto"><label style="display:block;padding:8px;border-bottom:1px solid {border};font-weight:700"><input type="checkbox" class="select-all-agent"> كل المحصلين</label>',
    ])
    for value in agent_options:
        parts.append(f'<label style="display:block;padding:8px 6px;border-radius:7px;cursor:pointer"><input type="checkbox" class="agent-option" value="{escape(value, quote=True)}"> {escape(value)}</label>')
    parts.extend([
        f'</div></div><div style="position:relative;display:flex;flex-direction:column;gap:7px;color:{text_dim};font-size:13px"><span>📊 الحالات الفرعية</span><button type="button" class="multi-trigger" data-target="state-menu" style="background:{background};color:{text};border:1px solid {border};border-radius:9px;padding:8px;font-size:13px;text-align:right;cursor:pointer"><span id="state-label">كل الحالات</span>⌄</button><div id="state-menu" class="multi-menu" style="display:none;position:absolute;z-index:20;top:74px;right:0;left:0;background:#FFFFFF;color:{text};border:1px solid {border};border-radius:10px;padding:8px;box-shadow:0 10px 24px rgba(15,23,42,.16);max-height:230px;overflow-y:auto"><label style="display:block;padding:8px;border-bottom:1px solid {border};font-weight:700"><input type="checkbox" class="select-all-state"> كل الحالات</label>',
    ])
    for value in state_options:
        parts.append(f'<label style="display:block;padding:8px 6px;border-radius:7px;cursor:pointer"><input type="checkbox" class="state-option" value="{escape(value, quote=True)}"> {escape(value)}</label>')
    parts.extend([
        f'</div></div><label style="display:flex;flex-direction:column;gap:7px;color:{text_dim};font-size:13px"><span>🏷️ التصنيف</span><select id="filter-class" style="background:{background};color:{text};border:1px solid {border};border-radius:9px;padding:8px;font-size:13px"><option value="">الكل</option><option value="success">ناجحة</option><option value="failure">غير ناجحة</option></select></label>',
        f'<label style="display:flex;flex-direction:column;gap:7px;color:{text_dim};font-size:13px"><span>📅 من تاريخ</span><input id="filter-date-from" type="date" value="{export_date_min}" min="{export_date_min}" max="{export_date_max}" style="background:{background};color:{text};border:1px solid {border};border-radius:9px;padding:8px;font-size:13px"></label>',
        f'<label style="display:flex;flex-direction:column;gap:7px;color:{text_dim};font-size:13px"><span>📅 إلى تاريخ</span><input id="filter-date-to" type="date" value="{export_date_max}" min="{export_date_min}" max="{export_date_max}" style="background:{background};color:{text};border:1px solid {border};border-radius:9px;padding:8px;font-size:13px"></label>',
        f'<div style="display:flex;gap:8px;align-items:end"><button id="reset-filters" type="button" style="flex:1;background:{export_accent};color:#fff;border:0;border-radius:9px;padding:8px;font-size:13px;cursor:pointer">↺ إعادة ضبط</button></div>',
        '</section><div id="filter-status" style="text-align:center;color:' + text_dim + ';font-size:12px;margin-top:12px">عرض كل البيانات</div></section>',
        '<section id="kpi-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:18px">',
        metric_card("kpi-agents", "👥 عدد المحصلين", f"{agent_count:,}", text),
        metric_card("kpi-total", "📞 إجمالي المكالمات", f"{total:,}", text),
        metric_card("kpi-success", "✅ المكالمات الناجحة", f"{success:,}", export_success),
        metric_card("kpi-rate", "📈 نسبة النجاح", f"{rate:.1f}%", export_accent),
        metric_card("kpi-wasted", "⏱️ إجمالي الوقت المهدر", f"{wasted:,.1f} دقيقة", export_warn),
        '</section>',
    ])

    figs = []
    if class_col and class_col in work.columns:
        donut_df = pd.DataFrame({"النتيجة": ["ناجحة", "غير ناجحة"], "العدد": [success, total - success]})
        fig = px.pie(donut_df, names="النتيجة", values="العدد", hole=0.62, color="النتيجة", color_discrete_map={"ناجحة": export_success, "غير ناجحة": export_fail}, template=export_template)
        fig.update_traces(textinfo="percent", textfont_size=15, marker={"line": {"color": surface, "width": 3}}, hovertemplate="<b>%{label}</b><br>العدد: %{value:,}<br>النسبة: %{percent}<extra></extra>")
        fig.update_layout(**_activity_layout(title="🎯 الناجحة مقابل غير الناجحة", title_x=0.5, height=410, margin={"t":68,"b":65,"l":20,"r":20}, legend={"orientation":"h","y":-0.1,"x":0.5,"xanchor":"center"}, annotations=[{"text":f"{rate:.1f}%<br>نجاح","x":0.5,"y":0.5,"font":{"size":22,"color":export_success},"showarrow":False}]))
        figs.append(("🎯 توزيع نتائج المكالمات", fig))

    if time_col and time_col in work.columns:
        work["_activity_time"] = pd.to_datetime(work[time_col], errors="coerce")
        timed = work.dropna(subset=["_activity_time"]).copy()
        if not timed.empty:
            timed["اليوم"] = timed["_activity_time"].dt.strftime("%Y-%m-%d")
            timed["_success_for_day"] = _activity_success_mask(timed, class_col)
            daily = timed.groupby(["اليوم", "_agent_display"], as_index=False).agg(**{"عدد المكالمات": ("_agent_display", "size"), "المكالمات الناجحة": ("_success_for_day", "sum")})
            daily_totals = daily.groupby("اليوم", as_index=False).agg(**{"إجمالي المكالمات": ("عدد المكالمات", "sum"), "إجمالي الناجحة": ("المكالمات الناجحة", "sum")})
            daily_totals["نسبة النجاح (%)"] = (daily_totals["إجمالي الناجحة"] / daily_totals["إجمالي المكالمات"].replace(0, pd.NA) * 100).fillna(0).round(1)
            ordered_days = sorted(daily_totals["اليوم"].tolist())
            daily["اليوم"] = pd.Categorical(daily["اليوم"], categories=ordered_days, ordered=True)
            daily = daily.sort_values(["اليوم", "_agent_display"])
            daily_totals["اليوم"] = pd.Categorical(daily_totals["اليوم"], categories=ordered_days, ordered=True)
            daily_totals = daily_totals.sort_values("اليوم")
            day_fig = px.bar(daily, x="اليوم", y="عدد المكالمات", color="_agent_display", barmode="group", text_auto=True, custom_data=["_agent_display"], template=export_template, labels={"_agent_display":"المحصل"}, color_discrete_sequence=ACTIVITY_AGENT_PALETTE)
            day_fig.add_trace(go.Scatter(x=daily_totals["اليوم"].astype(str), y=daily_totals["نسبة النجاح (%)"], name="نسبة النجاح", mode="lines+markers+text", text=daily_totals["نسبة النجاح (%)"].map(lambda value: f"{value:.1f}%"), textposition="top center", line={"color": export_accent, "width": 3}, marker={"color": export_accent, "size": 9}, yaxis="y2", hovertemplate="<b>%{x}</b><br>نسبة النجاح: %{y:.1f}%<extra></extra>"))
            day_fig.update_layout(**_activity_layout(title="📊 Combo Chart يومي: المكالمات ونسبة النجاح", title_x=0.5, xaxis_title="اليوم", yaxis_title="عدد المكالمات", height=430, bargap=0.18, margin={"t":68,"b":95,"l":55,"r":65}, xaxis={"type":"category","categoryorder":"array","categoryarray":ordered_days,"tickangle":-25}, yaxis={"rangemode":"tozero"}, yaxis2={"title":"نسبة النجاح (%)","overlaying":"y","side":"right","range":[0,100],"ticksuffix":"%","showgrid":False}, legend={"orientation":"h","y":-0.2,"x":0.5,"xanchor":"center"}))
            day_fig.update_traces(selector={"type":"bar"}, marker_line_width=0, hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>")
            figs.append(("📊 Combo Chart النشاط اليومي", day_fig))

            timed["الساعة"] = timed["_activity_time"].dt.hour
            hour_min = int(timed["الساعة"].min())
            hour_max = int(timed["الساعة"].max())
            hourly = timed.groupby(["الساعة", "_agent_display"], as_index=False).size().rename(columns={"size":"عدد المكالمات"})
            hour_fig = px.bar(hourly, x="الساعة", y="عدد المكالمات", color="_agent_display", barmode="stack", text_auto=True, custom_data=["_agent_display"], template=export_template, labels={"_agent_display":"المحصل"}, color_discrete_sequence=ACTIVITY_AGENT_PALETTE)
            hour_fig.update_layout(**_activity_layout(title="🕒 Histogram ساعي لنشاط المحصلين", title_x=0.5, xaxis_title="ساعة اليوم", yaxis_title="عدد المكالمات", height=430, bargap=0.08, margin={"t":68,"b":95,"l":55,"r":20}, xaxis={"dtick":1,"tickvals":list(range(hour_min, hour_max + 1)),"range":[max(-0.5, hour_min - 0.5), min(23.5, hour_max + 0.5)]}, legend={"orientation":"h","y":-0.2,"x":0.5,"xanchor":"center"}))
            hour_fig.update_traces(marker_line_width=0, hovertemplate="<b>الساعة %{x}:00</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>")
            figs.append(("🕒 Histogram النشاط الساعي", hour_fig))

    sub_col = find_column(work, PROMISE_SUB_STATE_CANDIDATES)
    if sub_col:
        work["_activity_state"] = work[sub_col].map(_classify_activity_sub_state)
        state_counts = work.pivot_table(index="_agent_display", columns="_activity_state", aggfunc="size", fill_value=0)
        for state_name in ACTIVITY_NO_ANSWER_STATES:
            if state_name not in state_counts.columns:
                state_counts[state_name] = 0
        state_counts = state_counts.reindex(columns=ACTIVITY_NO_ANSWER_STATES, fill_value=0).reset_index().rename(columns={"_agent_display":"المحصّل"})
        state_long = state_counts.melt(id_vars=["المحصّل"], var_name="الحالة", value_name="العدد")
        no_fig = px.bar(state_long, x="العدد", y="المحصّل", orientation="h", color="الحالة", barmode="stack", text_auto=True, template=export_template, category_orders={"الحالة":ACTIVITY_NO_ANSWER_STATES}, color_discrete_sequence=[export_success, export_accent, export_warn, "#A6B4B9"])
        no_fig.update_layout(**_activity_layout(title="📵 حالات لا يرد لكل محصل", title_x=0.5, xaxis_title="عدد الحالات", yaxis_title="", height=430, margin={"t":68,"b":80,"l":105,"r":20}, legend={"orientation":"h","y":-0.18,"x":0.5,"xanchor":"center"}, yaxis={"categoryorder":"total ascending"}))
        no_fig.update_traces(hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,}<extra></extra>")
        figs.append(("📵 تحليل حالات Sub State", no_fig))

    parts.append('<section style="display:grid;grid-template-columns:repeat(auto-fit,minmax(460px,1fr));gap:18px">')
    include_js = True
    for index, (heading, fig) in enumerate(figs):
        parts.append(f'<article id="chart-card-{index}" style="background:{surface};border:1px solid {border};border-radius:16px;padding:10px 14px 4px;min-width:0"><h2 style="font-size:17px;margin:8px 10px;color:{text};text-align:center">{heading}</h2>')
        parts.append(pio.to_html(fig, full_html=False, include_plotlyjs=include_js, config=PLOTLY_CONFIG, div_id=f"activity_plot_{index}", default_width="100%", default_height=f"{max(fig.layout.height or 430, 450)}px"))
        parts.append('</article>')
        include_js = False
    parts.append('</section>')

    if sales_col and sales_col in df.columns:
        agent_table, _, _ = _build_activity_summary(df, class_col, sales_col, time_col)
        columns = ["المحصّل", "إجمالي المكالمات", "المكالمات الناجحة", "نسبة النجاح (%)", "نسبة من إجمالي المكالمات (%)", "واعد بالسداد", "إجمالي لا يرد", "أيام النشاط", "متوسط ساعات العمل/اليوم", "إجمالي ساعات العمل", "إجمالي الوقت المهدر (دقيقة)"]
        columns = [column for column in columns if column in agent_table.columns]
        parts.append(f'<section id="activity-summary-table" style="background:{surface};border:1px solid {border};border-radius:16px;padding:18px;margin-top:18px"><h2 style="font-size:19px;margin:0 0 12px;text-align:center">📋 ملخص أداء كل محصل</h2><div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr>')
        for column in columns:
            parts.append(f'<th style="padding:10px;border-bottom:1px solid {border};color:{text_dim};white-space:nowrap;text-align:right">{escape(column)}</th>')
        parts.append('</tr></thead><tbody>')
        for _, row in agent_table.sort_values("إجمالي المكالمات", ascending=False).iterrows():
            parts.append('<tr>')
            for column in columns:
                value = row[column]
                if isinstance(value, float):
                    value = f"{value:,.2f}"
                parts.append(f'<td style="padding:9px;border-bottom:1px solid rgba(128,145,170,.18);white-space:nowrap">{escape(str(value))}</td>')
            parts.append('</tr>')
        parts.append('</tbody></table></div></section>')

    interactive_js = """
<script>
const activityData = __ACTIVITY_DATA__;
const agentColors = __AGENT_COLORS__;
const stateColors = __STATE_COLORS__;
const dailyPlot = document.getElementById('activity_plot_1');
const hourlyPlot = document.getElementById('activity_plot_2');
const donutPlot = document.getElementById('activity_plot_0');
const statePlot = document.getElementById('activity_plot_3');
const fmt = n => Number(n || 0).toLocaleString('en-US');
function updateMultiLabels() {
  const agent = [...document.querySelectorAll('.agent-option:checked')].map(o => o.value);
  const state = [...document.querySelectorAll('.state-option:checked')].map(o => o.value);
  document.getElementById('agent-label').textContent = agent.length ? `${agent.length} محصل محدد` : 'كل المحصلين';
  document.getElementById('state-label').textContent = state.length ? `${state.length} حالة محددة` : 'كل الحالات';
}
function selectedRows() {
  const values = cls => [...document.querySelectorAll('.' + cls + ':checked')].map(option => option.value).filter(Boolean);
  const agent = values('agent-option');
  const state = values('state-option');
  const cls = document.getElementById('filter-class').value;
  const from = document.getElementById('filter-date-from').value;
  const to = document.getElementById('filter-date-to').value;
  return activityData.filter(row => ((!agent.length) || agent.includes(row.agent)) && ((!state.length) || state.includes(row.state)) && (!cls || (cls === 'success' ? row.success : !row.success)) && (!from || !row.time || row.time.slice(0,10) >= from) && (!to || !row.time || row.time.slice(0,10) <= to));
}
function setKpi(id, value) { const el = document.querySelector('#' + id + ' [data-role=value]'); if (el) el.textContent = value; }
function refreshDashboard() {
  const rows = selectedRows();
  const agents = [...new Set(rows.map(r => r.agent))].sort();
  const success = rows.filter(r => r.success).length;
  const rate = rows.length ? success / rows.length * 100 : 0;
  setKpi('kpi-agents', fmt(agents.length)); setKpi('kpi-total', fmt(rows.length)); setKpi('kpi-success', fmt(success)); setKpi('kpi-rate', rate.toFixed(1) + '%'); setKpi('kpi-wasted', Number(rows.reduce((s,r) => s + (r.wasted || 0), 0)).toLocaleString('en-US', {maximumFractionDigits:1}) + ' دقيقة');
  document.getElementById('filter-status').textContent = `عرض ${fmt(rows.length)} مكالمة من أصل ${fmt(activityData.length)} — ${fmt(agents.length)} محصل`;
  const days = [...new Set(rows.filter(r => r.time).map(r => r.time.slice(0,10)))].sort();
  const dailyTraces = agents.map(agent => ({type:'bar', name:agent, x:days, y:days.map(day => rows.filter(r => r.agent===agent && r.time.slice(0,10)===day).length), marker:{color:agentColors[agent] || '#6F9FB5'}, texttemplate:'%{y}', textposition:'inside', hovertemplate:'<b>%{x}</b><br>%{fullData.name}: %{y} مكالمة<extra></extra>'}));
  const dailySuccess = days.map(day => { const d=rows.filter(r => r.time && r.time.slice(0,10)===day); return d.length ? d.filter(r=>r.success).length/d.length*100 : 0; });
  dailyTraces.push({type:'scatter', mode:'lines+markers+text', name:'نسبة النجاح', x:days, y:dailySuccess, text:dailySuccess.map(v=>v.toFixed(1)+'%'), textposition:'top center', line:{color:'__ACCENT__',width:3}, marker:{color:'__ACCENT__',size:8}, yaxis:'y2', hovertemplate:'<b>%{x}</b><br>نسبة النجاح: %{y:.1f}%<extra></extra>'});
  if (dailyPlot) Plotly.react(dailyPlot, dailyTraces, {...dailyPlot.layout, xaxis:{...(dailyPlot.layout?.xaxis||{}), type:'category', categoryarray:days}, yaxis2:{...(dailyPlot.layout?.yaxis2||{}), range:[0,100], ticksuffix:'%'}});
  const hours = Array.from({length:24},(_,i)=>i); const hourlyTraces = agents.map(agent=>({type:'bar',name:agent,x:hours,y:hours.map(h=>rows.filter(r=>r.agent===agent && r.time && new Date(r.time).getHours()===h).length),marker:{color:agentColors[agent]||'#6F9FB5'},texttemplate:'%{y}',textposition:'inside'}));
  if (hourlyPlot) Plotly.react(hourlyPlot, hourlyTraces, {...hourlyPlot.layout, barmode:'stack', xaxis:{...(hourlyPlot.layout?.xaxis||{}), dtick:1, range:[-0.5,23.5]}});
  if (donutPlot) Plotly.react(donutPlot, [{type:'pie',labels:['ناجحة','غير ناجحة'],values:[success, rows.length-success],hole:.62,marker:{colors:['__SUCCESS__','__FAIL__']},textinfo:'percent'}], donutPlot.layout);
  const stateNames = Object.keys(stateColors); const stateTraces = stateNames.map(state=>({type:'bar',name:state,x:agents,y:agents.map(a=>rows.filter(r=>r.agent===a && r.state===state).length),marker:{color:stateColors[state]},texttemplate:'%{y}',textposition:'inside'}));
  if (statePlot) Plotly.react(statePlot, stateTraces, {...statePlot.layout, barmode:'stack'});
}
document.querySelectorAll('.multi-trigger').forEach(trigger => trigger.addEventListener('click', event => { event.stopPropagation(); const menu = document.getElementById(trigger.dataset.target); document.querySelectorAll('.multi-menu').forEach(other => { if (other !== menu) other.style.display = 'none'; }); menu.style.display = menu.style.display === 'block' ? 'none' : 'block'; }));
document.addEventListener('click', () => document.querySelectorAll('.multi-menu').forEach(menu => menu.style.display = 'none'));
document.querySelectorAll('.agent-option,.state-option').forEach(option => option.addEventListener('change', () => { updateMultiLabels(); refreshDashboard(); }));
document.querySelector('.select-all-agent')?.addEventListener('change', event => { document.querySelectorAll('.agent-option').forEach(option => option.checked = event.target.checked); updateMultiLabels(); refreshDashboard(); });
document.querySelector('.select-all-state')?.addEventListener('change', event => { document.querySelectorAll('.state-option').forEach(option => option.checked = event.target.checked); updateMultiLabels(); refreshDashboard(); });
['filter-class','filter-date-from','filter-date-to'].forEach(id => document.getElementById(id)?.addEventListener('change', refreshDashboard));
document.getElementById('reset-filters')?.addEventListener('click', () => { document.querySelectorAll('.agent-option,.state-option,.select-all-agent,.select-all-state').forEach(option => option.checked=false); document.getElementById('filter-class').value=''; document.getElementById('filter-date-from').value='__DATE_MIN__'; document.getElementById('filter-date-to').value='__DATE_MAX__'; updateMultiLabels(); refreshDashboard(); });
updateMultiLabels();
refreshDashboard();
</script>
""".replace('__ACTIVITY_DATA__', raw_records_json).replace('__AGENT_COLORS__', agent_color_json).replace('__STATE_COLORS__', state_color_json).replace('__ACCENT__', json.dumps(export_accent)).replace('__SUCCESS__', json.dumps(export_success)).replace('__FAIL__', json.dumps(export_fail)).replace('__DATE_MIN__', export_date_min).replace('__DATE_MAX__', export_date_max)
    parts.append(interactive_js)
    parts.extend(['<footer style="color:' + text_dim + ';font-size:12px;text-align:center;margin-top:24px">تم إنشاء التقرير من لوحة تحليل نشاط المحصلين</footer></main></body></html>'])
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


def _render_period_results(stored, period_key):
    """عرض نتائج الفترة المحفوظة (جدول + كروت + شارتات + تنزيل) — تُستخدم بعد الضغط على زر التصنيف وبعد شيل الملف."""
    period_title = stored["period_title"]
    result_df = stored["df"]
    sales_col = stored["sales_col"]
    time_col = stored["time_col"]

    render_duplicate_summary(stored.get("duplicate_stats"))
    st.dataframe(result_df, use_container_width=True, hide_index=True)

    render_period_charts(result_df, sales_col, time_col, period_title)

    if result_df is not None:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="النتائج")
        st.download_button(
            f"⬇️ تحميل نتائج {period_title}",
            data=buffer.getvalue(),
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
            go.Bar(name="إجمالي المكالمات", x=names, y=agent_sorted["إجمالي المكالمات"], marker_color=THEME["text_dim"], customdata=names),
            go.Bar(name="المكالمات الناجحة", x=names, y=agent_sorted["ناجحة"], marker_color=COLOR_SUCCESS, customdata=names),
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
            color_continuous_scale=[COLOR_FAIL, COLOR_WARN, COLOR_SUCCESS],
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
        go.Bar(name="المكالمات الناجحة", x=names, y=ordered["ناجحة"], marker_color=COLOR_SUCCESS, customdata=names),
        go.Bar(name="المكالمات غير الناجحة", x=names, y=failed, marker_color=COLOR_FAIL, customdata=names),
    ])
    fig.update_layout(**PLOTLY_LAYOUT, title=f"✅ الناجحة مقابل غير الناجحة ({period_title})", barmode="group", xaxis_title="", yaxis_title="عدد المكالمات")
    fig.update_traces(hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:,} مكالمة<extra></extra>")
    render_selectable_chart(fig, f"classification_success_fail_{period_title}")


def render_avg_duration_chart(agent, period_title):
    key = "متوسط مدة المكالمة (دقيقة)"
    if key not in agent.columns:
        return
    ordered = agent.sort_values(key, ascending=True)
    fig = px.bar(ordered, x=key, y="المحصّل", orientation="h", text=key, color=key, color_continuous_scale=[COLOR_ACCENT, COLOR_WARN], template=PLOTLY_TEMPLATE)
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
    fig = px.bar(counts, x=sales_col, y="عدد إفادات لا يرد", color_discrete_sequence=[COLOR_FAIL], template=PLOTLY_TEMPLATE)
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
        _render_aggregate_results(stored, period_title)
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

def _render_aggregate_results(stored, period_title):
    result_df = stored["df"]
    sales_col = stored["sales_col"]
    time_col = stored["time_col"]
    render_duplicate_summary(stored.get("duplicate_stats"))
    st.dataframe(result_df, use_container_width=True, hide_index=True)
    render_period_charts(result_df, sales_col, time_col, period_title)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name=period_title)
    st.download_button(
        f"⬇️ تحميل نتائج {period_title}",
        data=buffer.getvalue(),
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
        _render_aggregate_results(stored, period_title)
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


def _run_neglect_followup_pipeline(new_file, old_file):
    try:
        df_new = read_uploaded_dataframe(new_file)
        df_old = read_uploaded_dataframe(old_file)
        
        # حذف أول صف
        if len(df_new) > 0: df_new = df_new.iloc[1:].reset_index(drop=True)
        if len(df_old) > 0: df_old = df_old.iloc[1:].reset_index(drop=True)
        
        # البحث عن الأعمدة
        id_col_new = find_column(df_new, ID_CANDIDATES)
        id_col_old = find_column(df_old, ID_CANDIDATES)
        last_date_new = find_column(df_new, NEGLECT_LAST_DATE_CANDIDATES)
        
        if not id_col_new or not id_col_old or not last_date_new:
            st.error("تعذر العثور على عمود الرقم التعريفي (ID) أو تاريخ آخر متابعة في الملفات.")
            return
            
        # تحويل المعرفات لنصوص لضمان المطابقة
        df_new[id_col_new] = df_new[id_col_new].astype(str).str.strip()
        df_old[id_col_old] = df_old[id_col_old].astype(str).str.strip()
        
        # جلب التواريخ الحديثة
        mapping = df_new.set_index(id_col_new)[last_date_new].to_dict()
        
        # تحديث ملف الإهمال القديم
        result_df = df_old.copy()
        result_df['تاريخ_متابعة_حديث'] = result_df[id_col_old].map(mapping)
        
        # حساب الملاحظات
        target_date = st.session_state.get(TODAY_KEY, datetime.now().date())
        
        def check_coverage(val):
            d = parse_date_cell(val)
            if not d: return "لم يتم التغطية"
            diff = (target_date - d).days
            return "تم التغطية" if diff < 8 else "لم يتم التغطية"
            
        result_df['الملاحظات'] = result_df['تاريخ_متابعة_حديث'].apply(check_coverage)
        
        st.session_state["neglect_followup_result"] = {"df": result_df}
        st.rerun()
    except Exception as e:
        st.error(f"خطأ في معالجة الملفات: {e}")

def _show_neglect_followup_results(df):
    st.subheader("📊 نتائج متابعة الإهمال")
    covered = int((df["الملاحظات"] == "تم التغطية").sum())
    not_covered = int((df["الملاحظات"] == "لم يتم التغطية").sum())
    pct = covered / len(df) * 100 if len(df) else 0
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ تم التغطية", covered)
    c2.metric("❌ لم يتم التغطية", not_covered)
    c3.metric("📈 نسبة التغطية", f"{pct:.1f}%")
    st.dataframe(df, use_container_width=True, hide_index=True)
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="متابعة الإهمال")
    st.download_button("⬇️ تحميل تقرير متابعة الإهمال المحدث (Excel)", data=out_excel.getvalue(), file_name=f"متابعة_الإهمال_{datetime.now():%Y-%m-%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")


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
            _show_neglect_followup_results(cached_followup["df"])

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
    sales_col = meta["sales_col"]
    net_col = meta["net_col"]
    total = len(df)
    total_amount = pd.to_numeric(df[net_col], errors="coerce").fillna(0).sum() if net_col and net_col in df.columns else 0
    agent_count = int(df[sales_col].nunique()) if sales_col and sales_col in df.columns else 0
    avg_days = pd.to_numeric(df["فرق_الأيام"], errors="coerce").mean() if "فرق_الأيام" in df.columns else None

    st.subheader("📊 ملخص الإهمال")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("⚠️ إجمالي الحالات", f"{total:,}")
    c2.metric("👥 عدد المحصّلين", f"{agent_count:,}")
    c3.metric("💰 إجمالي المديونية", f"{total_amount:,.0f}" if net_col else "—")
    c4.metric("📅 متوسط فرق الأيام", f"{avg_days:.1f}" if avg_days is not None and pd.notna(avg_days) else "—")

    if total and sales_col and sales_col in df.columns:
        agent_counts = df[sales_col].value_counts().head(15).sort_values().reset_index()
        agent_counts.columns = ["المحصّل", "عدد الحالات"]
        fig = px.bar(
            agent_counts, x="عدد الحالات", y="المحصّل", orientation="h",
            title="أعلى المحصّلين في حالات الإهمال",
            color="عدد الحالات", color_continuous_scale=[THEME["surface_2"], COLOR_WARN],
            template=PLOTLY_TEMPLATE,
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=420, yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    st.subheader("📋 جدول حالات الإهمال التفصيلي")
    display_cols = [c for c in [sales_col, meta["substate_col"], meta["duedate_col"], meta["lastdate_col"], "فرق_الأيام", net_col] if c and c in df.columns]
    st.dataframe(df[display_cols], use_container_width=True, hide_index=True)
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
        df[display_cols].to_excel(writer, index=False, sheet_name="حالات الإهمال")
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

    # لو اتشال الملف والنتيجة لسه في الذاكرة — نعرضها من الكاش
    if dash_file is None:
        _show_dashboard_from_cache(break_start, break_end)
        return

    # 💾 لو الملف ده اتعرج قبل كده — نعرض الكاش من غير إعادة معالجة
    current_file_hash = uploaded_file_hash(dash_file)
    if current_source_hash == current_file_hash and cached is not None:
        df_show, hint = _render_slicers(cached["df"], cached["sales_col"], cached["time_col"])
        _render_dashboard(df_show, cached["class_col"], cached["sales_col"],
                          cached["time_col"], dash_file.name, filter_hint=hint,
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
    df_show, hint = _render_slicers(df, sales_col, time_col)
    _render_dashboard(df_show, class_col, sales_col, time_col, dash_file.name, filter_hint=hint,
                      break_start=break_start, break_end=break_end)


def _render_dashboard(df, class_col, sales_col, time_col, source_name, filter_hint="", break_start=None, break_end=None):
    if filter_hint:
        st.info(f"الفلاتر المطبقة: {filter_hint}")
    render_full_dashboard(df, class_col=class_col, sales_col=sales_col, time_col=time_col,
                          break_start=break_start, break_end=break_end)
    dashboard_html = build_dashboard_html(
        df, class_col=class_col, sales_col=sales_col, time_col=time_col,
        source_name=source_name, filter_hint=filter_hint,
        filter_summary=st.session_state.get("dashboard_filter_summary", {}),
    )
    st.download_button("🌐 تحميل لوحة التحكم كصفحة ويب HTML", data=dashboard_html.encode("utf-8"), file_name="داشبورد_النشاط.html", mime="text/html", use_container_width=True, key="dash_html_download_v3", type="primary")
    st.download_button("⬇️ تحميل البيانات كـ CSV", data=df.to_csv(index=False).encode("utf-8-sig"), file_name="بيانات_النشاط.csv", mime="text/csv", use_container_width=True, key="dash_csv_download_v3")


def _clear_dashboard_chart_filter():
    st.session_state.pop(DASHBOARD_AGENT_FILTER_KEY, None)


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
    df, hint = _render_slicers(cached["df"], cached["sales_col"], cached["time_col"])
    _render_dashboard(df, cached["class_col"], cached["sales_col"],
                      cached["time_col"], cached["source_name"], filter_hint=hint,
                      break_start=break_start, break_end=break_end)


# ==========================================================
# التنقل (Sidebar Navigation)
# ==========================================================

PAGES = {
    "🎯 تصنيف المكالمات": page_classification,
    "📚 الوعود": page_promises,
    "⚠️ الإهمال والمتابعة": page_neglect,
    "🧾 أخطاء الحالات": lambda: page_placeholder(
        "CASE ERRORS", "أخطاء الحالات", "الحالات اللي فيها أخطاء في التسجيل أو المتابعة", "🧾"
    ),
    "📊 تحليل نشاط المحصّلين": page_dashboard,
}

DEFAULT_PAGE = next(iter(PAGES))
with st.sidebar:
    from pathlib import Path
    _logo_candidates = [
        Path(__file__).resolve().parent / "ahly_logo.png",
        Path("ahly_logo.png"),
    ]
    _logo_path = next((p for p in _logo_candidates if p.exists()), None)

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if _logo_path:
            st.image(str(_logo_path), width=110)
        else:
            st.markdown("<div style='text-align:center;font-size:40px'>🦅</div>", unsafe_allow_html=True)

    st.markdown(
        "<h2 style='text-align:center;margin:0.25rem 0 0.1rem'>🎙️ لوحة التحكم</h2>"
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
