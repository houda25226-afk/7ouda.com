"""تبويب «الإعفاءات» لتطبيق Streamlit.

الخطوات: (1) مصدر المطالبات والمرفقات: Odoo أو رفع ملفات  →  (2) قراءة المرفقات تلقائيا
→ (3) مراجعة وتعديل الحقول → (4) حساب الملاحظات + تنزيل Excel و zip المرفقات.

الأسرار في .streamlit/secrets.toml (أو Settings → Secrets في Streamlit Cloud):
    [odoo]
    url = "https://ejada.sa"
    db = "ejada"
    user = "يوزر_القراءة"
    api_key = "الـ API key"
"""
from __future__ import annotations
import ast
import io
import os
import re
import xmlrpc.client
import zipfile

import pandas as pd
import streamlit as st

from exemption_extract import DOC_LABELS, analyze_claim, to_images
from exemption_flow import AR_HEADERS, ROW_COLS, build_domain, build_row, run_review
from odoo_client import OdooClient
from odoo_export import fetch_attachments_by_lead, zip_files_by_claim

MODEL = "crm.lead"
NONE = "— لا يوجد —"
IMG_EXT = (".jpg", ".jpeg", ".png", ".jfif", ".webp")
ROLES = [("claim", "اسم/رقم المطالبة (مطلوب)"), ("driver_id", "هوية السائق"), ("driver_name", "اسم السائق"),
         ("owner_name", "مالك المركبة"), ("plate", "رقم اللوحة"), ("accident_date", "تاريخ الحادث")]


def _odoo_secrets():
    try:
        return dict(st.secrets["odoo"])
    except Exception:
        return None


def _client(url, db, user, key) -> OdooClient:
    # من غير cache: العميل خفيف (كل استدعاء باتصال جديد)، وكدا مفيش نسخة قديمة عالقة بعد تحديث الكود
    return OdooClient(url, db, user, key)


def _tesseract_ok() -> bool:
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _claim_from_filename(name: str) -> str:
    stem = os.path.splitext(os.path.basename(name))[0]
    return re.sub(r"_\d+$", "", stem)            # C-...-001_2.pdf → C-...-001


def _files_from_uploads(uploads) -> dict:
    out: dict[str, list] = {}
    for up in uploads:
        data = up.getvalue()
        if up.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for n in z.namelist():
                    if n.endswith("/") or os.path.basename(n).startswith(("_", ".")):
                        continue
                    out.setdefault(_claim_from_filename(n), []).append((os.path.basename(n), z.read(n)))
        else:
            out.setdefault(_claim_from_filename(up.name), []).append((up.name, data))
    return out


def _short_err(e) -> str:
    """رسالة خطأ Odoo قصيرة (آخر سطر من الـ traceback) بدل الصفحة كلها."""
    if isinstance(e, xmlrpc.client.Fault):
        lines = [x for x in str(e.faultString).strip().splitlines() if x.strip()]
        return lines[-1] if lines else str(e)
    return str(e)


def _norm(v) -> str:
    return str(v).strip().upper()


# ---------------------------------------------------------------- الخطوة 1: المصدر
def _source_odoo():
    sec = _odoo_secrets()
    if not sec:
        st.warning("مفيش إعدادات Odoo في الـ secrets. ضيف قسم [odoo] (الرابط، DB، اليوزر، الـ API key) من إعدادات Streamlit، "
                   "أو استخدم «رفع ملفات».")
        return None
    try:
        client = _client(sec["url"], sec["db"], sec["user"], sec["api_key"])
    except Exception as e:
        st.error(f"تعذر الاتصال بـ Odoo: {e}")
        return None
    if "ex_odoo_version" not in st.session_state:
        st.session_state["ex_odoo_version"] = client.version().get("server_version")
    st.caption(f"متصل ✅ | إصدار Odoo: {st.session_state['ex_odoo_version']}")
    if "ex_fields_df" not in st.session_state:
        st.session_state["ex_fields_df"] = client.fields(MODEL)
    fdf = st.session_state["ex_fields_df"]
    fdf = fdf[fdf["type"] != "binary"]
    label = {r.field: f"{r.field} — {r.string}" for r in fdf.itertuples()}
    with st.expander("استكشاف حقول crm.lead (الاسم التقني + الاسم الظاهر + النوع)"):
        st.dataframe(fdf[["field", "string", "type"]], use_container_width=True, height=260)
    textual = [r.field for r in fdf.itertuples() if r.type in ("char", "text", "integer")]
    guess = next((r.field for r in fdf.itertuples()
                  if r.type in ("char", "text") and re.search(r"claim|مطالب", str(r.string), re.I)), "name")
    c1, c2 = st.columns(2)
    claim_field = c1.selectbox("حقل رقم المطالبة في Odoo", textual, key="ex_claim_field", format_func=label.get,
                               index=textual.index(guess) if guess in textual else 0,
                               help="اختار الحقل اللي فيه رقم المطالبة (ابدأ اكتب جزء من اسمه أو اسمه الظاهر)")
    wanted = [f for f in ("name", "partner_id", "stage_id", "create_date") if f in label]
    if claim_field not in wanted:
        wanted.insert(0, claim_field)
    fields_sel = c2.multiselect("الحقول المطلوب سحبها", list(label), default=wanted, format_func=label.get,
                                key=f"ex_fields_{claim_field}",
                                help="ضيف حقول الهوية واللوحة وتاريخ الحادث لو موجودة في crm.lead")
    claims_text = st.text_area("أرقام مطالبات محددة (اختياري، كل رقم في سطر)", height=80, key="ex_claims")
    partial = st.checkbox("بحث جزئي (الرقم يكون جزء من الحقل)", key="ex_partial")
    domain_text = st.text_input("فلتر إضافي (Domain)", "[]", key="ex_domain",
                                help='مثال: [("create_date", ">=", "2026-10-01")]')
    limit = st.number_input("أقصى عدد مطالبات (0 = الكل)", min_value=0, value=200, step=50, key="ex_limit")
    if st.button("سحب المطالبات", key="ex_btn_pull", type="primary"):
        try:
            domain = build_domain(domain_text, claim_field, claims_text, partial)
            st.session_state["ex_leads"] = client.search_read_df(MODEL, domain, list(fields_sel), limit=int(limit) or None)
            st.session_state["ex_client_key"] = "odoo"
            st.session_state.pop("ex_rows", None)
        except Exception as e:
            st.error(f"خطأ في السحب: {_short_err(e)}")
    return st.session_state.get("ex_leads")


def _source_upload():
    ups = st.file_uploader("المرفقات (PDF / صور / zip). اسم الملف لازم يبدأ برقم المطالبة", key="ex_uploads",
                           type=["pdf", "jpg", "jpeg", "png", "jfif", "webp", "zip"], accept_multiple_files=True)
    sheet = st.file_uploader("شيت المطالبات (اختياري: Excel / CSV)", key="ex_sheet", type=["xlsx", "csv"])
    if not ups and not sheet:
        return None
    files = _files_from_uploads(ups or [])
    st.session_state["ex_files_up"] = files
    if sheet is not None:
        leads = pd.read_csv(sheet) if sheet.name.lower().endswith(".csv") else pd.read_excel(sheet)
    else:
        leads = pd.DataFrame({"name": sorted(files)})
    st.session_state["ex_leads"], st.session_state["ex_client_key"] = leads, "upload"
    st.caption(f"{len(files)} مطالبة فيها مرفقات، و{len(leads)} صف في الشيت/القايمة")
    return leads


# ---------------------------------------------------------------- الخطوة 2: القراءة
def _run_analysis(leads: pd.DataFrame, mapping: dict, max_claims: int):
    mode = st.session_state.get("ex_client_key")
    leads = leads.head(max_claims).reset_index(drop=True)
    names = [str(v) for v in leads[mapping["claim"]]]
    bar = st.progress(0.0, text="تجهيز المرفقات…")

    if mode == "odoo":
        sec = _odoo_secrets()
        client = _client(sec["url"], sec["db"], sec["user"], sec["api_key"])
        by_id = fetch_attachments_by_lead(client, [int(i) for i in leads["id"]],
                                          progress=lambda d, t: bar.progress(0.3 * d / max(t, 1), text=f"تنزيل المرفقات {d}/{t}"))
        files = {n: by_id.get(int(i), []) for n, i in zip(names, leads["id"])}
    else:
        up = {_norm(k): v for k, v in st.session_state.get("ex_files_up", {}).items()}
        files = {n: up.get(_norm(n), []) for n in names}

    rows, docs = [], {}
    for i, (n, lead) in enumerate(zip(names, leads.to_dict("records"))):
        bar.progress(0.3 + 0.7 * i / max(len(names), 1), text=f"قراءة المطالبة {i + 1}/{len(names)}: {n}")
        lead_vals = {k: lead.get(mapping[k]) for k in ("driver_id", "driver_name", "owner_name", "plate", "accident_date")
                     if mapping.get(k)}
        analysis = analyze_claim(files[n], driver_id=str(lead_vals.get("driver_id") or "") or None)
        rows.append(build_row(n, lead_vals, analysis))
        docs[n] = [{"file": d["file"], "type": d["type"], "method": d["method"]} for d in analysis["docs"]]
    bar.empty()
    st.session_state.update(ex_rows=pd.DataFrame(rows, columns=ROW_COLS), ex_files=files, ex_docs=docs,
                            ex_version=st.session_state.get("ex_version", 0) + 1)
    st.session_state.pop("ex_results", None)


# ---------------------------------------------------------------- معاينة
def _preview(claim: str):
    files = st.session_state.get("ex_files", {}).get(claim, [])
    docs = {d["file"]: d for d in st.session_state.get("ex_docs", {}).get(claim, [])}
    if not files:
        st.info("مفيش مرفقات للمطالبة دي.")
    for name, content in files:
        d = docs.get(name, {})
        st.markdown(f"**{name}** — {DOC_LABELS.get(d.get('type', 'unknown'))} ({'نص' if d.get('method') == 'text' else 'OCR'})")
        try:
            for img in to_images(content, name, dpi=90, max_pages=3):
                st.image(img, use_container_width=True)
        except Exception as e:
            st.warning(f"تعذر عرض الملف: {e}")


def _excel(results: pd.DataFrame, data: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        results.to_excel(w, sheet_name="النتائج", index=False)
        data.rename(columns=AR_HEADERS).to_excel(w, sheet_name="البيانات", index=False)
    return buf.getvalue()


# ---------------------------------------------------------------- الصفحة
def page_exemptions(page_header=None):
    title, sub = "🛡️ مراجعة الإعفاءات", "سحب المرفقات، قراءة الرخص والتأمين وخطابات الإلغاء، وكتابة ملاحظة كل مطالبة"
    if page_header:
        page_header("EXEMPTIONS", title, sub)
    else:
        st.title(title)
        st.caption(sub)
    if not _tesseract_ok():
        st.warning("الـ OCR مش متنصّب على السيرفر. ضيف ملف packages.txt فيه: tesseract-ocr و tesseract-ocr-ara. "
                   "من غيره الصور والمسح مش هتتقرا (ملفات PDF اللي فيها نص هتتقرا عادي).")

    st.markdown("### 1) مصدر المطالبات والمرفقات")
    source = st.radio("المصدر", ["Odoo", "رفع ملفات"], horizontal=True, key="ex_src", label_visibility="collapsed")
    leads = _source_odoo() if source == "Odoo" else _source_upload()
    if leads is None or leads.empty:
        return

    st.dataframe(leads.head(50), use_container_width=True, height=200)
    st.markdown("**ربط الأعمدة** (اختار العمود اللي فيه كل معلومة؛ اللي مش موجود سيبه «لا يوجد»)")
    cols = [c for c in leads.columns if c != "id"] if st.session_state.get("ex_client_key") == "odoo" else list(leads.columns)
    mapping, cc = {}, st.columns(3)
    for n, (role, label) in enumerate(ROLES):
        default = cols.index("name") if role == "claim" and "name" in cols else 0
        opts = cols if role == "claim" else [NONE] + cols
        pick = cc[n % 3].selectbox(label, opts, index=default if role == "claim" else 0, key=f"ex_map_{role}")
        mapping[role] = None if pick == NONE else pick

    max_claims = st.number_input("أقصى عدد مطالبات في المرة (القراءة بالـ OCR بطيئة: ~10-60 ثانية للمطالبة)",
                                 min_value=1, value=25, step=5, key="ex_max")
    if st.button("اقرا المرفقات", key="ex_btn_run", type="primary"):
        try:
            _run_analysis(leads, mapping, int(max_claims))
        except Exception as e:
            st.error(f"خطأ أثناء القراءة: {_short_err(e)}")

    rows = st.session_state.get("ex_rows")
    if rows is None:
        return

    st.markdown("### 2) مراجعة وتعديل")
    st.caption("اللي اتقرا بالـ OCR راجعه (خصوصا التواريخ). الخانات الفاضية اللي محتاجها المحرك املاها من الصورة. "
               "تاريخ الرخصة الهجري يتكتب كدا: 1448/03/19.")
    shown = st.data_editor(rows.rename(columns=AR_HEADERS), use_container_width=True, hide_index=True,
                           num_rows="fixed", key=f"ex_editor_{st.session_state.get('ex_version', 0)}")
    edited = shown.rename(columns={v: k for k, v in AR_HEADERS.items()})

    with st.expander("👁️ معاينة مرفقات مطالبة"):
        pick = st.selectbox("المطالبة", list(rows["claim"]), key="ex_preview_claim")
        _preview(pick)

    st.markdown("### 3) الملاحظات")
    if st.button("احسب الملاحظات", key="ex_btn_calc", type="primary"):
        st.session_state["ex_results"] = (run_review(edited), edited)
    if "ex_results" in st.session_state:
        res, data = st.session_state["ex_results"]
        c1, c2, c3 = st.columns(3)
        c1.metric("✅ جاهز", int((res["الحالة"] == "✅ جاهز").sum()))
        c2.metric("🟡 محتاج مراجعة", int((res["الحالة"] == "🟡 مراجعة").sum()))
        c3.metric("⚪ بدون مرفق", int((res["الحالة"] == "⚪ بدون مرفق").sum()))
        st.dataframe(res, use_container_width=True)
        st.download_button("⬇️ تنزيل النتائج (Excel)", _excel(res, data), "exemptions_review.xlsx", key="ex_dl_xlsx")

    st.markdown("### 4) ملف المرفقات")
    st.caption("zip فيه مرفقات كل مطالبة، وكل ملف باسم المطالبة (مرفق واحد: الاسم بس، أكتر: _1 و _2).")
    files = st.session_state.get("ex_files", {})
    if any(files.values()):
        st.download_button("⬇️ تنزيل مرفقات المطالبات (zip)", zip_files_by_claim(files), "claims_attachments.zip",
                           mime="application/zip", key="ex_dl_zip")
