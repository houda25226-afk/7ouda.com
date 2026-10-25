"""منطق تبويب الإعفاءات بدون واجهة (عشان يتختبر): بناء الصفوف، تحويلها لمطالبات، وتشغيل المحرك."""
from __future__ import annotations
import re

import pandas as pd

from exemption_engine import review_claim, parse_date
from exemption_extract import DOC_LABELS

ROW_COLS = [
    "claim", "driver_id", "driver_name", "owner_name", "plate", "accident_date", "case_number",
    "letter_id", "letter_accident_number", "letter_accident_date", "letter_date", "letter_company", "letter_amount",
    "lic_id", "lic_issue", "lic_expiry",
    "ins_holder_id", "ins_holder_name", "ins_plate", "ins_start", "ins_end",
    "najm_license_expiry", "ids_in_report", "القراءة",
]
AR_HEADERS = {
    "claim": "المطالبة", "driver_id": "هوية السائق", "driver_name": "اسم السائق", "owner_name": "مالك المركبة",
    "plate": "اللوحة", "accident_date": "تاريخ الحادث", "case_number": "رقم الحادث (نجم)",
    "letter_id": "هوية الخطاب", "letter_accident_number": "رقم الحادث بالخطاب", "letter_accident_date": "تاريخ الحادث بالخطاب",
    "letter_date": "تاريخ الخطاب", "letter_company": "شركة الخطاب", "letter_amount": "مبلغ الخطاب",
    "lic_id": "هوية الرخصة", "lic_issue": "إصدار الرخصة", "lic_expiry": "انتهاء الرخصة",
    "ins_holder_id": "هوية التأمين", "ins_holder_name": "اسم التأمين", "ins_plate": "لوحة التأمين",
    "ins_start": "بداية التأمين", "ins_end": "نهاية التأمين",
    "najm_license_expiry": "انتهاء الرخصة (نجم) للعلم", "ids_in_report": "هويات المحضر", "القراءة": "القراءة",
}


def _s(v) -> str:
    if isinstance(v, float) and v.is_integer():          # 1126070364.0 جاية من إكسل/pandas
        return str(int(v))
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip().lower() in ("nan", "none", "false") else str(v).strip()


def _date_only(v) -> str:
    """"2026-08-23 06:45:00" أو "23/08/2026 06:45" → التاريخ بس."""
    t = _s(v)
    m = re.search(r"\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}", t)
    return m.group(0) if m else t


def build_row(claim: str, lead: dict, analysis: dict) -> dict:
    """lead: قيم مأخوذة من Odoo/الشيت (driver_id, driver_name, owner_name, plate, accident_date). analysis: ناتج analyze_claim."""
    f = analysis["fields"]
    row = {c: "" for c in ROW_COLS}
    row.update({k: _s(f.get(k)) for k in ROW_COLS if k in f})
    row["claim"] = claim
    for k in ("driver_id", "driver_name", "owner_name", "plate"):
        row[k] = _s(lead.get(k))
    # تاريخ الحادث: الأولوية للشيت/Odoo ثم محضر نجم ثم الخطاب
    row["accident_date"] = _date_only(lead.get("accident_date")) or _s(f.get("accident_date")) or _s(f.get("letter_accident_date"))
    row["ins_plate"] = row["ins_plate"] or ""
    def label(d):
        if d["method"] == "error":
            return f"{d['file']} (تعذرت القراءة)"
        return f"{DOC_LABELS[d['type']]} ({'نص' if d['method'] == 'text' else 'OCR'})"
    row["القراءة"] = " | ".join(label(d) for d in analysis["docs"]) or "لا توجد مرفقات"
    return row


def row_to_claim(r: dict):
    """بيحوّل صف لمطالبة للمحرك، ولو مستند ناقص بيتشال منه ويتسجل تنبيه."""
    g = lambda k: _s(r.get(k))
    flags = []
    claim = {"accident": {"date": g("accident_date"), "driver_id": g("driver_id"), "driver_name": g("driver_name"),
                          "owner_name": g("owner_name"), "plate": g("plate"), "case_number": g("case_number")}}
    if not g("accident_date"):
        return None, ["تاريخ الحادث ناقص"]
    if not g("driver_id"):
        flags.append("هوية السائق ناقصة")

    if any(g(k) for k in ("letter_id", "letter_date")):
        if g("letter_id") and g("letter_date"):
            claim["cancellation_letter"] = {
                "id_number": g("letter_id"), "accident_number": g("letter_accident_number"),
                "accident_date": g("letter_accident_date"), "letter_date": g("letter_date"),
                "company": g("letter_company"), "amount": g("letter_amount")}
        else:
            flags.append("بيانات خطاب الإلغاء ناقصة (الهوية أو التاريخ)")
    if any(g(k) for k in ("lic_id", "lic_issue", "lic_expiry")):
        if g("lic_id") and g("lic_issue") and g("lic_expiry"):
            claim["license"] = {"id_number": g("lic_id"), "issue_date": g("lic_issue"), "expiry_date": g("lic_expiry")}
        else:
            flags.append("بيانات الرخصة ناقصة (الهوية أو الإصدار أو الانتهاء)")
    if any(g(k) for k in ("ins_start", "ins_end", "ins_holder_id", "ins_holder_name", "ins_plate")):
        if g("ins_start") and g("ins_end") and (g("ins_holder_id") or g("ins_holder_name") or g("ins_plate") or g("plate")):
            claim["insurance_print"] = {"holder_id": g("ins_holder_id"), "holder_name": g("ins_holder_name"),
                                        "plate": g("ins_plate") or g("plate"), "start_date": g("ins_start"),
                                        "end_date": g("ins_end")}
        else:
            flags.append("بيانات برنت التأمين ناقصة (البداية أو النهاية أو ما يربطه بالمركبة)")
    return claim, flags


def run_review(df: pd.DataFrame) -> pd.DataFrame:
    out = []
    for r in df.to_dict("records"):
        claim, flags = row_to_claim(r)
        if claim is None:
            res = {"note": "", "status": "review", "accepted": False, "flags": flags}
        else:
            try:
                res = review_claim(claim)
            except Exception as e:                                   # تاريخ غير مفهوم وغيره
                res = {"note": "", "status": "review", "accepted": False, "flags": [f"بيانات غير مفهومة: {e}"]}
            if flags:
                if res["status"] == "no_attachment":
                    res["note"] = ""
                res = {**res, "status": "review", "flags": flags + res["flags"]}
        out.append({"المطالبة": r.get("claim", ""), "الملاحظة": res["note"],
                    "مقبول": "نعم" if res["accepted"] else "لا",
                    "الحالة": {"ok": "✅ جاهز", "review": "🟡 مراجعة", "no_attachment": "⚪ بدون مرفق"}[res["status"]],
                    "سبب المراجعة": " ؛ ".join(res["flags"])})
    return pd.DataFrame(out)


def build_domain(domain_text: str, claim_field: str, claims_text: str, partial: bool = False) -> list:
    """فلتر Odoo: الفلتر اليدوي + أرقام المطالبات (تطابق تام، أو «يحتوي على» لو partial)."""
    import ast
    domain = list(ast.literal_eval(domain_text or "[]"))
    claims = [c.strip() for c in (claims_text or "").splitlines() if c.strip()]
    if claims:
        if partial:
            leaves = [(claim_field, "ilike", c) for c in claims]
            domain += ["|"] * (len(leaves) - 1) + leaves
        else:
            domain.append((claim_field, "in", claims))
    return domain
