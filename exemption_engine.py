"""محرك قواعد مراجعة حذف مسترد.

الموديل (الرؤية/OCR) يستخرج الحقول فقط، وكل المقارنات والحسابات هنا بكود ثابت.
المدخل: dict لكل مطالبة (انظر test_engine.py). التواريخ نصوص بالهجري أو الميلادي.
"""
from __future__ import annotations
import re
import difflib
from datetime import date
from hijridate import Hijri

_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "0123456789" * 2)
_EN2AR = {"A": "ا", "B": "ب", "J": "ح", "D": "د", "R": "ر", "S": "س", "X": "ص", "T": "ط",
          "E": "ع", "G": "ق", "K": "ك", "L": "ل", "Z": "م", "N": "ن", "H": "ه", "U": "و", "V": "ي"}
NAME_MATCH = 0.85  # أقل من كدا = مراجعة يدوية
GRACE_DAYS = 50   # أقصى فترة سماح بعد الحادث لبرنت الرخصة أو التأمين
ACCEPTED = ("valid", "grace")  # حالات السريان المقبولة


# ---------- أدوات ----------
def parse_date(v: str) -> date:
    """يقبل 2026-08-17 / 17/08/2026 / 1448/03/19 (السنة 13xx-14xx تعتبر هجري)."""
    parts = [int(p) for p in re.findall(r"\d+", str(v).translate(_DIGITS))]
    if len(parts) != 3:
        raise ValueError(f"تاريخ غير مفهوم: {v!r}")
    y, m, d = parts if parts[0] > 31 else (parts[2], parts[1], parts[0])
    if 1300 <= y <= 1500:
        g = Hijri(y, m, d).to_gregorian()
        return date(g.year, g.month, g.day)
    return date(y, m, d)


def norm_id(v) -> str:
    return re.sub(r"\D", "", str(v or "").translate(_DIGITS))


def _norm_ar(s: str) -> str:
    s = re.sub(r"[\u064B-\u0652\u0640]", "", s or "")
    s = re.sub("[أإآ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", s).strip().lower()


def name_score(a: str, b: str) -> float:
    a, b = _norm_ar(a), _norm_ar(b)
    if not a or not b:
        return 0.0
    flat = difflib.SequenceMatcher(None, a.replace(" ", ""), b.replace(" ", "")).ratio()
    tok = difflib.SequenceMatcher(None, " ".join(sorted(a.split())), " ".join(sorted(b.split()))).ratio()
    return max(flat, tok)


def _plate_parts(s: str):
    s = str(s or "").translate(_DIGITS).upper()
    digits = "".join(re.findall(r"\d", s))
    letters = [_EN2AR.get(c, c) for c in s if c in _EN2AR or "\u0621" <= c <= "\u064A"]
    return digits, sorted(_norm_ar("".join(letters)).replace(" ", ""))


def plates_match(a: str, b: str) -> bool:
    (da, la), (db, lb) = _plate_parts(a), _plate_parts(b)
    return bool(da) and da == db and la == lb


def _validity(acc: date, start: date | None, end: date | None):
    if start and acc < start:
        days = (start - acc).days
        return ("grace" if days <= GRACE_DAYS else "late"), days
    if end and acc > end:
        return "expired", (acc - end).days
    return "valid", 0


def _phrase(kind: str, days: int) -> str:
    return {"valid": "ساري وقت الحادث", "grace": f"بعد الحادث بـ {days} يوم",
            "late": f"بعد الحادث بـ {days} يوم (تجاوز فترة السماح {GRACE_DAYS} يوم)",
            "expired": f"منتهي قبل الحادث بـ {days} يوم"}[kind]


# ---------- الفروع ----------
def _letter(letter: dict, acc: dict, acc_date: date):
    problems = []
    if norm_id(letter.get("id_number")) != norm_id(acc.get("driver_id")):
        problems.append("هوية خطاب الإلغاء لا تطابق هوية السائق في المحضر")
    if letter.get("accident_number") and acc.get("case_number") and \
            letter["accident_number"].strip().upper() != acc["case_number"].strip().upper():
        problems.append("رقم الحادث في الخطاب لا يطابق المحضر")
    if letter.get("accident_date") and parse_date(letter["accident_date"]) != acc_date:
        problems.append("تاريخ الحادث في الخطاب لا يطابق المحضر")
    if problems:
        return None, problems
    what = "للهوية ورقم الحادث" if letter.get("accident_number") else "للهوية"
    note = (f"مرفق خطاب إلغاء مطالبة الاسترداد من {letter.get('company', '')} "
            f"بتاريخ {parse_date(letter['letter_date']):%d/%m/%Y}، مطابق {what}")
    if letter.get("amount"):
        note += f"، بمبلغ {letter['amount']}"
    return note, []


# الرخصة والتأمين يرجعوا (ملاحظة، مشاكل، حالة السريان: valid / grace / late / expired / None)
def _license(lic: dict, acc: dict, acc_date: date):
    if norm_id(lic.get("id_number")) != norm_id(acc.get("driver_id")):
        return None, ["هوية الرخصة لا تطابق هوية السائق في المحضر"], None
    kind, days = _validity(acc_date, parse_date(lic["issue_date"]), parse_date(lic["expiry_date"]))
    return f"مرفق برنت رخصة {_phrase(kind, days)}", [], kind


def _insurance(ins: dict, acc: dict, acc_date: date):
    kind, days = _validity(acc_date, parse_date(ins["start_date"]), parse_date(ins["end_date"]))
    ph = _phrase(kind, days)
    hid, hname = norm_id(ins.get("holder_id")), ins.get("holder_name", "")
    if hid and hid == norm_id(acc.get("driver_id")):
        return f"مرفق برنت تأمين باسم السائق {ph}", [], kind
    if hname and name_score(hname, acc.get("driver_name", "")) >= NAME_MATCH:
        return f"مرفق برنت تأمين باسم السائق {ph}", [], kind
    if hname and name_score(hname, acc.get("owner_name", "")) >= NAME_MATCH:
        return f"مرفق برنت تأمين باسم مالك المركبة {ph}", [], kind
    if ins.get("plate") and acc.get("plate") and plates_match(ins["plate"], acc["plate"]):
        return f"مرفق برنت تأمين {ph}، مطابق للوحة المركبة (المالك: {acc.get('owner_name', '')})", [], kind
    return None, ["برنت التأمين لا يطابق السائق ولا المالك ولا اللوحة"], None


def review_claim(claim: dict) -> dict:
    """الأولوية: خطاب إلغاء مطابق > برنت تأمين مقبول > رخصة مقبولة.
    المقبول = ساري وقت الحادث، أو صادر بعده في حدود فترة السماح (GRACE_DAYS).
    لو لا التأمين ولا الرخصة مقبولين، نكتب الاتنين (اللي موجود منهم).
    accepted = هل المستندات تدعم الإعفاء؟"""
    acc = claim["accident"]
    acc_date = parse_date(acc["date"])
    flags = []

    if claim.get("cancellation_letter"):
        note, probs = _letter(claim["cancellation_letter"], acc, acc_date)
        if note:
            return {"note": note, "status": "ok", "accepted": True, "flags": []}
        flags += probs

    docs = {}  # key -> (note, problems, kind)
    for key, fn in (("insurance_print", _insurance), ("license", _license)):
        if claim.get(key):
            docs[key] = fn(claim[key], acc, acc_date)

    if not docs and not claim.get("cancellation_letter"):
        return {"note": "لا يوجد مرفق", "status": "no_attachment", "accepted": False, "flags": []}

    valid = [k for k in ("insurance_print", "license") if k in docs and docs[k][2] in ACCEPTED]
    if valid:  # أول ساري حسب الأولوية يكفي، ومشاكل المستند التاني متتحسبش
        notes = [docs[valid[0]][0]]
    else:
        notes = [d[0] for d in docs.values() if d[0]]
        flags += [p for d in docs.values() for p in d[1]]
    return {"note": "؛ ".join(notes), "status": "review" if flags or not notes else "ok",
            "accepted": bool(valid), "flags": flags}
