"""قراءة مستندات الإعفاءات (PDF / صور) واستخراج الحقول بدون أي API.

- PDF فيه نص داخلي (زي محاضر نجم): بنقرا النص مباشرة، دقة كاملة.
- غير كدا: OCR بـ Tesseract (عربي + إنجليزي) مع محاولات تدوير بسيطة لو الحقول ناقصة.
- أي حقل مش مؤكد بيتسيب فاضي، وبيتراجع في شاشة المراجعة بدل ما يتخمن.
"""
from __future__ import annotations
import io
import re

import pymupdf
import pytesseract
from PIL import Image, ImageOps

from exemption_engine import parse_date

_AR = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "0123456789" * 2)
ID_RE = r"(?<!\d)[12]\d{9}(?!\d)"
DMY = r"(\d{2}/\d{2}/\d{4})"
YMD = r"(\d{4}-\d{2}-\d{2})"
HIJRI = r"(?<!\d)(14\d{2}[/-]\d{2}[/-]\d{2})(?!\d)"
COMPANIES = {"wataniya": "الوطنية للتأمين", "tawuniya": "التعاونية للتأمين", "medgulf": "ميدغلف",
             "malath": "ملاذ للتأمين", "salama": "سلامة للتأمين", "solidarity": "التضامن", "walaa": "ولاء",
             "amana": "أمانة", "saico": "سايكو", "chubb": "تشب", "axa": "أكسا"}
DOC_LABELS = {"najm": "محضر نجم", "traffic": "محضر مرور", "letter": "خطاب إلغاء",
              "license": "برنت رخصة", "insurance": "برنت تأمين", "unknown": "غير معروف"}


# ---------- قراءة الملفات ----------
def to_images(content: bytes, filename: str, dpi: int = 200, max_pages: int = 4) -> list[Image.Image]:
    if filename.lower().endswith(".pdf") or content[:4] == b"%PDF":
        doc = pymupdf.open(stream=content, filetype="pdf")
        out = []
        for p in list(doc)[:max_pages]:
            pix = p.get_pixmap(dpi=dpi)
            out.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        return out
    return [ImageOps.exif_transpose(Image.open(io.BytesIO(content))).convert("RGB")]


def pdf_text_layer(content: bytes) -> str:
    try:
        return "".join(p.get_text() for p in pymupdf.open(stream=content, filetype="pdf"))
    except Exception:
        return ""


def _ocr_page(img: Image.Image, angles=(0,)) -> str:
    g = ImageOps.autocontrast(ImageOps.grayscale(img))
    if max(g.size) > 2600:
        g.thumbnail((2600, 2600))
    try:
        osd = pytesseract.image_to_osd(g, config="--psm 0 -l osd")
        rot = int(re.search(r"Rotate:\s*(\d+)", osd).group(1))
        if rot:
            g = g.rotate(-rot, expand=True, fillcolor=255)
    except Exception:
        pass
    texts = []
    for ang in angles:
        r = g.rotate(ang, expand=True, fillcolor=255, resample=Image.BICUBIC) if ang else g
        texts.append(pytesseract.image_to_string(r, lang="ara+eng", config="--psm 6"))
    return "\n".join(texts)


def ocr_images(imgs: list[Image.Image], deep: bool = False) -> str:
    return "\n".join(_ocr_page(i, (0, 4, -4) if deep else (0,)) for i in imgs).translate(_AR)


# ---------- التصنيف ----------
def classify(text: str, filename: str = "") -> str:
    t, f = text.lower(), filename.lower()
    if "liability determination" in t or "تحديد المسؤولية" in t:
        return "najm"
    if "cancellation letter" in t or "claim recovery" in t or "خطاب الإلغاء" in t or "خطاب الالغاء" in t:
        return "letter"
    if "driving license" in t or "رخصة سياقة" in t or (re.search(r"no\.?\s*\d{10}", t) and "dob" in t):
        return "license"
    if "نظام الحوادث المرورية" in t or "تقرير حادث مروري" in t:
        return "traffic"
    if any(k in t for k in ("بوليصة", "بوليصه", "policy no", "vehicle-insurance", "طرف ثالث")):
        return "insurance"
    for kind, keys in (("license", ("رخصة", "رخصه", "license", "licence")),
                       ("insurance", ("تأمين", "تامين", "insurance")),
                       ("letter", ("الغاء", "إلغاء", "cancel")),
                       ("najm", ("نجم", "najm")), ("traffic", ("مرور", "traffic"))):
        if any(k in f for k in keys):
            return kind
    return "unknown"


# ---------- المحللات ----------
def _ok(d) -> bool:
    """تاريخ حقيقي (الـ OCR أحيانا بيطلع تواريخ مستحيلة زي 2026-02-93)."""
    try:
        parse_date(d)
        return True
    except Exception:
        return False


def _first(pattern, text, flags=0, group=1):
    m = re.search(pattern, text, flags)
    return m.group(group if m.re.groups else 0) if m else None


def pick_case_no(text: str) -> str | None:
    """رقم الحالة/الحادث بيبدأ بحرفين + تاريخ الحادث (ddmmyy)؛ ده بيفرّقه عن رقم الوثيقة."""
    cands = re.findall(r"\b([A-Z]{2}\d{8,10})\b", text)
    stamps = {d[:2] + d[3:5] + d[8:10] for d in re.findall(DMY, text)}
    for c in cands:
        if c[2:8] in stamps:
            return c
    return None


def parse_najm(text: str, driver_id: str | None = None) -> dict:
    out = {"case_number": pick_case_no(text)}
    if "Accident Time" in text:
        out["accident_date"] = _first(DMY + r"\s+\d{2}:\d{2}:\d{2}", text)
    ids = list(dict.fromkeys(re.findall(ID_RE, text)))
    out["ids_in_report"] = ",".join(ids)
    if driver_id and driver_id in text:           # تاريخ انتهاء الرخصة المسجل في نجم للسائق
        tail = text[text.index(driver_id): text.index(driver_id) + 160]
        out["najm_license_expiry"] = _first(DMY, tail)
    return out


def parse_letter(text: str, driver_id=None) -> dict:
    dates = sorted({d for d in re.findall(DMY, text) if _ok(d) and 2000 <= int(d[-4:]) <= 2100}, key=parse_date)
    amount = _first(r"Amount\s*of\W*([\d,]+(?:\.\d+)?)", text, re.I)
    if not amount:
        cands = [a for a in re.findall(r"(?<![\d/])(\d+\.\d+)(?![\d/])", text)]
        amount = cands[-1] if cands else None
    company = _first(r"([A-Za-z]+)\s+Insurance\s+Company", text, re.I)
    company = COMPANIES.get((company or "").lower(), company)
    return {"letter_id": _first(ID_RE, text),
            "letter_accident_number": pick_case_no(text),
            "letter_accident_date": dates[0] if len(dates) >= 2 else None,
            "letter_date": dates[-1] if dates else None,
            "letter_company": company, "letter_amount": amount}


def parse_license(text: str, driver_id=None) -> dict:
    exp = _first(r"Exp\w*\W*" + DMY, text)
    dob = _first(r"DOB\W*" + DMY, text)
    lic_id = _first(r"No\.?\s*[:\-]?\s*(\d{10})", text) or _first(ID_RE, text)
    exp = exp if exp and _ok(exp) else None
    skip = {parse_date(d) for d in (exp, dob) if d and _ok(d)}
    cands = []
    for h in re.findall(HIJRI, text):
        try:
            g = parse_date(h)
        except Exception:
            continue
        if g not in skip and (not exp or g < parse_date(exp)):
            cands.append((g, h))
    issue = max(cands)[1].replace("-", "/") if cands else None
    return {"lic_id": lic_id, "lic_issue": issue, "lic_expiry": exp}


def parse_insurance(text: str, driver_id=None) -> dict:
    dates = sorted(set(re.findall(YMD, text)))
    dates = [d for d in dates if 2000 <= int(d[:4]) <= 2100 and _ok(d)]
    start = dates[0] if dates else None
    end = dates[-1] if len(dates) >= 2 else None
    if start and end and (parse_date(end) - parse_date(start)).days < 120:
        end = None          # الـ OCR غالبا غلط في تاريخ النهاية: نسيبه فاضي للمراجعة بدل ما نخمّن
    return {"ins_start": start, "ins_end": end,
            "ins_policy": _first(r"(P/[\d/\-]+)", text),
            "ins_holder_id": _first(ID_RE, text)}


PARSERS = {"najm": parse_najm, "letter": parse_letter, "license": parse_license, "insurance": parse_insurance}
REQUIRED = {"letter": ("letter_id", "letter_date"), "license": ("lic_id", "lic_issue", "lic_expiry"),
            "insurance": ("ins_start", "ins_end"), "najm": ("case_number", "accident_date")}


# ---------- مستند واحد / مطالبة كاملة ----------
def analyze_document(content: bytes, filename: str, driver_id: str | None = None) -> dict:
    is_pdf = filename.lower().endswith(".pdf") or content[:4] == b"%PDF"
    text, method, imgs = (pdf_text_layer(content) if is_pdf else ""), "text", None
    if len(text.strip()) < 200:
        imgs = to_images(content, filename)
        text, method = ocr_images(imgs), "ocr"
    kind = classify(text, filename)
    fields = PARSERS[kind](text, driver_id) if kind in PARSERS else {}
    missing = [k for k in REQUIRED.get(kind, ()) if not fields.get(k)]
    if method == "ocr" and kind not in ("najm", "traffic") and (missing or kind == "unknown"):      # محاولة تانية بميل بسيط
        text += "\n" + ocr_images(imgs, deep=True)
        kind = classify(text, filename) if kind == "unknown" else kind
        fields = PARSERS[kind](text, driver_id) if kind in PARSERS else {}
    return {"file": filename, "type": kind, "method": method, "fields": fields}


def analyze_claim(attachments: list[tuple[str, bytes]], driver_id: str | None = None) -> dict:
    """attachments: [(اسم الملف, محتواه)]. بيرجّع الحقول المدمجة + تفاصيل كل مستند."""
    docs, merged = [], {}
    for name, content in attachments:
        try:
            d = analyze_document(content, name, driver_id)
        except Exception as e:                                   # ملف تالف/غير مدعوم
            d = {"file": name, "type": "unknown", "method": "error", "fields": {}, "error": str(e)}
        docs.append(d)
        for k, v in d["fields"].items():
            if v and not merged.get(k):
                merged[k] = v
    return {"fields": merged, "docs": docs}
