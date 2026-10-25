"""تجهيز ملف zip فيه مرفقات كل مطالبة، وكل ملف مسمّى باسم المطالبة.

لو المطالبة ليها مرفق واحد:   C-01-2026-4-416-130627-001.pdf
لو ليها أكتر من واحد:          C-01-2026-4-416-130627-001_1.jpg  ،  ..._2.pdf
وبيتضاف جوه الـ zip ملف _report.csv بيوضح كل مرفق اتسمّى إزاي ومين المطالبات اللي مالهاش مرفقات.
"""
from __future__ import annotations
import base64
import io
import mimetypes
import os
import re
import zipfile

import pandas as pd

_BAD = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def safe_name(s) -> str:
    return _BAD.sub("_", str(s)).strip(" .") or "بدون_اسم"


def _ext(att: dict) -> str:
    ext = os.path.splitext(att.get("name") or "")[1].lower()
    return ext or (mimetypes.guess_extension(att.get("mimetype") or "") or "")


def build_attachments_zip(client, leads: pd.DataFrame, name_col: str = "name", keywords=None,
                          model: str = "crm.lead", max_mb: int = 300, progress=None):
    """leads: DataFrame فيه عمود id وعمود اسم المطالبة. بيرجّع (bytes الـ zip, تقرير DataFrame)."""
    names = dict(zip(leads["id"], leads[name_col]))
    ids = list(names)
    kws = [k.strip().lower() for k in (keywords or []) if k and k.strip()]

    metas = []
    for i in range(0, len(ids), 200):
        metas += client.search_read(
            "ir.attachment", [("res_model", "=", model), ("res_id", "in", ids[i:i + 200])],
            ["name", "mimetype", "file_size", "res_id", "type"], order="res_id asc, id asc", batch=500)
    if kws:
        metas = [m for m in metas if any(k in (m.get("name") or "").lower() for k in kws)]

    links = [m for m in metas if m.get("type") == "url"]       # روابط مش ملفات: بتتسجل في التقرير بس
    metas = [m for m in metas if m.get("type") != "url"]

    by_lead: dict[int, list] = {}
    for m in metas:
        by_lead.setdefault(m["res_id"], []).append(m)

    used, plan = set(), []          # plan: (attachment, اسم الملف النهائي)
    for rid, atts in by_lead.items():
        base = safe_name(names[rid])
        for n, a in enumerate(atts, 1):
            stem = base if len(atts) == 1 else f"{base}_{n}"
            final, k = stem + _ext(a), 2
            while final.lower() in used:           # لو مطالبتين بنفس الاسم
                final, k = f"{stem}_dup{k}{_ext(a)}", k + 1
            used.add(final.lower())
            plan.append((a, final))

    rows, total_bytes, limit = [], 0, max_mb * 1024 * 1024
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i in range(0, len(plan), 5):                       # 5 ملفات في المرة عشان الذاكرة
            chunk = plan[i:i + 5]
            data = {r["id"]: r.get("datas") for r in client.call(
                "ir.attachment", "read", [a["id"] for a, _ in chunk], fields=["datas"])}
            for a, final in chunk:
                raw = data.get(a["id"])
                status = "تم"
                if not raw:
                    status = "ملف فاضي"
                elif total_bytes > limit:
                    status = f"اتجاوز الحد ({max_mb}MB)"
                else:
                    content = base64.b64decode(raw)
                    total_bytes += len(content)
                    z.writestr(final, content)
                rows.append({"claim": names[a["res_id"]], "original": a.get("name"),
                             "saved_as": final if status == "تم" else "", "status": status})
            if progress:
                progress(min(i + 5, len(plan)), len(plan))
        for m in links:
            rows.append({"claim": names[m["res_id"]], "original": m.get("name"), "saved_as": "",
                         "status": "رابط خارجي: مفيش ملف"})
        for rid, claim in names.items():
            if rid not in by_lead and not any(m["res_id"] == rid for m in links):
                rows.append({"claim": claim, "original": "", "saved_as": "", "status": "لا توجد مرفقات"})
        report = pd.DataFrame(rows)
        z.writestr("_report.csv", report.to_csv(index=False).encode("utf-8-sig"))
    return buf.getvalue(), report


def fetch_attachments_by_lead(client, lead_ids, model: str = "crm.lead", keywords=None,
                              max_mb: int = 300, progress=None) -> dict:
    """{id المطالبة: [(اسم الملف، المحتوى)]} — بيتخطى الروابط، ويقف عند الحد الأقصى للحجم."""
    ids, kws = list(lead_ids), [k.strip().lower() for k in (keywords or []) if k and k.strip()]
    metas = []
    for i in range(0, len(ids), 200):
        metas += client.search_read(
            "ir.attachment", [("res_model", "=", model), ("res_id", "in", ids[i:i + 200])],
            ["name", "res_id", "type"], order="res_id asc, id asc", batch=500)
    metas = [m for m in metas if m.get("type") != "url"
             and (not kws or any(k in (m.get("name") or "").lower() for k in kws))]
    out, total, limit = {}, 0, max_mb * 1024 * 1024
    for i in range(0, len(metas), 5):
        chunk = metas[i:i + 5]
        data = {r["id"]: r.get("datas") for r in client.call(
            "ir.attachment", "read", [m["id"] for m in chunk], fields=["datas"])}
        for m in chunk:
            raw = data.get(m["id"])
            if raw and total <= limit:
                content = base64.b64decode(raw)
                total += len(content)
                out.setdefault(m["res_id"], []).append((m.get("name") or "file", content))
        if progress:
            progress(min(i + 5, len(metas)), len(metas))
    return out


def zip_files_by_claim(files_by_claim: dict) -> bytes:
    """{اسم المطالبة: [(اسم الملف، المحتوى)]} → zip كل ملف فيه باسم المطالبة (نفس قواعد التسمية)."""
    used, buf = set(), io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for claim, files in files_by_claim.items():
            base = safe_name(claim)
            for n, (name, content) in enumerate(files, 1):
                ext = os.path.splitext(name)[1].lower()
                stem = base if len(files) == 1 else f"{base}_{n}"
                final, k = stem + ext, 2
                while final.lower() in used:
                    final, k = f"{stem}_dup{k}{ext}", k + 1
                used.add(final.lower())
                z.writestr(final, content)
    return buf.getvalue()
