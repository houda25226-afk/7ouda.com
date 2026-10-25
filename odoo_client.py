"""عميل بسيط لسحب البيانات من Odoo عبر XML-RPC (Odoo 14 → 18).

بيشتغل بدون streamlit، فتقدر تجربه من أي سكريبت.
الاستخدام:
    c = OdooClient("https://ejada.sa", "DB_NAME", "user@company.com", "API_KEY")
    df = c.search_read_df("crm.lead", [("create_date", ">=", "2026-10-01")], ["name", "stage_id"])
"""
from __future__ import annotations
import base64
import xmlrpc.client
from urllib.parse import urlparse

import pandas as pd


class _TimeoutMixin:
    timeout = 60

    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = self.timeout
        return conn


class _HTTP(_TimeoutMixin, xmlrpc.client.Transport):
    pass


class _HTTPS(_TimeoutMixin, xmlrpc.client.SafeTransport):
    pass


class OdooClient:
    def __init__(self, url: str, db: str, user: str, api_key: str, timeout: int = 60):
        self.url, self.db, self.user, self._key = url.rstrip("/"), db, user, api_key
        transport = (_HTTPS if urlparse(self.url).scheme == "https" else _HTTP)()
        transport.timeout = timeout
        self._common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common", transport=transport, allow_none=True)
        self._models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object", transport=transport, allow_none=True)
        self.uid = self._common.authenticate(db, user, api_key, {})
        if not self.uid:
            raise PermissionError("فشل تسجيل الدخول لـ Odoo: راجع اسم قاعدة البيانات واليوزر والـ API key")

    # ---------- استدعاء عام ----------
    def call(self, model: str, method: str, *args, **kwargs):
        return self._models.execute_kw(self.db, self.uid, self._key, model, method, list(args), kwargs)

    def version(self) -> dict:
        return self._common.version()

    # ---------- استكشاف ----------
    def fields(self, model: str) -> pd.DataFrame:
        """كل حقول الموديل: الاسم التقني، الاسم الظاهر، النوع."""
        raw = self.call(model, "fields_get", attributes=["string", "type", "relation"])
        return (pd.DataFrame([{"field": k, **v} for k, v in raw.items()])
                .sort_values("field").reset_index(drop=True))

    def count(self, model: str, domain: list | None = None) -> int:
        return self.call(model, "search_count", domain or [])

    # ---------- سحب البيانات ----------
    def search_read(self, model, domain=None, fields=None, order=None, batch=500, limit=None):
        """بيسحب على دفعات (batch) عشان ما يتقفلش على الداتا الكبيرة."""
        rows, offset = [], 0
        while True:
            size = batch if limit is None else min(batch, limit - len(rows))
            if size <= 0:
                break
            chunk = self.call(model, "search_read", domain or [], fields=fields or [],
                              offset=offset, limit=size, order=order or "id asc")
            rows += chunk
            if len(chunk) < size:
                break
            offset += size
        return rows

    def search_read_df(self, model, domain=None, fields=None, **kw) -> pd.DataFrame:
        return flatten_relations(pd.DataFrame(self.search_read(model, domain, fields, **kw)))

    # ---------- المرفقات ----------
    def attachments(self, model: str, res_ids: list[int], with_data: bool = False) -> list[dict]:
        """مرفقات سجلات معينة. with_data=True بينزل محتوى الملف (base64) فاستخدمه بحذر."""
        fields = ["name", "mimetype", "file_size", "res_id", "create_date"] + (["datas"] if with_data else [])
        domain = [("res_model", "=", model), ("res_id", "in", res_ids)]
        rows = self.search_read("ir.attachment", domain, fields, batch=50)
        if with_data:
            for r in rows:
                r["content"] = base64.b64decode(r.pop("datas") or b"")
        return rows


def flatten_relations(df: pd.DataFrame) -> pd.DataFrame:
    """Odoo بيرجع حقول many2one كـ [id, "الاسم"]؛ هنا بنحولها للاسم، و False للقيم الفاضية."""
    for col in df.columns:
        s = df[col]
        if s.map(lambda v: isinstance(v, list) and len(v) == 2).any():
            df[col] = s.map(lambda v: v[1] if isinstance(v, list) and len(v) == 2 else None)
        elif s.dtype == object:
            df[col] = s.map(lambda v: None if v is False else v)
    return df
