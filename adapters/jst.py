"""
APPS — JST Adapter (XLSX mode)
อ่าน XLSX export จาก JST แล้วแปลงเป็น dataclass มาตรฐาน

หมายเหตุสำคัญ:
- sales_history จาก JST เป็น AGGREGATED ต่อ SKU (ไม่มีรายวัน)
  → แปลงเป็น avg_daily_qty สำหรับ forecast engine
- stock_on_hand และ pending_po อาจไม่มีสิทธิ์ export
  → คืน empty list แทน crash
"""
import pandas as pd
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional

from adapters.base import ERPDataSource
from models.schema import SkuMaster, SalesRecord, StockRecord, PendingPO
from models.config_loader import AppConfig


class JSTAdapter(ERPDataSource):

    def __init__(self, cfg: AppConfig):
        self.cfg = cfg
        self.mode = cfg.jst_mode
        # ระยะเวลาประวัติขาย (วัน) — ใช้หาร qty_sold_total → avg_daily
        self._sales_days = cfg._raw.get("jst", {}).get("sales_history_days", 430)

    def get_skus(self) -> List[SkuMaster]:
        df = self._load("sku_master", "sku")
        skus = []
        for _, row in df.iterrows():
            lt = row.get("lead_time_days")
            uc = row.get("unit_cost", 0.0)
            try:
                uc = float(uc) if uc and not pd.isna(uc) else 0.0
            except (ValueError, TypeError):
                uc = 0.0
            skus.append(SkuMaster(
                sku_id=str(row["sku_id"]),
                sku_name=str(row.get("sku_name", "")),
                category=str(row.get("category", "")),
                supplier_id=str(row.get("supplier_id", "")),
                uom=str(row.get("uom", "pcs")),
                pack_size=int(row.get("pack_size", 1) or 1),
                moq=int(row.get("moq", 1) or 1),
                unit_cost=uc,
                lead_time_days=None if pd.isna(lt) else int(lt) if lt else None,
                status=str(row.get("status", "active")),
            ))
        return [s for s in skus if s.status != "discontinued"]

    def get_sales_history(self, start: date, end: date) -> List[SalesRecord]:
        """
        JST ส่ง aggregated total ต่อ SKU (ไม่มีรายวัน)
        → แปลงเป็น synthetic daily avg สำหรับ forecast engine
        โดยสร้าง SalesRecord หนึ่งรายการต่อวัน (qty = total / days)
        """
        df = self._load("sales_history", "sales")

        # qty_sold_total = รวมตลอด sales_history_days วัน
        if "qty_sold_total" not in df.columns:
            # ถ้า mapping ไม่ตรง ลอง fallback
            for col in df.columns:
                if "qty" in col.lower() or "จํานวน" in col or "จำนวน" in col:
                    df = df.rename(columns={col: "qty_sold_total"})
                    break

        records = []
        period_days = max(1, (end - start).days)

        for _, row in df.iterrows():
            sku_id = str(row.get("sku_id", ""))
            if not sku_id or sku_id == "nan":
                continue
            total_qty = float(row.get("qty_sold_total", 0) or 0)
            if total_qty <= 0:
                continue

            # avg daily qty
            avg_daily = total_qty / self._sales_days

            # สร้าง synthetic daily records ทุกวัน (ประหยัด: ทุก 7 วัน)
            # เพื่อให้ forecast engine มี time-series ใช้งานได้
            current = start
            while current <= end:
                records.append(SalesRecord(
                    sku_id=sku_id,
                    date=current,
                    qty_sold=avg_daily,
                    is_stockout=False,
                ))
                current += timedelta(days=7)  # weekly synthetic records

        return records

    def get_stock_on_hand(self) -> List[StockRecord]:
        """คืน empty list ถ้าไม่มีไฟล์ (เช่น ไม่มีสิทธิ์ export)"""
        try:
            df = self._load("stock_on_hand", "stock")
        except (FileNotFoundError, Exception) as e:
            print(f"   ⚠️  Stock on hand ไม่พร้อม: {e}")
            return []

        records = []
        for _, row in df.iterrows():
            oh = row.get("on_hand", 0)
            res = row.get("reserved", 0)
            try:
                records.append(StockRecord(
                    sku_id=str(row["sku_id"]),
                    on_hand=float(oh) if oh and not pd.isna(oh) else 0.0,
                    reserved=float(res) if res and not pd.isna(res) else 0.0,
                ))
            except Exception:
                continue
        return records

    def get_pending_po(self) -> List[PendingPO]:
        """คืน empty list ถ้าไม่มีไฟล์ (เช่น ไม่มีสิทธิ์ export)"""
        try:
            df = self._load("pending_po", "po")
        except (FileNotFoundError, Exception) as e:
            print(f"   ⚠️  Pending PO ไม่พร้อม: {e}")
            return []

        records = []
        for _, row in df.iterrows():
            eta_raw = row.get("eta_date")
            eta = None
            if eta_raw and not pd.isna(eta_raw):
                try:
                    eta = pd.to_datetime(eta_raw).date()
                except Exception:
                    pass
            try:
                records.append(PendingPO(
                    sku_id=str(row["sku_id"]),
                    po_number=str(row.get("po_number", "")),
                    qty_ordered=float(row.get("qty_ordered", 0) or 0),
                    qty_received=float(row.get("qty_received", 0) or 0),
                    eta_date=eta,
                ))
            except Exception:
                continue
        return records

    def _load(self, resource: str, table: str) -> pd.DataFrame:
        path = self.cfg.jst_connection(resource)
        p = Path(path)

        if not p.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์: {path}")

        ext = p.suffix.lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(path, engine="openpyxl")
        elif ext == ".csv":
            df = pd.read_csv(path, encoding="utf-8-sig")
        elif self.mode == "api":
            raise NotImplementedError("JST API mode: ยังไม่ได้ implement")
        elif self.mode == "db":
            raise NotImplementedError("JST DB mode: ยังไม่ได้ implement")
        else:
            raise ValueError(f"ไม่รู้จักนามสกุลไฟล์: {ext}")

        col_map = self.cfg.field_map(table)
        if col_map:
            df = df.rename(columns=col_map)
        return df
