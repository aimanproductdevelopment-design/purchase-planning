"""
APPS — JST Adapter (XLSX mode)
อ่าน XLSX export จาก JST แล้วแปลงเป็น dataclass มาตรฐาน

หมายเหตุสำคัญ:
- sales_history จาก JST เป็น AGGREGATED ต่อ SKU (ไม่มีรายวัน)
  → แปลงเป็น avg_daily_qty สำหรับ forecast engine
- stock_on_hand และ pending_po อาจไม่มีสิทธิ์ export
  → คืน empty list แทน crash
"""
import unicodedata
import pandas as pd
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional

from adapters.base import ERPDataSource
from models.schema import SkuMaster, SalesRecord, StockRecord, PendingPO
from models.config_loader import AppConfig
from typing import Dict


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
            # แปลง status_label → status
            status_label = str(row.get("status_label", ""))
            supplier_val  = str(row.get("supplier_id", ""))
            # เลิกจำหน่าย สามารถอยู่ใน status_label หรือ supplier_id
            if "เลิกจ" in status_label or supplier_val == "เลิกจำหน่าย":
                status = "discontinued"
            else:
                status = str(row.get("status", "active"))
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
                status=status,
            ))
        _exclude_suppliers = {"อื่นๆ", "ไม่ระบุ", "อื่นๆ ไม่ระบุ"}
        return [s for s in skus
                if s.status != "discontinued"
                and not str(s.supplier_id).startswith("BG")
                and str(s.supplier_id).strip() not in _exclude_suppliers]

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

    def get_sales_windows(self) -> Dict[str, Dict[str, int]]:
        """
        อ่านไฟล์ sales_{n}d.xlsx ที่ exportแยกไว้ และคืนยอดขายจริงต่อ SKU
        Return: {sku_id: {s3: qty, s7: qty, s15: qty, s30: qty, s60: qty, s90: qty}}
        """
        data_dir = Path(self.cfg.jst_connection("sku_master")).parent
        windows  = [("s3", "3d"), ("s7", "7d"), ("s15", "15d"),
                    ("s30", "30d"), ("s60", "60d"), ("s90", "90d")]
        result: Dict[str, Dict[str, int]] = {}

        for key, suffix in windows:
            fpath = data_dir / f"sales_{suffix}.xlsx"
            if not fpath.exists():
                print(f"   ⚠️  {fpath.name} ไม่พบ (ข้าม)ไป")
                continue
            try:
                df = pd.read_excel(str(fpath), engine="openpyxl")
                # Normalize Thai column names
                def _norm(s):
                    import unicodedata
                    return unicodedata.normalize("NFC", s).replace("ํา", "ำ") if isinstance(s, str) else s
                df.columns = [_norm(c) for c in df.columns]

                # หา sku_id column
                sku_col = None
                for c in df.columns:
                    if "รหัสsku" in c.lower().replace(" ", "") or "รหัสสินค้า" in c or "sku_id" in c.lower():
                        sku_col = c
                        break
                # หา qty column (จำนวนสินค้าที่ขาย)
                qty_col = None
                for c in df.columns:
                    cl = c.lower()
                    if "จำนวนสินค้าที่ขาย" in c or "qty" == cl or "qty_sold" in cl:
                        qty_col = c
                        break
                # fallback: เอาคอลัมน์แรกที่มี "qty" หรือ "จำนวน" แต่ไม่ใช่ send_sku_qty
                if not qty_col:
                    for c in df.columns:
                        if ("qty" in c.lower() or "จำนวน" in c) and "send" not in c.lower() and "after" not in c.lower():
                            qty_col = c
                            break

                if not sku_col or not qty_col:
                    print(f"   ⚠️  {fpath.name}: ไม่พบคอลัมน์ sku/qty (columns: {list(df.columns)[:8]})")
                    continue

                for _, row in df.iterrows():
                    sku_id = str(row[sku_col]).strip()
                    if not sku_id or sku_id in ("nan", "None", ""):
                        continue
                    qty = int(float(row[qty_col] or 0))
                    if sku_id not in result:
                        result[sku_id] = {}
                    result[sku_id][key] = qty

                print(f"   ✅ {fpath.name}: โหลด {len(df)} SKU (คอลัมน์ qty = {qty_col!r})")
            except Exception as e:
                print(f"   ⚠️  อ่าน {fpath.name} ล้มเหลว: {e}")

        return result

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

        # Normalize column names: Thai nikhahit+sara-aa (จํา) → SARA AM (จำ)
        # NFC alone doesn't convert these; need explicit substitution
        def _norm_thai(s: str) -> str:
            # U+0E4D (nikhahit ํ) + U+0E32 (sara aa า)  →  U+0E33 (SARA AM ำ)
            return unicodedata.normalize("NFC", s).replace("ํา", "ำ")

        df.columns = [_norm_thai(c) if isinstance(c, str) else c for c in df.columns]
        col_map = self.cfg.field_map(table)
        if col_map:
            col_map = {_norm_thai(k): v for k, v in col_map.items()}
            df = df.rename(columns=col_map)
        return df
