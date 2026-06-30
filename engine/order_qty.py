"""
APPS — Order Quantity Engine
คำนวณจากยอดขายจริง 3/7/15/30/60/90 วันที่ดึงจาก JST ERP
(Weighted average daily rate → × cover_days)

Logic:
  daily_rate = Σ (weight_n × sales_n / n)   สำหรับ n ที่มีข้อมูล
  order_qty  = daily_rate × cover_days       แล้วปัดขึ้นตาม MOQ / pack_size

Cover days ตามซัพลายเยอร์:
  - Import From CN  → 90 วัน  (3 เดือน — lead time 45 วัน)
  - ซัพลายเยอร์อื่น → 30 วัน  (1 เดือน)

Fallback: ถ้าไม่มีข้อมูล window เลย ใช้ fc.daily (forecast จาก 430 วัน)
"""
import math
from typing import Optional, Dict
from models.schema import SkuMaster, ForecastResult, ReorderResult, OrderQtyResult
from models.config_loader import AppConfig

_CN_SUPPLIER = "Import From CN"

# mapping: window key → (field ใน sales_window dict, จำนวนวัน)
_WINDOW_MAP = [
    ("w3",  "s3",   3),
    ("w7",  "s7",   7),
    ("w15", "s15", 15),
    ("w30", "s30", 30),
    ("w60", "s60", 60),
    ("w90", "s90", 90),
]


def _daily_from_windows(sw: Dict[str, int], weights: Dict[str, float]) -> tuple[float, str]:
    """
    คำนวณ weighted average daily rate จาก JST window data
    คืน (daily_rate, debug_string)

    - ถ้า window ใดไม่มีข้อมูล → ข้ามไป (ไม่นับน้ำหนัก)
    - normalize น้ำหนักตามช่องที่มีข้อมูลจริง
    """
    total_weight = 0.0
    weighted_sum = 0.0
    parts = []

    for wkey, skey, days in _WINDOW_MAP:
        w = weights.get(wkey, 0.0)
        if w <= 0:
            continue
        qty = sw.get(skey)
        if qty is None:           # ไม่มีข้อมูล window นี้เลย → ข้าม
            continue
        daily_n = qty / days
        weighted_sum += w * daily_n
        total_weight  += w
        parts.append(f"{skey}={qty}({daily_n:.2f}/วัน×{w})")

    if total_weight <= 0:
        return 0.0, "ไม่มีข้อมูล window"

    # normalize เผื่อบาง window ข้ามไป
    daily = weighted_sum / total_weight
    debug = " + ".join(parts) + f" → {daily:.3f}/วัน"
    return daily, debug


def compute_order_qty(
    sku: SkuMaster,
    fc: ForecastResult,
    reorder: ReorderResult,
    lead_time_days: int,
    cfg: AppConfig,
    sales_window: Optional[Dict[str, int]] = None,
) -> OrderQtyResult:

    if not reorder.need_to_order:
        return OrderQtyResult(
            sku_id=sku.sku_id, raw_qty=0, suggested_qty=0,
            target_level=0, reason="Inventory position ยังสูงกว่า ROP ไม่ต้องสั่ง",
        )

    # ── เลือก cover_days ตามซัพลายเยอร์ ────────────────────────────
    if sku.supplier_id == _CN_SUPPLIER:
        cover_days   = 90
        reason_label = "3เดือน(Import CN)"
    else:
        cover_days   = 30
        reason_label = "1เดือน"

    # ── คำนวณ daily rate จาก JST windows ────────────────────────────
    if sales_window:
        weights = cfg.order_window_weights
        daily, window_debug = _daily_from_windows(sales_window, weights)

        if daily > 0:
            source = f"JST[{window_debug}]"
        else:
            # ขายไม่ได้เลยใน 90 วัน → ใช้ fc.daily เป็น fallback
            daily  = fc.daily
            source = f"Forecast(ไม่มียอดขายใน window) {fc.daily:.3f}/วัน"
    else:
        # ยังไม่มีไฟล์ window → fallback
        daily  = fc.daily
        source = f"Forecast(ไม่มีไฟล์ window) {fc.daily:.3f}/วัน"

    raw = daily * cover_days
    reason_parts = [f"{reason_label}: {source} × {cover_days} วัน = {raw:.0f}"]

    qty = raw

    if qty < sku.moq:
        qty = float(sku.moq)
        reason_parts.append(f"ปรับขึ้นตาม MOQ={sku.moq}")

    if sku.pack_size > 1:
        qty = math.ceil(qty / sku.pack_size) * sku.pack_size
        reason_parts.append(f"ปัดขึ้นเป็น pack={sku.pack_size}")

    max_cap = cfg._raw.get("alert", {}).get("max_order_cap", {}).get(sku.sku_id)
    if max_cap and qty > max_cap:
        qty = max_cap
        reason_parts.append(f"ถูก cap ที่ {max_cap}")

    return OrderQtyResult(
        sku_id=sku.sku_id,
        raw_qty=round(raw, 2),
        suggested_qty=int(qty),
        target_level=round(daily * cover_days, 2),
        reason=" | ".join(reason_parts),
    )
