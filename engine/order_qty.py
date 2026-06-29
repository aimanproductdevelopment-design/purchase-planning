"""
APPS — Order Quantity Engine
Order-up-to level → ปรับ MOQ → ปัด pack_size
"""
import math
from models.schema import SkuMaster, ForecastResult, ReorderResult, OrderQtyResult
from models.config_loader import AppConfig


def compute_order_qty(
    sku: SkuMaster,
    fc: ForecastResult,
    reorder: ReorderResult,
    lead_time_days: int,
    cfg: AppConfig,
) -> OrderQtyResult:

    if not reorder.need_to_order:
        return OrderQtyResult(
            sku_id=sku.sku_id, raw_qty=0, suggested_qty=0,
            target_level=0, reason="Inventory position ยังสูงกว่า ROP ไม่ต้องสั่ง",
        )

    review = cfg.review_period_days
    target = fc.daily * (lead_time_days + review) + (reorder.rop - fc.lead_time_forecast)
    raw = max(0.0, target - reorder.inventory_position)

    qty = raw
    reason_parts = [f"target={target:.0f}, inv_pos={reorder.inventory_position:.0f}"]

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
        target_level=round(target, 2),
        reason=" | ".join(reason_parts),
    )
