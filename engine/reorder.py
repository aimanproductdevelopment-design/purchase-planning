"""
APPS — Reorder Point Engine
ROP = Forecast_LeadTime + SafetyStock
เทียบกับ Inventory Position (stock + on_order) ไม่ใช่แค่ stock
"""
from typing import List
from models.schema import StockRecord, PendingPO, ForecastResult, SafetyStockResult, ReorderResult


def compute_reorder(
    sku_id: str,
    stock: StockRecord,
    pending_pos: List[PendingPO],
    fc: ForecastResult,
    ss: SafetyStockResult,
    lead_time_days: int,
) -> ReorderResult:

    available = stock.available
    on_order = sum(po.qty_outstanding for po in pending_pos if po.sku_id == sku_id)
    inv_position = available + on_order - stock.backorder
    rop = fc.lead_time_forecast + ss.safety_stock
    need_to_order = inv_position <= rop

    if fc.daily > 0:
        dos = inv_position / fc.daily
    else:
        dos = float("inf")

    return ReorderResult(
        sku_id=sku_id,
        rop=round(rop, 2),
        inventory_position=round(inv_position, 2),
        available=round(available, 2),
        on_order=round(on_order, 2),
        need_to_order=need_to_order,
        days_of_supply=round(dos, 1) if dos != float("inf") else 9999,
    )
