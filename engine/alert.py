"""
APPS — Alert Engine
จัดระดับความเร่งด่วน 5 ระดับ จาก Days of Supply เทียบ Lead Time
"""
from models.schema import ReorderResult, SafetyStockResult, AlertResult
from models.config_loader import AppConfig


LEVELS = ["CRITICAL", "WARNING", "WATCH", "OK", "OVERSTOCK"]

EMOJI = {
    "CRITICAL":  "🔴",
    "WARNING":   "🟠",
    "WATCH":     "🟡",
    "OK":        "🟢",
    "OVERSTOCK": "⚫",
}


def compute_alert(
    sku_id: str,
    reorder: ReorderResult,
    ss: SafetyStockResult,
    lead_time_days: int,
    cfg: AppConfig,
) -> AlertResult:

    dos = reorder.days_of_supply
    available = reorder.available
    ss_val = ss.safety_stock
    on_order = reorder.on_order

    if available <= ss_val or dos <= lead_time_days:
        level = "CRITICAL"
        if on_order > 0:
            msg = (f"สต็อกวิกฤต ({dos:.0f} วัน < lead time {lead_time_days} วัน) "
                   f"— มี PO ค้างรับ {on_order:.0f} หน่วย ควรเร่ง supplier")
        else:
            msg = (f"สต็อกวิกฤต ({dos:.0f} วัน < lead time {lead_time_days} วัน) "
                   f"— ยังไม่มี PO ค้างรับ ต้องสั่งด่วน")

    elif reorder.need_to_order:
        level = "WARNING"
        msg = (f"ถึงจุดสั่งซื้อ (inv_pos={reorder.inventory_position:.0f} ≤ ROP={reorder.rop:.0f}) "
               f"— {dos:.0f} วัน")

    elif dos <= lead_time_days * cfg.watch_multiplier:
        level = "WATCH"
        msg = f"ใกล้ถึงจุดสั่งซื้อ — {dos:.0f} วัน (ติดตาม)"

    elif dos >= cfg.overstock_days:
        level = "OVERSTOCK"
        msg = f"สต็อกมากเกิน {dos:.0f} วัน — ชะลอการสั่ง"

    else:
        level = "OK"
        msg = f"ปกติ — {dos:.0f} วัน"

    return AlertResult(
        sku_id=sku_id,
        level=level,
        message=f"{EMOJI[level]} {msg}",
    )
