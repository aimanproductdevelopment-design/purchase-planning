"""
APPS — Orchestrator (run_planning.py)
รัน Planning Pipeline ครบสาย:
  Adapter → Data Prep → Forecast → Safety Stock → Reorder → Order Qty → Alert → Excel
"""
import sys
import os
import math
from datetime import date, datetime, timedelta
from collections import defaultdict
from typing import List, Dict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

from models.config_loader import AppConfig
from models.schema import (
    SkuMaster, SalesRecord, StockRecord, PendingPO, PlanningOutput
)
from adapters.jst import JSTAdapter
from engine.forecast import compute_forecast
from engine.safety_stock import compute_safety_stock
from engine.reorder import compute_reorder
from engine.order_qty import compute_order_qty
from engine.alert import compute_alert, LEVELS


def run(config_path: str = "config.yaml", output_path: str = "output/purchasing_plan.xlsx"):
    print("=" * 60)
    print("  APPS — Aiman Purchase Planning System")
    print(f"  รันวันที่: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    cfg = AppConfig(config_path)
    today = date.today()
    snapshot_at = datetime.now().isoformat(timespec="seconds")

    print("\n[1/5] ดึงข้อมูลจาก JST ERP...")
    adapter = JSTAdapter(cfg)

    skus: List[SkuMaster] = adapter.get_skus()
    print(f"  SKU: {len(skus)} รายการ")

    history_start = today - timedelta(days=430)
    sales: List[SalesRecord] = adapter.get_sales_history(history_start, today)
    print(f"  ยอดขาย: {len(sales):,} records ({history_start} → {today})")

    stocks: List[StockRecord] = adapter.get_stock_on_hand()
    stock_map: Dict[str, StockRecord] = {s.sku_id: s for s in stocks}
    print(f"  สต็อก: {len(stocks)} SKU")

    pos: List[PendingPO] = adapter.get_pending_po()
    po_map: Dict[str, List[PendingPO]] = defaultdict(list)
    for po in pos:
        po_map[po.sku_id].append(po)
    print(f"  PO ค้างรับ: {len(pos)} รายการ")

    print("\n[2/5] เตรียมข้อมูล...")
    sales_by_sku: Dict[str, List[SalesRecord]] = defaultdict(list)
    for s in sales:
        sales_by_sku[s.sku_id].append(s)

    print("\n[3/5] คำนวณ Forecast / Safety Stock / ROP / Order Qty...")
    results: List[PlanningOutput] = []

    for sku in skus:
        if sku.status == "discontinued":
            continue

        lead_time = cfg.lead_time(sku)
        fc = compute_forecast(sku, sales_by_sku.get(sku.sku_id, []), today, lead_time, cfg)
        ss = compute_safety_stock(sku, fc, lead_time, cfg)
        stock = stock_map.get(sku.sku_id) or StockRecord(sku_id=sku.sku_id, on_hand=0.0)
        reorder = compute_reorder(sku.sku_id, stock, po_map.get(sku.sku_id, []), fc, ss, lead_time)
        order = compute_order_qty(sku, fc, reorder, lead_time, cfg)
        alert = compute_alert(sku.sku_id, reorder, ss, lead_time, cfg)

        results.append(PlanningOutput(
            sku=sku, forecast=fc, safety_stock=ss,
            reorder=reorder, order_qty=order, alert=alert,
            lead_time_used=lead_time, snapshot_at=snapshot_at,
        ))

    print(f"  คำนวณเสร็จ {len(results)} SKU")

    print("\n[4/5] สรุป Alert:")
    level_count = defaultdict(int)
    for r in results:
        level_count[r.alert.level] += 1
    for lvl in LEVELS:
        if level_count[lvl]:
            print(f"  {lvl}: {level_count[lvl]} SKU")

    print(f"\n[5/5] บันทึก Excel → {output_path}")
    _export_excel(results, output_path, snapshot_at)
    print("\n✅ เสร็จสิ้น!")
    return results


def _export_excel(results: List[PlanningOutput], path: str, snapshot_at: str):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)

    rows_rec = []
    rows_alert = []
    rows_audit = []

    level_order = {l: i for i, l in enumerate(LEVELS)}
    sorted_results = sorted(
        results,
        key=lambda r: (level_order.get(r.alert.level, 99), r.reorder.days_of_supply)
    )

    for r in sorted_results:
        rows_rec.append({
            "Alert": r.alert.level,
            "SKU ID": r.sku.sku_id,
            "ชื่อสินค้า": r.sku.sku_name,
            "หมวด": r.sku.category,
            "Supplier": r.sku.supplier_id,
            "คงเหลือ (available)": r.reorder.available,
            "PO ค้างรับ (on_order)": r.reorder.on_order,
            "Inv. Position": r.reorder.inventory_position,
            "Forecast/วัน": r.forecast.daily,
            "Days of Supply": r.reorder.days_of_supply,
            "Safety Stock": r.safety_stock.safety_stock,
            "Reorder Point": r.reorder.rop,
            "ต้องสั่ง?": "✅ ใช่" if r.reorder.need_to_order else "—",
            "แนะนำสั่ง (หน่วย)": r.order_qty.suggested_qty if r.order_qty.suggested_qty > 0 else "—",
            "มูลค่าสั่ง (บาท)": (r.order_qty.suggested_qty * r.sku.unit_cost
                                   if r.order_qty.suggested_qty > 0 else 0),
            "หมายเหตุ": r.order_qty.reason,
        })

        if r.alert.level in ("CRITICAL", "WARNING"):
            rows_alert.append({
                "ระดับ": r.alert.level,
                "SKU ID": r.sku.sku_id,
                "ชื่อสินค้า": r.sku.sku_name,
                "Supplier": r.sku.supplier_id,
                "Lead Time": r.lead_time_used,
                "Days of Supply": r.reorder.days_of_supply,
                "แนะนำสั่ง": r.order_qty.suggested_qty,
                "ข้อความ": r.alert.message,
            })

        rows_audit.append({
            "SKU ID": r.sku.sku_id,
            "snapshot_at": snapshot_at,
            "lead_time_used": r.lead_time_used,
            "weighted_add": r.forecast.weighted_add,
            "seasonal_factor": r.forecast.seasonal_factor,
            "std_dev_daily": r.forecast.std_dev_daily,
            "forecast_daily": r.forecast.daily,
            "forecast_lead_time": r.forecast.lead_time_forecast,
            "ss_mode": r.safety_stock.mode_used,
            "safety_stock": r.safety_stock.safety_stock,
            "rop": r.reorder.rop,
            "available": r.reorder.available,
            "on_order": r.reorder.on_order,
            "inventory_position": r.reorder.inventory_position,
            "days_of_supply": r.reorder.days_of_supply,
            "suggested_qty": r.order_qty.suggested_qty,
            "alert_level": r.alert.level,
            "flags": ", ".join(r.forecast.flags),
        })

    df_rec = pd.DataFrame(rows_rec)
    df_alert = pd.DataFrame(rows_alert) if rows_alert else pd.DataFrame(
        columns=["ระดับ", "SKU ID", "ชื่อสินค้า", "Supplier", "ข้อความ"])
    df_audit = pd.DataFrame(rows_audit)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_rec.to_excel(writer, sheet_name="คำแนะนำสั่งซื้อ", index=False)
        df_alert.to_excel(writer, sheet_name="แจ้งเตือน", index=False)
        df_audit.to_excel(writer, sheet_name="Audit", index=False)

        for sheet_name, df in [("คำแนะนำสั่งซื้อ", df_rec),
                                 ("แจ้งเตือน", df_alert),
                                 ("Audit", df_audit)]:
            ws = writer.sheets[sheet_name]
            for col_idx, col in enumerate(df.columns, 1):
                max_len = max(len(str(col)), df[col].astype(str).map(len).max() if len(df) > 0 else 0)
                ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = min(max_len + 3, 40)


if __name__ == "__main__":
    cfg_path = sys.argv[1] if len(sys.argv) > 1 else "config.yaml"
    run(config_path=cfg_path)
