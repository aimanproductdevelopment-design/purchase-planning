"""
APPS — LINE Notify Integration
ส่งสรุปผลการวางแผนการสั่งซื้อทาง LINE หลัง pipeline เสร็จ

ตั้งค่า:
  เพิ่ม LINE_NOTIFY_TOKEN=xxxx ใน auto_export/.env หรือ .env

วิธีขอ Token:
  1. ไปที่ https://notify-bot.line.me/th/
  2. Login → "My page" → "Generate token"
  3. ตั้งชื่อ และเลือก Group หรือ Keep to myself

วิธีใช้จาก pipeline:
  from notify.line_notify import send_summary
  send_summary(results, output_path="output/purchasing_plan.xlsx")
"""
import os
import requests
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / "auto_export" / ".env")
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass

LINE_NOTIFY_TOKEN = os.getenv("LINE_NOTIFY_TOKEN", "")
LINE_NOTIFY_URL   = "https://notify-api.line.me/api/notify"


def send_line(message: str, token: str = "") -> bool:
    """ส่ง message ไป LINE Notify — คืน True ถ้าสำเร็จ"""
    tok = token or LINE_NOTIFY_TOKEN
    if not tok:
        print("   ⚠️  LINE_NOTIFY_TOKEN ไม่ได้ตั้งค่า — ข้าม LINE Notify")
        print("   → เพิ่ม LINE_NOTIFY_TOKEN=<token> ใน auto_export/.env")
        return False
    try:
        resp = requests.post(
            LINE_NOTIFY_URL,
            headers={"Authorization": f"Bearer {tok}"},
            data={"message": message},
            timeout=10,
        )
        if resp.status_code == 200:
            print("   ✅ ส่ง LINE Notify สำเร็จ")
            return True
        else:
            print(f"   ❌ LINE Notify error {resp.status_code}: {resp.text}")
            return False
    except Exception as e:
        print(f"   ❌ LINE Notify exception: {e}")
        return False


def send_summary(results, output_path: str = "") -> bool:
    """สร้างข้อความสรุปจาก results แล้วส่ง LINE Notify"""
    from collections import Counter

    level_count = Counter(r.alert.level for r in results)
    total       = len(results)

    # SKU ที่ต้องสั่งด่วน (CRITICAL + need_to_order)
    urgent_skus = [
        r for r in results
        if r.alert.level == "CRITICAL" and r.reorder.need_to_order
    ]
    urgent_skus.sort(key=lambda r: r.reorder.days_of_supply)

    # รวมมูลค่าสั่งซื้อ
    total_value = sum(
        r.order_qty.suggested_qty * r.sku.unit_cost
        for r in results if r.order_qty.suggested_qty > 0
    )

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    lines = [
        "",
        f"📦 APPS Purchase Planning — {now}",
        "─" * 28,
        f"🔴 CRITICAL : {level_count.get('CRITICAL', 0)} SKU",
        f"🟠 WARNING  : {level_count.get('WARNING', 0)} SKU",
        f"🟡 WATCH    : {level_count.get('WATCH', 0)} SKU",
        f"🟢 OK       : {level_count.get('OK', 0)} SKU",
        f"⚪ OVERSTOCK: {level_count.get('OVERSTOCK', 0)} SKU",
        f"📊 รวม: {total} SKU",
        f"💰 มูลค่าสั่งซื้อ: {total_value:,.0f} บาท",
    ]

    if urgent_skus:
        lines.append(f"\n⚠️ ต้องสั่งด่วน {len(urgent_skus)} รายการ (DOS ต่ำสุด):")
        for r in urgent_skus[:10]:
            dos = round(r.reorder.days_of_supply, 1)
            qty = int(r.order_qty.suggested_qty)
            lines.append(f"  • {r.sku.sku_id} DOS={dos}d สั่ง {qty} {r.sku.uom}")
        if len(urgent_skus) > 10:
            lines.append(f"  ... และอีก {len(urgent_skus) - 10} รายการ")

    if output_path:
        lines.append(f"\n📁 {Path(output_path).name}")

    message = "\n".join(lines)
    return send_line(message)


if __name__ == "__main__":
    # ทดสอบส่ง test message
    print("ทดสอบ LINE Notify...")
    ok = send_line("\n🤖 APPS — ทดสอบ LINE Notify ✅\nถ้าเห็นข้อความนี้แสดงว่าตั้งค่าถูกต้องแล้ว")
    if not ok:
        print("\nวิธีตั้งค่า:")
        print("  เพิ่ม LINE_NOTIFY_TOKEN=<token> ใน auto_export/.env")
        print("  ขอ token ได้ที่: https://notify-bot.line.me/th/")
