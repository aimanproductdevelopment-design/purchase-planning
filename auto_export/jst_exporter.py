"""
APPS — JST Auto Export (Playwright + Direct API)
Login ผ่าน Playwright แล้วใช้ requests เรียก API โดยตรง

วิธีใช้:
  python jst_exporter.py              # รันปกติ (browser ซ่อน)
  python jst_exporter.py --show       # แสดง browser (debug login)
  python jst_exporter.py --step sku  # ทดสอบเฉพาะ SKU export

ไฟล์ที่ได้: ../data/sku_master.xlsx, sales_history.xlsx, stock_on_hand.xlsx, pending_po.xlsx

API flow (ค้นพบจาก browser diagnostics):
  Login    : Playwright → extract cookies → requests.Session
  Export   : POST https://asia.jsterp.com/Common/FileExport/AsyncExport
             (ต้องมี exportJTableConfigs ไม่งั้น JST จะ error "导出列为空")
  Get URL  : POST https://asia.jsterp.com/Common/FileImport/DownloadExportFile
             → returns JSON {"Data": "https://oss-signed-url..."}
  Download : GET signed OSS URL → file binary
  Poll     : POST https://asia.jsterp.com/Common/FileImport/GetFileExports
             (ใช้เมื่อ IsExportSuccess=false — async queue)

หมายเหตุ:
  - Stock (SkuInventory) และ PO อาจไม่ได้รับอนุญาตในบัญชีนี้
  - ถ้า export ล้มเหลวจะ warn แทน crash เพื่อให้ pipeline ทำงานต่อได้
"""

import os
import sys
import time
import json
import argparse
import requests
from pathlib import Path
from datetime import datetime, timedelta, date
from dotenv import load_dotenv

# โหลด .env
load_dotenv(Path(__file__).parent / ".env")

JST_URL      = os.getenv("JST_URL", "https://asia-web.jsterp.com/")
API_BASE     = "https://asia.jsterp.com"
USERNAME     = os.getenv("JST_USERNAME", "")
PASSWORD     = os.getenv("JST_PASSWORD", "")
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "../data")).resolve()
HEADLESS     = os.getenv("HEADLESS", "false").lower() == "true"

# ─────────────────────────────────────────────────────────────
#  Column configs (exportJTableConfigs) — required by JST API
#  Captured from browser: /ReportCenter/SalesTheme/item
# ─────────────────────────────────────────────────────────────

SALES_ITEM_COLUMNS = [
    {"name": "sku_id",                     "label": "รหัสสินค้า",            "width": "120", "isShow": True,  "floor": 1},
    {"name": "item_id",                    "label": "รหัสรูปแบบ",            "width": "120", "isShow": True,  "floor": 1},
    {"name": "sku_name",                   "label": "ชื่อสินค้า",            "width": "120", "isShow": True,  "floor": 1},
    {"name": "property_value_string",      "label": "รูปแบบสินค้า",          "width": "200", "isShow": False, "floor": 1},
    {"name": "order_count",                "label": "จำนวนคำสั่งซื้อที่ขาย", "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "average_daily_sales",        "label": "ยอดขายรายวัน",          "width": "120", "isShow": False, "floor": 1},
    {"name": "qty",                        "label": "จํานวนสินค้าที่ขาย",    "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "send_sku_qty",               "label": "จำนวนสินค้าที่ส่งจริง", "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "amount",                     "label": "ราคาสินค้าทั้งหมด",    "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "f_amount",                   "label": "แพลตฟอร์มการชำระเงิน", "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "shop_free_amount",           "label": "ส่วนลดร้านค้า",         "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "platform_free_amount",       "label": "ส่วนลดแพลตฟอร์ม",      "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "freight_income",             "label": "ค่าจัดส่ง",             "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "freight_fee",                "label": "ค่าจัดส่ง (รายจ่าย)",  "width": "120", "isShow": False, "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "orther_amount",              "label": "ค่าธรรมเนียมอื่นๆ",    "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "pay_amount",                 "label": "จำนวนเงินที่ควรได้รับ", "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "paid_amount",                "label": "ยอดที่ชำระแล้ว",        "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "item_pay_amount",            "label": "ราคาสินค้า",            "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c"},
    {"name": "item_cost",                  "label": "ต้นทุนสินค้า",          "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "disabled": False, "sortable": "custom"},
    {"name": "gross_profit",               "label": "กำไรขั้นต้น",           "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "disabled": False},
    {"name": "gross_profit_rate",          "label": "อัตรากำไรขั้นต้น",     "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "disabled": False},
    {"name": "after_order_sku_qty",        "label": "จํานวนสินค้าตีกลับ",   "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "after_order_sku_return_qty", "label": "สินค้าตีกลับจริง",     "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "after_order_refund_amount",  "label": "จํานวนเงินที่คืน",      "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "after_order_amount",         "label": "ยอดเงินที่คืนจริง",     "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "sortable": "custom"},
    {"name": "after_order_item_cost",      "label": "ต้นทุนการส่งคืน",       "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "disabled": False, "sortable": "custom"},
    {"name": "after_order_refund_item_cost","label": "ต้นทุนส่งคืนตามจริง", "width": "120", "isShow": True,  "floor": 1, "customChosen": True, "template": "$c", "disabled": False, "sortable": "custom"},
]


# ─────────────────────────────────────────────────────────────
#  Public entry point
# ─────────────────────────────────────────────────────────────

def run(steps: list = None, headless: bool = None):
    """
    steps: รายการ ['sku','sales','stock','po'] หรือ None = รันทั้งหมด
    Stock และ PO อาจ fail เนื่องจากสิทธิ์บัญชี — จะ warn แทน crash
    """
    steps = steps or ["sku", "sales", "stock", "po"]
    use_headless = headless if headless is not None else HEADLESS
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 55)
    print("  APPS — JST Auto Export")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  URL: {JST_URL}")
    print("=" * 55)

    # ── Step 1: Login ─────────────────────────────────────────
    print("\n[1] 🔐 Login via Playwright...")
    session = _playwright_login(use_headless)
    print("   ✅ Login สำเร็จ — ได้ session cookies แล้ว")

    # ── Step 2: Export ────────────────────────────────────────
    date_range_days = 430
    start_date = (date.today() - timedelta(days=date_range_days)).strftime("%Y-%m-%d")
    end_date   = date.today().strftime("%Y-%m-%d")

    results = {}

    if "sku" in steps:
        print("\n[2] 📦 Export SKU Master...")
        path = _export_sku(session)
        results["sku"] = path
        print(f"   ✅ {path}")

    if "sales" in steps:
        print("\n[3] 📊 Export Sales History (per SKU, ~14 months)...")
        path = _export_sales(session, start_date, end_date)
        results["sales"] = path
        print(f"   ✅ {path}")

        print("\n[3b] 📊 Export ยอดขายแยกช่วง 3/7/15/30/60/90 วัน...")
        window_results = export_sales_windows(session)
        results["sales_windows"] = window_results

    if "stock" in steps:
        print("\n[4] 🏭 Export Stock on Hand...")
        try:
            path = _export_stock(session)
            results["stock"] = path
            print(f"   ✅ {path}")
        except Exception as e:
            print(f"   ⚠️  Stock export ล้มเหลว (ข้ามไป): {e}")
            results["stock"] = None

    if "po" in steps:
        print("\n[5] 📋 Export Pending PO...")
        try:
            path = _export_po(session)
            results["po"] = path
            print(f"   ✅ {path}")
        except Exception as e:
            print(f"   ⚠️  PO export ล้มเหลว (ข้ามไป): {e}")
            results["po"] = None

    print("\n✅ Export เสร็จสิ้น")
    return results


# ─────────────────────────────────────────────────────────────
#  Login
# ─────────────────────────────────────────────────────────────

def _playwright_login(headless: bool) -> requests.Session:
    """Login ผ่าน Playwright แล้ว return requests.Session พร้อม cookies"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ pip install playwright && playwright install chromium")
        sys.exit(1)

    if not USERNAME or not PASSWORD:
        raise ValueError("ยังไม่ได้ตั้งค่า JST_USERNAME / JST_PASSWORD ใน .env")

    pw_kwargs = {"headless": headless}
    # ใช้ pre-installed chromium ถ้ามี (remote environment)
    chromium_path = "/opt/pw-browsers/chromium"
    if Path(chromium_path).exists():
        pw_kwargs["executable_path"] = chromium_path

    with sync_playwright() as p:
        browser = p.chromium.launch(**pw_kwargs)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()

        page.goto(JST_URL, wait_until="networkidle", timeout=30_000)

        # หา input fields สำหรับ login
        page.wait_for_selector("input[type='text']", timeout=15_000)
        page.fill("input[type='text']", USERNAME)
        page.fill("input[type='password']", PASSWORD)

        # กดปุ่ม login — ใช้ evaluate เพราะ selector อาจไม่ตรง
        page.evaluate("document.querySelector('button').click()")

        # รอจน login สำเร็จ
        try:
            page.wait_for_url("**/", timeout=15_000)
        except Exception:
            pass
        page.wait_for_load_state("networkidle", timeout=15_000)

        # ตรวจสอบ login สำเร็จ
        current_url = page.url
        if "login" in current_url.lower():
            _save_screenshot(page, "login_failed")
            raise RuntimeError(f"Login ไม่สำเร็จ — URL ยังอยู่ที่ {current_url}")

        # ดึง cookies มาสร้าง requests.Session
        pw_cookies = context.cookies()
        browser.close()

    session = requests.Session()
    for c in pw_cookies:
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))

    # ตั้ง headers มาตรฐาน
    session.headers.update({
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Referer": JST_URL,
        "Origin": JST_URL.rstrip("/"),
    })

    return session


# ─────────────────────────────────────────────────────────────
#  Export functions
# ─────────────────────────────────────────────────────────────

def _export_sku(session: requests.Session) -> str:
    """Export SKU Master → sku_master.xlsx"""
    payload = {
        "exportWay": "SyncExport",
        "executeTime": None,
        "relativeId": "Item",
        "requestModel": {
            "column_config": {
                "elementId": "item",
                "key": "/Item/Goods/Item/ItemManager",
            },
            "source": "",
            "skuName": "",
            "skuShortName": "",
            "categoryId": 0,
            "categoryName": "",
            "skuIds": [],
            "itemIds": [],
            "startModified": None,
            "endModified": None,
            "openStatus": None,
            "isBundle": None,
        }
    }
    return _async_export_and_download(session, payload, "sku_master.xlsx")


def _export_sales(session: requests.Session, start_date: str, end_date: str,
                  filename: str = "sales_history.xlsx") -> str:
    """Export Sales History by SKU

    Payload structure confirmed by browser capture from /ReportCenter/SalesTheme/item
    Dates must be ISO format: "YYYY-MM-DDT00:00:00.000Z"
    exportJTableConfigs is REQUIRED (otherwise JST returns "导出列为空" error)
    """
    payload = {
        "exportWay": "SyncExport",
        "executeTime": None,
        "relativeId": "NewSaleOrderItem",
        "requestModel": {
            "column_config": {
                "elementId": "sku",
                "key": "/ReportCenter/SalesTheme/item",
            },
            "platform_ids": [],
            "shop_ids": [],
            "shop_group_ids": [],
            "warehouse_ids": [],
            "partner_ids": [],
            "salesman_ids": [],
            "statuses": [
                "WaitConfirm", "WaitFConfirm", "Question",
                "Delivering", "WaitOuterDeliver", "Delivered", "OuterDelivered",
            ],
            "label_type": "",
            "order_labels": [],
            "is_exclude_special": True,
            "is_combined_split": False,
            "is_exclude_sub_order": False,
            "time_type": "order_time",
            "start_date": f"{start_date}T00:00:00.000Z",
            "end_date":   f"{end_date}T23:59:59.000Z",
            "group_type": "sku",
            "sort_column": "",
            "sort_type": "",
            "include_sku_type": "",
            "is_exclude_gift": False,
            "is_only_gift": False,
            "mop_ids": [],
            "item_cost_type": "pay_time",
            "paraDrpId": False,
            "cps_ids": [],
            "buyer_phone": "",
            "buyer_name": "",
        },
        "exportJTableConfigs": SALES_ITEM_COLUMNS,
    }
    return _async_export_and_download(session, payload, filename)


def export_sales_windows(session: requests.Session) -> dict:
    """
    Export sales totals per SKU for each window: 3, 7, 15, 30, 60, 90 days.
    บันทึกเป็น ../data/sales_{n}d.xlsx แต่ละไฟล์
    คืน dict: {"3d": path, "7d": path, ...}
    """
    today = date.today()
    windows = [(3, "3d"), (7, "7d"), (15, "15d"), (30, "30d"), (60, "60d"), (90, "90d")]
    results = {}
    for days, key in windows:
        start = (today - timedelta(days=days)).strftime("%Y-%m-%d")
        end   = today.strftime("%Y-%m-%d")
        fname = f"sales_{key}.xlsx"
        print(f"   📅 Export ยอดขาย {days} วัน ({start} → {end})...")
        try:
            path = _export_sales(session, start, end, filename=fname)
            results[key] = path
            print(f"      ✅ {path}")
        except Exception as e:
            print(f"      ⚠️  sales_{key} ล้มเหลว: {e}")
            results[key] = None
    return results


def _export_stock(session: requests.Session) -> str:
    """Export Stock on Hand via /wms/Inventory/GetSkuInventorys
    ใช้ query API (ไม่ต้องการ WMS export permission)
    """
    import pandas as pd

    url       = f"{API_BASE}/wms/Inventory/GetSkuInventorys"
    all_rows  = []
    page      = 1
    page_size = 200

    print(f"   ดึง stock จาก API (ทีละ {page_size} รายการ)...")
    while True:
        payload = {
            "RequestModel": {
                "SkuIds": [], "ItemIds": [], "VirtualCategorys": [],
                "IsShowZero": True,
                "IsShowSale": False,
                "IsExcludeSpecialOrder": True,
                "BrandNames": [], "SupplierCodes": [], "Keywords": [],
            },
            "DataPage": {"pageSize": page_size, "pageIndex": page},
        }
        resp = session.post(url, json=payload, timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"GetSkuInventorys ตอบ {resp.status_code}: {resp.text[:200]}")

        data = resp.json()
        rows = data.get("Data", [])
        if not rows:
            break

        all_rows.extend(rows)
        dp      = data.get("DataPage", {})
        total   = dp.get("DataCount", 0)
        is_last = dp.get("IsLast", True)
        print(f"   หน้า {page}/{dp.get('PageCount','?')} — ได้ {len(all_rows)}/{total} รายการ")

        if is_last or len(all_rows) >= total:
            break
        page += 1

    if not all_rows:
        raise RuntimeError("GetSkuInventorys ไม่คืนข้อมูลเลย")

    df = pd.DataFrame(all_rows)
    df = df.rename(columns={
        "SkuId":          "รหัสSKU",
        "SkuName":        "ชื่อสินค้า",
        "AvailableQty":   "จำนวนคงเหลือ",
        "Qty":            "จำนวนทั้งหมด",
        "OrderLock":      "ล็อคคำสั่งซื้อ",
        "ShopLock":       "ล็อคร้านค้า",
        "PurchaseQty":    "สั่งซื้อค้างรับ",
        "SafetyStockQty": "Safety Stock",
        "WeekSales":      "ยอดขาย7วัน",
        "MonthSales":     "ยอดขาย30วัน",
        "SupplierName":   "ชื่อSupplier",
        "SupplierCode":   "รหัอSupplier",
        "BrandName":      "แบรนด์",
    })

    dest = DOWNLOAD_DIR / "stock_on_hand.xlsx"
    df.to_excel(str(dest), index=False, engine="openpyxl")
    return str(dest)


def _export_po(session: requests.Session) -> str:
    """Export Pending PO via /PurchaseOrder/Purchase/Query
    ดึง PO ที่ยังไม่ได้รับสินค้า (waitconfirm + confirmed + delivering)
    """
    import pandas as pd

    url      = f"{API_BASE}/PurchaseOrder/Purchase/Query"
    all_rows = []
    page     = 1
    page_size = 200

    # status ที่ยังค้างรับอยู่ (ไม่ได้รับครบ)
    pending_statuses = ["waitconfirm", "confirmed", "delivering"]

    print(f"   ดึง PO จาก API (status: {pending_statuses})...")
    while True:
        payload = {
            "PurchaseCreatedAfter": "", "PurchaseCreatedBefore": "",
            "ReceiptStatus": "", "SupplierCategory": "",
            "SupplierCode": "", "SupplierCodes": [],
            "GoodsQueryType": "SkuId", "GoodsQueryValue": [],
            "Remark": "", "GoodsType": "", "WarehouseIds": [],
            "OrderStatus": pending_statuses,
            "IsIn": True, "PurchaseId": None, "LogisticsId": "",
            "PurchaseIds": [], "ThirdOrderNos": [], "Creators": [],
            "LogisticsIds": [], "PlatformOrderStatus": "",
            "PlatformOrderId": "", "FreightType": "",
            "Labels": [], "Processes": [], "ProcessValue": None,
            "Value": None, "OrderByField": "", "OrderByIsAsc": False,
            "PageSize": page_size, "PageIndex": page, "CreateTypes": [],
        }
        resp = session.post(url, json=payload, timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"PO Query ตอบ {resp.status_code}: {resp.text[:200]}")

        data = resp.json()
        rows = data.get("Data", []) or []
        if not rows:
            break

        all_rows.extend(rows)
        dp      = data.get("DataPage", {})
        total   = dp.get("DataCount", 0)
        is_last = dp.get("IsLast", True)
        print(f"   หน้า {page}/{dp.get('PageCount',1)} — ได้ {len(all_rows)}/{total} PO")

        if is_last or len(all_rows) >= total:
            break
        page += 1

    # ไม่มี PO ค้างรับ — บันทึก empty file ให้ pipeline อ่านได้
    if not all_rows:
        print("   ไม่มี PO ค้างรับในขณะนี้")
        df = pd.DataFrame(columns=["เลขที่PO","รหัสSKU","จำนวนสั่งซื้อ","จำนวนรับแล้ว","วันที่รับสินค้า"])
        dest = DOWNLOAD_DIR / "pending_po.xlsx"
        df.to_excel(str(dest), index=False, engine="openpyxl")
        return str(dest)

    # flatten PO header + detail rows
    flat_rows = []
    for po in all_rows:
        po_no     = po.get("PurchaseId") or po.get("PurchaseNo") or ""
        supplier  = po.get("SupplierName") or po.get("SupplierCode") or ""
        status    = po.get("OrderStatus") or ""
        eta       = po.get("PlanArrivalDate") or po.get("ExpectDate") or ""
        details   = po.get("PurchaseItems") or po.get("Items") or po.get("GoodsList") or []
        if not details:
            flat_rows.append({
                "เลขที่PO": po_no, "ชื่อSupplier": supplier,
                "สถานะ": status, "วันที่รับสินค้า": eta,
                "รหัสSKU": "", "จำนวนสั่งซื้อ": 0, "จำนวนรับแล้ว": 0,
            })
        for item in details:
            flat_rows.append({
                "เลขที่PO":        po_no,
                "ชื่อSupplier":     supplier,
                "สถานะ":          status,
                "วันที่รับสินค้า":  eta,
                "รหัสSKU":         item.get("SkuId") or item.get("skuId") or "",
                "จำนวนสั่งซื้อ":  item.get("Qty") or item.get("qty") or 0,
                "จำนวนรับแล้ว":  item.get("ReceivedQty") or item.get("receivedQty") or 0,
            })

    df = pd.DataFrame(flat_rows)
    dest = DOWNLOAD_DIR / "pending_po.xlsx"
    df.to_excel(str(dest), index=False, engine="openpyxl")
    return str(dest)


# ─────────────────────────────────────────────────────────────
#  Core export + download helpers
# ─────────────────────────────────────────────────────────────

def _async_export_and_download(session: requests.Session, payload: dict, filename: str) -> str:
    """
    Flow:
    1. POST /Common/FileExport/AsyncExport → {Success, Data: {IsExportSuccess, OssId}}
    2. ถ้า IsExportSuccess=True → ได้ OssId ทันที (sync)
       ถ้า IsExportSuccess=False → async: poll GetFileExports จนไฟล์พร้อม
    3. POST /Common/FileImport/DownloadExportFile → {Data: "https://oss-signed-url..."}
    4. GET signed URL → ไฟล์จริง
    """
    export_url = f"{API_BASE}/Common/FileExport/AsyncExport"
    resp = session.post(export_url, json=payload, timeout=120)

    if resp.status_code != 200:
        raise RuntimeError(f"AsyncExport API ตอบ {resp.status_code}: {resp.text[:300]}")

    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"AsyncExport ตอบไม่ใช่ JSON: {resp.text[:200]}")

    if data.get("Success") is False or data.get("success") is False:
        msg = data.get("Message") or data.get("message") or str(data)
        raise RuntimeError(f"AsyncExport failed: {msg}")

    resp_data = data.get("Data") or data.get("data") or {}
    oss_id = (
        resp_data.get("OssId")
        or resp_data.get("FileOssId")
        or resp_data.get("fileOssId")
        or resp_data.get("ossId")
    )
    is_sync = resp_data.get("IsExportSuccess", resp_data.get("isExportSuccess", False))

    if oss_id and is_sync:
        return _download_oss_file(session, oss_id, filename)

    relative_id = payload.get("relativeId", "")
    print(f"   ⏳ Async export queued for {relative_id}, polling...")
    oss_id = _poll_get_file_exports(session, relative_id)
    return _download_oss_file(session, oss_id, filename)


def _poll_get_file_exports(
    session: requests.Session,
    relative_id: str,
    max_wait: int = 300,
) -> str:
    """Poll /Common/FileImport/GetFileExports จนได้ OssId ของ export ล่าสุด"""
    poll_url = f"{API_BASE}/Common/FileImport/GetFileExports"
    payload = {
        "RequestModel": {"RelativeIds": [relative_id]},
        "DataPage": {"pageSize": 5, "total": 0, "pageIndex": 1},
    }

    for attempt in range(max_wait // 10):
        time.sleep(10)
        try:
            resp = session.post(poll_url, json=payload, timeout=30)
            if resp.status_code != 200:
                continue
            pdata = resp.json()
            items = (pdata.get("Data") or {}).get("Items") or (pdata.get("Data") or {}).get("DataList") or []
            if not items:
                continue
            latest = items[0]
            status = latest.get("Status") or ""
            oss = latest.get("OssId") or latest.get("FileOssId") or latest.get("fileOssId")
            print(f"   ⏳ [{(attempt+1)*10}s] Status={status}, OssId={oss}")
            if oss:
                return oss
        except Exception as e:
            print(f"   ⚠️  Poll error: {e}")

    raise RuntimeError(f"Export {relative_id} timeout หลังรอ {max_wait}s")


def _download_oss_file(session: requests.Session, oss_id: str, filename: str) -> str:
    """
    1. POST /Common/FileImport/DownloadExportFile → {Data: "https://oss-signed-url..."}
    2. GET that signed URL → binary file content
    Note: API may return JSON without "Success" key — check if Data is a URL directly
    """
    url = f"{API_BASE}/Common/FileImport/DownloadExportFile"
    resp = session.post(url, json={"responseType": "blob", "FileOssId": oss_id}, timeout=60)

    if resp.status_code != 200:
        raise RuntimeError(f"DownloadExportFile ตอบ {resp.status_code}: {resp.text[:200]}")

    try:
        result = resp.json()
        signed_url = result.get("Data") or ""
        if isinstance(signed_url, str) and signed_url.startswith("http"):
            file_resp = requests.get(signed_url, timeout=120)
            if file_resp.status_code == 200 and len(file_resp.content) > 100:
                dest = DOWNLOAD_DIR / filename
                dest.write_bytes(file_resp.content)
                return str(dest)
            else:
                raise RuntimeError(f"OSS download failed: {file_resp.status_code}, url={signed_url[:80]}")
        if result.get("Success") is False:
            raise RuntimeError(f"DownloadExportFile failed: {result.get('Message', result)}")
    except (ValueError, KeyError):
        pass

    # Fallback: binary response
    if len(resp.content) > 100:
        dest = DOWNLOAD_DIR / filename
        dest.write_bytes(resp.content)
        return str(dest)

    raise RuntimeError(f"ไม่สามารถดาวน์โหลดไฟล์ {filename}: {resp.text[:200]}")


# ─────────────────────────────────────────────────────────────
#  Debug helpers
# ─────────────────────────────────────────────────────────────

def _save_screenshot(page, name: str):
    """บันทึก screenshot เพื่อ debug"""
    screenshot_dir = DOWNLOAD_DIR.parent / "screenshots"
    screenshot_dir.mkdir(exist_ok=True)
    path = screenshot_dir / f"{name}_{datetime.now().strftime('%H%M%S')}.png"
    try:
        page.screenshot(path=str(path))
        print(f"   📸 Screenshot: {path}")
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────
#  CLI Entry point
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="APPS — JST Auto Export")
    parser.add_argument("--show",  action="store_true", help="แสดง browser ขณะ login")
    parser.add_argument("--step",  choices=["sku", "sales", "stock", "po"],
                        action="append", dest="steps",
                        help="รันเฉพาะขั้นตอนที่ระบุ (ใช้ซ้ำได้ เช่น --step sku --step sales)")
    args = parser.parse_args()

    steps = args.steps if args.steps else None
    run(steps=steps, headless=not args.show)
