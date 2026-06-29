"""
APPS — Data Models (Dataclasses)
ชื่อ field มาตรฐานที่ engine ใช้ทั้งหมด
Adapter มีหน้าที่แปลงชื่อจาก JST ให้ตรงนี้
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Optional


# ─────────────────────────────────────────────
#  Input models (จาก JST ผ่าน Adapter)
# ─────────────────────────────────────────────

@dataclass
class SkuMaster:
    sku_id: str
    sku_name: str
    category: str
    supplier_id: str
    uom: str
    pack_size: int = 1
    moq: int = 1
    unit_cost: float = 0.0
    lead_time_days: Optional[int] = None
    status: str = "active"


@dataclass
class SalesRecord:
    sku_id: str
    date: date
    qty_sold: float
    is_stockout: bool = False


@dataclass
class StockRecord:
    sku_id: str
    on_hand: float
    reserved: float = 0.0
    backorder: float = 0.0

    @property
    def available(self) -> float:
        return max(0.0, self.on_hand - self.reserved)


@dataclass
class PendingPO:
    sku_id: str
    po_number: str
    qty_ordered: float
    qty_received: float = 0.0
    eta_date: Optional[date] = None

    @property
    def qty_outstanding(self) -> float:
        return max(0.0, self.qty_ordered - self.qty_received)


# ─────────────────────────────────────────────
#  Intermediate / output models (จาก Engine)
# ─────────────────────────────────────────────

@dataclass
class ForecastResult:
    sku_id: str
    daily: float
    weighted_add: float
    seasonal_factor: float
    std_dev_daily: float
    lead_time_forecast: float
    flags: list = field(default_factory=list)


@dataclass
class SafetyStockResult:
    sku_id: str
    safety_stock: float
    mode_used: str
    z_value: Optional[float] = None
    safety_days_used: Optional[float] = None


@dataclass
class ReorderResult:
    sku_id: str
    rop: float
    inventory_position: float
    available: float
    on_order: float
    need_to_order: bool
    days_of_supply: float


@dataclass
class OrderQtyResult:
    sku_id: str
    raw_qty: float
    suggested_qty: int
    target_level: float
    reason: str


@dataclass
class AlertResult:
    sku_id: str
    level: str
    message: str


@dataclass
class PlanningOutput:
    """ผลลัพธ์รวมต่อ 1 SKU — ใช้สร้าง recommendation table"""
    sku: SkuMaster
    forecast: ForecastResult
    safety_stock: SafetyStockResult
    reorder: ReorderResult
    order_qty: OrderQtyResult
    alert: AlertResult
    lead_time_used: int
    snapshot_at: str
