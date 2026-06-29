"""
APPS — ERP Data Source Interface
Engine เรียกผ่านนี้เท่านั้น ไม่รู้จัก JST โดยตรง
"""
from abc import ABC, abstractmethod
from datetime import date
from typing import List
from models.schema import SkuMaster, SalesRecord, StockRecord, PendingPO


class ERPDataSource(ABC):

    @abstractmethod
    def get_skus(self) -> List[SkuMaster]:
        """คืน SKU master ทั้งหมด"""

    @abstractmethod
    def get_sales_history(self, start: date, end: date) -> List[SalesRecord]:
        """คืนยอดขายรายวันในช่วงที่กำหนด"""

    @abstractmethod
    def get_stock_on_hand(self) -> List[StockRecord]:
        """คืนสต็อกคงเหลือ ณ ปัจจุบัน"""

    @abstractmethod
    def get_pending_po(self) -> List[PendingPO]:
        """คืน PO ค้างรับทั้งหมด"""
