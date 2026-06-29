"""
APPS — Safety Stock Engine
รองรับ 2 โหมด: statistical (Z×σ√LT) และ days_of_cover
"""
import math
from models.schema import SkuMaster, ForecastResult, SafetyStockResult
from models.config_loader import AppConfig


def compute_safety_stock(
    sku: SkuMaster,
    fc: ForecastResult,
    lead_time_days: int,
    cfg: AppConfig,
) -> SafetyStockResult:

    mode = cfg.ss_mode(sku)

    if mode == "statistical" and fc.std_dev_daily > 0:
        ss = _statistical(sku, fc, lead_time_days, cfg)
    else:
        ss = _days_of_cover(sku, fc, cfg)
        mode = "days_of_cover"

    ss_val = max(ss.safety_stock, cfg.min_safety_stock)
    return SafetyStockResult(
        sku_id=sku.sku_id,
        safety_stock=round(ss_val, 2),
        mode_used=mode,
        z_value=ss.z_value,
        safety_days_used=ss.safety_days_used,
    )


def _statistical(sku, fc, lead_time_days, cfg) -> SafetyStockResult:
    z = cfg.service_level_z(sku)
    sigma_d = fc.std_dev_daily

    if cfg.lead_time_variable:
        sigma_lt_days = 0
        variance = lead_time_days * sigma_d**2 + fc.daily**2 * sigma_lt_days**2
        ss = z * math.sqrt(variance)
    else:
        ss = z * sigma_d * math.sqrt(lead_time_days)

    return SafetyStockResult(
        sku_id=sku.sku_id, safety_stock=ss,
        mode_used="statistical", z_value=z, safety_days_used=None,
    )


def _days_of_cover(sku, fc, cfg) -> SafetyStockResult:
    safety_days = cfg.safety_days(sku)
    ss = fc.daily * safety_days
    return SafetyStockResult(
        sku_id=sku.sku_id, safety_stock=ss,
        mode_used="days_of_cover", z_value=None, safety_days_used=safety_days,
    )
