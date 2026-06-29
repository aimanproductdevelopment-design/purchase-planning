"""
APPS — Config Loader
โหลด config.yaml และ resolve ค่า per-SKU override อัตโนมัติ
"""
import yaml
import math
from pathlib import Path

Z_TABLE = {
    0.90: 1.28,
    0.95: 1.65,
    0.975: 1.96,
    0.99: 2.33,
}


class AppConfig:
    def __init__(self, path: str = "config.yaml"):
        cfg_path = Path(path)
        with open(cfg_path) as f:
            self._raw = yaml.safe_load(f)
        self._base_dir = cfg_path.parent

    @property
    def jst_mode(self) -> str:
        return self._raw["jst"]["mode"]

    def jst_connection(self, resource: str) -> str:
        path = self._raw["jst"]["connection"][resource]
        return str(self._base_dir / path)

    def field_map(self, table: str) -> dict:
        return self._raw["jst"]["field_map"].get(table, {})

    @property
    def forecast_windows(self) -> dict:
        return self._raw["forecast"]["windows"]

    @property
    def stockout_handling(self) -> str:
        return self._raw["forecast"]["stockout_handling"]

    @property
    def outlier_sigma(self) -> float:
        return self._raw["forecast"].get("outlier_sigma", 3.0)

    @property
    def min_history_days_seasonal(self) -> int:
        return self._raw["forecast"].get("min_history_days_seasonal", 395)

    def lead_time(self, sku) -> int:
        override = self._raw.get("per_sku", {}).get(sku.sku_id, {})
        if "lead_time_days" in override:
            return override["lead_time_days"]
        if sku.lead_time_days is not None:
            return sku.lead_time_days
        by_sup = self._raw["lead_time"].get("by_supplier", {})
        if sku.supplier_id in by_sup:
            return by_sup[sku.supplier_id]
        return self._raw["lead_time"]["default_lead_time_days"]

    def ss_mode(self, sku) -> str:
        override = self._raw.get("per_sku", {}).get(sku.sku_id, {})
        return override.get("ss_mode", self._raw["safety_stock"]["global_mode"])

    def service_level_z(self, sku) -> float:
        override = self._raw.get("per_sku", {}).get(sku.sku_id, {})
        sl = override.get("service_level", self._raw["safety_stock"]["service_level"])
        closest = min(Z_TABLE.keys(), key=lambda k: abs(k - sl))
        return Z_TABLE[closest]

    def safety_days(self, sku) -> float:
        override = self._raw.get("per_sku", {}).get(sku.sku_id, {})
        return override.get("safety_days", self._raw["safety_stock"]["default_safety_days"])

    @property
    def min_safety_stock(self) -> float:
        return self._raw["safety_stock"].get("min_safety_stock", 0)

    @property
    def lead_time_variable(self) -> bool:
        return self._raw["safety_stock"].get("lead_time_variable", False)

    @property
    def review_period_days(self) -> int:
        return self._raw["order"]["review_period_days"]

    @property
    def watch_multiplier(self) -> float:
        return self._raw["alert"]["watch_multiplier"]

    @property
    def overstock_days(self) -> float:
        return self._raw["alert"]["overstock_days"]
