"""
APPS — Forecast Engine
Moving Average (3 หน้าต่าง) + Seasonal Adjustment
"""
import math
import statistics
from collections import defaultdict
from datetime import date, timedelta
from typing import List, Dict, Optional

from models.schema import SkuMaster, SalesRecord, ForecastResult
from models.config_loader import AppConfig


def compute_forecast(
    sku: SkuMaster,
    sales: List[SalesRecord],
    today: date,
    lead_time_days: int,
    cfg: AppConfig,
) -> ForecastResult:

    flags = []
    records = sorted([s for s in sales if s.sku_id == sku.sku_id], key=lambda x: x.date)

    if len(records) == 0:
        flags.append("no_history")
        return _zero_forecast(sku.sku_id, lead_time_days, flags)

    if len(records) < 30:
        flags.append("new_sku")

    daily = _build_daily_series(records, cfg.stockout_handling)
    daily, capped = _winsorize(daily, cfg.outlier_sigma)
    if capped:
        flags.append("outlier_capped")

    w = cfg.forecast_windows
    add30 = _mean_last_n(daily, today, 30)
    add60 = _mean_last_n(daily, today, 60)
    add90 = _mean_last_n(daily, today, 90)

    weighted_add = _weighted_avg(
        [(add30, w["w30"]), (add60, w["w60"]), (add90, w["w90"])]
    )

    total_days = (records[-1].date - records[0].date).days + 1
    if total_days >= cfg.min_history_days_seasonal:
        seasonal = _compute_seasonal_index(daily)
    else:
        seasonal = {m: 1.0 for m in range(1, 13)}
        flags.append("low_history_seasonal")

    horizon = lead_time_days + cfg.review_period_days
    seasonal_factor = _horizon_seasonal_factor(today, horizon, seasonal)

    fc_daily = weighted_add * seasonal_factor
    fc_lt = fc_daily * lead_time_days

    vals = list(daily.values())
    std_dev = statistics.stdev(vals) if len(vals) >= 2 else 0.0

    return ForecastResult(
        sku_id=sku.sku_id,
        daily=round(fc_daily, 4),
        weighted_add=round(weighted_add, 4),
        seasonal_factor=round(seasonal_factor, 4),
        std_dev_daily=round(std_dev, 4),
        lead_time_forecast=round(fc_lt, 4),
        flags=flags,
    )


def _build_daily_series(records: List[SalesRecord], handling: str) -> Dict[date, float]:
    raw: Dict[date, float] = {}
    stockout_dates = set()

    for r in records:
        raw[r.date] = raw.get(r.date, 0.0) + r.qty_sold
        if r.is_stockout:
            stockout_dates.add(r.date)

    if handling == "exclude":
        return {d: v for d, v in raw.items() if d not in stockout_dates}
    elif handling == "impute":
        normal_vals = [v for d, v in raw.items() if d not in stockout_dates]
        avg_normal = statistics.mean(normal_vals) if normal_vals else 0.0
        result = dict(raw)
        for d in stockout_dates:
            result[d] = avg_normal
        return result
    return raw


def _winsorize(daily: Dict[date, float], sigma: float):
    vals = list(daily.values())
    if len(vals) < 4:
        return daily, False

    mu = statistics.mean(vals)
    sd = statistics.stdev(vals)
    cap = mu + sigma * sd

    capped = False
    result = {}
    for d, v in daily.items():
        if v > cap:
            result[d] = cap
            capped = True
        else:
            result[d] = v
    return result, capped


def _mean_last_n(daily: Dict[date, float], today: date, n: int) -> float:
    cutoff = today - timedelta(days=n)
    vals = [v for d, v in daily.items() if d > cutoff and d <= today]
    if not vals:
        vals = list(daily.values())
    return statistics.mean(vals) if vals else 0.0


def _weighted_avg(pairs) -> float:
    total_w = sum(w for _, w in pairs if _ is not None and _ > 0)
    if total_w == 0:
        return 0.0
    return sum(v * w for v, w in pairs if v is not None and v > 0) / total_w


def _compute_seasonal_index(daily: Dict[date, float]) -> Dict[int, float]:
    month_totals: Dict[int, List[float]] = defaultdict(list)
    for d, v in daily.items():
        month_totals[d.month].append(v)

    month_avg = {m: statistics.mean(vals) for m, vals in month_totals.items()}
    grand = statistics.mean(month_avg.values()) if month_avg else 1.0
    if grand == 0:
        return {m: 1.0 for m in range(1, 13)}

    raw_idx = {m: avg / grand for m, avg in month_avg.items()}
    result = {m: raw_idx.get(m, 1.0) for m in range(1, 13)}

    mean_idx = statistics.mean(result.values())
    if mean_idx > 0:
        result = {m: v / mean_idx for m, v in result.items()}
    return result


def _horizon_seasonal_factor(today: date, horizon_days: int, seasonal: Dict[int, float]) -> float:
    if horizon_days <= 0:
        return seasonal.get(today.month, 1.0)
    factors = []
    for i in range(horizon_days):
        d = today + timedelta(days=i + 1)
        factors.append(seasonal.get(d.month, 1.0))
    return statistics.mean(factors) if factors else 1.0


def _zero_forecast(sku_id: str, lead_time: int, flags: list) -> ForecastResult:
    return ForecastResult(
        sku_id=sku_id, daily=0.0, weighted_add=0.0, seasonal_factor=1.0,
        std_dev_daily=0.0, lead_time_forecast=0.0, flags=flags,
    )
