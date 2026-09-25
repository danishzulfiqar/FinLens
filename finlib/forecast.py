"""Monte Carlo balance forecast: project the liquid balance forward with bands.

The data is monthly and short, so instead of fitting a fragile single curve we
bootstrap the historical monthly net cash-flow and simulate many possible futures,
reporting the P10 / P50 / P90 spread. This is honest about uncertainty rather than
implying false precision.
"""

from __future__ import annotations

import numpy as np
import pandas as pd


def _latest_balance(tx: pd.DataFrame) -> float:
    """Combined current balance: the most recent balance of each account, summed."""
    df = tx.sort_values("date")
    if "account" in df.columns and df["account"].nunique() > 1:
        return float(df.groupby("account")["balance"].last().sum())
    return float(df["balance"].iloc[-1])


def _monthly_net(tx: pd.DataFrame) -> pd.Series:
    """Net cash-flow (credits - debits) per calendar month, oldest -> newest."""
    m = tx[["date", "amount"]].copy()
    m["_period"] = pd.to_datetime(m["date"]).dt.to_period("M")
    return m.groupby("_period")["amount"].sum().sort_index()


def forecast_balance(
    tx: pd.DataFrame,
    horizon: int = 6,
    n_sims: int = 2000,
    seed: int | None = 42,
) -> dict:
    """Project the combined liquid balance ``horizon`` months ahead.

    Returns ``{"history", "forecast", "summary"}`` where ``forecast`` has columns
    ``month, p10, p50, p90, mean`` and ``summary`` carries headline numbers.
    """
    empty_hist = pd.DataFrame(columns=["month", "balance"])
    empty_fc = pd.DataFrame(columns=["month", "p10", "p50", "p90", "mean"])
    if tx is None or tx.empty or "balance" not in tx or "amount" not in tx:
        return {"history": empty_hist, "forecast": empty_fc, "summary": {}}

    nets = _monthly_net(tx)
    months = list(nets.index)
    current = _latest_balance(tx)

    # Rebuild internally-consistent month-end balances that terminate at `current`,
    # so the history line and the forecast start from the same point.
    hist_bal: dict = {}
    running = current
    for period in reversed(months):
        hist_bal[period] = running
        running -= float(nets[period])
    history = pd.DataFrame(
        {
            "month": [p.to_timestamp(how="end").normalize() for p in months],
            "balance": [hist_bal[p] for p in months],
        }
    )

    pool = nets.to_numpy(dtype=float)
    rng = np.random.default_rng(seed)
    draws = rng.choice(pool, size=(n_sims, horizon), replace=True)
    paths = current + np.cumsum(draws, axis=1)  # (n_sims, horizon)

    p10, p50, p90 = (np.percentile(paths, q, axis=0) for q in (10, 50, 90))
    mean = paths.mean(axis=0)

    last = months[-1]
    future = [(last + i).to_timestamp(how="end").normalize() for i in range(1, horizon + 1)]
    forecast = pd.DataFrame(
        {"month": future, "p10": p10, "p50": p50, "p90": p90, "mean": mean}
    )

    summary = {
        "current_balance": current,
        "horizon": horizon,
        "months_history": len(months),
        "monthly_net_mean": float(pool.mean()),
        "monthly_net_std": float(pool.std(ddof=1)) if len(pool) > 1 else 0.0,
        "proj_median_end": float(p50[-1]),
        "proj_p10_end": float(p10[-1]),
        "proj_p90_end": float(p90[-1]),
        "prob_negative": float((paths.min(axis=1) < 0).mean()),
        "reliable": len(months) >= 3,
    }
    return {"history": history, "forecast": forecast, "summary": summary}
