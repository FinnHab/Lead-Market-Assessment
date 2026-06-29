"""
Price transmission - intra-EU vs world market.

Input  : ../inputs/price_transmission_panel.xlsx, sheet 01_Panel_for_PT
Output : ../outputs/05_price_transmission.xlsx

Method: winsorised log-log Pearson correlation between intra-EU producer
price and the corresponding world unit value, OLS elasticity (beta), a
2021-2023 sub-period check and a six-month lag sensitivity.

Coverage: 15 sectors, Jan 2010 - Dec 2023 (CN4-to-sector mapping in the
panel workbook, sheet 00_CN4_Scope_Audit).
"""

import os
import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import mstats

# -- Configuration ---
BASE_DIR    = "../inputs"
INPUT_FILE  = os.path.join(BASE_DIR, "price_transmission_panel.xlsx")
SHEET_NAME  = "01_Panel_for_PT"
OUTPUT_FILE = "../outputs/05_price_transmission.xlsx"

RECENT_PERIOD_START = 202101   # 3-year subperiod: Jan 2021–Dec 2023

# -- Helper functions ---
def winsorize_series(s, limits=(0.01, 0.01)):
    """Winsorize at 1%/99% to reduce influence of remaining extreme values."""
    s = pd.Series(s).reset_index(drop=True)
    if len(s) < 5:
        return s
    return pd.Series(mstats.winsorize(s, limits=limits))


def compute_pt(sector_df):
    """
    Compute all PT statistics for one sector.
    Input: DataFrame with columns uv_intraEU_eur_per_t, uv_world_eur_per_t,
           valid_for_PT, period (YYYYMM integers).
    Returns: dict of statistics.
    """
    valid = sector_df['valid_for_PT'].astype(bool)
    vi    = sector_df.loc[valid, 'uv_intraEU_eur_per_t'].reset_index(drop=True)
    vw    = sector_df.loc[valid, 'uv_world_eur_per_t'].reset_index(drop=True)
    per   = sector_df.loc[valid, 'period'].reset_index(drop=True)

    n = len(vi)
    if n < 10:
        return {'N_Obs': n, 'Corr': np.nan, 'R2': np.nan,
                'P_Value': np.nan, 'Beta_Elasticity': np.nan,
                'Corr_3Y': np.nan, 'Max_Lagged_Corr': np.nan, 'Best_Lag': np.nan}

    # Winsorize
    vi_w = winsorize_series(vi)
    vw_w = winsorize_series(vw)

    # Log-transform
    ln_i = np.log(vi_w.replace(0, np.nan))
    ln_w = np.log(vw_w.replace(0, np.nan))
    both = ln_i.notna() & ln_w.notna()
    n_both = int(both.sum())

    if n_both < 10:
        return {'N_Obs': n_both, 'Corr': np.nan, 'R2': np.nan,
                'P_Value': np.nan, 'Beta_Elasticity': np.nan,
                'Corr_3Y': np.nan, 'Max_Lagged_Corr': np.nan, 'Best_Lag': np.nan}

    # Pearson correlation + p-value
    corr, pval = stats.pearsonr(ln_i[both], ln_w[both])

    # OLS beta elasticity (log-log)
    slope, _, r, _, _ = stats.linregress(ln_w[both], ln_i[both])
    r2   = r ** 2
    beta = slope

    # 3-year subperiod
    mask3y = per.values >= RECENT_PERIOD_START
    if mask3y.sum() > 10:
        vi3 = winsorize_series(vi.values[mask3y])
        vw3 = winsorize_series(vw.values[mask3y])
        ln_i3 = np.log(pd.Series(vi3).replace(0, np.nan))
        ln_w3 = np.log(pd.Series(vw3).replace(0, np.nan))
        b3 = ln_i3.notna() & ln_w3.notna()
        corr3y = float(ln_i3[b3].corr(ln_w3[b3])) if b3.sum() > 5 else np.nan
    else:
        corr3y = np.nan

    # Lag robustness (lags 0–6 months)
    corrs = [float(ln_i.corr(ln_w.shift(lag))) for lag in range(7)]
    valid_corrs = [c for c in corrs if not np.isnan(c)]
    max_corr = max(valid_corrs) if valid_corrs else np.nan
    best_lag = corrs.index(max(corrs)) if valid_corrs else np.nan

    return {
        'N_Obs':           n_both,
        'Corr':            float(corr),
        'R2':              float(r2),
        'P_Value':         float(pval),
        'Beta_Elasticity': float(beta),
        'Corr_3Y':         corr3y,
        'Max_Lagged_Corr': max_corr,
        'Best_Lag':        best_lag,
    }


# -- Main ---
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Input not found: {INPUT_FILE}")

print(f"Loading: {INPUT_FILE}  ->  Sheet: {SHEET_NAME}")
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

sectors = sorted(df['sector'].unique())
print(f"Sectors found ({len(sectors)}): {sectors}\n")

# -- Verify sector list ---
EXPECTED_SECTORS = {
    'Meat & Meat Products', 'Barley', 'Beer', 'Dairy',
    'Fertilisers', 'Fruit & Vegetable (Processed)', 'Grain-Mill and Bakery Products',
    'Maize', 'Oilseed Oils', 'Oilseeds', 'Other Cereals and grains',
    'Potatoes', 'Prepared Animal Feeds', 'Refined Sugar', 'Wheat',
}
missing  = EXPECTED_SECTORS - set(sectors)
extra    = set(sectors) - EXPECTED_SECTORS
if missing:  print(f"WARNING — Missing sectors: {missing}")
if extra:    print(f"WARNING — Unexpected sectors: {extra}")
if not missing and not extra:
    print("OK Sector list matches expected 15 sectors.\n")

# -- Run PT per sector ---
results = []
for sector in sectors:
    ds   = df[df['sector'] == sector].sort_values('period').reset_index(drop=True)
    res  = compute_pt(ds)
    res['Sector'] = sector
    results.append(res)
    sig = ('***' if not np.isnan(res['P_Value']) and res['P_Value'] < 0.01 else
           '**'  if not np.isnan(res['P_Value']) and res['P_Value'] < 0.05 else
           '*'   if not np.isnan(res['P_Value']) and res['P_Value'] < 0.10 else 'n.s.')
    corr_s = f"{res['Corr']:.4f}" if not np.isnan(res.get('Corr', np.nan)) else '   NaN'
    r2_s   = f"{res['R2']:.4f}"   if not np.isnan(res.get('R2',   np.nan)) else '   NaN'
    pv_s   = f"{res['P_Value']:.4f}" if not np.isnan(res.get('P_Value', np.nan)) else '      NaN'
    print(f"  {sector:<42} N={res['N_Obs']:>3}  Corr={corr_s}  R²={r2_s}  p={pv_s}  {sig}")

df_res = pd.DataFrame(results)[[
    'Sector','N_Obs','Corr','R2','P_Value',
    'Beta_Elasticity','Corr_3Y','Max_Lagged_Corr','Best_Lag'
]].sort_values('Corr', ascending=False).reset_index(drop=True)

# -- Export ---
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    df_res.to_excel(writer, sheet_name='PT_Results', index=False)

# Apply unified visual styling for the replication package (see _styling.py)
import openpyxl as _openpyxl_for_styling
from _styling import apply_unified_styling
_wb_style = _openpyxl_for_styling.load_workbook(OUTPUT_FILE)
apply_unified_styling(_wb_style)
_wb_style.save(OUTPUT_FILE)


print(f"\nOK Saved: {OUTPUT_FILE}")
