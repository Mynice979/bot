import pandas as pd
import numpy as np
from typing import Dict, List

def merge_and_calculate(master_df: pd.DataFrame, trans_df: pd.DataFrame,
                        kode_col: str, target_col: str, realtime_col: str,
                        ach_col: str, type_col: str, modul_type: str) -> pd.DataFrame:
    """
    Gabungkan master dengan transaksi untuk modul tertentu,
    update REALTIME, hitung ACH.
    """
    # Filter master berdasarkan modul
    master_mod = master_df[master_df[type_col] == modul_type].copy()
    # Left join
    merged = master_mod.merge(trans_df, left_on=kode_col, right_on='Kode', how='left')
    # Update REALTIME
    merged[realtime_col] = merged['Qty']
    # Hitung ACH jika target numerik
    merged[target_col] = pd.to_numeric(merged[target_col], errors='coerce')
    merged[ach_col] = np.where(
        (merged[target_col] > 0) & (merged['Qty'].notna()),
        ((merged['Qty'] / merged[target_col]) * 100).round(1),
        np.nan
    )
    # Kosongkan REALTIME untuk toko yang tidak muncul di laporan
    merged.loc[merged['Qty'].isna(), realtime_col] = np.nan
    # Drop kolom bantu jika perlu, biarkan saja
    return merged

def aggregate_am(df: pd.DataFrame, am_col: str, as_col: str,
                 realtime_col: str) -> pd.DataFrame:
    """Agregasi per AM: jumlah toko, toko berdata, realtime max/min."""
    agg = df.groupby(am_col).agg(
        total_toko=(am_col, 'count'),
        toko_berdata=(realtime_col, lambda x: x.notna().sum()),
        realtime_max=(realtime_col, 'max'),
        realtime_min=(realtime_col, 'min')
    ).reset_index()
    agg['realtime_max'] = agg['realtime_max'].fillna(0)
    agg['realtime_min'] = agg['realtime_min'].fillna(0)
    return agg

def aggregate_as(df: pd.DataFrame, am_col: str, as_col: str,
                 realtime_col: str, kode_am: str) -> pd.DataFrame:
    """Agregasi per AS untuk AM tertentu."""
    sub = df[df[am_col] == kode_am]
    agg = sub.groupby(as_col).agg(
        total_toko=(as_col, 'count'),
        toko_berdata=(realtime_col, lambda x: x.notna().sum()),
        realtime_max=(realtime_col, 'max'),
        realtime_min=(realtime_col, 'min')
    ).reset_index()
    agg['realtime_max'] = agg['realtime_max'].fillna(0)
    agg['realtime_min'] = agg['realtime_min'].fillna(0)
    return agg

def top_bottom_toko(df: pd.DataFrame, realtime_col: str, nama_toko_col: str,
                    kode_toko_col: str, n: int = 10, ascending: bool = True):
    """
    Urutkan toko berdasarkan REALTIME, jika ascending=True mulai dari terendah.
    Kembalikan n data teratas setelah sort.
    """
    sorted_df = df.sort_values(by=realtime_col, ascending=ascending,
                               na_position='last')
    return sorted_df[[kode_toko_col, nama_toko_col, realtime_col]].head(n)