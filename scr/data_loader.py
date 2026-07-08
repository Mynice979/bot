import pandas as pd
import re
import os
from typing import Optional, Tuple

def load_master_excel(file_path: str) -> pd.DataFrame:
    """Baca file master Excel, pastikan kolom minimal ada."""
    df = pd.read_excel(file_path, dtype=str)
    # Bersihkan spasi di nama kolom
    df.columns = df.columns.str.strip()
    return df

def parse_laporan_text(file_path: str, pattern: str) -> Tuple[Optional[pd.DataFrame], str, Optional[str]]:
    """
    Parse file teks laporan. Kembalikan (DataFrame hasil, modul, last_update).
    modul: 'FRIED CHICKEN' atau 'HOT SAUSAGE' atau None.
    """
    try:
        with open(file_path, 'r') as f:
            content = f.read()
    except Exception as e:
        return None, None, f"Gagal baca file: {e}"

    # Deteksi modul
    if 'FRIED CHICKEN' in content:
        modul = 'FRIED CHICKEN'
    elif 'HOT SAUSAGE' in content:
        modul = 'HOT SAUSAGE'
    else:
        return None, None, "Modul tidak dikenali (harus FRIED CHICKEN atau HOT SAUSAGE)"

    # Cari Last Update
    last_update = None
    match = re.search(r'Last Update:\s*(\d{2}-\d{2}-\d{4}\s\d{2}:\d{2}:\d{2})', content)
    if match:
        last_update = match.group(1)

    # Parse baris data
    rows = []
    for line in content.splitlines():
        m = re.search(pattern, line)
        if m:
            rows.append({
                'Kode': m.group('kode'),
                'Qty': int(m.group('qty')),
                'Rp': int(m.group('rp').replace(',', '')),
                'Stock': int(m.group('stock'))
            })
    if not rows:
        return None, modul, "Tidak ada data toko ditemukan di file."

    df = pd.DataFrame(rows)
    return df, modul, last_update