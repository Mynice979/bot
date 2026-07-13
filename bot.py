import io
import os
import re
import pickle
import logging
from io import BytesIO

import pandas as pd
import numpy as np
import yaml
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from dotenv import load_dotenv

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)
from telegram.request import HTTPXRequest

# -------------------------------------------------------------------
# 1. Load environment & config
# -------------------------------------------------------------------
load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise ValueError("Token bot tidak ditemukan. Set di .env atau environment variable.")

with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

MASTER_COLS = config["master"]
PARSING_PATTERN = config["parsing"]["pattern"]
MODUL_LABEL = config["modul"]
TOP_N = config["ui"]["top_n"]
BOTTOM_N = config["ui"]["bottom_n"]

# -------------------------------------------------------------------
# 2. Helper functions
# -------------------------------------------------------------------
def load_master_excel(file_bytes: bytes) -> pd.DataFrame:
    """Baca file master Excel, cari sheet yang mengandung kolom 'Kode Toko'."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    for sheet in sheet_names:
        try:
            df_preview = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=50, dtype=str)
        except Exception:
            continue
        header_row = None
        for idx, row in df_preview.iterrows():
            for cell in row:
                if isinstance(cell, str):
                    cleaned = cell.replace(' ', '').lower()
                    if 'kodetoko' in cleaned:
                        header_row = idx
                        break
            if header_row is not None:
                break
        if header_row is not None:
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=header_row, dtype=str)
            df.columns = df.columns.str.strip()

            target_col = MASTER_COLS['target']
            if target_col in df.columns:
                df[target_col] = pd.to_numeric(df[target_col], errors='coerce').round().astype('Int64')

            required = [MASTER_COLS['kode_toko'], MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['type_col']]
            missing = [col for col in required if col not in df.columns]
            if not missing:
                return df

    raise ValueError(
        f"Tidak ditemukan sheet yang valid di file Excel.\n"
        f"Sheet tersedia: {', '.join(sheet_names)}\n"
        "Pastikan salah satu sheet memiliki header: Kode Toko, AM, AS, TYPE, dll."
    )

def parse_laporan_text(content: str):
    if 'FRIED CHICKEN' in content:
        modul = 'FRIED CHICKEN'
    elif 'HOT SAUSAGE' in content:
        modul = 'HOT SAUSAGE'
    else:
        return None, None, "Modul tidak dikenali (harus FRIED CHICKEN/HOT SAUSAGE)."

    last_update = None
    m = re.search(r'Last Update:\s*(\d{2}-\d{2}-\d{4}\s\d{2}:\d{2}:\d{2})', content)
    if m:
        last_update = m.group(1)

    rows = []
    for line in content.splitlines():
        m = re.search(PARSING_PATTERN, line)
        if m:
            rows.append({
                'Kode': m.group('kode'),
                'Qty': int(m.group('qty')),
                'Rp': int(m.group('rp').replace(',', '')),
                'Stock': int(m.group('stock'))
            })
    if not rows:
        return None, modul, "Tidak ada data toko."
    return pd.DataFrame(rows), modul, last_update

# -------------------------------------------------------------------
# 3. Merge & calculation
# -------------------------------------------------------------------
def merge_and_calc(master_df, trans_df, type_val):
    kd = MASTER_COLS['kode_toko']
    tp = MASTER_COLS['type_col']
    tg = MASTER_COLS['target']
    rt = MASTER_COLS['realtime']
    ac = MASTER_COLS['ach']

    sub = master_df[master_df[tp] == type_val].copy()
    if sub.empty:
        raise ValueError(f"Tidak ada toko dengan TYPE '{type_val}' di master.")

    merged = sub.merge(trans_df, left_on=kd, right_on='Kode', how='left')
    merged[rt] = merged['Qty']
    target_num = pd.to_numeric(merged[tg], errors='coerce')
    merged[ac] = np.where(
        (target_num > 0) & merged['Qty'].notna(),
        ((merged['Qty'] / target_num) * 100).round(1),
        np.nan
    )
    merged.loc[merged['Qty'].isna(), rt] = np.nan
    merged.drop(columns=['Kode', 'Qty', 'Rp', 'Stock'], inplace=True, errors='ignore')
    return merged

# -------------------------------------------------------------------
# 4. ASCII table formatter
# -------------------------------------------------------------------
def format_table(headers, rows, col_widths=None, title=None):
    if not rows and not headers:
        return ""
    if col_widths is None:
        col_widths = []
        for i, h in enumerate(headers):
            max_w = len(str(h))
            for row in rows:
                if i < len(row):
                    max_w = max(max_w, len(str(row[i])))
            col_widths.append(max_w + 2)
    def format_row(values, widths, sep='|'):
        cells = [f" {str(v):<{w-1}}" for v, w in zip(values, widths)]
        return sep.join(cells) + sep
    total_width = sum(col_widths) + len(headers) - 1
    line = "-" * total_width
    double_line = "=" * total_width
    lines = []
    if title:
        lines.append(title)
        lines.append(double_line)
    lines.append(format_row(headers, col_widths, sep='|'))
    lines.append(line)
    for row in rows:
        lines.append(format_row(row, col_widths, sep='|'))
    lines.append(line if title else double_line)
    return "\n".join(lines)

def df_summary(df, modul_name):
    am_col = MASTER_COLS['am']
    as_col = MASTER_COLS['as']
    rt_col = MASTER_COLS['realtime']
    nm_col = MASTER_COLS['nama_toko']
    kd_col = MASTER_COLS['kode_toko']

    grp_am = df.groupby(am_col).agg(
        total_toko=(am_col, 'count'),
        toko_berdata=(rt_col, lambda x: x.notna().sum()),
    ).reset_index()
    grp_am.columns = [am_col, 'Total Toko', 'Toko Ada Transaksi']

    df_as = df[df[rt_col].notna()].copy()
    top_as = bot_as = pd.DataFrame()
    if not df_as.empty:
        grp_as = df_as.groupby(as_col).agg(avg_realtime=(rt_col, 'mean')).reset_index()
        top_as = grp_as.nlargest(TOP_N, 'avg_realtime')
        bot_as = grp_as.nsmallest(BOTTOM_N, 'avg_realtime')

    df_toko = df[df[rt_col].notna()].sort_values(rt_col, ascending=False)
    top_toko = df_toko.head(TOP_N)
    bot_toko = df_toko.tail(BOTTOM_N)

    html = f"<b>Data: {modul_name}</b>\n<pre>"

    headers_am = ['AM', 'Total Toko', 'Toko Ada Transaksi']
    rows_am = [(r[am_col], int(r['Total Toko']), int(r['Toko Ada Transaksi'])) for _, r in grp_am.iterrows()]
    html += format_table(headers_am, rows_am, title="Area Manager")
    html += "\n\n"

    if not top_as.empty:
        headers_as = ['AS', 'Avg Realtime']
        rows_as = [(r[as_col], f"{r['avg_realtime']:.1f}") for _, r in top_as.iterrows()]
        html += format_table(headers_as, rows_as, title=f"Top {TOP_N} AS (rata-rata realtime)")
        html += "\n\n"

    if not bot_as.empty:
        headers_as = ['AS', 'Avg Realtime']
        rows_as = [(r[as_col], f"{r['avg_realtime']:.1f}") for _, r in bot_as.iterrows()]
        html += format_table(headers_as, rows_as, title=f"Bottom {BOTTOM_N} AS (rata-rata realtime)")
        html += "\n\n"

    if not top_toko.empty:
        headers_toko = ['Kode Toko', 'Nama Toko', 'Realtime']
        rows_top = [(r[kd_col], str(r[nm_col])[:20], f"{r[rt_col]:.0f}" if pd.notna(r[rt_col]) else "-") for _, r in top_toko.iterrows()]
        html += format_table(headers_toko, rows_top, title=f"Top {TOP_N} Toko (realtime tertinggi)")
        html += "\n\n"

    if not bot_toko.empty:
        headers_toko = ['Kode Toko', 'Nama Toko', 'Realtime']
        rows_bot = [(r[kd_col], str(r[nm_col])[:20], f"{r[rt_col]:.0f}" if pd.notna(r[rt_col]) else "-") for _, r in bot_toko.iterrows()]
        html += format_table(headers_toko, rows_bot, title=f"Bottom {BOTTOM_N} Toko (realtime terendah)")

    html += "</pre>"
    return html

# -------------------------------------------------------------------
# 5. JPEG table image (HD, sangat rapat) – untuk Top/Bottom
# -------------------------------------------------------------------
def create_table_image(df, title, last_update="", filename='temp.jpg', max_rows_per_page=100):
    n_rows = len(df)
    files = []

    df = df.reset_index(drop=True)
    df.insert(0, 'No', range(1, len(df) + 1))

    for page, start in enumerate(range(0, n_rows, max_rows_per_page)):
        end = min(start + max_rows_per_page, n_rows)
        page_df = df.iloc[start:end]
        page_n_rows, page_n_cols = page_df.shape

        if page_n_rows > 50:
            font_size = 6.2
            header_font_size = 6.8
            scale_y = 0.9
        elif page_n_rows > 25:
            font_size = 7.5
            header_font_size = 8.0
            scale_y = 1.0
        elif page_n_rows > 15:
            font_size = 8.5
            header_font_size = 9.0
            scale_y = 1.1
        else:
            font_size = 9.5
            header_font_size = 10.0
            scale_y = 1.2

        col_widths = []
        for col in page_df.columns:
            max_len = max(
                len(str(col)),
                page_df[col].astype(str).str.len().max() if len(page_df) > 0 else 0
            )
            col_widths.append(min(max_len, 15))
        total_width = sum(col_widths) * 0.12 + 1.0
        fig_width = max(9, min(total_width, 16))

        fig_height = max(3.0, page_n_rows * 0.28 + 2.0)

        fig = plt.figure(figsize=(fig_width, fig_height), facecolor='white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        title_lines = [title]
        if last_update:
            title_lines.append(f"Last Update: {last_update}")
        if n_rows > max_rows_per_page:
            total_pages = (n_rows - 1) // max_rows_per_page + 1
            title_lines.append(f"Hal {page+1}/{total_pages}")

        y_title = 0.99
        for i, line in enumerate(title_lines):
            if i == 0:
                ax.text(0.5, y_title, line, transform=fig.transFigure, ha='center',
                        fontsize=11, weight='bold', color='#1A3C5E')
            else:
                y_title -= 0.015
                ax.text(0.5, y_title, line, transform=fig.transFigure, ha='center',
                        fontsize=6.5, color='#5D6D7E', style='italic')

        table = ax.table(
            cellText=page_df.values,
            colLabels=page_df.columns,
            cellLoc='center',
            loc='center',
            bbox=[0, 0, 1, 1]
        )
        table.auto_set_font_size(False)
        table.set_fontsize(font_size)

        header_color = '#1A3C5E'
        for j in range(page_n_cols):
            cell = table[0, j]
            cell.set_facecolor(header_color)
            cell.set_text_props(color='white', weight='bold', fontsize=header_font_size)
            cell.set_edgecolor('#0F2A44')
            cell.set_linewidth(0.4)
            cell.set_height(0.04)

        row_colors = ['#FFFFFF', '#F4F6F7']
        highlight_green = '#E8F8F5'
        highlight_red = '#FDEDEC'

        realtime_col_idx = None
        for j, col_name in enumerate(page_df.columns):
            if 'REALTIME' in str(col_name).upper() and j > 0:
                realtime_col_idx = j
                break

        for i in range(1, page_n_rows + 1):
            is_top = is_bottom = False
            if realtime_col_idx is not None and page_n_rows >= 5:
                try:
                    vals = []
                    for idx in range(page_n_rows):
                        try:
                            v = str(page_df.iloc[idx, realtime_col_idx]).replace(',', '').replace('-', '')
                            vals.append(float(v) if v != '' else 0)
                        except:
                            vals.append(0)
                    sorted_idx = sorted(range(len(vals)), key=lambda x: vals[x], reverse=True)
                    top3 = sorted_idx[:3]
                    bottom3 = sorted_idx[-3:] if len(sorted_idx) >= 3 else []
                    if (i-1) in top3 and vals[i-1] > 0:
                        is_top = True
                    if (i-1) in bottom3 and vals[i-1] > 0 and (i-1) not in top3:
                        is_bottom = True
                except:
                    pass

            for j in range(page_n_cols):
                cell = table[i, j]
                cell.set_edgecolor('#BDC3C7')
                cell.set_linewidth(0.3)
                if is_top:
                    cell.set_facecolor(highlight_green)
                elif is_bottom:
                    cell.set_facecolor(highlight_red)
                else:
                    cell.set_facecolor(row_colors[(i-1) % 2])

        fig.text(0.5, 0.12, f"Total: {page_n_rows} toko", ha='center', fontsize=6.5, color='#7F8C8D')

        plt.tight_layout(rect=[0, 0.05, 1, 1.0], pad=0.05)
        page_filename = f"{filename.replace('.jpg','')}_p{page+1}.jpg"
        plt.savefig(page_filename, format='jpg', dpi=300, bbox_inches='tight',
                    pad_inches=0.03, facecolor='white', edgecolor='none',
                    pil_kwargs={'quality': 95, 'optimize': True})
        plt.close()
        files.append(page_filename)

    return files

# -------------------------------------------------------------------
# 6. JPEG detail AM/AS dengan banner navy + ringkasan kiri + tabel rapi
# -------------------------------------------------------------------
def create_detail_jpeg(df, title, last_update, summary, filename='temp.jpg', max_rows_per_page=80):
    """
    Buat JPEG detail AM/AS dengan layout presisi:
    - Banner navy, ringkasan, dan tabel mengisi penuh tanpa celah.
    - Tabel dipaksa mengisi area yang dihitung via bbox=[0,0,1,1].
    """
    n_rows = len(df)
    files = []

    df = df.reset_index(drop=True)
    df.insert(0, 'No', range(1, len(df) + 1))

    # Bobot lebar kolom relatif
    default_weights = {
        'No': 0.5, 'Kode Toko': 0.9, 'Nama Toko': 3.0, 'AS': 0.6,
        'Type': 1.6, 'Target': 0.9, 'Realtime': 1.0, '+/-': 0.8, 'ACH': 0.8,
    }
    if 'AM' in df.columns:
        default_weights['AM'] = 0.6
    weights = [default_weights.get(col, 1.0) for col in df.columns]
    total_weight = sum(weights)
    col_widths = [w / total_weight for w in weights]

    for page, start in enumerate(range(0, n_rows, max_rows_per_page)):
        end = min(start + max_rows_per_page, n_rows)
        page_df = df.iloc[start:end]
        page_n_rows = len(page_df)

        # Ukuran font dinamis
        if page_n_rows > 50:
            font_size = 6.5
            header_font_size = 7.0
        elif page_n_rows > 25:
            font_size = 8.0
            header_font_size = 8.5
        elif page_n_rows > 15:
            font_size = 9.0
            header_font_size = 9.5
        else:
            font_size = 10.0
            header_font_size = 10.5

        # Lebar figure
        col_widths_chars = []
        for col in page_df.columns:
            max_len = max(len(str(col)), page_df[col].astype(str).str.len().max() if len(page_df) > 0 else 0)
            col_widths_chars.append(min(max_len, 25))
        total_char_width = sum(col_widths_chars) * 0.14 + 2.0
        fig_width = max(11, min(total_char_width, 22))

        # ---------- UKURAN PRESISI (tanpa padding tak berguna) ----------
        banner_height_in = 0.7          # tinggi banner navy
        summary_height_in = 0.3         # tinggi area ringkasan (cukup 3 baris)
        table_row_height_in = 0.32      # tinggi per baris tabel
        top_pad_in = 0.0                # tidak ada padding ekstra
        bottom_pad_in = 0.08            # ruang footer kecil

        fig_height = (banner_height_in + summary_height_in + top_pad_in
                      + page_n_rows * table_row_height_in + bottom_pad_in)

        fig = plt.figure(figsize=(fig_width, fig_height), facecolor='white')
        left_margin = 0.035
        right_margin = 0.965

        banner_frac = banner_height_in / fig_height
        summary_frac = summary_height_in / fig_height

        # ===== 1. BANNER NAVY SOLID =====
        fig.add_artist(Rectangle(
            (0, 1 - banner_frac), 1, banner_frac,
            transform=fig.transFigure, facecolor='#1A3C5E', edgecolor='none', zorder=0
        ))
        fig.text(left_margin, 1 - banner_frac * 0.35, title,
                  ha='left', va='center', fontsize=14, weight='bold', color='white')
        fig.text(left_margin, 1 - banner_frac * 0.75, f"Last Update: {last_update}",
                  ha='left', va='center', fontsize=9, color='white')

        # ===== 2. BARIS RINGKASAN =====
        summary_top = 1 - banner_frac - 0.01
        line_gap = summary_frac / 3.2
        fig.text(left_margin, summary_top, f"Target : {summary['total_target']:.0f}",
                  ha='left', va='top', fontsize=9, weight='bold', color='#1A1A1A')
        fig.text(left_margin, summary_top - line_gap, f"Realtime : {summary['total_realtime']:.0f}",
                  ha='left', va='top', fontsize=9, weight='bold', color='#1A1A1A')
        fig.text(left_margin, summary_top - 2 * line_gap, f"ACH Total : {summary['ach_total']:.2f}",
                  ha='left', va='top', fontsize=9, weight='bold', color='#1A1A1A')
        fig.text(right_margin, summary_top, f"Jumlah Toko : {summary['jumlah_toko']}",
                  ha='right', va='top', fontsize=9, weight='bold', color='#1A1A1A')

        # ===== 3. TABEL (mengisi penuh area yang dihitung) =====
        table_top = 1 - banner_frac - summary_frac - 0.02
        table_bottom = bottom_pad_in / fig_height * 0.4
        table_height = table_top - table_bottom

        ax = fig.add_subplot(111)
        ax.axis('off')
        ax.set_position([0.02, table_bottom, 0.96, table_height])

        table = ax.table(
            cellText=page_df.values,
            colLabels=page_df.columns,
            cellLoc='center',
            loc='center',
            colWidths=col_widths,
            bbox=[0, 0, 1, 1]          # <-- KUNCI: tabel mengisi penuh area ax
        )
        table.auto_set_font_size(False)
        table.set_fontsize(font_size)
        # table.scale(1, 1.35)   <-- DIHAPUS, karena bbox sudah mengatur tinggi

        n_cols = len(page_df.columns)

        # Header tabel: abu-abu terang, teks hitam tebal
        for j in range(n_cols):
            cell = table[0, j]
            cell.set_facecolor('#EAEAEA')
            cell.set_text_props(color='#1A1A1A', weight='bold', fontsize=header_font_size)
            cell.set_edgecolor('#B0B0B0')
            cell.set_linewidth(0.8)

        # Baris data: selang-seling putih & biru muda
        row_colors = ['#FFFFFF', '#F0F4FA']
        for i in range(1, page_n_rows + 1):
            for j in range(n_cols):
                cell = table[i, j]
                cell.set_facecolor(row_colors[(i - 1) % 2])
                cell.set_edgecolor('#D0D5DD')
                cell.set_linewidth(0.4)
                cell.set_text_props(color='#2C3E50')

        # Garis tebal navy di bawah header
        for j in range(n_cols):
            table[0, j].set_edgecolor('#1A3C5E')
            table[0, j].set_linewidth(1.2)

        page_filename = f"{filename.replace('.jpg', '')}_p{page + 1}.jpg"
        plt.savefig(page_filename, format='jpg', dpi=300,
                    facecolor='white', edgecolor='none',
                    pil_kwargs={'quality': 95, 'optimize': True})
        plt.close()
        files.append(page_filename)

    return files
# -------------------------------------------------------------------
# 7. State & handlers
# -------------------------------------------------------------------
WAITING_MASTER_FILE = 0
WAITING_SOSIS = 1
WAITING_AYAM = 2

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    master_path = f"data/masters/{chat_id}.pkl"

    if ('master' not in context.user_data or context.user_data['master'] is None) and os.path.exists(master_path):
        try:
            with open(master_path, 'rb') as f:
                context.user_data['master'] = pickle.load(f)
        except Exception as e:
            logging.warning(f"Gagal memuat master: {e}")

    if 'master' not in context.user_data or context.user_data['master'] is None:
        await update.message.reply_text(
            "Anda belum mengunggah file struktur master.\n"
            "Silakan gunakan perintah /upload_struktur_master terlebih dahulu."
        )
        return ConversationHandler.END

    context.user_data['sosis_list'] = []
    context.user_data['ayam_list'] = []

    await update.message.reply_text(
        "Kirim data HOT SAUSAGE (copy‑paste teks) atau ketik *skip* jika tidak ada.\n"
        "Anda dapat mengirim beberapa kali. Setelah selesai, ketik **done**."
    )
    return WAITING_SOSIS

async def upload_master_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Silakan kirim file master toko (.xlsx).")
    return WAITING_MASTER_FILE

async def receive_master_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.endswith('.xlsx'):
        await update.message.reply_text("File harus .xlsx. Kirim ulang atau /cancel.")
        return WAITING_MASTER_FILE

    file = await context.bot.get_file(doc.file_id)
    fb = await file.download_as_bytearray()
    try:
        df = load_master_excel(fb)
        context.user_data['master'] = df

        chat_id = update.effective_chat.id
        os.makedirs("data/masters", exist_ok=True)
        master_path = f"data/masters/{chat_id}.pkl"
        with open(master_path, 'wb') as f:
            pickle.dump(df, f)

        await update.message.reply_text(
            f"Master berhasil disimpan ({len(df)} toko).\n"
            "Sekarang Anda dapat menggunakan /start untuk memulai input data penjualan."
        )
    except Exception as e:
        await update.message.reply_text(f"Gagal membaca master: {e}\nKirim ulang file yang benar atau /cancel.")
        return WAITING_MASTER_FILE

    return ConversationHandler.END

async def cancel_master_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Upload master dibatalkan.")
    return ConversationHandler.END

async def receive_sosis(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == 'skip':
        context.user_data['sosis_df'] = None
        context.user_data['sosis_list'] = []
        await update.message.reply_text("Data Sosis dilewati.\nSekarang kirim data FRIED CHICKEN (copy‑paste teks) atau ketik *skip*.\nAnda dapat mengirim beberapa kali, ketik **done** jika selesai.")
        return WAITING_AYAM

    if text.lower() == 'done':
        sosis_list = context.user_data.get('sosis_list', [])
        if sosis_list:
            context.user_data['sosis_df'] = pd.concat(sosis_list, ignore_index=True)
        else:
            context.user_data['sosis_df'] = None
        await update.message.reply_text(
            "Data Sosis selesai.\nSekarang kirim data FRIED CHICKEN (copy‑paste teks) atau ketik *skip*.\nAnda dapat mengirim beberapa kali, ketik **done** jika selesai."
        )
        return WAITING_AYAM

    df_trans, modul, info = parse_laporan_text(text)
    if df_trans is None or modul != 'HOT SAUSAGE':
        await update.message.reply_text(f"Gagal: {info}\nPastikan teks mengandung 'HOT SAUSAGE'. Coba lagi atau ketik *skip* / *done*.")
        return WAITING_SOSIS

    context.user_data.setdefault('sosis_list', []).append(df_trans)
    if info and isinstance(info, str):
        context.user_data['sosis_last'] = info

    cnt = len(context.user_data['sosis_list'])
    await update.message.reply_text(f"Data Sosis ke-{cnt} disimpan. Kirim lagi atau ketik *done* untuk selesai.")
    return WAITING_SOSIS

async def receive_ayam(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == 'skip':
        context.user_data['ayam_df'] = None
        context.user_data['ayam_list'] = []
        return await process_data(update, context)

    if text.lower() == 'done':
        ayam_list = context.user_data.get('ayam_list', [])
        if ayam_list:
            context.user_data['ayam_df'] = pd.concat(ayam_list, ignore_index=True)
        else:
            context.user_data['ayam_df'] = None
        return await process_data(update, context)

    df_trans, modul, info = parse_laporan_text(text)
    if df_trans is None or modul != 'FRIED CHICKEN':
        await update.message.reply_text(f"Gagal: {info}\nPastikan teks mengandung 'FRIED CHICKEN'. Coba lagi atau ketik *skip* / *done*.")
        return WAITING_AYAM

    context.user_data.setdefault('ayam_list', []).append(df_trans)
    if info and isinstance(info, str):
        context.user_data['ayam_last'] = info

    cnt = len(context.user_data['ayam_list'])
    await update.message.reply_text(f"Data Ayam ke-{cnt} disimpan. Kirim lagi atau ketik *done* untuk selesai.")
    return WAITING_AYAM

async def process_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    master = context.user_data['master']
    summaries = []
    available_moduls = []

    if context.user_data.get('sosis_df') is not None:
        try:
            merged_sosis = merge_and_calc(master, context.user_data['sosis_df'], MASTER_COLS['type_sosis'])
            context.user_data['merged_sosis'] = merged_sosis
            summaries.append(df_summary(merged_sosis, "HOT SAUSAGE"))
            available_moduls.append('sosis')
        except Exception as e:
            await update.message.reply_text(f"Peringatan (Sosis): {e}")

    if context.user_data.get('ayam_df') is not None:
        try:
            merged_ayam = merge_and_calc(master, context.user_data['ayam_df'], MASTER_COLS['type_ayam'])
            context.user_data['merged_ayam'] = merged_ayam
            summaries.append(df_summary(merged_ayam, "FRIED CHICKEN"))
            available_moduls.append('ayam')
        except Exception as e:
            await update.message.reply_text(f"Peringatan (Ayam): {e}")

    if not summaries:
        await update.message.reply_text("Tidak ada data yang berhasil diolah. Selesai.")
        return ConversationHandler.END

    for s in summaries:
        await update.message.reply_text(s, parse_mode='HTML')

    keyboard = []
    for mod in available_moduls:
        label = MODUL_LABEL[mod]['label']
        keyboard.append([InlineKeyboardButton(label, callback_data=f"mod:{mod}")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Pilih modul untuk detail lebih lanjut:", reply_markup=reply_markup)
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Proses dibatalkan.")
    context.user_data.clear()
    return ConversationHandler.END

# -------------------------------------------------------------------
# 8. Inline keyboard handlers
# -------------------------------------------------------------------
async def modul_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    mod = data.split(":")[1]
    context.user_data['selected_modul'] = mod

    keyboard = [
        [InlineKeyboardButton("1. Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
         InlineKeyboardButton("2. Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
        [InlineKeyboardButton("3. Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
         InlineKeyboardButton("4. Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
        [InlineKeyboardButton("5. Top 5 Toko Atas by AM", callback_data="opt:top_am")],
        [InlineKeyboardButton("6. Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("7. Top 5 Toko Atas by AS", callback_data="opt:top_as")],
        [InlineKeyboardButton("8. Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
        [InlineKeyboardButton("Kembali", callback_data="back_to_modul")],
    ]
    await query.edit_message_text(f"Opsi untuk {MODUL_LABEL[mod]['label']}:", reply_markup=InlineKeyboardMarkup(keyboard))

async def option_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "back_to_modul":
        keyboard = []
        for m in ['ayam', 'sosis']:
            if f'merged_{m}' in context.user_data:
                keyboard.append([InlineKeyboardButton(MODUL_LABEL[m]['label'], callback_data=f"mod:{m}")])
        if keyboard:
            await query.edit_message_text("Pilih modul:", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            await query.edit_message_text("Tidak ada modul tersedia.")
        return

    opt = data.split(":")[1]
    mod = context.user_data.get('selected_modul')
    if not mod:
        await query.edit_message_text("Sesi habis, silakan mulai ulang dengan /start.")
        return

    if opt in ['detail_am_excel', 'detail_as_excel', 'detail_am_jpeg', 'detail_as_jpeg']:
        merged = context.user_data.get(f'merged_{mod}')
        if merged is None:
            await query.edit_message_text("Data tidak tersedia.")
            return
        col = MASTER_COLS['am'] if 'am' in opt else MASTER_COLS['as']
        sorted_df = merged.sort_values(by=MASTER_COLS['realtime'], ascending=False)

        if 'excel' in opt:
            filename = f"{mod}_detail_{col}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine='openpyxl') as writer:
                for name, group in sorted_df.groupby(col):
                    cols_show = [MASTER_COLS['kode_toko'], MASTER_COLS['nama_toko'],
                                 MASTER_COLS['am'], MASTER_COLS['as'],
                                 MASTER_COLS['realtime'], MASTER_COLS['ach'],
                                 MASTER_COLS['target']]
                    group_show = group[cols_show].copy()
                    group_show.to_excel(writer, sheet_name=str(name)[:31], index=False)
            bio.seek(0)
            await query.message.reply_document(document=bio, filename=filename, caption=f"Detail per {col.upper()} - {MODUL_LABEL[mod]['label']} (Excel)")
        else:
            last_update = context.user_data.get(f'{mod}_last', '')
            for name, group in sorted_df.groupby(col):
                if 'am' in opt:
                    display_cols = {
                        'Kode Toko': MASTER_COLS['kode_toko'],
                        'Nama Toko': MASTER_COLS['nama_toko'],
                        'AS': MASTER_COLS['as'],
                        'Type': MASTER_COLS['type_col'],
                        'Target': MASTER_COLS['target'],
                        'Realtime': MASTER_COLS['realtime'],
                        'ACH': MASTER_COLS['ach']
                    }
                    col_order = ['Kode Toko', 'Nama Toko', 'AS', 'Type', 'Target', 'Realtime', '+/-', 'ACH']
                else:
                    display_cols = {
                        'Kode Toko': MASTER_COLS['kode_toko'],
                        'Nama Toko': MASTER_COLS['nama_toko'],
                        'AM': MASTER_COLS['am'],
                        'AS': MASTER_COLS['as'],
                        'Type': MASTER_COLS['type_col'],
                        'Target': MASTER_COLS['target'],
                        'Realtime': MASTER_COLS['realtime'],
                        'ACH': MASTER_COLS['ach']
                    }
                    col_order = ['Kode Toko', 'Nama Toko', 'AM', 'AS', 'Type', 'Target', 'Realtime', '+/-', 'ACH']

                raw_df = group[list(display_cols.values())].copy()
                raw_df.columns = list(display_cols.keys())

                target_num = pd.to_numeric(raw_df['Target'], errors='coerce').fillna(0)
                realtime_num = pd.to_numeric(raw_df['Realtime'], errors='coerce').fillna(0)
                ach_num = pd.to_numeric(raw_df['ACH'], errors='coerce').fillna(0)
                selisih = realtime_num - target_num

                raw_df['+/-'] = [f"{int(x)}" if x >= 0 else f"{int(x)}" for x in selisih]
                raw_df['Target'] = [f"{x:.0f}" for x in target_num]
                raw_df['Realtime'] = [f"{x:.0f}" for x in realtime_num]
                raw_df['ACH'] = [f"{x:.1f}%" for x in ach_num]

                detail_df = raw_df[col_order]

                total_target = target_num.sum()
                total_realtime = realtime_num.sum()
                ach_total = (total_realtime / total_target * 100) if total_target > 0 else 0
                summary = {
                    'total_target': total_target,
                    'total_realtime': total_realtime,
                    'ach_total': ach_total,
                    'jumlah_toko': len(detail_df)
                }

                unit_type = 'AM' if 'am' in opt else 'AS'
                title = f"REPORT AREA {unit_type} {name} - {MODUL_LABEL[mod]['label']}"

                img_files = create_detail_jpeg(
                    detail_df, title, last_update, summary,
                    max_rows_per_page=80
                )
                if len(img_files) == 1:
                    await query.message.reply_photo(photo=open(img_files[0], 'rb'), caption=title[:1024])
                else:
                    media = [InputMediaPhoto(open(f, 'rb')) for f in img_files]
                    await query.message.reply_media_group(media=media)
                for f in img_files:
                    os.remove(f)

        keyboard = [
            [InlineKeyboardButton("1. Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
             InlineKeyboardButton("2. Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
            [InlineKeyboardButton("3. Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
             InlineKeyboardButton("4. Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
            [InlineKeyboardButton("5. Top 5 Toko Atas by AM", callback_data="opt:top_am")],
            [InlineKeyboardButton("6. Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
            [InlineKeyboardButton("7. Top 5 Toko Atas by AS", callback_data="opt:top_as")],
            [InlineKeyboardButton("8. Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
            [InlineKeyboardButton("Kembali", callback_data="back_to_modul")],
        ]
        await query.message.reply_text("Pilih opsi lain:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif opt in ['top_am', 'bottom_am', 'top_as', 'bottom_as']:
        context.user_data['selected_opt'] = opt
        context.user_data['awaiting_code'] = True
        pesan = "Masukkan kode AM:" if 'am' in opt else "Masukkan kode AS:"
        await query.edit_message_text(pesan)
    else:
        await query.edit_message_text("Opsi tidak dikenal.")

async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get('awaiting_code'):
        return
    text = update.message.text.strip()
    mod = context.user_data.get('selected_modul')
    opt = context.user_data.get('selected_opt')
    if not mod or not opt:
        await update.message.reply_text("Sesi habis, /start ulang.")
        context.user_data.pop('awaiting_code', None)
        return

    merged = context.user_data.get(f'merged_{mod}')
    if merged is None:
        await update.message.reply_text("Data tidak tersedia.")
        context.user_data.pop('awaiting_code', None)
        return

    col = MASTER_COLS['am'] if 'am' in opt else MASTER_COLS['as']
    if text not in merged[col].values:
        await update.message.reply_text(f"Kode '{text}' tidak ditemukan. Coba lagi:")
        return

    filtered = merged[merged[col] == text]
    ascending = 'bottom' in opt
    n = TOP_N if 'top' in opt else BOTTOM_N
    sorted_df = filtered.sort_values(by=MASTER_COLS['realtime'], ascending=ascending).head(n)

    cols_show = [MASTER_COLS['kode_toko'], MASTER_COLS['nama_toko'], MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['realtime'], MASTER_COLS['ach']]
    display_df = sorted_df[cols_show].copy()
    display_df[MASTER_COLS['realtime']] = display_df[MASTER_COLS['realtime']].apply(lambda x: f"{x:.0f}" if pd.notna(x) else "-")
    display_df[MASTER_COLS['ach']] = display_df[MASTER_COLS['ach']].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")

    title = f"Top {n} {'Atas' if 'top' in opt else 'Bawah'} - {col.upper()} {text} ({MODUL_LABEL[mod]['label']})"
    img_files = create_table_image(display_df, title, max_rows_per_page=25)
    if len(img_files) == 1:
        await update.message.reply_photo(photo=open(img_files[0], 'rb'))
    else:
        media = [InputMediaPhoto(open(f, 'rb')) for f in img_files]
        await update.message.reply_media_group(media=media)
    for f in img_files:
        os.remove(f)

    context.user_data.pop('awaiting_code', None)
    keyboard = [
        [InlineKeyboardButton("1. Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
         InlineKeyboardButton("2. Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
        [InlineKeyboardButton("3. Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
         InlineKeyboardButton("4. Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
        [InlineKeyboardButton("5. Top 5 Toko Atas by AM", callback_data="opt:top_am")],
        [InlineKeyboardButton("6. Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("7. Top 5 Toko Atas by AS", callback_data="opt:top_as")],
        [InlineKeyboardButton("8. Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
        [InlineKeyboardButton("Kembali", callback_data="back_to_modul")],
    ]
    await update.message.reply_text("Pilih opsi lain:", reply_markup=InlineKeyboardMarkup(keyboard))

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logging.error(msg="Exception while handling an update:", exc_info=context.error)

# -------------------------------------------------------------------
# 9. Main
# -------------------------------------------------------------------
def main():
    os.makedirs("data/masters", exist_ok=True)
    logging.basicConfig(level=logging.INFO)

    request = HTTPXRequest(connect_timeout=30, read_timeout=60, write_timeout=60)
    app = Application.builder().token(TOKEN).request(request).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            WAITING_SOSIS: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sosis)],
            WAITING_AYAM: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_ayam)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    conv_master = ConversationHandler(
        entry_points=[CommandHandler('upload_struktur_master', upload_master_start)],
        states={
            WAITING_MASTER_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), receive_master_file)],
        },
        fallbacks=[CommandHandler('cancel', cancel_master_upload)],
    )

    app.add_handler(conv_master)
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(modul_selected, pattern='^mod:'))
    app.add_handler(CallbackQueryHandler(option_selected, pattern='^opt:|^back_to_modul'))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, receive_code))
    app.add_handler(CommandHandler('help', lambda u,c: u.message.reply_text(
        "/start - Mulai input data penjualan (Sosis & Ayam)\n"
        "/upload_struktur_master - Upload file master toko (.xlsx)\n"
        "/cancel - Batalkan proses"
    )))
    app.add_error_handler(error_handler)

    logging.info("Bot berjalan...")
    app.run_polling()

if __name__ == "__main__":
    main()