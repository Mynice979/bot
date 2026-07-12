import io
import os
import re
import logging
from io import BytesIO
from collections import defaultdict
import html as html_mod
import pandas as pd
import numpy as np
import yaml
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

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
# 2. Helper: baca master & parse teks
# -------------------------------------------------------------------
def load_master_excel(file_bytes: bytes) -> pd.DataFrame:
    """Baca file master Excel, cari sheet yang mengandung kolom 'Kode Toko' (tanpa spasi, case‑insensitive)."""
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
            # Baca sheet ini dengan header yang ditemukan
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=header_row, dtype=str)
            df.columns = df.columns.str.strip()

            # ---------- bulatkan TARGET menjadi integer ----------
            target_col = MASTER_COLS['target']
            if target_col in df.columns:
                df[target_col] = pd.to_numeric(df[target_col], errors='coerce').round().astype('Int64')

            # Validasi kolom wajib
            required = [MASTER_COLS['kode_toko'], MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['type_col']]
            missing = [col for col in required if col not in df.columns]
            if not missing:
                return df
            # Jika kolom tidak lengkap, lanjutkan ke sheet berikutnya

    # Tidak ditemukan sheet yang valid
    raise ValueError(
        f"Tidak ditemukan sheet yang valid di file Excel.\n"
        f"Sheet tersedia: {', '.join(sheet_names)}\n"
        "Pastikan salah satu sheet memiliki header: Kode Toko, AM, AS, TYPE, dll."
    )

def escape_html(text):
    """Escape karakter HTML dalam string."""
    if text is None:
        return ""
    return html_mod.escape(str(text))

def parse_laporan_text(content: str):
    """Kembalikan (DataFrame, modul, last_update) atau None."""
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
# 3. Penggabungan & perhitungan
# -------------------------------------------------------------------
def merge_and_calc(master_df, trans_df, type_val):
    """Gabungkan, update REALTIME, hitung ACH tanpa menimpa kolom TARGET asli."""
    kd = MASTER_COLS['kode_toko']
    tp = MASTER_COLS['type_col']
    tg = MASTER_COLS['target']
    rt = MASTER_COLS['realtime']
    ac = MASTER_COLS['ach']

    sub = master_df[master_df[tp] == type_val].copy()
    if sub.empty:
        raise ValueError(f"Tidak ada toko dengan TYPE '{type_val}' di master.")

    merged = sub.merge(trans_df, left_on=kd, right_on='Kode', how='left')

    # Update REALTIME hanya dari Qty
    merged[rt] = merged['Qty']

    # Hitung ACH menggunakan konversi target sementara (tidak mengubah kolom target)
    target_num = pd.to_numeric(merged[tg], errors='coerce')
    merged[ac] = np.where(
        (target_num > 0) & merged['Qty'].notna(),
        ((merged['Qty'] / target_num) * 100).round(1),
        np.nan
    )

    # Kosongkan REALTIME untuk toko yang tidak muncul di laporan
    merged.loc[merged['Qty'].isna(), rt] = np.nan

    # Hapus kolom dari data transaksi yang sudah tidak diperlukan
    merged.drop(columns=['Kode', 'Qty', 'Rp', 'Stock'], inplace=True, errors='ignore')

    return merged

# -------------------------------------------------------------------
# 4. Ringkasan teks per modul
# -------------------------------------------------------------------
def df_summary(df, modul_name):
    """Buat ringkasan dalam format HTML untuk satu modul."""
    am_col = MASTER_COLS['am']
    as_col = MASTER_COLS['as']
    rt_col = MASTER_COLS['realtime']
    nm_col = MASTER_COLS['nama_toko']
    kd_col = MASTER_COLS['kode_toko']

    # --- Agregasi AM ---
    grp_am = df.groupby(am_col).agg(
        total_toko=(am_col, 'count'),
        toko_berdata=(rt_col, lambda x: x.notna().sum()),
    ).reset_index()
    grp_am.columns = [am_col, 'Total Toko', 'Toko Ada Transaksi']

    # --- Agregasi AS (rata-rata realtime) ---
    df_as = df[df[rt_col].notna()].copy()
    if not df_as.empty:
        grp_as = df_as.groupby(as_col).agg(
            avg_realtime=(rt_col, 'mean')
        ).reset_index()
        top_as = grp_as.nlargest(TOP_N, 'avg_realtime')
        bot_as = grp_as.nsmallest(BOTTOM_N, 'avg_realtime')
    else:
        top_as = bot_as = pd.DataFrame()

    # --- Top / Bottom Toko ---
    df_toko = df[df[rt_col].notna()].sort_values(rt_col, ascending=False)
    top_toko = df_toko.head(TOP_N)
    bot_toko = df_toko.tail(BOTTOM_N)

    # === Bangun HTML ===
    html = f"<b>📊 {modul_name}</b>\n\n"

    # Tabel AM
    html += "<b>📋 Area Manager</b>\n<pre>"
    html += f"{'AM':<10} {'Total Toko':<12} {'Toko Ada Transaksi':<18}\n"
    for _, r in grp_am.iterrows():
        html += f"{escape_html(r[am_col]):<10} {int(r['Total Toko']):<12} {int(r['Toko Ada Transaksi']):<18}\n"
    html += "</pre>\n"

    # Top AS
    if not top_as.empty:
        html += f"<b>🔝 Top {TOP_N} AS (rata‑rata realtime)</b>\n<pre>"
        for _, r in top_as.iterrows():
            html += f"{escape_html(r[as_col]):<10} {r['avg_realtime']:>8.1f}\n"
        html += "</pre>\n"

    # Bottom AS
    if not bot_as.empty:
        html += f"<b>🔻 Bottom {BOTTOM_N} AS (rata‑rata realtime)</b>\n<pre>"
        for _, r in bot_as.iterrows():
            html += f"{escape_html(r[as_col]):<10} {r['avg_realtime']:>8.1f}\n"
        html += "</pre>\n"

    # Top Toko
    if not top_toko.empty:
        html += f"<b>🏆 Top {TOP_N} Toko (realtime tertinggi)</b>\n<pre>"
        for _, r in top_toko.iterrows():
            nama = escape_html(r[nm_col])[:20] if pd.notna(r[nm_col]) else ""
            realtime = f"{r[rt_col]:.0f}" if pd.notna(r[rt_col]) else "-"
            html += f"{escape_html(r[kd_col]):<6} {nama:<20} {realtime:>6}\n"
        html += "</pre>\n"

    # Bottom Toko
    if not bot_toko.empty:
        html += f"<b>🔻 Bottom {BOTTOM_N} Toko (realtime terendah)</b>\n<pre>"
        for _, r in bot_toko.iterrows():
            nama = escape_html(r[nm_col])[:20] if pd.notna(r[nm_col]) else ""
            realtime = f"{r[rt_col]:.0f}" if pd.notna(r[rt_col]) else "-"
            html += f"{escape_html(r[kd_col]):<6} {nama:<20} {realtime:>6}\n"
        html += "</pre>"

    return html

# -------------------------------------------------------------------
# 5. Gambar JPEG untuk opsi 3-6
# -------------------------------------------------------------------
def create_table_image(df, title, last_update="", filename='temp.jpg', max_rows_per_page=100):
    """
    Buat gambar JPEG profesional dari DataFrame.
    - ≤100 toko: 1 file
    - >100 toko: pagination per 50 toko
    - Menampilkan Last Update di header
    """
    n_rows = len(df)
    files = []
    
    # Tambahkan kolom nomor urut
    df = df.reset_index(drop=True)
    df.insert(0, 'No', range(1, len(df) + 1))
    
    for page, start in enumerate(range(0, n_rows, max_rows_per_page)):
        end = min(start + max_rows_per_page, n_rows)
        page_df = df.iloc[start:end]
        page_n_rows, page_n_cols = page_df.shape
        
        # Ukuran font menyesuaikan
        if page_n_rows > 50:
            font_size = 6
            scale_y = 0.9
        elif page_n_rows > 25:
            font_size = 8
            scale_y = 1.1
        elif page_n_rows > 15:
            font_size = 9
            scale_y = 1.3
        else:
            font_size = 10
            scale_y = 1.5
        
        # Ukuran figure
        fig_width = max(10, page_n_cols * 2.2)
        fig_height = max(4, page_n_rows * 0.45 + 2.5)
        
        fig = plt.figure(figsize=(fig_width, fig_height), facecolor='white')
        
        # Background gradient
        ax = fig.add_subplot(111)
        ax.set_facecolor('#F8F9FA')
        gradient = np.linspace(0.98, 0.85, 256).reshape(1, -1)
        gradient = np.vstack([gradient, gradient])
        ax.imshow(gradient, aspect='auto', extent=[0, 1, 0, 1], alpha=0.3, cmap='Greys')
        ax.axis('off')
        
        # Judul dengan Last Update
        if last_update:
            title_full = f"{title}\n📅 Last Update: {last_update}"
        else:
            title_full = title
        
        # Nomor halaman
        if n_rows > max_rows_per_page:
            total_pages = (n_rows - 1) // max_rows_per_page + 1
            title_full += f" | Halaman {page+1}/{total_pages}"
        
        ax.set_title(title_full, fontsize=13, weight='bold', pad=25, color='#2C3E50', loc='center')
        
        # Buat tabel
        table = ax.table(
            cellText=page_df.values,
            colLabels=page_df.columns,
            cellLoc='center',
            loc='center',
            bbox=[0.05, 0.1, 0.9, 0.75]  # [left, bottom, width, height]
        )
        table.auto_set_font_size(False)
        table.set_fontsize(font_size)
        table.scale(1, scale_y)
        
        # Gaya header
        header_color = '#1A5276'  # biru navy
        header_font_color = 'white'
        alt_row_colors = ['#FFFFFF', '#F2F4F4']  # putih & abu sangat muda
        
        for j in range(page_n_cols):
            cell = table[0, j]
            cell.set_facecolor(header_color)
            cell.set_text_props(color=header_font_color, weight='bold', fontsize=font_size+1)
            cell.set_edgecolor('#154360')
            cell.set_linewidth(1.2)
            cell.set_height(0.08)
        
        # Gaya baris data
        highlight_max_col = None
        highlight_min_col = None
        
        # Cari kolom REALTIME untuk highlight
        for j, col_name in enumerate(page_df.columns):
            if 'REALTIME' in str(col_name).upper():
                highlight_max_col = j
                break
        
        for i in range(1, page_n_rows + 1):
            is_top = False
            is_bottom = False
            
            # Highlight top 3 & bottom 3
            if highlight_max_col is not None and page_n_rows > 0:
                try:
                    val = float(str(page_df.iloc[i-1, highlight_max_col]).replace(',', ''))
                    all_vals = page_df.iloc[:, highlight_max_col].apply(
                        lambda x: float(str(x).replace(',', '')) if str(x).replace(',', '').replace('.', '').replace('-', '').isdigit() else 0
                    )
                    sorted_vals = all_vals.sort_values(ascending=False)
                    if i-1 in all_vals.nlargest(3).index:
                        is_top = True
                    if i-1 in all_vals.nsmallest(3).index and val > 0:
                        is_bottom = True
                except:
                    pass
            
            for j in range(page_n_cols):
                cell = table[i, j]
                if is_top:
                    cell.set_facecolor('#D5F5E3')  # hijau muda
                elif is_bottom:
                    cell.set_facecolor('#FADBD8')  # merah muda
                else:
                    cell.set_facecolor(alt_row_colors[(i-1) % 2])
                cell.set_edgecolor('#BDC3C7')
                cell.set_linewidth(0.5)
        
        # Footer: total toko
        fig.text(0.5, 0.02, f"Total: {page_n_rows} toko", ha='center', fontsize=9, color='#7F8C8D', style='italic')
        
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        page_filename = f"{filename.replace('.jpg','')}_p{page+1}.jpg"
        plt.savefig(page_filename, format='jpg', dpi=200, bbox_inches='tight', facecolor='white')
        plt.close()
        files.append(page_filename)
    
    return files
# -------------------------------------------------------------------
# 6. State untuk ConversationHandler
# -------------------------------------------------------------------
WAITING_MASTER = 0
WAITING_SOSIS = 1
WAITING_AYAM = 2

# -------------------------------------------------------------------
# 7. Handlers
# -------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Selamat datang!\n\n"
        "📌 **Langkah pertama:** unggah file **master toko** dalam format Excel (.xlsx).\n\n"
        "File master HARUS berisi kolom berikut:\n"
        "`Kode Toko`, `Nama Toko`, `AM`, `AS`, `TYPE`, `TARGET`, dll.\n\n"
        "⚠️ **Bukan** file laporan/rekapan, melainkan file daftar seluruh toko dengan target penjualan.\n\n"
        "Silakan kirim file tersebut sekarang."
    )
    context.user_data.clear()
    return WAITING_MASTER

async def receive_master(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.endswith('.xlsx'):
        await update.message.reply_text("File harus .xlsx, kirim ulang.")
        return WAITING_MASTER

    file = await context.bot.get_file(doc.file_id)
    fb = await file.download_as_bytearray()
    try:
        df = load_master_excel(fb)
        context.user_data['master'] = df
        await update.message.reply_text(
            f"✅ Master disimpan ({len(df)} toko).\n"
            "Sekarang kirim data **HOT SAUSAGE** (copy‑paste teks) atau ketik *skip* jika tidak ada."
        )
        return WAITING_SOSIS
    except Exception as e:
        await update.message.reply_text(f"❌ Gagal: {e}\nKirim ulang file XLSX yang benar (master toko).")
        return WAITING_MASTER

async def receive_sosis(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text.lower() == 'skip':
        context.user_data['sosis_df'] = None
        await update.message.reply_text("Data Sosis dilewati.\nSekarang kirim data **FRIED CHICKEN** (copy‑paste teks) atau ketik *skip*.")
        return WAITING_AYAM

    df_trans, modul, info = parse_laporan_text(text)
    if df_trans is None or modul != 'HOT SAUSAGE':
        await update.message.reply_text(f"❌ {info}\nPastikan teks mengandung 'HOT SAUSAGE'. Coba lagi atau ketik *skip*.")
        return WAITING_SOSIS

    context.user_data['sosis_df'] = df_trans
    context.user_data['sosis_last'] = info
    await update.message.reply_text("✅ Data Sosis diterima.\nSekarang kirim data **FRIED CHICKEN** (copy‑paste teks) atau ketik *skip*.")
    return WAITING_AYAM

async def receive_ayam(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text.lower() == 'skip':
        context.user_data['ayam_df'] = None
    else:
        df_trans, modul, info = parse_laporan_text(text)
        if df_trans is None or modul != 'FRIED CHICKEN':
            await update.message.reply_text(f"❌ {info}\nPastikan teks mengandung 'FRIED CHICKEN'. Coba lagi atau ketik *skip*.")
            return WAITING_AYAM
        context.user_data['ayam_df'] = df_trans
        context.user_data['ayam_last'] = info

    # --- Proses data ---
    master = context.user_data['master']
    summaries = []
    available_moduls = []

    # Proses Sosis
    if context.user_data.get('sosis_df') is not None:
        try:
            merged_sosis = merge_and_calc(master, context.user_data['sosis_df'], MASTER_COLS['type_sosis'])
            context.user_data['merged_sosis'] = merged_sosis
            summaries.append(df_summary(merged_sosis, "HOT SAUSAGE"))
            available_moduls.append('sosis')
        except Exception as e:
            await update.message.reply_text(f"⚠️ Gagal proses Sosis: {e}")

    # Proses Ayam
    if context.user_data.get('ayam_df') is not None:
        try:
            merged_ayam = merge_and_calc(master, context.user_data['ayam_df'], MASTER_COLS['type_ayam'])
            context.user_data['merged_ayam'] = merged_ayam
            summaries.append(df_summary(merged_ayam, "FRIED CHICKEN"))
            available_moduls.append('ayam')
        except Exception as e:
            await update.message.reply_text(f"⚠️ Gagal proses Ayam: {e}")

    if not summaries:
        await update.message.reply_text("Tidak ada data yang berhasil diolah. Selesai.")
        return ConversationHandler.END

    # Kirim ringkasan
    for s in summaries:
        await update.message.reply_text(s, parse_mode='HTML')

    # Buat inline keyboard pilih modul
    keyboard = []
    for mod in available_moduls:
        label = MODUL_LABEL[mod]['label']
        keyboard.append([InlineKeyboardButton(label, callback_data=f"mod:{mod}")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("👇 Pilih modul untuk detail lebih lanjut:", reply_markup=reply_markup)

    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Proses dibatalkan.")
    context.user_data.clear()
    return ConversationHandler.END

# -------------------------------------------------------------------
# 8. Inline keyboard handler (opsi setelah ringkasan)
# -------------------------------------------------------------------
async def modul_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data  # "mod:ayam" atau "mod:sosis"
    mod = data.split(":")[1]
    context.user_data['selected_modul'] = mod

    keyboard = [
        [InlineKeyboardButton("📊 Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
         InlineKeyboardButton("🖼️ Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
        [InlineKeyboardButton("📊 Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
         InlineKeyboardButton("🖼️ Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
        [InlineKeyboardButton("🏆 Top 5 Toko Atas by AM", callback_data="opt:top_am")],
        [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("🏆 Top 5 Toko Atas by AS", callback_data="opt:top_as")],
        [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
        [InlineKeyboardButton("🔙 Kembali", callback_data="back_to_modul")],
    ]
    await query.edit_message_text(f"📋 Opsi untuk {MODUL_LABEL[mod]['label']}:", reply_markup=InlineKeyboardMarkup(keyboard))

async def option_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "back_to_modul":
        # Kembali ke pilihan modul
        keyboard = []
        for m in ['ayam', 'sosis']:
            if f'merged_{m}' in context.user_data:
                keyboard.append([InlineKeyboardButton(MODUL_LABEL[m]['label'], callback_data=f"mod:{m}")])
        if keyboard:
            await query.edit_message_text("👇 Pilih modul:", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            await query.edit_message_text("Tidak ada modul tersedia.")
        return

    opt = data.split(":")[1]  # detail_am, top_as, dll
    mod = context.user_data.get('selected_modul')
    if not mod:
        await query.edit_message_text("Sesi habis, silakan mulai ulang dengan /start.")
        return

    if opt in ['detail_am_excel', 'detail_as_excel', 'detail_am_jpeg', 'detail_as_jpeg']:
        merged = context.user_data.get(f'merged_{mod}')
        if merged is None:
            await query.edit_message_text("Data tidak tersedia.")
            return

        # Tentukan kolom filter & sort
        if 'am' in opt:
            col = MASTER_COLS['am']
        else:
            col = MASTER_COLS['as']

        sorted_df = merged.sort_values(by=MASTER_COLS['realtime'], ascending=False)

        if 'excel' in opt:
            # --- VERSI EXCEL ---
            filename = f"{mod}_detail_{col}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine='openpyxl') as writer:
                for name, group in sorted_df.groupby(col):
                    # Ambil kolom yang relevan saja
                    cols_show = [MASTER_COLS['kode_toko'], MASTER_COLS['nama_toko'],
                                 MASTER_COLS['am'], MASTER_COLS['as'],
                                 MASTER_COLS['realtime'], MASTER_COLS['ach'],
                                 MASTER_COLS['target']]
                    group_show = group[cols_show].copy()
                    group_show.to_excel(writer, sheet_name=str(name)[:31], index=False)
            bio.seek(0)
            await query.message.reply_document(
                document=bio, filename=filename,
                caption=f"📊 Detail per {col.upper()} - {MODUL_LABEL[mod]['label']} (Excel)"
            )
        else:
            # --- VERSI JPEG (1 file jika ≤100 toko, pagination jika >100) ---
            last_update = context.user_data.get(f'{mod}_last', '')
            for name, group in sorted_df.groupby(col):
                cols_show = [MASTER_COLS['kode_toko'], MASTER_COLS['nama_toko'],
                             MASTER_COLS['realtime'], MASTER_COLS['ach']]
                display_df = group[cols_show].copy()
                display_df[MASTER_COLS['realtime']] = display_df[MASTER_COLS['realtime']].apply(
                    lambda x: f"{x:.0f}" if pd.notna(x) else "-")
                display_df[MASTER_COLS['ach']] = display_df[MASTER_COLS['ach']].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
                
                title = f"📊 Detail {col.upper()} {name} - {MODUL_LABEL[mod]['label']}"
                img_files = create_table_image(display_df, title, last_update=last_update, max_rows_per_page=100)
                
                if len(img_files) == 1:
                    caption = f"🖼️ {title}\n📅 {last_update}" if last_update else f"🖼️ {title}"
                    await query.message.reply_photo(photo=open(img_files[0], 'rb'), caption=caption[:1024])
                else:
                    media = []
                    for i, f in enumerate(img_files):
                        cap = f"{title} - Halaman {i+1}/{len(img_files)}" if i == 0 else ""
                        media.append(InputMediaPhoto(open(f, 'rb'), caption=cap[:1024]))
                    await query.message.reply_media_group(media=media)
                
                for f in img_files:
                    os.remove(f)
        # Kembalikan keyboard opsi
        keyboard = [
            [InlineKeyboardButton("📊 Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
             InlineKeyboardButton("🖼️ Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
            [InlineKeyboardButton("📊 Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
             InlineKeyboardButton("🖼️ Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
            [InlineKeyboardButton("🏆 Top 5 Toko Atas by AM", callback_data="opt:top_am")],
            [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
            [InlineKeyboardButton("🏆 Top 5 Toko Atas by AS", callback_data="opt:top_as")],
            [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
            [InlineKeyboardButton("🔙 Kembali", callback_data="back_to_modul")],
        ]
        await query.message.reply_text("✅ File terkirim. Pilih opsi lain:", reply_markup=InlineKeyboardMarkup(keyboard))

# Handler untuk input kode setelah opsi 3-6
async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get('awaiting_code'):
        return  # Biarkan handler lain yang menangani
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

    # Tentukan kolom filter
    if 'am' in opt:
        col = MASTER_COLS['am']
    else:
        col = MASTER_COLS['as']

    if text not in merged[col].values:
        await update.message.reply_text(f"Kode '{text}' tidak ditemukan. Coba lagi:")
        return

    filtered = merged[merged[col] == text]
    # Urutkan
    ascending = True if 'bottom' in opt else False
    n = TOP_N if 'top' in opt else BOTTOM_N
    sorted_df = filtered.sort_values(by=MASTER_COLS['realtime'], ascending=ascending).head(n)

    # Siapkan kolom gambar: Kode Toko, Nama Toko, AM, AS, Realtime, ACH
    cols_show = [MASTER_COLS['kode_toko'], MASTER_COLS['nama_toko'], MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['realtime'], MASTER_COLS['ach']]
    display_df = sorted_df[cols_show].copy()
    display_df[MASTER_COLS['realtime']] = display_df[MASTER_COLS['realtime']].apply(lambda x: f"{x:.0f}" if pd.notna(x) else "-")
    display_df[MASTER_COLS['ach']] = display_df[MASTER_COLS['ach']].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")

    # Buat gambar (mendukung multi-halaman jika perlu)
    title = f"Top {n} {'Atas' if 'top' in opt else 'Bawah'} - {col.upper()} {text} ({MODUL_LABEL[mod]['label']})"
    img_files = create_table_image(display_df, title, max_rows_per_page=25)
    if len(img_files) == 1:
        await update.message.reply_photo(photo=open(img_files[0], 'rb'))
    else:
        media = []
        for i, f in enumerate(img_files):
            media.append(InputMediaPhoto(open(f, 'rb')))
        await update.message.reply_media_group(media=media)
    for f in img_files:
        os.remove(f)

    # Reset flag
    context.user_data.pop('awaiting_code', None)
    # Kembalikan keyboard opsi
    keyboard = [
        [InlineKeyboardButton("📊 Detail Per AM (Excel)", callback_data="opt:detail_am_excel"),
         InlineKeyboardButton("🖼️ Detail Per AM (JPEG)", callback_data="opt:detail_am_jpeg")],
        [InlineKeyboardButton("📊 Detail Per AS (Excel)", callback_data="opt:detail_as_excel"),
         InlineKeyboardButton("🖼️ Detail Per AS (JPEG)", callback_data="opt:detail_as_jpeg")],
        [InlineKeyboardButton("🏆 Top 5 Toko Atas by AM", callback_data="opt:top_am")],
        [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AM", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("🏆 Top 5 Toko Atas by AS", callback_data="opt:top_as")],
        [InlineKeyboardButton("🔻 Top 5 Toko Bawah by AS", callback_data="opt:bottom_as")],
        [InlineKeyboardButton("🔙 Kembali", callback_data="back_to_modul")],
    ]
    await update.message.reply_text("Pilih opsi lain:", reply_markup=InlineKeyboardMarkup(keyboard))

# -------------------------------------------------------------------
# 9. Error handler
# -------------------------------------------------------------------
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Log error tanpa menghentikan bot."""
    logging.error(msg="Exception while handling an update:", exc_info=context.error)

# -------------------------------------------------------------------
# 10. Main
# -------------------------------------------------------------------
def main():
    logging.basicConfig(level=logging.INFO)
    app = Application.builder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            WAITING_MASTER: [MessageHandler(filters.Document.FileExtension("xlsx"), receive_master)],
            WAITING_SOSIS: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sosis)],
            WAITING_AYAM: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_ayam)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(modul_selected, pattern='^mod:'))
    app.add_handler(CallbackQueryHandler(option_selected, pattern='^opt:|^back_to_modul'))
    # Handler untuk menerima kode setelah opsi 3-6
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, receive_code))

    app.add_handler(CommandHandler('help', lambda u,c: u.message.reply_text("/start untuk memulai.\nProses: upload master, input Sosis, input Ayam, lalu pilih opsi.")))

    app.add_error_handler(error_handler)

    logging.info("Bot berjalan...")
    app.run_polling()

if __name__ == "__main__":
    main()