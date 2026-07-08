import io
import os
import re
import logging
from io import BytesIO
from collections import defaultdict

import pandas as pd
import numpy as np
import yaml
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from dotenv import load_dotenv

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
    """Baca file master Excel, cari sheet yang mengandung kolom 'Kode Toko'."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    for sheet in sheet_names:
        # Baca 50 baris pertama tanpa header untuk cek
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
            # Validasi kolom wajib
            required = [MASTER_COLS['kode_toko'], MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['type_col']]
            missing = [col for col in required if col not in df.columns]
            if not missing:
                return df
            else:
                # Kolom tidak lengkap, lanjut cek sheet lain
                continue

    # Jika tidak ada sheet yang cocok, tampilkan pesan jelas
    raise ValueError(
        f"Tidak ditemukan sheet yang valid di file Excel.\n"
        f"Sheet tersedia: {', '.join(sheet_names)}\n"
        "Pastikan salah satu sheet memiliki header: Kode Toko, AM, AS, TYPE, dll."
    )
    
    
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
    """Gabungkan, update REALTIME, hitung ACH tanpa mengubah kolom TARGET asli."""
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
# 4. Agregasi & ringkasan
# -------------------------------------------------------------------
def df_summary(df, modul_name):
    """Buat teks ringkasan untuk satu modul."""
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

    # --- Top / Bottom Toko (nilai realtime) ---
    df_toko = df[df[rt_col].notna()].sort_values(rt_col, ascending=False)
    top_toko = df_toko.head(TOP_N)
    bot_toko = df_toko.tail(BOTTOM_N)

    # --- Bangun teks ---
    txt = f"📊 **{modul_name}**\n"
    txt += "```\n"
    txt += f"{'AM':<10} {'Total Toko':<12} {'Toko Ada Transaksi':<18}\n"
    for _, r in grp_am.iterrows():
        txt += f"{r[am_col]:<10} {int(r['Total Toko']):<12} {int(r['Toko Ada Transaksi']):<18}\n"
    txt += "```\n"

    if not top_as.empty:
        txt += f"\n🔝 **Top {TOP_N} AS (rata‑rata realtime)**\n```\n"
        for _, r in top_as.iterrows():
            txt += f"{r[as_col]:<10} {r['avg_realtime']:>8.1f}\n"
        txt += "```\n"

    if not bot_as.empty:
        txt += f"\n🔻 **Bottom {BOTTOM_N} AS (rata‑rata realtime)**\n```\n"
        for _, r in bot_as.iterrows():
            txt += f"{r[as_col]:<10} {r['avg_realtime']:>8.1f}\n"
        txt += "```\n"

    if not top_toko.empty:
        txt += f"\n🏆 **Top {TOP_N} Toko (realtime tertinggi)**\n```\n"
        for _, r in top_toko.iterrows():
            txt += f"{r[kd_col]:<6} {r[nm_col][:20]:<20} {r[rt_col]:>6.0f}\n"
        txt += "```\n"

    if not bot_toko.empty:
        txt += f"\n🔻 **Bottom {BOTTOM_N} Toko (realtime terendah)**\n```\n"
        for _, r in bot_toko.iterrows():
            txt += f"{r[kd_col]:<6} {r[nm_col][:20]:<20} {r[rt_col]:>6.0f}\n"
        txt += "```\n"

    return txt

# -------------------------------------------------------------------
# 5. Gambar JPEG untuk opsi 3-6
# -------------------------------------------------------------------
def create_table_image(df, title, filename='temp.jpg'):
    """Buat gambar JPEG dari DataFrame."""
    fig, ax = plt.subplots(figsize=(10, 2 + 0.4 * len(df)))
    ax.axis('off')
    table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.2)
    plt.title(title, fontsize=12, weight='bold')
    plt.tight_layout()
    plt.savefig(filename, format='jpg', dpi=150)
    plt.close()
    return filename

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
        "⚠️ **Bukan** file laporan/rekapan (seperti 'REPORT PENCAPAIAN...'), "
        "melainkan file daftar seluruh toko dengan target penjualan.\n\n"
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
        context.user_data['sosis'] = None
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
        # Pisahkan jika terlalu panjang? untuk aman kita kirim per modul
        await update.message.reply_text(s, parse_mode='Markdown')

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
        [InlineKeyboardButton("1. Detail Per AM (Excel)", callback_data="opt:detail_am")],
        [InlineKeyboardButton("2. Detail Per AS (Excel)", callback_data="opt:detail_as")],
        [InlineKeyboardButton("3. Top 5 Toko Atas by AM (JPEG)", callback_data="opt:top_am")],
        [InlineKeyboardButton("4. Top 5 Toko Bawah by AM (JPEG)", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("5. Top 5 Toko Atas by AS (JPEG)", callback_data="opt:top_as")],
        [InlineKeyboardButton("6. Top 5 Toko Bawah by AS (JPEG)", callback_data="opt:bottom_as")],
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

    if opt in ['detail_am', 'detail_as']:
        # Langsung kirim Excel
        merged = context.user_data.get(f'merged_{mod}')
        if merged is None:
            await query.edit_message_text("Data tidak tersedia.")
            return
        filename = f"{mod}_{opt}.xlsx"
        # Filter & sort
        if 'am' in opt:
            col = MASTER_COLS['am']
        else:
            col = MASTER_COLS['as']
        sorted_df = merged.sort_values(by=MASTER_COLS['realtime'], ascending=False)
        # Kirim per grup? Atau satu file dengan semua data, sudah terurut. 
        # Opsi: kirim satu file besar. Agar lebih rapi, kita bisa mengelompokkan dan menambahkan sheet, tapi untuk sederhana satu file.
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine='openpyxl') as writer:
            for name, group in sorted_df.groupby(col):
                group.to_excel(writer, sheet_name=str(name)[:31], index=False)
        bio.seek(0)
        await query.message.reply_document(document=bio, filename=filename, caption=f"Detail per {col}")
    elif opt in ['top_am', 'bottom_am', 'top_as', 'bottom_as']:
        # Minta input kode AM/AS
        context.user_data['selected_opt'] = opt
        if 'am' in opt:
            pesan = "Masukkan **kode AM** yang diinginkan:"
        else:
            pesan = "Masukkan **kode AS** yang diinginkan:"
        # Simpan status menunggu input
        context.user_data['awaiting_code'] = True
        await query.edit_message_text(pesan)
    else:
        await query.edit_message_text("Opsi tidak dikenal.")

# Handler untuk input kode setelah opsi 3-6
async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get('awaiting_code'):
        return  # Biarkan handler lain yang menangani (tidak ada fallback, abaikan)
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

    # Buat gambar
    title = f"Top {n} {'Atas' if 'top' in opt else 'Bawah'} - {col.upper()} {text} ({MODUL_LABEL[mod]['label']})"
    img_path = create_table_image(display_df, title)
    await update.message.reply_photo(photo=open(img_path, 'rb'))
    os.remove(img_path)

    # Reset flag
    context.user_data.pop('awaiting_code', None)
    # Kembalikan keyboard opsi
    keyboard = [
        [InlineKeyboardButton("1. Detail Per AM", callback_data="opt:detail_am")],
        [InlineKeyboardButton("2. Detail Per AS", callback_data="opt:detail_as")],
        [InlineKeyboardButton("3. Top 5 Atas by AM", callback_data="opt:top_am")],
        [InlineKeyboardButton("4. Top 5 Bawah by AM", callback_data="opt:bottom_am")],
        [InlineKeyboardButton("5. Top 5 Atas by AS", callback_data="opt:top_as")],
        [InlineKeyboardButton("6. Top 5 Bawah by AS", callback_data="opt:bottom_as")],
        [InlineKeyboardButton("🔙 Kembali", callback_data="back_to_modul")],
    ]
    await update.message.reply_text("Pilih opsi lain:", reply_markup=InlineKeyboardMarkup(keyboard))

# -------------------------------------------------------------------
# 9. Main
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

    logging.info("Bot berjalan...")
    app.run_polling()

if __name__ == "__main__":
    main()