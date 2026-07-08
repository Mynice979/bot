import os
import logging
import tempfile
from io import BytesIO
import pandas as pd
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler,
    filters, ContextTypes
)
import yaml
from src.data_loader import load_master_excel, parse_laporan_text
from src.aggregator import merge_and_calculate, aggregate_am, aggregate_as, top_bottom_toko
from src.user_state import set_master, get_master, set_transaction, get_transaction, get_last_update, user_data

load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise ValueError("Token bot tidak ditemukan. Set di .env atau environment variable.")

# Load konfigurasi
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

MASTER_COLS = config['master']
PARSING_PATTERN = config['parsing']['pattern']
MODUL_MAP = {
    'ayam': 'ayam',
    'sosis': 'sosis'
}

logging.basicConfig(level=logging.INFO)

# States untuk ConversationHandler
UPLOAD_MASTER = 0
UPLOAD_REPORT = 1

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Selamat datang di Bot Monitoring Ayam & Sosis!\n"
        "Gunakan /help untuk melihat perintah yang tersedia."
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "📋 *Perintah yang tersedia:*\n"
        "/start - Memulai bot\n"
        "/upload_master - Unggah file master toko (Excel)\n"
        "/upload_report - Unggah laporan penjualan harian (file .txt)\n"
        "/status - Ringkasan data terbaru\n"
        "/am `<ayam/sosis>` - Daftar Area Manager\n"
        "/as `<ayam/sosis> <kode_am>` - Daftar AS di bawah AM\n"
        "/top `<ayam/sosis> <kode_am|as>` - 10 toko realtime tertinggi\n"
        "/bottom `<ayam/sosis> <kode_am|as>` - 10 toko realtime terendah\n"
        "/download `<ayam/sosis> <kode_am|as>` - Unduh CSV detail\n\n"
        "Contoh: /am ayam\n"
        "        /as sosis RFK\n"
        "        /top ayam RFK"
    )
    await update.message.reply_text(help_text, parse_mode='Markdown')

async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    master = get_master(chat_id)
    if master is None:
        await update.message.reply_text("Belum ada master diunggah. Gunakan /upload_master.")
        return

    msg = "📊 Status Data:\n"
    msg += f"• Master: {len(master)} toko\n"
    for mod in ['ayam', 'sosis']:
        last = get_last_update(chat_id, mod)
        if last:
            df = get_transaction(chat_id, mod)
            toko_ada = df[MASTER_COLS['realtime']].notna().sum()
            msg += f"• {mod.capitalize()}: {toko_ada} toko berdata (update {last})\n"
        else:
            msg += f"• {mod.capitalize()}: belum ada data\n"
    await update.message.reply_text(msg)

# ----- Conversation: Upload Master -----
async def upload_master_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Silakan kirim file master Excel (.xlsx).")
    return UPLOAD_MASTER

async def receive_master(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.endswith('.xlsx'):
        await update.message.reply_text("File harus berformat .xlsx. Coba lagi.")
        return UPLOAD_MASTER

    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        await file.download_to_drive(tmp.name)
        try:
            df = load_master_excel(tmp.name)
            set_master(update.effective_chat.id, df)
            await update.message.reply_text(f"✅ Master berhasil diunggah: {len(df)} toko.")
        except Exception as e:
            await update.message.reply_text(f"❌ Gagal membaca master: {str(e)}")
        finally:
            os.unlink(tmp.name)
    return ConversationHandler.END

async def cancel_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Upload dibatalkan.")
    return ConversationHandler.END

# ----- Conversation: Upload Report -----
async def upload_report_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Kirim file laporan (.txt) untuk modul Ayam atau Sosis.")
    return UPLOAD_REPORT

async def receive_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.endswith('.txt'):
        await update.message.reply_text("File harus berformat .txt.")
        return UPLOAD_REPORT

    chat_id = update.effective_chat.id
    master = get_master(chat_id)
    if master is None:
        await update.message.reply_text("⚠️ Harap unggah master terlebih dahulu dengan /upload_master.")
        return ConversationHandler.END

    file = await context.bot.get_file(document.file_id)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp:
        await file.download_to_drive(tmp.name)
        df_trans, modul, info = parse_laporan_text(tmp.name, PARSING_PATTERN)
        os.unlink(tmp.name)

    if df_trans is None:
        await update.message.reply_text(f"❌ Gagal parsing: {info}")
        return ConversationHandler.END

    # Tentukan key modul
    if modul == 'FRIED CHICKEN':
        mod_key = 'ayam'
    elif modul == 'HOT SAUSAGE':
        mod_key = 'sosis'
    else:
        await update.message.reply_text("Modul tidak dikenal.")
        return ConversationHandler.END

    # Gabungkan dengan master
    merged = merge_and_calculate(
        master, df_trans,
        MASTER_COLS['kode_toko'], MASTER_COLS['target'],
        MASTER_COLS['realtime'], MASTER_COLS['ach'],
        MASTER_COLS['type_col'],
        MASTER_COLS['type_ayam'] if mod_key == 'ayam' else MASTER_COLS['type_sosis']
    )
    set_transaction(chat_id, mod_key, merged, info if isinstance(info, str) else "")
    toko_ada = merged[MASTER_COLS['realtime']].notna().sum()
    await update.message.reply_text(
        f"✅ Laporan {modul} berhasil diunggah.\n"
        f"Total toko di master: {len(merged)}\n"
        f"Toko dengan data: {toko_ada}\n"
        f"Last Update: {info if isinstance(info, str) else '-'}"
    )
    return ConversationHandler.END

# ----- Perintah Agregasi -----
def get_merged(chat_id, modul: str):
    """Ambil DataFrame merged, validasi."""
    if modul not in ['ayam', 'sosis']:
        raise ValueError("Modul harus 'ayam' atau 'sosis'")
    master = get_master(chat_id)
    if master is None:
        return None, "Master belum diunggah."
    df = get_transaction(chat_id, modul)
    if df is None:
        return None, f"Belum ada data laporan untuk {modul}. Gunakan /upload_report."
    return df, None

async def am_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        modul = context.args[0].lower()
    except IndexError:
        await update.message.reply_text("Gunakan: /am <ayam/sosis>")
        return
    if modul not in ['ayam', 'sosis']:
        await update.message.reply_text("Modul tidak valid.")
        return

    df, err = get_merged(update.effective_chat.id, modul)
    if err:
        await update.message.reply_text(err)
        return

    agg = aggregate_am(df, MASTER_COLS['am'], MASTER_COLS['as'], MASTER_COLS['realtime'])
    # Format teks
    lines = [f"📊 Area Manager - {modul.upper()}"]
    for _, row in agg.iterrows():
        lines.append(
            f"{row[MASTER_COLS['am']]}: {int(row['total_toko'])} toko, "
            f"Data: {int(row['toko_berdata'])}, "
            f"Max: {row['realtime_max']:.0f}, Min: {row['realtime_min']:.0f}"
        )
    await update.message.reply_text("\n".join(lines))

async def as_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        modul = context.args[0].lower()
        kode_am = context.args[1]
    except IndexError:
        await update.message.reply_text("Gunakan: /as <ayam/sosis> <kode_am>")
        return
    df, err = get_merged(update.effective_chat.id, modul)
    if err:
        await update.message.reply_text(err)
        return

    if kode_am not in df[MASTER_COLS['am']].values:
        await update.message.reply_text(f"Kode AM '{kode_am}' tidak ditemukan.")
        return

    agg = aggregate_as(df, MASTER_COLS['am'], MASTER_COLS['as'],
                       MASTER_COLS['realtime'], kode_am)
    lines = [f"📌 Area Supervisor untuk AM {kode_am} ({modul})"]
    for _, row in agg.iterrows():
        lines.append(
            f"{row[MASTER_COLS['as']]}: {int(row['total_toko'])} toko, "
            f"Data: {int(row['toko_berdata'])}, "
            f"Max: {row['realtime_max']:.0f}, Min: {row['realtime_min']:.0f}"
        )
    await update.message.reply_text("\n".join(lines))

async def top_bottom(update: Update, context: ContextTypes.DEFAULT_TYPE, top=True):
    try:
        modul = context.args[0].lower()
        kode = context.args[1]
    except IndexError:
        await update.message.reply_text("Gunakan: /top atau /bottom <ayam/sosis> <kode_am|as>")
        return
    df, err = get_merged(update.effective_chat.id, modul)
    if err:
        await update.message.reply_text(err)
        return

    # Cek apakah kode adalah AM atau AS
    if kode in df[MASTER_COLS['am']].values:
        filtered = df[df[MASTER_COLS['am']] == kode]
        unit = f"AM {kode}"
    elif kode in df[MASTER_COLS['as']].values:
        filtered = df[df[MASTER_COLS['as']] == kode]
        unit = f"AS {kode}"
    else:
        await update.message.reply_text(f"Kode '{kode}' tidak ditemukan sebagai AM/AS.")
        return

    n = config['ui']['top_n'] if top else config['ui']['bottom_n']
    ascending = not top  # top = descending (ascending=False)
    res = top_bottom_toko(
        filtered, MASTER_COLS['realtime'], MASTER_COLS['nama_toko'],
        MASTER_COLS['kode_toko'], n=n, ascending=ascending
    )
    title = f"🔝 {n} Tertinggi" if top else f"🔻 {n} Terendah"
    lines = [f"{title} - {unit} ({modul})"]
    for _, row in res.iterrows():
        val = row[MASTER_COLS['realtime']]
        val_str = f"{val:.0f}" if pd.notna(val) else "Tidak Ada"
        lines.append(f"{row[MASTER_COLS['kode_toko']]} {row[MASTER_COLS['nama_toko']]}: {val_str}")
    await update.message.reply_text("\n".join(lines))

async def top_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await top_bottom(update, context, top=True)

async def bottom_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await top_bottom(update, context, top=False)

async def download(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        modul = context.args[0].lower()
        kode = context.args[1]
    except IndexError:
        await update.message.reply_text("Gunakan: /download <ayam/sosis> <kode_am|as>")
        return
    df, err = get_merged(update.effective_chat.id, modul)
    if err:
        await update.message.reply_text(err)
        return

    if kode in df[MASTER_COLS['am']].values:
        filtered = df[df[MASTER_COLS['am']] == kode]
    elif kode in df[MASTER_COLS['as']].values:
        filtered = df[df[MASTER_COLS['as']] == kode]
    else:
        await update.message.reply_text("Kode tidak valid.")
        return

    # Buat CSV dalam memori
    csv_buffer = BytesIO()
    filtered.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    await update.message.reply_document(
        document=csv_buffer,
        filename=f"{modul}_{kode}_detail.csv",
        caption=f"Detail toko untuk {kode} ({modul})"
    )

def main():
    app = Application.builder().token(TOKEN).build()

    # Conversation master
    conv_master = ConversationHandler(
        entry_points=[CommandHandler('upload_master', upload_master_start)],
        states={
            UPLOAD_MASTER: [MessageHandler(filters.Document.FileExtension("xlsx"), receive_master)]
        },
        fallbacks=[CommandHandler('cancel', cancel_upload)]
    )
    # Conversation report
    conv_report = ConversationHandler(
        entry_points=[CommandHandler('upload_report', upload_report_start)],
        states={
            UPLOAD_REPORT: [MessageHandler(filters.Document.FileExtension("txt"), receive_report)]
        },
        fallbacks=[CommandHandler('cancel', cancel_upload)]
    )

    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_command))
    app.add_handler(CommandHandler('status', status))
    app.add_handler(conv_master)
    app.add_handler(conv_report)
    app.add_handler(CommandHandler('am', am_list))
    app.add_handler(CommandHandler('as', as_list))
    app.add_handler(CommandHandler('top', top_command))
    app.add_handler(CommandHandler('bottom', bottom_command))
    app.add_handler(CommandHandler('download', download))

    logging.info("Bot dimulai...")
    app.run_polling()

if __name__ == "__main__":
    main()