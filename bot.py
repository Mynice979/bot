import io
import os
import re
import logging
from io import BytesIO
from collections import defaultdict
from typing import Optional

import pandas as pd
import numpy as np
import yaml
from dotenv import load_dotenv

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
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
MODUL_AYAM = config["modul"]["ayam"]
MODUL_SOSIS = config["modul"]["sosis"]
UI_TOP = config["ui"]["top_n"]
UI_BOTTOM = config["ui"]["bottom_n"]

# -------------------------------------------------------------------
# 2. Data loader functions
# -------------------------------------------------------------------
def load_master_excel(file_bytes: bytes) -> pd.DataFrame:
    """Baca file master Excel, otomatis cari baris header yang mengandung 'Kode Toko'."""
    # Baca dulu 20 baris pertama tanpa header untuk mencari baris header
    df_preview = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=20)
    header_row = None
    for idx, row in df_preview.iterrows():
        if row.astype(str).str.contains('Kode Toko', case=False).any():
            header_row = idx
            break
    if header_row is None:
        raise ValueError("Kolom 'Kode Toko' tidak ditemukan di 20 baris pertama file master.")
    
    # Baca ulang mulai dari baris header
    df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_row, dtype=str)
    df.columns = df.columns.str.strip()
    return df

def parse_laporan_text_from_content(content: str, pattern: str):
    """
    Parse laporan dari string content.
    Return (DataFrame, modul, last_update) atau (None, None, error_msg)
    """
    if "FRIED CHICKEN" in content:
        modul = "FRIED CHICKEN"
    elif "HOT SAUSAGE" in content:
        modul = "HOT SAUSAGE"
    else:
        return None, None, "Modul tidak dikenali (harus FRIED CHICKEN atau HOT SAUSAGE)"

    # Cari Last Update
    last_update = None
    match = re.search(r"Last Update:\s*(\d{2}-\d{2}-\d{4}\s\d{2}:\d{2}:\d{2})", content)
    if match:
        last_update = match.group(1)

    # Parse baris data
    rows = []
    for line in content.splitlines():
        m = re.search(pattern, line)
        if m:
            rows.append(
                {
                    "Kode": m.group("kode"),
                    "Qty": int(m.group("qty")),
                    "Rp": int(m.group("rp").replace(",", "")),
                    "Stock": int(m.group("stock")),
                }
            )

    if not rows:
        return None, modul, "Tidak ada data toko ditemukan."

    df = pd.DataFrame(rows)
    return df, modul, last_update


# -------------------------------------------------------------------
# 3. Aggregator functions
# -------------------------------------------------------------------
def merge_and_calculate(
    master_df: pd.DataFrame,
    trans_df: pd.DataFrame,
    kode_col: str,
    target_col: str,
    realtime_col: str,
    ach_col: str,
    type_col: str,
    modul_type: str,
) -> pd.DataFrame:
    """Gabungkan master dengan transaksi, update REALTIME, hitung ACH."""
    # Cek apakah kolom TYPE ada
    if type_col not in master_df.columns:
        raise KeyError(f"Kolom '{type_col}' tidak ditemukan di master. Pastikan file master memiliki kolom TYPE.")
    master_mod = master_df[master_df[type_col] == modul_type].copy()
    if master_mod.empty:
        raise ValueError(f"Tidak ada toko dengan TYPE='{modul_type}' di master. Cek isi kolom TYPE.")
    merged = master_mod.merge(trans_df, left_on=kode_col, right_on="Kode", how="left")

    # Update REALTIME
    merged[realtime_col] = merged["Qty"]

    # Hitung ACH
    merged[target_col] = pd.to_numeric(merged[target_col], errors="coerce")
    merged[ach_col] = np.where(
        (merged[target_col] > 0) & (merged["Qty"].notna()),
        ((merged["Qty"] / merged[target_col]) * 100).round(1),
        np.nan,
    )

    # Kosongkan REALTIME untuk toko yang tidak muncul di laporan
    merged.loc[merged["Qty"].isna(), realtime_col] = np.nan
    return merged


def aggregate_am(df: pd.DataFrame, am_col: str, as_col: str, realtime_col: str) -> pd.DataFrame:
    """Agregasi per AM."""
    agg = (
        df.groupby(am_col)
        .agg(
            total_toko=(am_col, "count"),
            toko_berdata=(realtime_col, lambda x: x.notna().sum()),
            realtime_max=(realtime_col, "max"),
            realtime_min=(realtime_col, "min"),
        )
        .reset_index()
    )
    agg["realtime_max"] = agg["realtime_max"].fillna(0)
    agg["realtime_min"] = agg["realtime_min"].fillna(0)
    return agg


def aggregate_as(
    df: pd.DataFrame, am_col: str, as_col: str, realtime_col: str, kode_am: str
) -> pd.DataFrame:
    """Agregasi per AS untuk AM tertentu."""
    sub = df[df[am_col] == kode_am]
    agg = (
        sub.groupby(as_col)
        .agg(
            total_toko=(as_col, "count"),
            toko_berdata=(realtime_col, lambda x: x.notna().sum()),
            realtime_max=(realtime_col, "max"),
            realtime_min=(realtime_col, "min"),
        )
        .reset_index()
    )
    agg["realtime_max"] = agg["realtime_max"].fillna(0)
    agg["realtime_min"] = agg["realtime_min"].fillna(0)
    return agg


def top_bottom_toko(
    df: pd.DataFrame,
    realtime_col: str,
    nama_toko_col: str,
    kode_toko_col: str,
    n: int = 10,
    ascending: bool = True,
):
    """Ambil n toko teratas/terbawah berdasarkan REALTIME."""
    sorted_df = df.sort_values(by=realtime_col, ascending=ascending, na_position="last")
    return sorted_df[[kode_toko_col, nama_toko_col, realtime_col]].head(n)


# -------------------------------------------------------------------
# 4. User state
# -------------------------------------------------------------------
user_data = defaultdict(
    lambda: {
        "master": None,
        "ayam": {"df": None, "last_update": None},
        "sosis": {"df": None, "last_update": None},
    }
)


def set_master(chat_id, df):
    user_data[chat_id]["master"] = df


def get_master(chat_id):
    return user_data[chat_id]["master"]


def set_transaction(chat_id, modul, df, last_update):
    user_data[chat_id][modul]["df"] = df
    user_data[chat_id][modul]["last_update"] = last_update


def get_transaction(chat_id, modul):
    return user_data[chat_id][modul]["df"]


def get_last_update(chat_id, modul):
    return user_data[chat_id][modul].get("last_update")


# -------------------------------------------------------------------
# 5. Conversation states
# -------------------------------------------------------------------
UPLOAD_MASTER = 0
UPLOAD_REPORT = 1
INPUT_REPORT_TEXT = 2   # state untuk menerima teks laporan

# -------------------------------------------------------------------
# 6. Command & conversation handlers
# -------------------------------------------------------------------
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
        "/report_text - Masukkan laporan penjualan via copy-paste teks\n"
        "/status - Ringkasan data terbaru\n"
        "/am `<ayam/sosis>` - Daftar Area Manager\n"
        "/as `<ayam/sosis> <kode_am>` - Daftar AS di bawah AM\n"
        "/top `<ayam/sosis> <kode_am|as>` - 10 toko realtime tertinggi\n"
        "/bottom `<ayam/sosis> <kode_am|as>` - 10 toko realtime terendah\n"
        "/download `<ayam/sosis> <kode_am|as>` - Unduh CSV detail\n\n"
        "Contoh: /am ayam\n"
        "        /as sosis RFK\n"
        "        /top ayam RFK\n"
        "        /report_text → lalu tempel teks laporan"
    )
    await update.message.reply_text(help_text, parse_mode="Markdown")


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    master = get_master(chat_id)
    if master is None:
        await update.message.reply_text("Belum ada master diunggah. Gunakan /upload_master.")
        return

    msg = "📊 Status Data:\n"
    msg += f"• Master: {len(master)} toko\n"
    for mod in ["ayam", "sosis"]:
        last = get_last_update(chat_id, mod)
        if last:
            df = get_transaction(chat_id, mod)
            if df is not None:
                toko_ada = df[MASTER_COLS["realtime"]].notna().sum()
                msg += f"• {mod.capitalize()}: {toko_ada} toko berdata (update {last})\n"
        else:
            msg += f"• {mod.capitalize()}: belum ada data\n"
    await update.message.reply_text(msg)


# ----- Upload Master -----
async def upload_master_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Silakan kirim file master Excel (.xlsx).")
    return UPLOAD_MASTER


async def receive_master(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.endswith(".xlsx"):
        await update.message.reply_text("File harus berformat .xlsx. Coba lagi.")
        return UPLOAD_MASTER

    file = await context.bot.get_file(document.file_id)
    file_bytes = await file.download_as_bytearray()
    try:
        df = load_master_excel(file_bytes)
        # Validasi kolom wajib
        missing = []
        for col in [MASTER_COLS["kode_toko"], MASTER_COLS["am"], MASTER_COLS["as"], MASTER_COLS["type_col"]]:
            if col not in df.columns:
                missing.append(col)
        if missing:
            await update.message.reply_text(
                f"❌ Kolom berikut tidak ditemukan di file master: {', '.join(missing)}\n"
                "Pastikan kolom tersebut ada dan namanya sesuai."
            )
            return ConversationHandler.END

        set_master(update.effective_chat.id, df)
        await update.message.reply_text(f"✅ Master berhasil diunggah: {len(df)} toko.")
    except Exception as e:
        await update.message.reply_text(f"❌ Gagal membaca master: {str(e)}")
    return ConversationHandler.END


async def cancel_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Upload dibatalkan.")
    return ConversationHandler.END


# ----- Upload Report (file) -----
async def upload_report_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Kirim file laporan (.txt) untuk modul Ayam atau Sosis.")
    return UPLOAD_REPORT


async def receive_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.endswith(".txt"):
        await update.message.reply_text("File harus berformat .txt.")
        return UPLOAD_REPORT

    chat_id = update.effective_chat.id
    master = get_master(chat_id)
    if master is None:
        await update.message.reply_text(
            "⚠️ Harap unggah master terlebih dahulu dengan /upload_master."
        )
        return ConversationHandler.END

    file = await context.bot.get_file(document.file_id)
    file_bytes = await file.download_as_bytearray()
    content = file_bytes.decode("utf-8")
    return await proses_laporan_dari_teks(update, chat_id, master, content)


# ----- Input Report via Copy-Paste (teks langsung) -----
async def report_text_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Silakan tempel (paste) teks laporan penjualan (Ayam/Sosis) di sini.\n"
        "Pastikan teks mengandung 'FRIED CHICKEN' atau 'HOT SAUSAGE'."
    )
    return INPUT_REPORT_TEXT


async def receive_report_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    master = get_master(chat_id)
    if master is None:
        await update.message.reply_text(
            "⚠️ Harap unggah master terlebih dahulu dengan /upload_master."
        )
        return ConversationHandler.END

    content = update.message.text
    return await proses_laporan_dari_teks(update, chat_id, master, content)


async def proses_laporan_dari_teks(update: Update, chat_id, master, content: str):
    """Fungsi bersama untuk memproses konten laporan, baik dari file atau teks."""
    df_trans, modul, info = parse_laporan_text_from_content(content, PARSING_PATTERN)

    if df_trans is None:
        await update.message.reply_text(f"❌ Gagal parsing: {info}")
        return ConversationHandler.END if update.message.text is None else INPUT_REPORT_TEXT
        # Untuk input teks, tetap stay di state agar bisa coba lagi? Kita atur: jika dari file -> END, jika dari teks -> END juga agar simple.
    # Sebenarnya kita bisa langsung return END, tapi karena dipanggil dari dua tempat,
    # kita tetap harus mengembalikan state yang tepat. Biarkan saja return ConversationHandler.END.

    if modul == "FRIED CHICKEN":
        mod_key = "ayam"
    elif modul == "HOT SAUSAGE":
        mod_key = "sosis"
    else:
        await update.message.reply_text("Modul tidak dikenal.")
        return ConversationHandler.END

    try:
        merged = merge_and_calculate(
            master,
            df_trans,
            MASTER_COLS["kode_toko"],
            MASTER_COLS["target"],
            MASTER_COLS["realtime"],
            MASTER_COLS["ach"],
            MASTER_COLS["type_col"],
            MASTER_COLS["type_ayam"] if mod_key == "ayam" else MASTER_COLS["type_sosis"],
        )
    except (KeyError, ValueError) as e:
        await update.message.reply_text(f"❌ Error saat menggabungkan data: {str(e)}")
        return ConversationHandler.END

    set_transaction(chat_id, mod_key, merged, info if isinstance(info, str) else "")
    toko_ada = merged[MASTER_COLS["realtime"]].notna().sum()
    await update.message.reply_text(
        f"✅ Laporan {modul} berhasil diunggah.\n"
        f"Total toko di master: {len(merged)}\n"
        f"Toko dengan data: {toko_ada}\n"
        f"Last Update: {info if isinstance(info, str) else '-'}"
    )
    return ConversationHandler.END


# ----- Agregasi commands -----
def get_merged(chat_id, modul: str):
    """Ambil DataFrame merged, validasi."""
    if modul not in ["ayam", "sosis"]:
        raise ValueError("Modul harus 'ayam' atau 'sosis'")
    master = get_master(chat_id)
    if master is None:
        return None, "Master belum diunggah."
    df = get_transaction(chat_id, modul)
    if df is None:
        return None, f"Belum ada data laporan untuk {modul}. Gunakan /upload_report atau /report_text."
    return df, None


async def am_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        modul = context.args[0].lower()
    except IndexError:
        await update.message.reply_text("Gunakan: /am <ayam/sosis>")
        return
    if modul not in ["ayam", "sosis"]:
        await update.message.reply_text("Modul tidak valid.")
        return

    df, err = get_merged(update.effective_chat.id, modul)
    if err:
        await update.message.reply_text(err)
        return

    agg = aggregate_am(df, MASTER_COLS["am"], MASTER_COLS["as"], MASTER_COLS["realtime"])
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

    if kode_am not in df[MASTER_COLS["am"]].values:
        await update.message.reply_text(f"Kode AM '{kode_am}' tidak ditemukan.")
        return

    agg = aggregate_as(
        df, MASTER_COLS["am"], MASTER_COLS["as"], MASTER_COLS["realtime"], kode_am
    )
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

    if kode in df[MASTER_COLS["am"]].values:
        filtered = df[df[MASTER_COLS["am"]] == kode]
        unit = f"AM {kode}"
    elif kode in df[MASTER_COLS["as"]].values:
        filtered = df[df[MASTER_COLS["as"]] == kode]
        unit = f"AS {kode}"
    else:
        await update.message.reply_text(f"Kode '{kode}' tidak ditemukan sebagai AM/AS.")
        return

    n = UI_TOP if top else UI_BOTTOM
    ascending = not top
    res = top_bottom_toko(
        filtered,
        MASTER_COLS["realtime"],
        MASTER_COLS["nama_toko"],
        MASTER_COLS["kode_toko"],
        n=n,
        ascending=ascending,
    )
    title = f"🔝 {n} Tertinggi" if top else f"🔻 {n} Terendah"
    lines = [f"{title} - {unit} ({modul})"]
    for _, row in res.iterrows():
        val = row[MASTER_COLS["realtime"]]
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

    if kode in df[MASTER_COLS["am"]].values:
        filtered = df[df[MASTER_COLS["am"]] == kode]
    elif kode in df[MASTER_COLS["as"]].values:
        filtered = df[df[MASTER_COLS["as"]] == kode]
    else:
        await update.message.reply_text("Kode tidak valid.")
        return

    csv_buffer = BytesIO()
    filtered.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    await update.message.reply_document(
        document=csv_buffer,
        filename=f"{modul}_{kode}_detail.csv",
        caption=f"Detail toko untuk {kode} ({modul})",
    )


# -------------------------------------------------------------------
# 7. Error handler
# -------------------------------------------------------------------
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Log error tanpa menghentikan bot."""
    logging.error(msg="Exception while handling an update:", exc_info=context.error)


# -------------------------------------------------------------------
# 8. Main
# -------------------------------------------------------------------
def main():
    logging.basicConfig(level=logging.INFO)
    app = Application.builder().token(TOKEN).build()

    # Conversation master
    conv_master = ConversationHandler(
        entry_points=[CommandHandler("upload_master", upload_master_start)],
        states={
            UPLOAD_MASTER: [MessageHandler(filters.Document.FileExtension("xlsx"), receive_master)]
        },
        fallbacks=[CommandHandler("cancel", cancel_upload)],
    )
    # Conversation report (file .txt)
    conv_report = ConversationHandler(
        entry_points=[CommandHandler("upload_report", upload_report_start)],
        states={
            UPLOAD_REPORT: [MessageHandler(filters.Document.FileExtension("txt"), receive_report)]
        },
        fallbacks=[CommandHandler("cancel", cancel_upload)],
    )
    # Conversation report text (copy-paste)
    conv_report_text = ConversationHandler(
        entry_points=[CommandHandler("report_text", report_text_start)],
        states={
            INPUT_REPORT_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_report_text)]
        },
        fallbacks=[CommandHandler("cancel", cancel_upload)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(conv_master)
    app.add_handler(conv_report)
    app.add_handler(conv_report_text)   # <-- handler baru
    app.add_handler(CommandHandler("am", am_list))
    app.add_handler(CommandHandler("as", as_list))
    app.add_handler(CommandHandler("top", top_command))
    app.add_handler(CommandHandler("bottom", bottom_command))
    app.add_handler(CommandHandler("download", download))
    app.add_error_handler(error_handler)

    logging.info("Bot dimulai...")
    app.run_polling()


if __name__ == "__main__":
    main()