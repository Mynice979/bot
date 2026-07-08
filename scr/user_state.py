from collections import defaultdict
import pandas as pd

# Struktur: {chat_id: {"master": DataFrame, "ayam": {...}, "sosis": {...}}}
user_data = defaultdict(lambda: {
    "master": None,
    "ayam": {"df": None, "last_update": None},
    "sosis": {"df": None, "last_update": None}
})

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