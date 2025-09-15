import os
import re
import zipfile
import logging
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from git import Repo  # pip install gitpython

# === Пути ===
DATA_DIR = "data"
ARCHIVE_PATH = os.path.join(DATA_DIR, "aggregated.zip")
UPLOADS_PATH = os.path.join(DATA_DIR, "new_uploads")

# === Вспомогательные функции ===
def safe_filename(s: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", str(s))

def detect_sklad_from_filename(file_path: str) -> str:
    name = os.path.basename(file_path).lower()
    if "моск" in name:
        return "Москва"
    elif "хаб" in name or "khab" in name:
        return "Хабаровск"
    else:
        return "Неизвестно"

def parse_date_from_cell(cell_value, file_path):
    if cell_value is None:
        date = datetime.fromtimestamp(os.path.getctime(file_path))
        logging.warning(f"[parse_date_from_cell] Дата в файле {file_path} отсутствует, подставлена дата создания {date}")
        return date
    date_str = str(cell_value).strip()
    date = pd.to_datetime(date_str, errors="coerce", dayfirst=True)
    if pd.isna(date):
        date = datetime.fromtimestamp(os.path.getctime(file_path))
        logging.warning(f"[parse_date_from_cell] Дата '{date_str}' в файле {file_path} не распознана, подставлена дата создания {date}")
    return date

# === Чтение Excel ===
def read_excel_file(file_path, sklad_name="auto"):
    logging.info(f"[read_excel_file] Чтение файла: {file_path}")
    extension = os.path.splitext(file_path)[1].lower()
    data = []

    try:
        if extension == ".xls":
            df_raw = pd.read_excel(file_path, header=None, engine="xlrd")
            date_cell = df_raw.iloc[1, 0]
            date = parse_date_from_cell(date_cell, file_path)
            for idx in range(4, len(df_raw)):
                row = df_raw.iloc[idx]
                if pd.isna(row[5]):
                    continue
                data.append({
                    "Дата": date,
                    "Номенклатура": row[1],
                    "Остаток": row[2] or 0,
                    "Цена": row[3] or 0,
                    "Производитель": row[4],
                    "Артикул": str(row[5]).strip(),
                })

        elif extension == ".xlsx":
            wb = load_workbook(filename=file_path, data_only=True)
            ws = wb.active
            date_cell = ws["A2"].value
            date = parse_date_from_cell(date_cell, file_path)
            for i in range(5, ws.max_row + 1):
                article = ws[f"F{i}"].value
                if article is None:
                    continue
                data.append({
                    "Дата": date,
                    "Номенклатура": ws[f"B{i}"].value,
                    "Остаток": ws[f"C{i}"].value or 0,
                    "Цена": ws[f"D{i}"].value or 0,
                    "Производитель": ws[f"E{i}"].value,
                    "Артикул": str(article).strip(),
                })

        df = pd.DataFrame(data)
        if df.empty:
            logging.warning(f"[read_excel_file] Файл {file_path} пуст после обработки")
            return None

        if sklad_name == "auto":
            df["Склад"] = detect_sklad_from_filename(file_path)
        else:
            df["Склад"] = sklad_name

        logging.info(f"[read_excel_file] Успешно загружено строк: {len(df)} из {file_path}")
        return df

    except Exception as e:
        logging.error(f"[read_excel_file] Ошибка при чтении {file_path}: {e}", exc_info=True)
        return None

# --- Обработка нового файла Excel и пуш в GitHub ---
def process_new_file(file_path, target_repo_path="data/new_uploads"):
    logging.info(f"[process_new_file] Начата обработка {file_path}")
    df_new = read_excel_file(file_path, sklad_name="auto")
    if df_new is None or df_new.empty:
        msg = f"[process_new_file] Файл {file_path} пуст или не содержит данных"
        logging.warning(msg)
        return 0, msg

    required = ["Артикул", "Номенклатура", "Остаток", "Цена", "Дата", "Склад"]
    missing = [c for c in required if c not in df_new.columns]
    if missing:
        msg = f"[process_new_file] В файле {file_path} отсутствуют колонки: {missing}"
        logging.error(msg)
        return 0, msg

    target_path = os.path.join(target_repo_path, os.path.basename(file_path))
    commit_message = f"Добавлен новый файл {os.path.basename(file_path)}"
    success = github_upload_file(file_path, target_path, commit_message)
    added_rows = len(df_new) if success else 0
    logging.info(f"[process_new_file] Завершено. Добавлено строк: {added_rows}")
    return added_rows, None if success else "Ошибка при загрузке файла в GitHub"



