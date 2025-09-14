import os
import re
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from git import Repo  # pip install gitpython

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
        return datetime.fromtimestamp(os.path.getctime(file_path))
    return pd.to_datetime(str(cell_value).strip(), errors="coerce", dayfirst=True)

def read_excel_file(file_path, sklad_name="auto"):
    extension = os.path.splitext(file_path)[1].lower()
    data = []

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
                "Количество": row[2] or 0,
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
                "Количество": ws[f"C{i}"].value or 0,
                "Цена": ws[f"D{i}"].value or 0,
                "Производитель": ws[f"E{i}"].value,
                "Артикул": str(article).strip(),
            })

    df = pd.DataFrame(data)
    if df.empty:
        return None

    if sklad_name == "auto":
        df["Склад"] = detect_sklad_from_filename(file_path)
    else:
        df["Склад"] = sklad_name
    return df

# === Обработка новых файлов ===
def process_new_file(file_path, output_folder="tmp_aggregated", repo_path="."):
    """
    Обрабатывает загруженный Excel-файл, сохраняет CSV и делает автокоммит в репозиторий GitHub.
    """
    df_new = read_excel_file(file_path, sklad_name="auto")

    if df_new is None or df_new.empty:
        return 0, "Файл пуст или не содержит обязательных колонок"

    # --- Убедимся, что все обязательные колонки есть ---
    for col in ["Артикул", "Номенклатура", "Остаток", "Цена"]:
        if col not in df_new.columns:
            return 0, f"В файле отсутствует обязательная колонка: {col}"

    added_rows = 0
    for (sklad, article), group in df_new.groupby(["Склад", "Артикул"]):
        folder = os.path.join(output_folder, safe_filename(sklad))
        os.makedirs(folder, exist_ok=True)
        out_file = os.path.join(folder, f"{safe_filename(article)}.csv")

        if os.path.exists(out_file):
            df_old = pd.read_csv(out_file)
            if "Дата" not in df_old.columns:
                continue
            max_date = pd.to_datetime(df_old["Дата"]).max()
            group = group[pd.to_datetime(group["Дата"]) > max_date]

            if not group.empty:
                group.to_csv(out_file, mode="a", header=False, index=False, encoding="utf-8-sig")
                added_rows += len(group)
        else:
            group.to_csv(out_file, index=False, encoding="utf-8-sig")
            added_rows += len(group)

    if added_rows > 0:
        try:
            git_autocommit(repo_path, file_path, added_rows)
        except Exception as e:
            return added_rows, f"Данные добавлены, но ошибка автокоммита: {e}"

    return added_rows, None
# === Автокоммит ===
def git_autocommit(repo_path, file_path, added_rows):
    """
    Делает git add, commit и push в указанный репозиторий.
    """
    try:
        repo = Repo(repo_path)
        repo.git.add(A=True)
        commit_msg = f"Автодобавление новых данных: {os.path.basename(file_path)}, строк {added_rows}"
        repo.index.commit(commit_msg)
        origin = repo.remote(name="origin")
        origin.push()
        print(f"Автокоммит выполнен: {commit_msg}")
    except Exception as e:
        print(f"Ошибка автокоммита: {e}")
