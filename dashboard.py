import dash
from dash import dcc, html, Input, Output, State, dash_table, no_update
import plotly.express as px
import plotly.graph_objs as go
import pandas as pd
import os
from dash import ctx
import io
import dash_bootstrap_components as dbc
import logging
import glob
import numpy as np
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import zipfile
import requests
from io import BytesIO
import pyarrow
import pyarrow.parquet as pq
import polars as pl
import traceback
import base64
from dash.exceptions import PreventUpdate
from data_uploader import process_new_file, read_excel_file, safe_filename
from github import Github, InputGitTreeElement
import uuid
import time
import threading
# --------------------
# НАСТРОЙКИ
# --------------------
HEIGHT_PER_BAR = 30  # высота одной строки в px
MAX_VISIBLE_BARS = 50  # сколько строк показывать без прокрутки
MAX_HEIGHT = HEIGHT_PER_BAR * MAX_VISIBLE_BARS  # высота контейнера в px

# --------------------
# ЗАГРУЗКА И ПРЕДОБРАБОТКА (один раз при старте)
# --------------------
def safe_read_excel(path):
    try:
        if path and os.path.exists(path):
            return pd.read_excel(path)
    except Exception:
        pass
    return pd.DataFrame()

df_result = safe_read_excel('итог_по_месяцу.xlsx')
df_fast = safe_read_excel('самые_ходовые.xlsx')
df_restock = safe_read_excel('чаще_всего_пополнялись.xlsx')
df_peaks = pd.read_excel('всплески_продаж1.xlsx')
df_peaks['Дата'] = pd.to_datetime(df_peaks['Дата'])

# Опционально: привести колонку Всплеск к булевому типу, если нужно
df_peaks['Всплеск'] = df_peaks['Всплеск'].astype(bool)

# Приведение числовых колонок
if not df_fast.empty:
    df_fast['Всего_продано'] = pd.to_numeric(df_fast.get('Всего_продано', 0), errors='coerce').fillna(0)
    df_fast = df_fast.dropna(subset=['Номенклатура'])

if not df_restock.empty:
    df_restock['Всего_пополнено'] = pd.to_numeric(df_restock.get('Всего_пополнено', df_restock.get('Всего_продано', 0)), errors='coerce').fillna(0)
    df_restock = df_restock.dropna(subset=['Номенклатура'])


# Группировки для топов
fast_grouped = df_fast.groupby(['Склад', 'Номенклатура', 'Артикул'], as_index=False)['Всего_продано'].sum() if not df_fast.empty else pd.DataFrame()
restock_grouped = df_restock.groupby(['Склад', 'Номенклатура', 'Артикул'], as_index=False)['Всего_пополнено'].sum() if not df_restock.empty else pd.DataFrame()

# Уникальные значения для фильтров
unique_sklads = df_result['Склад'].dropna().unique().tolist() if not df_result.empty else []
unique_peak_sklads = sorted(df_peaks['Склад'].dropna().unique()) if not df_peaks.empty else []
unique_peak_articles = sorted(df_peaks['Артикул'].dropna().unique()) if not df_peaks.empty else []
unique_peak_noms = sorted(df_peaks['Номенклатура'].dropna().unique()) if not df_peaks.empty else []


# --- Функции подготовки данных ---

def add_canonical_name(df: pd.DataFrame) -> pd.DataFrame:
    """Для каждого (Склад, Артикул, Номенклатура) выбираем каноническое название номенклатуры (мода)."""
    df = df.copy()
    df["Артикул_товар"] = df["Артикул"] + "|" + df["Номенклатура"]

    mode_map = (
        df.groupby(["Склад", "Артикул_товар"])["Номенклатура"]
        .agg(lambda s: s.mode().iat[0] if not s.mode().empty else s.dropna().iloc[0])
    )
    variants_map = (
        df.groupby(["Склад", "Артикул_товар"])["Номенклатура"]
        .agg(lambda s: ", ".join(sorted(set(s.dropna()))))
    )

    idx = df.set_index(["Склад", "Артикул_товар"]).index
    df["Номенклатура_канон"] = idx.map(mode_map.to_dict())
    df["Номенклатура_варианты"] = idx.map(variants_map.to_dict())
    df["Смена_наименования"] = df["Номенклатура"] != df["Номенклатура_канон"]
    return df


def calculate_daily_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Считаем 'Продано' и 'Пришло' по уникальным товарам (Артикул_товар), агрегируем по дате."""
    if df.empty:
        for c in ["Продано", "Пришло", "Цена_изменилась", "Аномалия"]:
            df[c] = pd.Series(dtype=float if c in ["Продано", "Пришло"] else bool)
        return df

    req = ["Склад", "Артикул_товар", "Дата", "Остаток", "Цена"]
    miss = [c for c in req if c not in df.columns]
    if miss:
        raise ValueError(f"Отсутствуют колонки: {miss}")

    df["Дата_только"] = df["Дата"].dt.normalize()

    df_daily = (
        df.sort_values("Дата")
        .groupby(["Склад", "Артикул_товар", "Дата_только"], as_index=False)
        .agg({
            "Остаток": "first",
            "Цена": "first",
            "Номенклатура": "first",
            "Номенклатура_канон": "first",
            "Номенклатура_варианты": "first"
        })
    )
    df_daily.rename(columns={"Дата_только": "Дата"}, inplace=True)

    g = df_daily.groupby(["Склад", "Артикул_товар"], group_keys=False)
    delta_stock = g["Остаток"].diff()

    df_daily["Продано"] = (-delta_stock.clip(upper=0)).fillna(0)
    df_daily["Пришло"] = (delta_stock.clip(lower=0)).fillna(0)
    df_daily["Цена_изменилась"] = g["Цена"].diff().fillna(0) != 0
    same_ost = delta_stock.fillna(0) == 0
    df_daily["Аномалия"] = ((df_daily["Продано"] > 0) | (df_daily["Пришло"] > 0)) & same_ost

    return df_daily


GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")  # токен с правом push
GITHUB_REPO = "AleksHertz/my-dash"
GITHUB_BRANCH = "main"
ARCHIVE_URL = "https://github.com/AleksHertz/my-dash/raw/refs/heads/main/data/aggregated.zip"
TMP_UPLOAD_PATH = "tmp_uploaded"  # временная папка для загруженных файлов

# --- Вспомогательная функция очистки и унификации DataFrame ---
def unify_and_clean_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    # Унификация колонки Остаток
    if "Количество" in df.columns and "Остаток" not in df.columns:
        df.rename(columns={"Количество": "Остаток"}, inplace=True)

    for col in ["Дата", "Артикул", "Номенклатура", "Остаток", "Склад"]:
        if col not in df.columns:
            df[col] = np.nan

    df["Дата"] = pd.to_datetime(df["Дата"], errors="coerce")
    df["Артикул"] = df["Артикул"].astype(str).str.strip()
    df["Номенклатура"] = df["Номенклатура"].astype(str).str.strip()
    df["Остаток"] = pd.to_numeric(df["Остаток"], errors="coerce")
    df["Цена"] = pd.to_numeric(df["Цена"], errors="coerce") if "Цена" in df.columns else np.nan

    df = df.dropna(subset=["Дата", "Артикул", "Остаток"])
    df = df.drop_duplicates(subset=["Дата", "Артикул", "Склад"], keep="last")
    return df


# --- Загрузка архива с GitHub ---
def load_and_prepare_2025_from_url(url: str = ARCHIVE_URL) -> pd.DataFrame:
    frames = []
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
            for name in z.namelist():
                if not name.lower().endswith(".csv"):
                    continue
                with z.open(name) as f:
                    df = pd.read_csv(f)
                    df.columns = [c.strip() for c in df.columns]

                    if "Москва" in name:
                        df["Склад"] = "Москва"
                    elif "Хабаровск" in name:
                        df["Склад"] = "Хабаровск"
                    else:
                        df["Склад"] = "Неизвестно"

                    frames.append(df)
        logging.info(f"[load_and_prepare_2025_from_url] Загружено файлов: {len(frames)}")
        print(f"[load_and_prepare_2025_from_url] Загружено файлов: {len(frames)}")
    except Exception as e:
        logging.error(f"[load_and_prepare_2025_from_url] Ошибка загрузки архива: {e}", exc_info=True)
        print(f"[load_and_prepare_2025_from_url] Ошибка загрузки архива: {e}")

    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    return unify_and_clean_df(df)


# --- Объединённая загрузка данных (архив + новый CSV с GitHub) ---
def load_combined_2025() -> pd.DataFrame:
    print("[load_combined_2025] Вызов функции объединённой загрузки данных")
    logging.info("[load_combined_2025] Вызов функции объединённой загрузки данных")

    try:
        # --- Загружаем архив ---
        df_archive = load_and_prepare_2025_from_url()
        print(f"[load_combined_2025] Архив загружен, строк: {len(df_archive)}")
        logging.info(f"[load_combined_2025] Архив загружен, строк: {len(df_archive)}")

        if df_archive.empty:
            print("[load_combined_2025] Нет данных в архиве")
            logging.warning("[load_combined_2025] Нет данных в архиве")
            return pd.DataFrame()

        # --- Загружаем новые CSV из GitHub ---
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO)

        all_new_parts = []
        try:
            contents = repo.get_contents("data/new_uploads", ref=GITHUB_BRANCH)
        except Exception as e:
            print(f"[load_combined_2025] Папка data/new_uploads не найдена: {e}")
            logging.warning(f"[load_combined_2025] Папка data/new_uploads не найдена: {e}")
            contents = []

        # Проходим по всем файлам и поддиректориям
        while contents:
            file_content = contents.pop(0)
            if file_content.type == "dir":
                contents.extend(repo.get_contents(file_content.path, ref=GITHUB_BRANCH))
            elif file_content.type == "file" and file_content.path.endswith(".csv"):
                try:
                    csv_data = file_content.decoded_content.decode("utf-8")
                    df_part = pd.read_csv(io.StringIO(csv_data))
                    print(f"[load_combined_2025] Загружен файл {file_content.path}, строк: {len(df_part)}")
                    logging.info(f"[load_combined_2025] Загружен файл {file_content.path}, строк: {len(df_part)}")
                    all_new_parts.append(df_part)
                except Exception as e:
                    print(f"[load_combined_2025] Ошибка чтения {file_content.path}: {e}")
                    logging.error(f"[load_combined_2025] Ошибка чтения {file_content.path}: {e}", exc_info=True)

        df_new = pd.concat(all_new_parts, ignore_index=True) if all_new_parts else pd.DataFrame()
        print(f"[load_combined_2025] Всего новых строк: {len(df_new)}")
        logging.info(f"[load_combined_2025] Всего новых строк: {len(df_new)}")

        # --- Объединяем архив и новые CSV ---
        df = pd.concat([df_archive, df_new], ignore_index=True)
        print(f"[load_combined_2025] После объединения строк: {len(df)}")
        logging.info(f"[load_combined_2025] После объединения строк: {len(df)}")

        # --- Приведение типов ---
        if "Дата" in df.columns:
            df["Дата"] = pd.to_datetime(df["Дата"], errors="coerce")

        # --- Убираем дубликаты ---
        if all(col in df.columns for col in ["Склад", "Артикул", "Дата"]):
            before = len(df)
            df.drop_duplicates(subset=["Склад", "Артикул", "Дата"], inplace=True)
            print(f"[load_combined_2025] Убрано дубликатов: {before - len(df)}")
            logging.info(f"[load_combined_2025] Убрано дубликатов: {before - len(df)}")

        # --- Дополнительная обработка ---
        df = add_canonical_name(df)
        df = calculate_daily_metrics(df)

        print(f"[load_combined_2025] Финальный размер: {len(df)} строк")
        logging.info(f"[load_combined_2025] Финальный размер: {len(df)} строк")
        return df

    except Exception as e:
        print(f"[load_combined_2025] Общая ошибка: {e}")
        logging.error(f"[load_combined_2025] Общая ошибка: {e}", exc_info=True)
        return pd.DataFrame()

# --- Функция загрузки файла в GitHub через API ---
def github_upload_file(local_path: str, target_path: str, commit_message: str):
    g = Github(GITHUB_TOKEN)
    repo = g.get_repo(GITHUB_REPO)
    try:
        with open(local_path, "rb") as f:
            content = f.read()
        try:
            file = repo.get_contents(target_path, ref=GITHUB_BRANCH)
            repo.update_file(path=target_path, message=commit_message, content=content, sha=file.sha, branch=GITHUB_BRANCH)
        except:
            repo.create_file(path=target_path, message=commit_message, content=content, branch=GITHUB_BRANCH)
        logging.info(f"[github_upload_file] Файл {target_path} успешно загружен в GitHub")
        return True
    except Exception as e:
        logging.error(f"[github_upload_file] Ошибка загрузки {target_path} в GitHub: {e}", exc_info=True)
        return False

# --- Инициализация глобальных переменных ---
df_2025 = load_combined_2025()
df_2025_clean = df_2025[~df_2025["Аномалия"]].copy() if not df_2025.empty else pd.DataFrame()

unique_sklads_2025 = sorted(df_2025_clean["Склад"].dropna().unique().tolist()) if not df_2025_clean.empty else []
unique_articles_2025 = sorted(df_2025_clean["Артикул_товар"].dropna().astype(str).unique().tolist()) if not df_2025_clean.empty else []
unique_noms_2025 = sorted(df_2025_clean["Номенклатура_канон"].dropna().unique().tolist()) if not df_2025_clean.empty else []
# --------------------
# DASH APP
# --------------------
app = dash.Dash(__name__)
server = app.server

app.layout = html.Div([
    html.H1("Анализ складских данных"),

    dcc.Tabs([
        # ===================== Основной анализ =====================
        dcc.Tab(label="Основной анализ", children=[
            # ===================== Блок ТОПЫ =====================
            html.Div([
                html.H2("ТОПы по складам"),
                html.Label("Выберите склад:"),
                dcc.Dropdown(
                    id='sklad-filter',
                    options=[{'label': s, 'value': s} for s in unique_sklads],
                    value=unique_sklads,
                    multi=True,
                    placeholder="Выберите один или несколько складов",
                    clearable=True,
                    style={'marginBottom': '20px'}
                ),
                html.Label("Выберите количество позиций для отображения ходовых товаров:"),
                dcc.RadioItems(
                    id='top-n-selector',
                    options=[
                        {'label': 'Топ 100', 'value': 100},
                        {'label': 'Топ 500', 'value': 500},
                        {'label': 'Топ 1000', 'value': 1000},
                    ],
                    value=100,
                    labelStyle={'display': 'inline-block', 'marginRight': '15px'},
                    style={'marginBottom': '20px'}
                ),
                html.H3("Топ самых ходовых товаров"),
                html.Div(
                    dcc.Graph(id='graph-top-fast'),
                    style={'height': '700px', 'overflowY': 'scroll',
                           'border': '1px solid #ddd', 'padding': '5px',
                           'marginBottom': '10px', 'backgroundColor': 'white'}
                ),
                dbc.Button("📥 Выгрузить топ ходовых в Excel", id="download-top-fast-btn", color="success", className="mb-4"),

                html.Label("Выберите количество позиций для отображения товаров по пополнениям:"),
                dcc.RadioItems(
                    id='top-n-selector-restock',
                    options=[
                        {'label': 'Топ 100', 'value': 100},
                        {'label': 'Топ 500', 'value': 500},
                        {'label': 'Топ 1000', 'value': 1000},
                    ],
                    value=100,
                    labelStyle={'display': 'inline-block', 'marginRight': '15px'},
                    style={'marginBottom': '20px'}
                ),
                html.H3("Топ товаров по пополнениям"),
                html.Div(
                    dcc.Graph(id='graph-top-restock'),
                    style={'height': '700px', 'overflowY': 'scroll',
                           'border': '1px solid #ddd', 'padding': '5px',
                           'marginBottom': '10px', 'backgroundColor': 'white'}
                ),
                dbc.Button("📥 Выгрузить топ пополнений в Excel", id="download-top-restock-btn", color="success"),

                dcc.Download(id="download-top-fast"),
                dcc.Download(id="download-top-restock"),
            ], style={'marginBottom': 40}),

            # ===================== Блок ВСПЛЕСКИ =====================
            html.Div([
                html.H2("Всплески продаж"),
                html.Div([
                    html.Label("Склад:"),
                    dcc.Dropdown(
                        id='peak-sklad-filter',
                        options=[{'label': s, 'value': s} for s in unique_peak_sklads],
                        multi=False,
                        placeholder="Выберите склад для всплесков",
                        clearable=True,
                    ),
                    html.Label("Артикул:"),
                    dcc.Dropdown(
                        id='peak-article-filter',
                        options=[{'label': a, 'value': a} for a in unique_peak_articles],
                        multi=False,
                        placeholder="Выберите артикул",
                        clearable=True,
                    ),
                    html.Label("Номенклатура:"),
                    dcc.Dropdown(
                        id='peak-nom-filter',
                        options=[],
                        multi=False,
                        placeholder="Выберите номенклатуру",
                        clearable=True,
                        searchable=True,
                        style={'width': '100%'}
                    ),
                    html.Button("📥 Скачать в Excel", id="btn-download-peaks", n_clicks=0),
                    dcc.Download(id="download-peaks-xlsx"),
                ], style={'maxWidth': 450, 'marginBottom': 30, 'display': 'flex', 'flexDirection': 'column', 'gap': '10px'}),

                dcc.Graph(id='graph-peaks'),

                html.Div([
                    html.P("График отображает:"),
                    html.Ul([
                        html.Li("Продажи (оси слева)"),
                        html.Li("Средняя цена (пунктирная линия, правая ось)"),
                        html.Li("Изменение цены в процентах (штриховая линия, правая ось)"),
                    ]),
                ], style={'maxWidth': 600, 'fontStyle': 'italic', 'color': 'gray', 'marginTop': 10}),
            ]),
        ]),  # 👈 Закрыл первую вкладку

          # ===================== Новая вкладка 2025 =====================
        dcc.Tab(label="Анализ 2025", children=[
            html.Div([
                # ===================== Загрузка новых данных =====================
                html.Div([
                    html.H3("Загрузить новые данные"),
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            'Перетащите файл сюда или ',
                            html.A('выберите файл')
                        ]),
                        style={
                            'width': '100%',
                            'height': '60px',
                            'lineHeight': '60px',
                            'borderWidth': '1px',
                            'borderStyle': 'dashed',
                            'borderRadius': '5px',
                            'textAlign': 'center',
                            'marginBottom': '20px'
                        },
                        multiple=False
                    ),
                    dcc.Loading(
                        id="loading-upload",
                        type="circle",
                        children=html.Div(
                            id='upload-status',
                            style={'marginTop': '10px', 'color': 'green'}
                        )
                    )
                ], style={'marginBottom': '30px'}),

                # ===================== Фильтры =====================
                html.Div([
                    html.Label("Склад:"),
                    dcc.Dropdown(
                        id='sklad-2025-filter',
                        options=[{'label': s, 'value': s} for s in unique_sklads_2025],
                        value=unique_sklads_2025,
                        multi=True,
                        placeholder="Выберите склад",
                        clearable=True,
                        style={'marginBottom': '15px'}
                    ),
                    html.Label("Артикул:"),
                    dcc.Dropdown(
                        id='article-2025-filter',
                        options=[{'label': a, 'value': a} for a in unique_articles_2025],
                        multi=False,
                        placeholder="Выберите артикул",
                        clearable=True,
                        style={'marginBottom': '15px'}
                    ),
                    html.Label("Номенклатура:"),
                    dcc.Dropdown(
                        id='nom-2025-filter',
                        options=[{'label': n, 'value': n} for n in unique_noms_2025],
                        multi=False,
                        placeholder="Выберите номенклатуру",
                        clearable=True,
                        style={'marginBottom': '15px'}
                    ),
                    html.Label("Месяц:"),
                    dcc.Dropdown(
                        id='month-2025-filter',
                        options=[
                            {'label': 'Январь', 'value': 1},
                            {'label': 'Февраль', 'value': 2},
                            {'label': 'Март', 'value': 3},
                            {'label': 'Апрель', 'value': 4},
                            {'label': 'Май', 'value': 5},
                            {'label': 'Июнь', 'value': 6},
                            {'label': 'Июль', 'value': 7},
                            {'label': 'Август', 'value': 8},
                            {'label': 'Сентябрь', 'value': 9},
                            {'label': 'Октябрь', 'value': 10},
                            {'label': 'Ноябрь', 'value': 11},
                            {'label': 'Декабрь', 'value': 12},
                        ],
                        multi=False,
                        placeholder="Выберите месяц",
                        clearable=True,
                        style={'marginBottom': '20px'}
                    ),
                ], style={'maxWidth': 500, 'marginBottom': 30}),

                # ===================== Линейный график =====================
                html.H3("Динамика продаж, пополнений и цены выбранного товара"),
                dcc.Graph(id='graph-2025-line'),

                # ===================== Таблица ТОП-100 =====================
                html.H3("ТОП-100 товаров по продажам (2025)", style={"marginTop": "20px"}),
                dash_table.DataTable(
                    id="top-100-table",
                    columns=[
                        {"name": "Артикул", "id": "Артикул"},
                        {"name": "Номенклатура", "id": "Номенклатура"},
                        {"name": "Продано", "id": "Продано"},
                        {"name": "Склад", "id": "Склад"},
                    ],
                    style_table={
                        "overflowX": "auto",
                        "maxHeight": "500px",
                        "overflowY": "scroll",
                        "width": "100%",
                    },
                    style_cell={
                        "textAlign": "left",
                        "padding": "5px",
                        "textDecoration": "none",
                        "whiteSpace": "normal",
                        "height": "auto",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#f0f0f0",
                        "textDecoration": "none",
                    },
                    page_size=20,
                    row_selectable="single",
                )
            ])
        ])  # Закрыли вторую вкладку
    ])  # Закрыли dcc.Tabs
])  # Закрыли html.Div(app.layout)
# --------------------
# КОЛБЭКИ
# --------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# --- Утилиты ---
def _to_list(x):
    """Нормализуем значение из дропдауна: str -> [str], None -> []"""
    if x is None:
        return []
    if isinstance(x, (list, tuple, set)):
        return list(x)
    return [x]

# ===================== Функции =====================

def get_item_line(df, article=None, nom=None, sklad_filter=None):
    dff = df.copy()
    sklads = _to_list(sklad_filter)
    if sklads:
        dff = dff[dff["Склад"].isin(sklads)]
    if article:
        dff = dff[dff["Артикул_товар"].astype(str) == str(article)]
    if nom:
        dff = dff[dff["Номенклатура_канон"] == nom]
    dff = dff.sort_values("Дата")

    keep = ["Дата", "Склад", "Артикул_товар", "Номенклатура_канон", "Остаток",
            "Продано", "Пришло", "Цена", "Цена_изменилась", "Аномалия"]
    return dff[keep]


def normalize_dataframe(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """Очистка и нормализация данных"""
    rows_before = len(df)

    # Приведение дат
    df["Дата"] = pd.to_datetime(df["Дата"], errors="coerce")
    bad_dates = df["Дата"].isna().sum()
    if bad_dates > 0:
        logging.warning(f"[{filename}] {bad_dates} строк удалено из-за некорректной даты")
        df = df.dropna(subset=["Дата"])

    # Заполнение пустых артикулов
    df["Артикул"] = df["Артикул"].fillna("").astype(str)
    missing_articles = (df["Артикул"] == "").sum()
    if missing_articles > 0:
        logging.warning(f"[{filename}] {missing_articles} строк с пустым артикулом")
        df.loc[df["Артикул"] == "", "Артикул"] = [
            f"UNKNOWN_{i}" for i in range(missing_articles)
        ]

    rows_after = len(df)
    logging.info(f"[{filename}] нормализация завершена: {rows_before} → {rows_after} строк")
    return df


# --- Колбэк загрузки файла через Dash с локальным сохранением CSV ---
@app.callback(
    Output("upload-status", "children"),
    Output("sklad-2025-filter", "options"),
    Output("article-2025-filter", "options"),
    Output("nom-2025-filter", "options"),
    Input("upload-data", "contents"),
    State("upload-data", "filename")
)
def upload_2025_file(contents, filename):
    if contents is None:
        raise dash.exceptions.PreventUpdate

    logging.info(f"[upload_2025_file] Вызов: filename={filename}, contents={'Есть' if contents else 'Нет'}")
    print(f"[upload_2025_file] Вызов: filename={filename}")

    # --- Декодируем файл ---
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    os.makedirs(TMP_UPLOAD_PATH, exist_ok=True)
    tmp_path = os.path.join(TMP_UPLOAD_PATH, filename)
    with open(tmp_path, "wb") as f:
        f.write(decoded)

    # --- Чтение Excel ---
    df_new = read_excel_file(tmp_path, sklad_name="auto")
    if df_new is None or df_new.empty:
        msg = f"Файл {filename} пуст или некорректен"
        logging.warning(f"[upload_2025_file] {msg}")
        print(f"[upload_2025_file] {msg}")
        return msg, dash.no_update, dash.no_update, dash.no_update

    logging.info(f"[upload_2025_file] Прочитано строк из Excel: {len(df_new)}")
    print(f"[upload_2025_file] Прочитано строк из Excel: {len(df_new)}")

    # --- Разбор даты (ячейка A2) ---
    if 'Дата' not in df_new.columns:
        df_new['Дата'] = pd.to_datetime(df_new.iloc[1, 0], dayfirst=True)
        logging.info(f"[upload_2025_file] Дата взята из ячейки A2: {df_new['Дата'].iloc[0]}")
        print(f"[upload_2025_file] Дата взята из ячейки A2: {df_new['Дата'].iloc[0]}")

    # --- Сохраняем все новые строки в один CSV ---
    new_csv_path = os.path.join(TMP_UPLOAD_PATH, "new_uploads.csv")
    os.makedirs(os.path.dirname(new_csv_path), exist_ok=True)
    df_new.to_csv(new_csv_path, index=False, encoding="utf-8-sig")
    logging.info(f"[upload_2025_file] Все строки сохранены в один CSV: {new_csv_path}")
    print(f"[upload_2025_file] Все строки сохранены в один CSV: {new_csv_path}")

    # --- Обновляем глобальный DataFrame ---
    global df_2025_clean
    df_2025 = load_combined_2025()
    if df_2025.empty:
        msg = "Ошибка: после обработки файла данные отсутствуют"
        logging.error(f"[upload_2025_file] {msg}")
        print(f"[upload_2025_file] {msg}")
        return msg, dash.no_update, dash.no_update, dash.no_update
    df_2025_clean = df_2025[~df_2025["Аномалия"]].copy()

    # --- Формируем options для фильтров ---
    sklads_options = [{"label": s, "value": s} for s in sorted(df_2025_clean['Склад'].unique())]
    articles_options = [{"label": a, "value": a} for a in sorted(df_2025_clean['Артикул_товар'].astype(str).unique())]
    noms_options = [{"label": n, "value": n} for n in sorted(df_2025_clean['Номенклатура_канон'].unique())]

    msg = f"Файл {filename} обработан: добавлено {len(df_new)} строк в единый CSV"
    logging.info(f"[upload_2025_file] {msg}")
    print(f"[upload_2025_file] {msg}")

    # --- Запуск фоновой загрузки на GitHub ---
    threading.Thread(target=upload_new_csv_to_github, args=(new_csv_path,), daemon=True).start()

    return msg, sklads_options, articles_options, noms_options


# --- Фоновая функция загрузки CSV на GitHub ---
def upload_new_csv_to_github(csv_path: str):
    logging.info("[upload_new_csv_to_github] Старт фоновой загрузки CSV на GitHub")
    print("[upload_new_csv_to_github] Старт фоновой загрузки CSV на GitHub")

    g = Github(GITHUB_TOKEN)
    repo = g.get_repo(GITHUB_REPO)

    filename = os.path.basename(csv_path)
    relative_path = f"data/new_uploads/{filename}"

    with open(csv_path, "rb") as f:
        content = f.read()

    commit_msg = f"Добавление новых данных: {filename}"
    try:
        try:
            existing_file = repo.get_contents(relative_path, ref=GITHUB_BRANCH)
            repo.update_file(relative_path, commit_msg, content, sha=existing_file.sha, branch=GITHUB_BRANCH)
            logging.info(f"[upload_new_csv_to_github] Обновлён файл {relative_path}")
            print(f"[upload_new_csv_to_github] Обновлён файл {relative_path}")
        except Exception:
            repo.create_file(relative_path, commit_msg, content, branch=GITHUB_BRANCH)
            logging.info(f"[upload_new_csv_to_github] Создан файл {relative_path}")
            print(f"[upload_new_csv_to_github] Создан файл {relative_path}")
    except Exception as e:
        logging.error(f"[upload_new_csv_to_github] Ошибка загрузки {relative_path}: {e}", exc_info=True)
        print(f"[upload_new_csv_to_github] Ошибка загрузки {relative_path}: {e}")
# ------------------- Колбэк графика -------------------
@app.callback(
    Output("graph-2025-line", "figure"),
    Input("article-2025-filter", "value"),
    Input("nom-2025-filter", "value"),
    Input("sklad-2025-filter", "value")
)
def update_line_graph(selected_article, selected_nom, selected_sklads):
    # Если товар не выбран, возвращаем пустой график с подсказкой
    if not selected_article or not selected_nom:
        return go.Figure(
            layout=go.Layout(
                title="Выберите товар из таблицы ТОП-100 для отображения графика",
                xaxis_title="Дата",
                yaxis_title="Остаток"
            )
        )

    dff = df_2025_clean.copy()

    # Фильтры
    if selected_sklads:
        dff = dff[dff["Склад"].isin(_to_list(selected_sklads))]
    dff = dff[(dff["Артикул_товар"].astype(str) == str(selected_article)) &
              (dff["Номенклатура_канон"] == selected_nom)]

    if dff.empty:
        return go.Figure(
            layout=go.Layout(
                title="Нет данных для выбранного товара",
                xaxis_title="Дата",
                yaxis_title="Остаток"
            )
        )

    fig = go.Figure()
    for sklad in dff["Склад"].unique():
        df_s = dff[dff["Склад"] == sklad].sort_values("Дата").copy()
        df_s["Продано_fix"] = (df_s["Остаток"].shift(1) - df_s["Остаток"]).clip(lower=0).fillna(0)
        df_s["Пополнено_fix"] = (df_s["Остаток"] - df_s["Остаток"].shift(1)).clip(lower=0).fillna(0)
        df_s["Среднее_Продано"] = df_s["Продано_fix"].rolling(7, min_periods=1).mean()
        df_s["Всплеск"] = df_s["Продано_fix"] > 1.5 * df_s["Среднее_Продано"]
        df_s["Цена_изменилась"] = df_s["Цена"].diff().fillna(0) != 0
        df_s["Цвет"] = df_s.apply(lambda row: "purple" if row["Всплеск"] and row["Цена_изменилась"]
                                   else "red" if row["Всплеск"]
                                   else "orange" if row["Цена_изменилась"]
                                   else "blue", axis=1)
        df_s["Размер"] = df_s["Всплеск"].apply(lambda x: 10 if x else 5)

        fig.add_trace(go.Scatter(
            x=df_s["Дата"],
            y=df_s["Остаток"],
            mode="lines+markers",
            name=str(sklad),
            marker=dict(size=df_s["Размер"], color=df_s["Цвет"]),
            text=[sklad]*len(df_s),
            customdata=df_s[[
                "Продано_fix", "Пополнено_fix", "Цена",
                "Артикул_товар", "Номенклатура_канон", "Всплеск", "Цена_изменилась"
            ]].values,
            hovertemplate=(
                "<b>Склад:</b> %{text}<br>"
                "<b>Дата:</b> %{x|%d-%m-%Y}<br>"
                "<b>Остаток:</b> %{y}<br>"
                "<b>Продано:</b> %{customdata[0]}<br>"
                "<b>Пополнено:</b> %{customdata[1]}<br>"
                "<b>Цена:</b> %{customdata[2]}<br>"
                "<b>Артикул:</b> %{customdata[3]}<br>"
                "<b>Номенклатура:</b> %{customdata[4]}<br>"
                "<b>Всплеск:</b> %{customdata[5]}<br>"
                "<b>Изм. цены:</b> %{customdata[6]}<br><extra></extra>"
            ),
            showlegend=False
        ))

    # Легенда
    legend_colors = {
        "Всплеск": "red",
        "Изменение цены": "orange",
        "Всплеск + Изм. цены": "purple",
        "Обычный день": "blue"
    }
    for label, color in legend_colors.items():
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(size=8, color=color), name=label))

    fig.update_layout(
        title="Динамика остатков, продаж и цен (2025)",
        xaxis_title="Дата",
        yaxis_title="Остаток",
        hovermode="closest",
        legend=dict(orientation="h", y=-0.2)
    )
    return fig

# ------------------- Таблица топ-100 -------------------
@app.callback(
    Output("top-100-table", "data"),
    Input("sklad-2025-filter", "value")
)
def update_top_100_table(selected_sklads):
    dff = df_2025_clean.copy()

    # Фильтр по складам
    if selected_sklads:
        dff = dff[dff["Склад"].isin(_to_list(selected_sklads))]

    if dff.empty:
        return []

    # Группировка по артикулу + номенклатуре + складу
    top_df = (
        dff.groupby(["Артикул_товар", "Номенклатура_канон", "Склад"], as_index=False)
           .agg({"Продано": "sum"})  # ✅ используем готовую колонку
           .sort_values("Продано", ascending=False)
           .head(100)
    )

    # Переименовываем для таблицы
    top_df = top_df.rename(
        columns={
            "Артикул_товар": "Артикул",
            "Номенклатура_канон": "Номенклатура",
        }
    )

    return top_df.to_dict("records")

# ------------------- Колбэк выбора товара из таблицы -------------------
@app.callback(
    Output("article-2025-filter", "value"),
    Output("nom-2025-filter", "value"),
    Output("month-2025-filter", "value"),
    Input("top-100-table", "selected_rows"),
    State("top-100-table", "data")
)
def select_from_table(selected_rows, table_data):
    if selected_rows and table_data:
        row = table_data[selected_rows[0]]
        # Сбрасываем месяц при выборе из таблицы
        return row["Артикул"], row["Номенклатура"], None
    return None, None, None
# --- Выгрузка топ-ходовых ---

def format_excel(dff, writer, sheet_name):
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    # Форматы
    money_fmt = workbook.add_format({'num_format': '#,##0.00 ₽'})
    integer_fmt = workbook.add_format({'num_format': '#,##0'})
    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    # Автоширина колонок и форматы
    for i, col in enumerate(dff.columns):
        max_len = max(
            dff[col].astype(str).map(len).max(),
            len(col)
        ) + 2

        fmt = None
        if col in ['Цена_в_начале', 'Цена_в_конце', 'Средняя_цена', 'Мин_цена', 'Макс_цена']:
            fmt = money_fmt
        elif col in ['Продано', 'Всего_пополнено', 'Средний_остаток']:
            fmt = integer_fmt
        elif col == 'Изменение_цены_%':
            fmt = percent_fmt
        elif col == 'Оборачиваемость':
            fmt = integer_fmt

        worksheet.set_column(i, i, max_len, fmt)

# --- Callback для топ-ходовых ---
@app.callback(
    Output("download-top-fast", "data"),
    Input("download-top-fast-btn", "n_clicks"),
    State("sklad-filter", "value"),
    State("top-n-selector", "value"),
    prevent_initial_call=True
)
def export_top_fast_to_excel(n_clicks, selected_sklads, top_n):
    if df_fast.empty or not selected_sklads:
        return None

    dff = df_fast[df_fast['Склад'].isin(selected_sklads)]
    dff = dff.sort_values('Всего_продано', ascending=False).head(top_n)

    # Новые расчёты
    for col in ['Средняя_цена', 'Мин_цена', 'Макс_цена']:
        if col in dff.columns:
            dff[col] = dff[col].round(2)

    if 'Цена_в_начале' in dff.columns and 'Цена_в_конце' in dff.columns:
        dff['Изменение_цены_%'] = (
            (dff['Цена_в_конце'] - dff['Цена_в_начале']) / dff['Цена_в_начале']
        ).round(4)

    if 'Средний_остаток' in dff.columns:
        dff['Оборачиваемость'] = (dff['Всего_продано'] / dff['Средний_остаток']).round(2)

    # Переименование колонок
    dff = dff.rename(columns={
        'Дней_продаж': 'Количество раз продаж',
        'Дней_в_наличии': 'Количество раз в наличии'
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dff.to_excel(writer, index=False, sheet_name="Топ_ходовые")
        format_excel(dff, writer, sheet_name="Топ_ходовые")
    output.seek(0)

    return dcc.send_bytes(output.getvalue(), filename=f"топ_{top_n}_ходовые.xlsx")

# --- Callback для топ-пополнений ---
@app.callback(
    Output("download-top-restock", "data"),
    Input("download-top-restock-btn", "n_clicks"),
    State("sklad-filter", "value"),
    State("top-n-selector-restock", "value"),
    prevent_initial_call=True
)
def export_top_restock_to_excel(n_clicks, selected_sklads, top_n):
    if df_restock.empty or not selected_sklads:
        return None

    dff = df_restock[df_restock['Склад'].isin(selected_sklads)]
    dff = dff.sort_values('Всего_пополнено', ascending=False).head(top_n)

    # Новые расчёты
    for col in ['Средняя_цена', 'Мин_цена', 'Макс_цена']:
        if col in dff.columns:
            dff[col] = dff[col].round(2)

    if 'Цена_в_начале' in dff.columns and 'Цена_в_конце' in dff.columns:
        dff['Изменение_цены_%'] = (
            (dff['Цена_в_конце'] - dff['Цена_в_начале']) / dff['Цена_в_начале']
        ).round(4)

    if 'Средний_остаток' in dff.columns:
        dff['Оборачиваемость'] = (dff['Всего_продано'] / dff['Средний_остаток']).round(2)

    # Переименование колонок
    dff = dff.rename(columns={
        'Дней_продаж': 'Количество раз продаж',
        'Дней_в_наличии': 'Количество раз в наличии'
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dff.to_excel(writer, index=False, sheet_name="Топ_пополнения")
        format_excel(dff, writer, sheet_name="Топ_пополнения")
    output.seek(0)

    return dcc.send_bytes(output.getvalue(), filename=f"топ_{top_n}_пополнения.xlsx")

HEIGHT_PER_BAR = 25  # Высота одной строки (можно подкорректировать)
MAX_CONTAINER_HEIGHT = 700  # Максимальная высота контейнера в px (как в layout)
@app.callback(
    Output('graph-top-fast', 'figure'),
    Input('sklad-filter', 'value'),
    Input('top-n-selector', 'value'),
)
def update_top_fast(selected_sklad, top_n):
    if not selected_sklad:
        return go.Figure()

    dff = fast_grouped[fast_grouped['Склад'].isin(selected_sklad)]
    dff = dff.sort_values('Всего_продано', ascending=False).head(top_n)

    graph_height = HEIGHT_PER_BAR * len(dff)
    graph_height = min(graph_height, MAX_CONTAINER_HEIGHT)

    fig = px.bar(
        dff,
        y='Номенклатура',
        x='Всего_продано',
        color='Склад',
        orientation='h',
        height=graph_height,
        title=f'Топ-{top_n} самых ходовых товаров'
    )

    fig.update_layout(
        yaxis={
            'categoryorder': 'array',
            'categoryarray': dff['Номенклатура'][::-1]  # переворачиваем порядок
        },
        template='plotly_white',
        margin=dict(l=250),
    )
    return fig


@app.callback(
    Output('graph-top-restock', 'figure'),
    Input('sklad-filter', 'value'),
    Input('top-n-selector-restock', 'value'),
)
def update_top_restock(selected_sklads, top_n):
    if not selected_sklads:
        return go.Figure()

    dff = restock_grouped[restock_grouped['Склад'].isin(selected_sklads)]
    dff = dff.sort_values('Всего_пополнено', ascending=False).head(top_n)

    graph_height = HEIGHT_PER_BAR * len(dff)
    graph_height = min(graph_height, MAX_CONTAINER_HEIGHT)

    fig = px.bar(
        dff,
        y='Номенклатура',
        x='Всего_пополнено',
        color='Склад',
        orientation='h',
        height=graph_height,
        title=f'Топ-{top_n} товаров по пополнениям'
    )

    fig.update_layout(
        yaxis={
            'categoryorder': 'array',
            'categoryarray': dff['Номенклатура'][::-1]  # переворачиваем порядок
        },
        template='plotly_white',
        margin=dict(l=250),
    )
    return fig

@app.callback(
    Output("peak-nom-filter", "options"),
    Input("peak-sklad-filter", "value"),
    Input("peak-article-filter", "value")
)
def update_nom_options(selected_sklad, selected_article):
    if not selected_sklad and not selected_article:
        return []

    dff = df_peaks.copy()
    if selected_sklad:
        dff = dff[dff["Склад"] == selected_sklad]
    if selected_article:
        dff = dff[dff["Артикул"] == selected_article]

    return [{"label": nom, "value": nom} for nom in sorted(dff["Номенклатура"].unique())]

@app.callback(
    Output('graph-peaks', 'figure'),
    Input('peak-sklad-filter', 'value'),
    Input('peak-article-filter', 'value'),
    Input('peak-nom-filter', 'value'),
)
def update_peaks_graph(sklad, article, nom):
    dff = df_peaks.copy()
    if sklad:
        dff = dff[dff['Склад'] == sklad]
    if article:
        dff = dff[dff['Артикул'] == article]
    if nom:
        dff = dff[dff['Номенклатура'] == nom]

    if dff.empty:
        return go.Figure()

    dff = dff.sort_values('Дата').tail(200)  # ограничение последних 200 строк

    fig = go.Figure()

    for sklad_name, group in dff.groupby('Склад'):
        fig.add_trace(go.Scatter(
            x=group['Дата'],
            y=group['Всего_продано'],
            mode='markers+lines',
            name=f'Продано - {sklad_name}',
            hovertemplate='Дата: %{x}<br>Продано: %{y}<br>Артикул: %{customdata[0]}<br>Номенклатура: %{customdata[1]}<extra></extra>',
            customdata=group[['Артикул', 'Номенклатура']],
            yaxis='y1',
        ))

        fig.add_trace(go.Scatter(
            x=group['Дата'],
            y=group['Средняя_цена'],
            mode='lines+markers',
            name=f'Средняя цена - {sklad_name}',
            line=dict(dash='dot'),
            hovertemplate='Дата: %{x}<br>Средняя цена: %{y}<extra></extra>',
            yaxis='y2',
        ))

        fig.add_trace(go.Scatter(
            x=group['Дата'],
            y=group['Изменение_цены_%'],
            mode='lines+markers',
            name=f'Изменение цены % - {sklad_name}',
            line=dict(dash='dash'),
            hovertemplate='Дата: %{x}<br>Изменение цены %: %{y}<extra></extra>',
            yaxis='y3',
        ))

    fig.update_layout(
        title='Всплески продаж и динамика цен',
        xaxis=dict(title='Дата'),
        yaxis=dict(title='Продано', side='left', showgrid=False, zeroline=False),
        yaxis2=dict(title='Средняя цена', overlaying='y', side='right', showgrid=False, zeroline=False, position=0.95),
        yaxis3=dict(title='Изменение цены %', overlaying='y', side='right', showgrid=False, zeroline=False,
                    position=1.0, anchor='free'),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=40, r=80, t=80, b=40),
        template='plotly_white'
    )
    return fig

@app.callback(
    Output("download-peaks-xlsx", "data"),
    Input("btn-download-peaks", "n_clicks"),
    State("peak-sklad-filter", "value"),
    State("peak-article-filter", "value"),
    State("peak-nom-filter", "value"),
    prevent_initial_call=True,
)
def download_peaks_excel(n_clicks, sklad, article, nom):
    dff = df_peaks.copy()
    if sklad:
        dff = dff[dff['Склад'] == sklad]
    if article:
        dff = dff[dff['Артикул'] == article]
    if nom:
        dff = dff[dff['Номенклатура'] == nom]

    if dff.empty:
        return dash.no_update

    # Добавим столбец с оборачиваемостью (если нет - считаем как пример)
    # Например: Оборачиваемость = Всего_продано / Среднее количество на складе (пример)
    # Здесь подставь свою логику, если нужно
    if 'Оборачиваемость' not in dff.columns:
        dff['Оборачиваемость'] = dff['Всего_продано'] / 10  # пример

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dff.to_excel(writer, index=False, sheet_name='Всплески_продаж')
    output.seek(0)

    return dcc.send_bytes(output.read(), filename="всплески_продаж.xlsx")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))  # Используем порт из переменной окружения или 10000 по умолчанию
    app.run_server(debug=False, host='0.0.0.0', port=port)
