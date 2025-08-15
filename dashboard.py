import dash
from dash import dcc, html, Input, Output, State, dash_table
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
from zipfile import ZipFile
import requests
from io import BytesIO
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
    # Уникальные комбинации Артикул + Номенклатура
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

    # Агрегируем по уникальному товару (Склад + Артикул_товар) и дате
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


def load_and_prepare_2025(base_path: str = "data/агрегированные") -> pd.DataFrame:
    frames = []
    columns_needed = ["Дата", "Артикул", "Номенклатура", "Остаток", "Цена"]

    for sklad in ("Москва", "Хабаровск"):
        pattern = os.path.join(base_path, "**", sklad, "*.csv")
        files = glob.glob(pattern, recursive=True)
        if not files:
            print(f"Внимание: нет файлов для склада '{sklad.lower()}' в {base_path}")
            continue

        for f in files:
            try:
                for chunk in pd.read_csv(f, usecols=columns_needed, chunksize=10000):
                    chunk["Склад"] = sklad
                    chunk["Дата"] = pd.to_datetime(chunk["Дата"], errors="coerce")
                    chunk["Артикул"] = chunk["Артикул"].astype(str).str.strip()
                    chunk["Номенклатура"] = chunk["Номенклатура"].astype(str).str.strip()
                    chunk["Остаток"] = pd.to_numeric(chunk["Остаток"], errors="coerce")
                    chunk["Цена"] = pd.to_numeric(chunk["Цена"], errors="coerce")
                    chunk = chunk.dropna(subset=["Дата", "Артикул", "Остаток"])
                    frames.append(chunk)
            except Exception as e:
                print(f"Ошибка при чтении файла {f}: {e}")

    if not frames:
        print(f"Нет файлов для объединения в {base_path}, возвращаем пустой DataFrame")
        return pd.DataFrame()

    # Объединяем только после обработки всех чанков
    df = pd.concat(frames, ignore_index=True)
    df = add_canonical_name(df)
    df = calculate_daily_metrics(df)
    return df

# --- Пример скачивания ZIP с GitHub прямо в память ---
def download_and_extract_zip(github_url: str, extract_path: str):
    os.makedirs(extract_path, exist_ok=True)
    try:
        r = requests.get(github_url)
        r.raise_for_status()
        with ZipFile(BytesIO(r.content)) as zip_ref:
            zip_ref.extractall(extract_path)
        print(f"Архив успешно распакован в {extract_path}")
    except Exception as e:
        print(f"Ошибка при скачивании или распаковке архива: {e}")


# --- Использование ---
github_zip_url = "https://github.com/AleksHertz/my-dash/raw/refs/heads/main/data/aggregated.zip"  # замените на реальный URL
extract_path = "data/агрегированные"
download_and_extract_zip(github_zip_url, extract_path)

df_2025 = load_and_prepare_2025(extract_path)
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
        ]),

        # ===================== Новая вкладка 2025 =====================
        dcc.Tab(label="Анализ 2025", children=[
            html.Div([
                html.H2("Анализ продаж за 2025 год"),

                # Фильтры
                html.Div([
                    html.Label("Склад:"),
                    dcc.Dropdown(
                        id='sklad-2025-filter',
                        options=[{'label': s, 'value': s} for s in unique_sklads_2025],
                        value=unique_sklads_2025,  # по умолчанию все склады
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
                        style={'marginBottom': '20px'}
                    ),
                ], style={'maxWidth': 500, 'marginBottom': 30}),

                # Линейный график
                html.H3("Динамика продаж, пополнений и цены выбранного товара"),
                dcc.Graph(id='graph-2025-line'),

                # Таблица ТОП-100 товаров
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
                        "textDecoration": "none",  # убираем подчеркивание
                        "whiteSpace": "normal",
                        "height": "auto",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#f0f0f0",
                        "textDecoration": "none",
                    },
                    page_size=20,
                    row_selectable="single",  # для клика по строке
                )
            ])
        ])
    ])
])
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

# ===================== Колбэки =====================

## ------------------- График остатков -------------------
@app.callback(
    Output("graph-2025-line", "figure"),
    Input("sklad-2025-filter", "value"),
    Input("article-2025-filter", "value"),
    Input("nom-2025-filter", "value")
)
def update_line_graph(selected_sklads, selected_article, selected_nom):
    if not selected_article and not selected_nom:
        return go.Figure(
            layout=go.Layout(
                title="Выберите артикул или номенклатуру (и, при необходимости, склад)",
                xaxis_title="Дата",
                yaxis_title="Остаток"
            )
        )

    dff = df_2025_clean.copy()

    if selected_sklads:
        sklads = _to_list(selected_sklads)
        dff = dff[dff["Склад"].isin(sklads)]

    if selected_article:
        dff = dff[dff["Артикул_товар"].astype(str) == str(selected_article)]
    if selected_nom:
        dff = dff[dff["Номенклатура_канон"] == selected_nom]

    if dff.empty:
        return go.Figure(
            layout=go.Layout(
                title="Нет данных для выбранных фильтров",
                xaxis_title="Дата",
                yaxis_title="Остаток"
            )
        )

    fig = go.Figure()

    for sklad in dff["Склад"].unique():
        df_s = dff[dff["Склад"] == sklad].sort_values("Дата").copy()

        # Расчёт Продано и Пополнено с учётом пропусков дат
        df_s["Продано_fix"] = (df_s["Остаток"].shift(1) - df_s["Остаток"]).clip(lower=0).fillna(0)
        df_s["Пополнено_fix"] = (df_s["Остаток"] - df_s["Остаток"].shift(1)).clip(lower=0).fillna(0)

        # Скользящее среднее и всплески
        df_s["Среднее_Продано"] = df_s["Продано_fix"].rolling(window=7, min_periods=1).mean()
        df_s["Всплеск"] = df_s["Продано_fix"] > 1.5 * df_s["Среднее_Продано"]
        df_s["Цена_изменилась"] = df_s["Цена"].diff().fillna(0) != 0

        # Цвет маркеров
        df_s["Цвет"] = df_s.apply(
            lambda row: "purple" if row["Всплеск"] and row["Цена_изменилась"]
                        else "red" if row["Всплеск"]
                        else "orange" if row["Цена_изменилась"]
                        else "blue",
            axis=1
        )
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
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", marker=dict(size=8, color=color), name=label))

    fig.update_layout(
        title="Динамика остатков, продаж и цен (2025)",
        xaxis_title="Дата",
        yaxis_title="Остаток",
        hovermode="closest",
        legend=dict(orientation="h", y=-0.2)
    )
    return fig

# ------------------- Таблица топ-100 -------------------
# ------------------- Таблица ТОП-100 -------------------
@app.callback(
    Output("top-100-table", "data"),
    Input("sklad-2025-filter", "value")
)
def update_top_100_table(selected_sklads):
    df_filtered = df_2025.copy()
    if selected_sklads:
        df_filtered = df_filtered[df_filtered["Склад"].isin(selected_sklads)]

    # Группировка по артикулу + номенклатуре + складу
    top_100 = (
        df_filtered.groupby(["Артикул_товар", "Номенклатура_канон", "Склад"], as_index=False)["Продано"]
        .sum()
        .sort_values("Продано", ascending=False)
        .head(100)
    )

    # Переименовываем колонки под таблицу
    top_100 = top_100.rename(
        columns={
            "Артикул_товар": "Артикул",
            "Номенклатура_канон": "Номенклатура",
        }
    )

    return top_100.to_dict("records")

# ------------------- Выбор из таблицы -------------------
# ------------------- Выбор из таблицы -------------------
@app.callback(
    Output("article-2025-filter", "value"),
    Output("nom-2025-filter", "value"),
    Input("top-100-table", "selected_rows"),
    State("top-100-table", "data")
)
def select_from_table(selected_rows, table_data):
    if selected_rows:
        row = table_data[selected_rows[0]]
        return row["Артикул"], row["Номенклатура"]
    return None, None

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
