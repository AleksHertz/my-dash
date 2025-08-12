import dash
from dash import dcc, html, Input, Output, State
import plotly.express as px
import plotly.graph_objs as go
import pandas as pd
import os
from dash import ctx
import io
import dash_bootstrap_components as dbc
import logging
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

# --------------------
# DASH APP
# --------------------
app = dash.Dash(__name__)
server = app.server
app.layout = html.Div([
    html.H1("Анализ складских данных"),

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

        # Выбор топа по количеству для ходовых
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
            style={
                'height': '700px',
                'overflowY': 'scroll',
                'border': '1px solid #ddd',
                'padding': '5px',
                'marginBottom': '10px',
                'backgroundColor': 'white'
            }
        ),
        dbc.Button("📥 Выгрузить топ ходовых в Excel", id="download-top-fast-btn", color="success", className="mb-4"),

        # Выбор топа по количеству для пополнений (добавлен отдельный селектор)
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
            style={
                'height': '700px',
                'overflowY': 'scroll',
                'border': '1px solid #ddd',
                'padding': '5px',
                'marginBottom': '10px',
                'backgroundColor': 'white'
            }
        ),
        dbc.Button("📥 Выгрузить топ пополнений в Excel", id="download-top-restock-btn", color="success"),

        # Компоненты для скачивания файлов
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
        ], style={
            'maxWidth': 450,
            'marginBottom': 30,
            'display': 'flex',
            'flexDirection': 'column',
            'gap': '10px'
        }),

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
])
# --------------------
# КОЛБЭКИ
# --------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Выгрузка топ-ходовых ---
@app.callback(
    Output("download-top-fast", "data"),
    Input("download-top-fast-btn", "n_clicks"),
    State("sklad-filter", "value"),  # заменили ID
    State("top-n-selector", "value"),  # заменили ID
    prevent_initial_call=True
)
def export_top_fast_to_excel(n_clicks, selected_sklads, top_n):
    if df_fast.empty or not selected_sklads:
        return None

    dff = df_fast[df_fast['Склад'].isin(selected_sklads)]
    dff = dff.sort_values('Всего_продано', ascending=False).head(top_n)

    if 'Цена_в_начале' in dff.columns and 'Цена_в_конце' in dff.columns:
        dff['Изменение_цены_%'] = (
            (dff['Цена_в_конце'] - dff['Цена_в_начале']) / dff['Цена_в_начале'] * 100
        ).round(2)

    if 'Средний_остаток' in dff.columns:
        dff['Оборачиваемость'] = (dff['Всего_продано'] / dff['Средний_остаток']).round(2)

    return dcc.send_data_frame(dff.to_excel, f"топ_{top_n}_ходовые.xlsx", index=False)


@app.callback(
    Output("download-top-restock", "data"),
    Input("download-top-restock-btn", "n_clicks"),
    State("sklad-filter", "value"),  # заменили ID
    State("top-n-selector-restock", "value"),
    prevent_initial_call=True
)
def export_top_restock_to_excel(n_clicks, selected_sklads, top_n):
    if df_restock.empty or not selected_sklads:
        return None

    dff = df_restock[df_restock['Склад'].isin(selected_sklads)]
    dff = dff.sort_values('Всего_пополнено', ascending=False).head(top_n)

    if 'Цена_в_начале' in dff.columns and 'Цена_в_конце' in dff.columns:
        dff['Изменение_цены_%'] = (
            (dff['Цена_в_конце'] - dff['Цена_в_начале']) / dff['Цена_в_начале'] * 100
        ).round(2)

    if 'Средний_остаток' in dff.columns:
        dff['Оборачиваемость'] = (dff['Всего_продано'] / dff['Средний_остаток']).round(2)

    return dcc.send_data_frame(dff.to_excel, f"топ_{top_n}_пополнения.xlsx", index=False)


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

    # высота графика пропорциональна количеству элементов
    graph_height = HEIGHT_PER_BAR * len(dff)
    # ограничение сверху максимальной высотой контейнера
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
        yaxis={'categoryorder': 'total ascending'},
        template='plotly_white',
        margin=dict(l=250),
    )
    return fig


@app.callback(
    Output('graph-top-restock', 'figure'),
    Input('sklad-filter', 'value'),
    Input('top-n-selector', 'value'),
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
        yaxis={'categoryorder': 'total ascending'},
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
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8050)))
