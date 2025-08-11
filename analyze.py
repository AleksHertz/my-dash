import os
import pandas as pd
import glob
from openpyxl import load_workbook
from datetime import datetime
import plotly.express as px
import logging
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import numpy as np
import time
import functools

# === НАСТРОЙКА ЛОГИРОВАНИЯ ===
os.makedirs("логи", exist_ok=True)
logging.basicConfig(
    filename='логи/анализ_склада.log',
    level=logging.INFO,
    format='%(asctime)s — %(levelname)s — %(message)s'
)

def timing_decorator(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        duration = end - start
        msg = f"Время выполнения функции '{func.__name__}': {duration:.3f} секунд"
        print(msg)
        logging.info(msg)
        return result
    return wrapper

@timing_decorator
def parse_date_from_cell(cell_value, file_path):
    """
    Безопасно парсит дату из значения ячейки.
    Если дата не распознана, берёт дату создания файла.
    Логирует результат.
    """
    if cell_value is None:
        date = datetime.fromtimestamp(os.path.getctime(file_path))
        logging.warning(f'Дата в файле {file_path} отсутствует, подставлена дата создания файла: {date}')
        return date

    date_str = str(cell_value).strip()
    date = pd.to_datetime(date_str, errors='coerce', dayfirst=True)
    if pd.isna(date):
        date = datetime.fromtimestamp(os.path.getctime(file_path))
        logging.warning(f'Дата "{date_str}" в файле {file_path} не распознана, подставлена дата создания файла: {date}')
    else:
        logging.info(f'Дата из файла {file_path}: "{date_str}" распознана как {date}')
    return date

@timing_decorator
def read_excel_file(file_path, sklad_name):
    extension = os.path.splitext(file_path)[1].lower()
    data = []

    try:
        if extension == '.xls':
            df_raw = pd.read_excel(file_path, header=None, engine='xlrd')
            date_cell = df_raw.iloc[1, 0]
            date = parse_date_from_cell(date_cell, file_path)

            for idx in range(4, len(df_raw)):
                row = df_raw.iloc[idx]
                if pd.isna(row[5]):
                    continue
                data.append({
                    'Дата': date,
                    'Номенклатура': row[1],
                    'Количество': row[2] or 0,
                    'Цена': row[3] or 0,
                    'Производитель': row[4],
                    'Артикул': str(row[5]).strip(),
                    'Склад': sklad_name
                })

        elif extension == '.xlsx':
            wb = load_workbook(filename=file_path, data_only=True)
            ws = wb.active
            date_cell = ws['A2'].value
            date = parse_date_from_cell(date_cell, file_path)

            for i in range(5, ws.max_row + 1):
                article = ws[f'F{i}'].value
                if article is None:
                    continue
                data.append({
                    'Дата': date,
                    'Номенклатура': ws[f'B{i}'].value,
                    'Количество': ws[f'C{i}'].value or 0,
                    'Цена': ws[f'D{i}'].value or 0,
                    'Производитель': ws[f'E{i}'].value,
                    'Артикул': str(article).strip(),
                    'Склад': sklad_name
                })

        logging.info(f'Успешно прочитан файл: {file_path}')
        return pd.DataFrame(data)

    except Exception as e:
        logging.error(f'Ошибка при чтении файла {file_path}: {e}')
        return None
@timing_decorator
def process_folder(folder_path, sklad_name):
    files = sorted(glob.glob(os.path.join(folder_path, '*.xls*')))
    dfs = []
    for file in files:
        df = read_excel_file(file, sklad_name)
        if df is not None and not df.empty:
            dfs.append(df)
    if not dfs:
        logging.warning(f'Нет данных в папке: {folder_path}')
        return None
    return pd.concat(dfs, ignore_index=True)
@timing_decorator
def generate_daily_sales_file(df_all: pd.DataFrame, output_path: str = 'итог_дневные_продажи.csv'):
    try:
        if 'Дата' not in df_all.columns:
            logging.error("Колонка 'Дата' отсутствует в данных — невозможно сформировать файл ежедневных продаж.")
            return

        if 'Количество' not in df_all.columns:
            logging.error("Колонка 'Количество' отсутствует в данных — невозможно сформировать файл ежедневных продаж.")
            return

        if 'Цена' not in df_all.columns:
            logging.error("Колонка 'Цена' отсутствует в данных — невозможно сформировать файл с ценами.")
            return

        # Агрегация количества и цены (цена — первая за день)
        df_daily = (
            df_all
            .sort_values('Дата')  # чтобы first() работал корректно
            .groupby(['Дата', 'Артикул', 'Склад'], as_index=False)
            .agg({
                'Количество': 'sum',
                'Цена': 'first'
            })
            .rename(columns={'Количество': 'Всего_продано', 'Цена': 'Цена_в_начале_дня'})
        )

        # Сортировка
        df_daily.sort_values(['Дата', 'Склад', 'Артикул'], inplace=True)

        # ✅ Сохранение в CSV
        df_daily.to_csv(output_path, index=False, encoding='utf-8-sig')
        logging.info(f"📁 CSV-файл с дневными продажами и ценами сохранён: {output_path}")

    except Exception as e:
        logging.error(f"❌ Ошибка при создании файла с дневными продажами: {e}")


@timing_decorator
def analyze_with_restock_vectorized_monthly(df_all):
    # Преобразуем дату в datetime и добавим колонки Год и Месяц
    df_all = df_all.sort_values(['Артикул', 'Склад', 'Дата']).copy()
    df_all['Дата'] = pd.to_datetime(df_all['Дата'], format='%d-%m-%Y')
    df_all['Год'] = df_all['Дата'].dt.year
    df_all['Месяц'] = df_all['Дата'].dt.month

    # Разница по количеству для подсчёта продано / пополнение
    df_all['diff_qty'] = df_all.groupby(['Артикул', 'Склад'])['Количество'].diff().fillna(0)

    # Продано — отрицательная разница
    df_all['Продано'] = (-df_all['diff_qty']).clip(lower=0)
    # Пополнение — положительная разница
    df_all['Пополнение'] = df_all['diff_qty'].clip(lower=0)

    # Поиск перемещений — пополнение на одном складе и уменьшение на другом в тот же день и с тем же артикулом и количеством
    restock_rows = df_all[df_all['diff_qty'] > 0][['Дата', 'Артикул', 'Склад', 'diff_qty']].copy()
    restock_rows.rename(columns={'Склад': 'Склад_куда', 'diff_qty': 'Кол-во'}, inplace=True)

    sold_rows = df_all[df_all['diff_qty'] < 0][['Дата', 'Артикул', 'Склад', 'diff_qty']].copy()
    sold_rows.rename(columns={'Склад': 'Склад_откуда', 'diff_qty': 'Кол-во'}, inplace=True)
    sold_rows['Кол-во'] = -sold_rows['Кол-во']

    merged_moves = pd.merge(
        restock_rows,
        sold_rows,
        on=['Дата', 'Артикул', 'Кол-во'],
        how='inner',
        suffixes=('_куда', '_откуда')
    )
    merged_moves = merged_moves[merged_moves['Склад_куда'] != merged_moves['Склад_откуда']]

    # Убираем перемещения из Пополнения и Продано
    for idx, row in merged_moves.iterrows():
        mask_restock = (
            (df_all['Дата'] == row['Дата']) &
            (df_all['Артикул'] == row['Артикул']) &
            (df_all['Склад'] == row['Склад_куда']) &
            (df_all['diff_qty'] == row['Кол-во'])
        )
        df_all.loc[mask_restock, 'Пополнение'] = 0

        mask_sold = (
            (df_all['Дата'] == row['Дата']) &
            (df_all['Артикул'] == row['Артикул']) &
            (df_all['Склад'] == row['Склад_откуда']) &
            (df_all['diff_qty'] == -row['Кол-во'])
        )
        df_all.loc[mask_sold, 'Продано'] = 0

    перемещения = merged_moves.rename(columns={
        'Дата': 'Дата',
        'Артикул': 'Артикул',
        'Склад_откуда': 'Склад_откуда',
        'Склад_куда': 'Склад_куда',
        'Кол-во': 'Кол-во'
    })[['Дата', 'Артикул', 'Склад_откуда', 'Склад_куда', 'Кол-во']].to_dict(orient='records')

    # Группировка по году, месяцу, складу и артикулу
    df_sales = df_all.groupby(['Артикул', 'Склад', 'Год', 'Месяц']).agg(
        Номенклатура=('Номенклатура', 'first'),
        Производитель=('Производитель', 'first'),
        Всего_продано=('Продано', 'sum'),
        Всего_пополнено=('Пополнение', 'sum'),
        Дней_продаж=('Продано', lambda x: (x > 0).sum()),
        Средняя_цена=('Цена', 'mean'),
        Мин_цена=('Цена', 'min'),
        Макс_цена=('Цена', 'max'),
        Дней_в_наличии=('Количество', lambda x: (x > 0).sum()),
        Последний_остаток=('Количество', 'last'),
        Уникальных_дней=('Дата', 'nunique')
    ).reset_index()

    # Коррекция: дней в наличии не может быть больше уникальных дней
    df_sales['Дней_в_наличии'] = df_sales[['Дней_в_наличии', 'Уникальных_дней']].min(axis=1)

    # Оборачиваемость по месяцу (чтобы не делить на 0, заменяем 0 на 1)
    df_sales['Оборачиваемость'] = df_sales['Всего_продано'] / df_sales['Дней_продаж'].replace(0, 1)

    # Цены в начале и конце месяца
    df_price = df_all.sort_values('Дата').groupby(['Артикул', 'Склад', 'Год', 'Месяц']).agg(
        Цена_в_начале=('Цена', 'first'),
        Цена_в_конце=('Цена', 'last')
    ).reset_index()

    # Подсчёт количества изменений цены по дням внутри месяца
    def count_price_changes_daily(group):
        daily_prices = group.groupby(group['Дата'].dt.date)['Цена'].first()
        # Считаем количество смен цены (считаем количество отличий между соседними днями)
        return daily_prices.ne(daily_prices.shift()).sum() - 1  # -1, чтобы не считать первое значение как изменение

    price_changes = (
        df_all.groupby(['Артикул', 'Склад', 'Год', 'Месяц'], group_keys=False)
        .apply(lambda g: pd.Series({'Кол_во_изменений_цены': count_price_changes_daily(g)}))
        .reset_index()
    )

    # Итоговое объединение
    df_result = pd.merge(df_sales, df_price, on=['Артикул', 'Склад', 'Год', 'Месяц'], how='left')
    df_result = pd.merge(df_result, price_changes, on=['Артикул', 'Склад', 'Год', 'Месяц'], how='left')

    df_result['Кол_во_изменений_цены'] = df_result['Кол_во_изменений_цены'].fillna(0).astype(int)
    df_result['Изменение_цены'] = df_result['Цена_в_конце'] - df_result['Цена_в_начале']
    df_result['Изменение_цены_%'] = df_result.apply(
        lambda row: ((row['Изменение_цены'] / row['Цена_в_начале']) * 100) if row['Цена_в_начале'] else 0,
        axis=1
    ).round(2)

    return df_result, перемещения
@timing_decorator
def detect_sales_peaks(df_all: pd.DataFrame,
                       output_path: str = 'всплески_продаж.xlsx',
                       multiplier: float = 2.0,
                       min_days_for_stats: int = 3) -> pd.DataFrame:
    """
    Находит всплески: дни, когда Всего_продано > Среднее + multiplier * СтдОтклонение
    Для стабильности:
      - СтдОтклонение для групп с < min_days_for_stats наблюдений не используется (пик не считается).
      - Номенклатура подтягивается как наиболее частая для пары (Артикул, Склад).
    Возвращает DataFrame найденных всплесков и сохраняет в Excel.
    """
    try:
        logging.info("📈 Поиск всплесков продаж...")

        # Проверка необходимых колонок
        required_cols = ['Дата', 'Артикул', 'Склад', 'Количество']
        if not all(col in df_all.columns for col in required_cols):
            logging.error(f"❌ Отсутствуют необходимые колонки: {required_cols}")
            return pd.DataFrame()

        df = df_all.copy()

        # Приводим типы и очищаем
        df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')
        df['Артикул'] = df['Артикул'].astype(str).str.strip()
        df['Склад'] = df['Склад'].astype(str).str.strip()
        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)

        df = df.dropna(subset=['Дата', 'Артикул', 'Склад'])
        if df.empty:
            logging.warning("Данные пусты после очистки (нет корректных Дата/Артикул/Склад).")
            return pd.DataFrame()

        # Агрегируем продажи по дате, артикулу, складу
        df_daily = (
            df.groupby(['Дата', 'Артикул', 'Склад'], as_index=False)['Количество']
              .sum()
              .rename(columns={'Количество': 'Всего_продано'})
        )

        # Статистика по паре (Артикул, Склад): mean, std, count
        stats = (
            df_daily.groupby(['Артикул', 'Склад'], as_index=False)['Всего_продано']
                    .agg(Среднее='mean', СтдОтклонение='std', Count='count')
        )
        # Заполняем NaN в std нулями (например, если только одно наблюдение)
        stats['СтдОтклонение'] = stats['СтдОтклонение'].fillna(0)

        # Объединяем stats с дневными продажами
        df_peaks = df_daily.merge(stats, on=['Артикул', 'Склад'], how='left')

        # Порог: если недостаточно наблюдений — ставим очень большой порог (чтобы не считать пик)
        df_peaks['threshold'] = df_peaks.apply(
            lambda r: (r['Среднее'] + multiplier * r['СтдОтклонение'])
                      if r['Count'] >= min_days_for_stats else np.inf,
            axis=1
        )

        df_peaks['Всплеск'] = df_peaks['Всего_продано'] > df_peaks['threshold']

        df_peaks_only = df_peaks[df_peaks['Всплеск']].copy()
        if df_peaks_only.empty:
            logging.info("Всплесков не найдено.")
            # Сохраняем пустой файл (по желанию) и возвращаем пустой DF
            pd.DataFrame().to_excel(output_path, index=False)
            return df_peaks_only

        # Выбираем надёжную Номенклатуру: наиболее частая для пары (Артикул, Склад)
        def most_common_nom(series):
            vals = series.dropna().astype(str)
            if vals.empty:
                return np.nan
            return vals.value_counts().index[0]

        sku_nom = (
            df.groupby(['Артикул', 'Склад'], as_index=False)['Номенклатура']
              .agg(most_common_nom)
              .rename(columns={'Номенклатура': 'Номенклатура_выбранная'})
        )

        # Мержим единственную номенклатуру — не будет множить строки
        df_peaks_only = df_peaks_only.merge(sku_nom, on=['Артикул', 'Склад'], how='left')

        # Удаляем возможные оставшиеся дубликаты по ключам (на всякий случай)
        df_peaks_only = df_peaks_only.drop_duplicates(subset=['Дата', 'Артикул', 'Склад'])

        # Форматируем дату для Excel
        df_peaks_only['Дата'] = df_peaks_only['Дата'].dt.strftime('%d/%m/%Y')

        # Переименуем колонку с выбранной номенклатурой в привычное имя
        df_peaks_only = df_peaks_only.rename(columns={'Номенклатура_выбранная': 'Номенклатура'})

        # Упорядочим колонки для читаемости
        cols_order = ['Дата', 'Артикул', 'Номенклатура', 'Склад', 'Всего_продано', 'Среднее', 'СтдОтклонение', 'Count', 'threshold', 'Всплеск']
        cols = [c for c in cols_order if c in df_peaks_only.columns] + [c for c in df_peaks_only.columns if c not in cols_order]
        df_peaks_only = df_peaks_only[cols]

        # Сохраняем
        df_peaks_only.to_excel(output_path, index=False)
        logging.info(f"✅ Файл со всплесками продаж сохранен: {output_path} (строк: {len(df_peaks_only)})")

        return df_peaks_only

    except Exception as e:
        logging.exception(f"❌ Ошибка при анализе всплесков продаж: {e}")
        return pd.DataFrame()

@timing_decorator
def run_month_analysis():
    logging.info("🔍 Начало анализа месяца")

    df_moscow = process_folder('data/moscow', 'Москва')
    df_khabarovsk = process_folder('data/khabarovsk', 'Хабаровск')

    if df_moscow is None and df_khabarovsk is None:
        logging.error("❌ Нет данных для анализа")
        return

    # Объединение и базовая очистка
    df_all = pd.concat([df for df in [df_moscow, df_khabarovsk] if df is not None], ignore_index=True)
    df_all.dropna(subset=['Артикул'], inplace=True)
    df_all['Артикул'] = df_all['Артикул'].astype(str).str.strip()

    # Обработка даты
    if 'Дата' in df_all.columns:
        df_all['Дата'] = pd.to_datetime(df_all['Дата'], errors='coerce')

    # Основной анализ
    df_result, df_flags = analyze_with_restock_vectorized_monthly(df_all)
    if df_result is None:
        logging.error("❌ Анализ не выполнен — df_result пуст")
        return

    if isinstance(df_flags, list):
        if len(df_flags) > 0:
            df_flags = pd.DataFrame(df_flags)
        else:
            df_flags = pd.DataFrame()

    if not df_flags.empty and 'Дата' in df_flags.columns:
        df_flags['Дата'] = pd.to_datetime(df_flags['Дата'], errors='coerce').dt.strftime('%d/%m/%Y')

    # Сохранение итоговых файлов
    df_result.to_excel('итог_по_месяцу.xlsx', index=False)
    df_flags.to_excel('фиксация_перемещений.xlsx', index=False)

    # Топы
    top_fast = df_result[df_result['Всего_продано'] > 0].sort_values('Оборачиваемость', ascending=False).head(1000)
    top_slow = df_result[df_result['Всего_продано'] == 0].sort_values('Дней_в_наличии', ascending=False).head(1000)
    top_restocked = df_result.sort_values('Всего_пополнено', ascending=False).head(1000)

    top_fast.to_excel('самые_ходовые.xlsx', index=False)
    top_slow.to_excel('залежалые.xlsx', index=False)
    top_restocked.to_excel('чаще_всего_пополнялись.xlsx', index=False)


    # Сохранение ежедневных продаж
    generate_daily_sales_file(df_all, output_path='итог_дневные_продажи.csv')

    # Всплески продаж
    detect_sales_peaks(df_all, output_path='всплески_продаж.xlsx')

    logging.info("✅ Анализ месяца завершен")


if __name__ == '__main__':
    run_month_analysis()