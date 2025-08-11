import pandas as pd

def load_monthly_data(filepath='итог_по_месяцу.xlsx'):
    df = pd.read_excel(filepath)
    df['Дата'] = pd.to_datetime(df['Год'].astype(str) + '-' + df['Месяц'].astype(str) + '-01')
    df = df.sort_values(['Артикул', 'Склад', 'Дата']).reset_index(drop=True)
    return df

def find_sales_spikes(df, window=3, spike_factor=1.5):
    df = df.copy()
    df['Среднее_последних_месяцев'] = df.groupby(['Артикул', 'Склад'])['Всего_продано']\
        .transform(lambda x: x.rolling(window=window, min_periods=1).mean().shift(1))
    df['Всплеск'] = df['Всего_продано'] > (df['Среднее_последних_месяцев'] * spike_factor)
    df['Всплеск'] = df['Всплеск'].fillna(False)
    return df

def prepare_spike_analysis(filepath='итог_по_месяцу.xlsx', window=3, spike_factor=1.5):
    df = load_monthly_data(filepath)
    df_spikes = find_sales_spikes(df, window, spike_factor)
    return df_spikes

def save_analysis(df, output_path='всплески_продаж1.xlsx'):
    """
    Сохраняет DataFrame с анализом всплесков в Excel файл.
    """
    df.to_excel(output_path, index=False)
    print(f'Результат сохранён в файл: {output_path}')

if __name__ == '__main__':
    df_spikes = prepare_spike_analysis()
    save_analysis(df_spikes)