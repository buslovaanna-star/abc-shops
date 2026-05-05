"""
shop_engine.py — рушій ABC-аналізу по магазинах
================================================
Читає файл: аркуш 'акції' + місячні аркуші (кожен = окремий місяць)
Структура місячного аркуша:
  Row 12 (0-indexed): заголовки — col0=Назва, col1=Артикул, col2=Група, col3=Клас, col4..N-1=Магазини, colN=Підсумок
  Row 15+: дані
  Остання row: Підсумок (пропускається)
"""

import re
import pandas as pd
import numpy as np
from datetime import datetime

HEADER_ROW   = 12   # 0-indexed
DATA_START   = 15
SERVICE_COLS = 4    # Назва, Артикул, Група, Клас

UA_MONTH_MAP = {
    'Січень': ('Січ', 1), 'Лютий': ('Лют', 2), 'Березень': ('Бер', 3),
    'Квітень': ('Кві', 4), 'Травень': ('Тра', 5), 'Червень': ('Чер', 6),
    'Липень': ('Лип', 7), 'Серпень': ('Сер', 8), 'Вересень': ('Вер', 9),
    'Жовтень': ('Жов', 10), 'Листопад': ('Лис', 11), 'Грудень': ('Гру', 12),
}


def parse_sheet_name(name: str):
    """Parse 'Квітень 2026 р.' -> ('Кві 26', datetime(2026,4,1))"""
    clean = name.replace('\xa0', ' ').replace('р.', '').strip()
    parts = clean.split()
    if len(parts) < 2:
        return None, None
    ua_month = parts[0]
    year_str = parts[1]
    if ua_month not in UA_MONTH_MAP:
        return None, None
    short, month_num = UA_MONTH_MAP[ua_month]
    year = int(year_str)
    label = f"{short} {str(year)[2:]}"
    dt = datetime(year, month_num, 1)
    return label, dt


def load_promo(path: str) -> pd.DataFrame:
    """Load promo sheet. Returns DataFrame with columns: Артикул, Старт, Кінець"""
    try:
        promo_sheets = ['акції', 'Sheet2', 'Аркуш2', 'акция', 'promo']
        xl = pd.ExcelFile(path)
        sheet = next((s for s in xl.sheet_names
                      if s.lower().replace('ї','і') in [p.lower() for p in promo_sheets]), None)
        if sheet is None:
            return pd.DataFrame(columns=['Артикул', 'Старт', 'Кінець'])
        df = pd.read_excel(path, sheet_name=sheet, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        col_art   = next((c for c in df.columns if 'АРТИКУЛ' in c.upper()), None)
        col_start = next((c for c in df.columns if 'ПОЧАТК' in c.upper() or 'START' in c.upper()), None)
        col_end   = next((c for c in df.columns if 'ЗАКІНЧ' in c.upper() or 'END' in c.upper()), None)
        if not col_art and len(df.columns) > 3: col_art = df.columns[3]
        if not col_start: col_start = df.columns[0]
        if not col_end and len(df.columns) > 1: col_end = df.columns[1]

        records = []
        for _, row in df.iterrows():
            raw = str(row[col_art]).replace('\n', ' ')
            for art in raw.split():
                art = art.strip()
                if art and art.lower() not in ('nan', 'none', ''):
                    try:
                        records.append({
                            'Артикул': art,
                            'Старт':   pd.to_datetime(row[col_start]),
                            'Кінець':  pd.to_datetime(row[col_end]),
                        })
                    except Exception:
                        pass
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame(columns=['Артикул', 'Старт', 'Кінець'])


def build_promo_map(promo_df: pd.DataFrame, month_dates: dict) -> dict:
    """Returns {article: set_of_promo_month_labels}"""
    result = {}
    for art in promo_df['Артикул'].unique():
        pm = set()
        for _, pr in promo_df[promo_df['Артикул'] == art].iterrows():
            for label, (ms, me) in month_dates.items():
                if pr['Старт'] <= me and pr['Кінець'] >= ms:
                    pm.add(label)
        result[art] = pm
    return result


def load_all_months(path: str):
    """
    Returns:
      df_long     — long DataFrame: Артикул, Назва, Група, Місяць, Магазин, ВП
      shops       — list of shop names (from last month sheet)
      month_labels — list of month labels in order
      month_dates  — {label: (start_ts, end_ts)}
      promo_df    — promo DataFrame
    """
    xl = pd.ExcelFile(path)

    # Detect month sheets (all except promo)
    month_sheets = []
    for sname in xl.sheet_names:
        label, dt = parse_sheet_name(sname)
        if label and dt:
            month_sheets.append((sname, label, dt))
    # Sort chronologically
    month_sheets.sort(key=lambda x: x[2])

    if not month_sheets:
        raise ValueError("Не знайдено жодного місячного аркуша")

    # Build month_dates dict
    month_dates = {}
    for _, label, dt in month_sheets:
        # last day of month
        if dt.month == 12:
            end = datetime(dt.year + 1, 1, 1) - pd.Timedelta(days=1)
        else:
            end = datetime(dt.year, dt.month + 1, 1) - pd.Timedelta(days=1)
        month_dates[label] = (pd.Timestamp(dt), pd.Timestamp(end))

    # Use LAST month's sheet for canonical shop list
    last_sname = month_sheets[-1][0]
    df_last = pd.read_excel(path, sheet_name=last_sname, header=None)
    shops_row = df_last.iloc[HEADER_ROW].tolist()
    shops_canonical = []
    for j in range(SERVICE_COLS, df_last.shape[1] - 1):
        v = str(shops_row[j]).replace('\xa0', ' ').strip()
        if v not in ('nan', 'None', ''):
            shops_canonical.append(v.replace('Магазин - ', '').strip())

    # Load promo
    promo_df = load_promo(path)
    promo_map = build_promo_map(promo_df, month_dates)

    # Load each month
    all_rows = []
    month_labels = []

    for sname, label, dt in month_sheets:
        month_labels.append(label)
        df_raw = pd.read_excel(path, sheet_name=sname, header=None)

        # Get shop columns for this month (may differ from canonical)
        shops_row_m = df_raw.iloc[HEADER_ROW].tolist()
        col_to_shop = {}
        for j in range(SERVICE_COLS, df_raw.shape[1] - 1):
            v = str(shops_row_m[j]).replace('\xa0', ' ').strip()
            if v not in ('nan', 'None', ''):
                col_to_shop[j] = v.replace('Магазин - ', '').strip()

        for i in range(DATA_START, df_raw.shape[0] - 1):
            row = df_raw.iloc[i]
            art = str(row.iloc[1]).replace('\xa0', '').strip()
            if art in ('nan', 'None', '', 'Підсумок'):
                continue
            nazva = str(row.iloc[0]).replace('\xa0', ' ').strip()
            grupa = str(row.iloc[2]).replace('\xa0', ' ').strip()
            is_promo = label in promo_map.get(art, set())

            for j, shop_name in col_to_shop.items():
                raw_val = str(row.iloc[j]).replace('\xa0', '').replace(' ', '').replace(',', '.')
                try:
                    vp = float(raw_val)
                except Exception:
                    vp = 0.0
                if vp != 0.0:
                    all_rows.append({
                        'Артикул':  art,
                        'Назва':    nazva,
                        'Група':    grupa,
                        'Місяць':   label,
                        'Магазин':  shop_name,
                        'ВП':       vp,
                        'is_promo': is_promo,
                    })

    df_long = pd.DataFrame(all_rows)
    return df_long, shops_canonical, month_labels, month_dates, promo_df


def compute_abc(df_items: pd.DataFrame, value_col='ВП_clean',
                a_thresh=80, b_thresh=95) -> pd.DataFrame:
    """
    Input: DataFrame with value_col per SKU.
    Returns same DataFrame with added columns: Частка%, Кумул%, ABC.
    """
    df = df_items.copy()
    df = df[df[value_col] > 0].sort_values(value_col, ascending=False).reset_index(drop=True)
    total = df[value_col].sum()
    if total == 0:
        df['Частка%'] = 0.0
        df['Кумул%']  = 0.0
        df['ABC']     = 'C'
        return df
    df['Частка%'] = df[value_col] / total * 100
    df['Кумул%']  = df['Частка%'].cumsum()
    df['ABC'] = df['Кумул%'].apply(
        lambda x: 'A' if x <= a_thresh else ('B' if x <= b_thresh else 'C'))
    df['№'] = range(1, len(df) + 1)
    return df


def build_global_abc(df_long: pd.DataFrame, promo_map: dict,
                     a_thresh=80, b_thresh=95) -> pd.DataFrame:
    """ABC for all shops combined, excluding promo months."""
    rows = []
    for art, grp in df_long.groupby('Артикул'):
        pm = promo_map.get(art, set())
        clean = grp[~grp['Місяць'].isin(pm)]
        vp = clean['ВП'].sum()
        rows.append({
            'Артикул': art,
            'Назва':   grp['Назва'].iloc[0],
            'Група':   grp['Група'].iloc[0],
            'ВП_clean': round(vp, 2),
            'Акц_міс': ', '.join(sorted(pm)) if pm else '',
            'N_promo':  len(pm),
        })
    df = pd.DataFrame(rows)
    return compute_abc(df, 'ВП_clean', a_thresh, b_thresh)


def build_shop_abc(df_long: pd.DataFrame, shop: str, promo_map: dict,
                   a_thresh=80, b_thresh=95) -> pd.DataFrame:
    """ABC for a single shop, excluding promo months."""
    df_shop = df_long[df_long['Магазин'] == shop]
    rows = []
    for art, grp in df_shop.groupby('Артикул'):
        pm = promo_map.get(art, set())
        clean = grp[~grp['Місяць'].isin(pm)]
        vp = clean['ВП'].sum()
        if vp > 0:
            rows.append({
                'Артикул': art,
                'Назва':   grp['Назва'].iloc[0],
                'Група':   grp['Група'].iloc[0],
                'ВП_clean': round(vp, 2),
                'Акц_міс': ', '.join(sorted(pm)) if pm else '',
                'N_promo':  len(pm),
            })
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    return compute_abc(df, 'ВП_clean', a_thresh, b_thresh)


def build_group_abc(global_abc: pd.DataFrame, a_thresh=80, b_thresh=95) -> pd.DataFrame:
    """ABC of nomenclature groups based on global ABC data."""
    grp = global_abc.groupby('Група').agg(
        SKU=('Артикул', 'count'),
        ВП=('ВП_clean', 'sum'),
        A_sku=('ABC', lambda x: (x == 'A').sum()),
        B_sku=('ABC', lambda x: (x == 'B').sum()),
        C_sku=('ABC', lambda x: (x == 'C').sum()),
    ).reset_index().sort_values('ВП', ascending=False)
    total = grp['ВП'].sum()
    grp['Частка%'] = grp['ВП'] / total * 100
    grp['Кумул%']  = grp['Частка%'].cumsum()
    grp['ABC_гр']  = grp['Кумул%'].apply(
        lambda x: 'A' if x <= a_thresh else ('B' if x <= b_thresh else 'C'))
    grp['№'] = range(1, len(grp) + 1)
    return grp


def build_monthly_by_shop(df_long: pd.DataFrame, promo_map: dict,
                           month_labels: list, shops: list) -> pd.DataFrame:
    """Monthly VP matrix: index=month, columns=shops."""
    result = pd.DataFrame(0.0, index=month_labels, columns=shops)
    for art, grp in df_long.groupby('Артикул'):
        pm = promo_map.get(art, set())
        clean = grp[~grp['Місяць'].isin(pm)]
        for _, row in clean.iterrows():
            m = row['Місяць']
            s = row['Магазин']
            if m in result.index and s in result.columns:
                result.loc[m, s] += row['ВП']
    return result.round(2)


def build_monthly_by_group(df_long: pd.DataFrame, promo_map: dict,
                            month_labels: list) -> pd.DataFrame:
    """Monthly VP matrix: index=month, columns=groups."""
    groups = sorted(df_long['Група'].unique())
    result = pd.DataFrame(0.0, index=month_labels, columns=groups)
    for art, grp in df_long.groupby('Артикул'):
        pm = promo_map.get(art, set())
        clean = grp[~grp['Місяць'].isin(pm)]
        for _, row in clean.iterrows():
            m = row['Місяць']
            g = row['Група']
            if m in result.index and g in result.columns:
                result.loc[m, g] += row['ВП']
    return result.round(2)


def build_shop_summary(df_long: pd.DataFrame, promo_map: dict,
                        shops: list, a_thresh=80, b_thresh=95) -> pd.DataFrame:
    """Summary table: one row per shop with VP, SKU counts, ABC breakdown."""
    rows = []
    for shop in shops:
        sdf = build_shop_abc(df_long, shop, promo_map, a_thresh, b_thresh)
        if sdf.empty:
            continue
        sm = sdf.groupby('ABC')['ВП_clean'].agg(['count', 'sum'])
        total_vp = sdf['ВП_clean'].sum()
        rows.append({
            'Магазин':   shop,
            'SKU':       len(sdf),
            'A SKU':     int(sm.loc['A', 'count']) if 'A' in sm.index else 0,
            'B SKU':     int(sm.loc['B', 'count']) if 'B' in sm.index else 0,
            'C SKU':     int(sm.loc['C', 'count']) if 'C' in sm.index else 0,
            'ВП (USD)':  round(total_vp, 2),
            'A ВП':      round(sm.loc['A', 'sum'] if 'A' in sm.index else 0, 2),
            'B ВП':      round(sm.loc['B', 'sum'] if 'B' in sm.index else 0, 2),
            'C ВП':      round(sm.loc['C', 'sum'] if 'C' in sm.index else 0, 2),
        })
    df = pd.DataFrame(rows).sort_values('ВП (USD)', ascending=False).reset_index(drop=True)
    total = df['ВП (USD)'].sum()
    df['Частка%'] = df['ВП (USD)'] / total * 100
    return df
