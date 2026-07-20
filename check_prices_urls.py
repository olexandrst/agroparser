"""Аналіз цін між нашим магазином (Агродрузі) та конкурентами.

Логіка (див. README.md):

1. Для кожного рядка Input.xlsx беремо значення колонки "Артикул" і шукаємо
   співпадіння в таблицях конкурентів (agrodoctor_prices.xlsx, tvk_ramos_prices.xlsx).
   1.1 Шукаємо співпадіння в колонках name, references, description.
   1.2 Якщо повного співпадіння немає і артикул містить крапку "." — шукаємо по
       значенню артикула до крапки (напр. "AZ10331.41" -> "AZ10331").
2. Якщо товар знайдено і найнижча ціна конкурента ("price") нижча за нашу
   ("Ціна Агродрузі") — пишемо цю ціну в "Рекомендація", "availability" в
   "Наявність", назву конкурента в "Джерело", посилання в "Посилання".
   2.1 "Наявність" == "Немає в наявності" -> рядок голубий (#80bcd7).
   2.2 "Наявність" != "Немає в наявності" -> рядок зелений (#82d200).
3. Товар знайдено, але найнижча ціна >= нашій -> "Ціна оптимальна", жовтий (#fffc00).
4. Товар не знайдено -> "Не знайдено", червоний (#ff6f6d).
"""

import re
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- Конфігурація ---------------------------------------------------------

INPUT_FILE = 'Input.xlsx'
OUTPUT_FILE = 'Input_analyzed_with_urls.xlsx'

COMPETITORS = [
    ('agrodoctor', 'agrodoctor_prices.xlsx'),
    ('tvk-ramos', 'tvk_ramos_prices.xlsx'),
]

# Колонки конкурентів, у яких шукаємо артикул.
SEARCH_COLUMNS = ['name', 'references', 'description']

# Колонки нашого файлу.
COL_ARTICLE = 'Артикул'
COL_OUR_PRICE = 'Ціна Агродрузі'
COL_RECOMMENDATION = 'Рекомендація'
COL_AVAILABILITY = 'Наявність'
COL_SOURCE = 'Джерело'
COL_URL = 'Посилання'

OUT_OF_STOCK = 'Немає в наявності'

# Кольори (ARGB для openpyxl).
FILL_BLUE = PatternFill('solid', fgColor='FF80BCD7')    # знайдено, немає в наявності
FILL_GREEN = PatternFill('solid', fgColor='FF82D200')   # знайдено, в наявності
FILL_YELLOW = PatternFill('solid', fgColor='FFFFFC00')  # ціна оптимальна
FILL_RED = PatternFill('solid', fgColor='FFFF6F6D')     # не знайдено

# Мінімальна довжина "базового" (до крапки) артикула. Захищає від надто
# загальних токенів на кшталт "38" (з "38.4V/2K1/JA/J2A"), які дали б тисячі
# хибних співпадінь.
MIN_BASE_LEN = 3


# --- Допоміжні функції ----------------------------------------------------

def clean_price(price):
    """Конвертація ціни у число. Непарсабельне значення -> inf."""
    if pd.isna(price):
        return float('inf')
    if isinstance(price, (int, float)):
        return float(price)
    price_str = str(price).replace(' ', '').replace('\xa0', '').replace(',', '.')
    try:
        return float(price_str)
    except ValueError:
        return float('inf')


def split_articles(raw):
    """Один рядок 'Артикул' може містити кілька артикулів через кому/крапку з комою."""
    if pd.isna(raw):
        return []
    text = str(raw).replace(';', ',')
    return [part.strip(' *') for part in text.split(',') if part.strip(' *')]


def base_article(article):
    """Значення артикула до останньої крапки (п.1.2 README).

    Повертає None, якщо крапки немає або база надто коротка/загальна.
    """
    if '.' not in article:
        return None
    base = article.rsplit('.', 1)[0].strip()
    if len(base) < MIN_BASE_LEN or base == article:
        return None
    return base


def combined_text(df):
    """Обʼєднаний текст пошукових колонок для кожного рядка конкурента."""
    parts = [df[col].fillna('').astype(str) for col in SEARCH_COLUMNS]
    joined = parts[0]
    for extra in parts[1:]:
        joined = joined + ' ' + extra
    return joined.tolist()


# --- Основна логіка -------------------------------------------------------

def build_token_map(df_input):
    """token -> список (row_idx, article_idx, level).

    level 0 = повне співпадіння артикула, level 1 = співпадіння по базі (до крапки).
    Рішення "повне чи база" приймається окремо для кожного артикула.
    """
    token_map = defaultdict(list)
    row_art_count = {}

    for row_idx, raw in enumerate(df_input[COL_ARTICLE]):
        articles = split_articles(raw)
        row_art_count[row_idx] = len(articles)
        for art_idx, article in enumerate(articles):
            token_map[article].append((row_idx, art_idx, 0))
            base = base_article(article)
            if base:
                token_map[base].append((row_idx, art_idx, 1))

    return token_map, row_art_count


def build_matcher(tokens):
    """Один регекс на всі токени. Довші — раніше, щоб префікс не "з'їдав" повний.

    Межі \\w з обох боків: артикул має бути цілим токеном, а не частиною довшого
    числа/слова (напр. "560210" не має ловитись усередині "0005602100").
    """
    ordered = sorted(tokens, key=len, reverse=True)
    pattern = r'(?<!\w)(?:' + '|'.join(re.escape(t) for t in ordered) + r')(?!\w)'
    return re.compile(pattern)


def collect_hits(token_map, matcher):
    """Прохід по всіх конкурентах. Повертає (row_idx, art_idx, level) -> [candidate]."""
    hits = defaultdict(list)

    for source, filename in COMPETITORS:
        df = pd.read_excel(filename)
        texts = combined_text(df)
        prices = df['price'].tolist()
        avails = df['availability'].tolist()
        urls = df['product_url'].tolist()

        for text, price, avail, url in zip(texts, prices, avails, urls):
            price = clean_price(price)
            if price == float('inf'):
                continue
            matched = {m.group(0) for m in matcher.finditer(text)}
            if not matched:
                continue
            candidate = (price, avail, source, url)
            for token in matched:
                for key in token_map.get(token, ()):
                    hits[key].append(candidate)

    return hits


def analyze(df_input, hits, row_art_count):
    """Формує колонки результату та колір для кожного рядка."""
    recommendations, availabilities, sources, urls, fills = [], [], [], [], []

    for row_idx in range(len(df_input)):
        our_price = clean_price(df_input.at[row_idx, COL_OUR_PRICE])

        # Для кожного артикула рядка: повні співпадіння, або (якщо їх нема) — база.
        pool = []
        for art_idx in range(row_art_count[row_idx]):
            full = hits.get((row_idx, art_idx, 0), [])
            base = hits.get((row_idx, art_idx, 1), [])
            pool.extend(full if full else base)

        if not pool:
            recommendations.append('Не знайдено')
            availabilities.append(None)
            sources.append(None)
            urls.append(None)
            fills.append(FILL_RED)
            continue

        best_price, best_avail, best_source, best_url = min(pool, key=lambda c: c[0])

        if best_price < our_price:
            recommendations.append(best_price)
            availabilities.append(best_avail)
            sources.append(best_source)
            urls.append(best_url)
            in_stock = str(best_avail).strip() != OUT_OF_STOCK
            fills.append(FILL_GREEN if in_stock else FILL_BLUE)
        else:
            recommendations.append('Ціна оптимальна')
            availabilities.append(None)
            sources.append(None)
            urls.append(None)
            fills.append(FILL_YELLOW)

    df_input[COL_RECOMMENDATION] = recommendations
    df_input[COL_AVAILABILITY] = availabilities
    df_input[COL_SOURCE] = sources
    df_input[COL_URL] = urls
    return fills


def write_output(df_input, fills):
    """Записує результат і фарбує всі клітинки кожного рядка."""
    df_input.to_excel(OUTPUT_FILE, index=False)

    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active
    n_cols = len(df_input.columns)
    for row_idx, fill in enumerate(fills):
        excel_row = row_idx + 2  # +1 заголовок, +1 бо openpyxl рахує з 1
        for col_idx in range(1, n_cols + 1):
            ws.cell(row=excel_row, column=col_idx).fill = fill
    wb.save(OUTPUT_FILE)


def main():
    df_input = pd.read_excel(INPUT_FILE)

    token_map, row_art_count = build_token_map(df_input)
    matcher = build_matcher(token_map.keys())
    hits = collect_hits(token_map, matcher)
    fills = analyze(df_input, hits, row_art_count)
    write_output(df_input, fills)

    # Коротка статистика.
    recs = df_input[COL_RECOMMENDATION]
    found_lower = sum(1 for r in recs if isinstance(r, (int, float)))
    optimal = int((recs == 'Ціна оптимальна').sum())
    not_found = int((recs == 'Не знайдено').sum())
    print(f'Аналіз завершено. Результати збережено у {OUTPUT_FILE}')
    print(f'  Знайдено дешевше в конкурентів: {found_lower}')
    print(f'  Ціна оптимальна: {optimal}')
    print(f'  Не знайдено: {not_found}')


if __name__ == '__main__':
    main()
