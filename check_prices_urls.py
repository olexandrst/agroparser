import pandas as pd

def parse_skus(text):
    """Парсинг артикулів за заданими правилами."""
    if pd.isna(text):
        return []
    
    text = str(text).replace('Код товару:', '').replace('Складський код:', '').replace('Дод. артикули:', '')
    text = text.replace(';', ',')
    
    parts = text.split(',')
    res = []
    for p in parts:
        p = p.strip(' *')
        if not p:
            continue
            
        if '.' in p:
            p = p.rsplit('.', 1)[0]
            
        res.append(p.strip())
        
    return res

def clean_price(price):
    """Конвертація ціни у числовий формат."""
    if pd.isna(price):
        return float('inf')
    if isinstance(price, (int, float)):
        return float(price)
    
    price_str = str(price).replace(' ', '').replace('\xa0', '').replace(',', '.')
    try:
        return float(price_str)
    except ValueError:
        return float('inf')

def main():
    df_input = pd.read_excel('Input.xlsx')
    df_agro = pd.read_excel('agrodoctor_prices.xlsx')
    df_tvk = pd.read_excel('tvk_ramos_prices.xlsx')
    
    competitors_data = {} 
    
    # Тепер словник зберігає кортеж (price, source, url)
    def add_to_dict(sku_list, price, source, url):
        for sku in sku_list:
            if not sku: continue
            if sku in competitors_data:
                if price < competitors_data[sku][0]:
                    competitors_data[sku] = (price, source, url)
            else:
                competitors_data[sku] = (price, source, url)

    # Збір даних Agrodoctor
    for _, row in df_agro.iterrows():
        price = clean_price(row.get('price'))
        url = row.get('product_url')
        if price == float('inf'): continue
        
        skus = parse_skus(row.get('sku')) + parse_skus(row.get('references'))
        add_to_dict(skus, price, 'agrodoctor', url)
        
    # Збір даних TVK-Ramos
    for _, row in df_tvk.iterrows():
        price_col = 'price_with_vat' if 'price_with_vat' in row else 'price'
        price = clean_price(row.get(price_col))
        url = row.get('product_url')
        if price == float('inf'): continue
        
        skus = parse_skus(row.get('sku')) + parse_skus(row.get('references'))
        add_to_dict(skus, price, 'tvk-ramos', url)

    # Порівняння
    recommended_prices = []
    sources = []
    urls = []
    
    for _, row in df_input.iterrows():
        our_price = clean_price(row.get('Ціна Агродрузі'))
        our_skus = parse_skus(row.get('Артикул'))
        
        best_comp_price = float('inf')
        best_source = None
        best_url = None
        
        for sku in our_skus:
            if sku in competitors_data:
                comp_price, comp_source, comp_url = competitors_data[sku]
                if comp_price < best_comp_price:
                    best_comp_price = comp_price
                    best_source = comp_source
                    best_url = comp_url
                    
        # Записуємо рекомендацію, якщо ціна нижча
        if best_comp_price < our_price:
            recommended_prices.append(best_comp_price)
            sources.append(best_source)
            urls.append(best_url)
        else:
            recommended_prices.append(None)
            sources.append(None)
            urls.append(None)
            
    # Запис результатів
    df_input['Рекомендація'] = recommended_prices
    df_input['Джерело'] = sources
    df_input['Посилання'] = urls
    
    output_filename = 'Input_analyzed_with_urls.xlsx'
    df_input.to_excel(output_filename, index=False)
    print(f"Аналіз успішно завершено. Результати збережено у {output_filename}")

if __name__ == '__main__':
    main()
    
    
