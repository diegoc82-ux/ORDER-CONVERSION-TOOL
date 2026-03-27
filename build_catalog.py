import pandas as pd
import json
import os

EXCEL_PATH = r"C:\Users\dcastro\OneDrive - Ultragroup\Documents\U1P\Ultra1Plus\PRODUCT LISTS\SKU LIST - FOR WEBSITE  SERVER - UPDATED 03.11.2026.xlsx"

def build():
    df_sku = pd.read_excel(EXCEL_PATH, sheet_name='SKU', dtype=str).fillna('')
    df_name = pd.read_excel(EXCEL_PATH, sheet_name='NAME', dtype=str).fillna('')

    name_map = {}
    for _, row in df_name.iterrows():
        code = str(row.iloc[0]).strip()
        name = str(row.iloc[1]).strip()
        if code:
            name_map[code] = name

    catalog = {}
    for _, row in df_sku.iterrows():
        code = str(row.iloc[0]).strip()
        presentation = str(row.iloc[1]).strip()
        sku = str(row.iloc[2]).strip()
        if not code or not presentation or not sku:
            continue
        if code not in catalog:
            catalog[code] = {'name': name_map.get(code, ''), 'presentations': {}}
        catalog[code]['presentations'][presentation] = sku

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'products.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)
    print(f"Built catalog: {len(catalog)} products -> products.json")

if __name__ == '__main__':
    build()
