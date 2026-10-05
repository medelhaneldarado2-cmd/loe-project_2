import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET
import re
from datetime import datetime

# --- НАСТРОЙКИ СТРАНИЦЫ И ДИЗАЙН ---

st.markdown("""
<style>
    .main-title { font-size: 36px; font-weight: 800; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #6B7280; margin-bottom: 30px; }
    .section-title { font-size: 20px; font-weight: 600; color: #374151; border-bottom: 2px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 20px; margin-top: 20px;}
    .stButton>button { background-color: #2563EB; color: white; border-radius: 8px; font-size: 16px; padding: 25px 0px; font-weight: bold; transition: all 0.3s;}
    .stButton>button:hover { background-color: #1D4ED8; border-color: #1D4ED8; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Генератор отчета LOE</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Умная система автоматического сведения продаж, себестоимости, складских остатков и брендов.</div>', unsafe_allow_html=True)


# --- ТЕХНИЧЕСКИЕ ФУНКЦИИ ---
def get_col_index(cell_ref):
    match = re.match(r"([A-Z]+)[0-9]+", cell_ref)
    if not match: return -1
    letters = match.group(1)
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def parse_1c_excel(file_bytes):
    data = []
    try:
        with zipfile.ZipFile(file_bytes, 'r') as z:
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    strings = [t.text if t.text is not None else '' for t in root.findall('.//ss:t', ns)]
            
            with z.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ws': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                rows = root.findall('.//ws:row', ns)
                for row in rows:
                    row_dict = {}
                    for c in row.findall('.//ws:c', ns):
                        r_attr = c.attrib.get('r')
                        col_idx = get_col_index(r_attr) if r_attr else len(row_dict)
                        val_tag = c.find('ws:v', ns)
                        val = val_tag.text if val_tag is not None else ''
                        t = c.attrib.get('t', '')
                        if t == 's' and val.isdigit():
                            idx = int(val)
                            val = strings[idx] if idx < len(strings) else val
                        row_dict[col_idx] = val
                    if row_dict:
                        max_col = max(row_dict.keys())
                        row_data = [row_dict.get(i, '') for i in range(max_col + 1)]
                        data.append(row_data)
    except Exception as e:
        st.error(f"Ошибка чтения файла 1С: {e}")
    return pd.DataFrame(data)


# --- ИНТЕРФЕЙС ЗАГРУЗКИ ---
st.markdown('<div class="section-title">📂 Шаг 1. Загрузите файлы выгрузок</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("📑 1. Общий отчет (loe проооггр.xlsx)", type=["xlsx"])
    sales_files = st.file_uploader("💰 2. Ежедневные продажи (Валовая прибыль.xlsx)", type=["xlsx"], accept_multiple_files=True)
    brand_file = st.file_uploader("🏷️ 3. Файл с брендами/категориями (SP12.xlsx)", type=["xlsx"], help="Файл, откуда берутся 'Кеды Осень' из тегов <b>")
with col2:
    cost_file = st.file_uploader("💳 4. Файл с себестоимостью (СЕБЕС.xlsx)", type=["xlsx"])
    stock_file = st.file_uploader("📦 5. Остатки товара (Остатки товара.xlsx)", type=["xlsx"])

st.markdown('<div class="section-title">⚙️ Шаг 2. Запуск обработки</div>', unsafe_allow_html=True)

if st.button("🚀 СФОРМИРОВАТЬ ИТОГОВЫЙ ОТЧЕТ", use_container_width=True):
    if not (base_file and sales_files and cost_file and stock_file and brand_file):
        st.error("⚠️ Пожалуйста, загрузите все 5 типов файлов перед запуском!")
    else:
        with st.spinner("🔄 Обработка данных и сведение таблиц..."):
            try:
                # 1. Читаем Главный отчет
                df_main = pd.read_excel(base_file)
                new_cols = []
                for c in df_main.columns:
                    if isinstance(c, datetime):
                        new_cols.append(c.strftime("%d.%m.%Y"))
                    else:
                        new_cols.append(c)
                df_main.columns = new_cols
                df_main = df_main.astype(object) 
                
                # 2. Читаем Бренды (вытаскиваем текст между <b> и </b>)
                df_brand = pd.read_excel(brand_file, header=None)
                brand_data = {}
                for _, row in df_brand.iterrows():
                    art = str(row[0]).strip()
                    desc = str(row[1])
                    # Ищем текст между <b> и </b>
                    match = re.search(r'<b>(.*?)</b>', desc)
                    if match:
                        brand_data[art] = match.group(1).strip()
                
                # 3. Читаем Себестоимость
                df_cost = pd.read_excel(cost_file)
                cost_data = {}
                for _, row in df_cost.iterrows():
                    art = str(row.get('Артикул', '')).strip()
                    if art and art != 'nan':
                        cost_data[art] = {
                            'Себест-сть': row.get('Себест-сть'),
                            'Доход': row.get('Доход'),
                            'Прибыль': row.get('Прибыль'),
                            'Цена реализации': row.get('Цена реализации')
                        }

                # 4. Читаем Остатки
                df_stock = parse_1c_excel(stock_file)
                stock_data = {}
                header_idx = -1
                for i, row in df_stock.iterrows():
                    if "Артикул" in row.values:
                        header_idx = i
                        break
                if header_idx != -1:
                    headers = df_stock.iloc[header_idx].tolist()
                    for i in range(header_idx + 1, len(df_stock)):
                        row = df_stock.iloc[i].tolist()
                        if len(row) > 1 and row[1]:
                            art = str(row[1]).strip()
                            if not art or art == 'None' or art.startswith('1-100'): continue
                            try:
                                ost_idx = headers.index('Остаток')
                                qty_str = str(row[ost_idx]).strip()
                                if qty_str: stock_data[art] = float(qty_str)
                            except: pass

                # 5. Читаем Продажи
                sales_data = {}
                for f in sales_files:
                    df_val = parse_1c_excel(f)
                    date_str = None
                    
                    match_name = re.search(r"(\d{2}\.\d{2}(?:\.\d{4})?)", f.name)
                    if match_name:
                        d = match_name.group(1)
                        date_str = f"{d}.2026" if len(d) == 5 else d
                            
                    if not date_str:
                        for i, row in df_val.head(15).iterrows():
                            for val in row.values:
                                if isinstance(val, str):
                                    matches = re.findall(r"(\d{2}\.\d{2}\.\d{4})", val)
                                    if matches:
                                        date_str = matches[0]
                                        break
                            if date_str: break
                    
                    if date_str:
                        header_idx = -1
                        for i, row in df_val.iterrows():
                            if "Артикул" in row.values:
                                header_idx = i
                                break
                        if header_idx != -1:
                            headers = df_val.iloc[header_idx].tolist()
                            for i in range(header_idx + 1, len(df_val)):
                                row = df_val.iloc[i].tolist()
                                if len(row) > 1 and row[1]:
                                    art = str(row[1]).strip()
                                    if not art or art == 'None' or art.startswith(('Итого', 'Средняя', '1.', '2.')): continue
                                    try:
                                        kol_idx = headers.index('Кол.')
                                        qty_str = str(row[kol_idx]).strip()
                                        qty = float(qty_str) if qty_str else 0
                                        if date_str not in sales_data: sales_data[date_str] = {}
                                        sales_data[date_str][art] = qty
                                    except: pass

                # --- 6. ОБНОВЛЕНИЕ ГЛАВНОГО ОТЧЕТА ---
                for dt in sales_data.keys():
                    if dt not in df_main.columns:
                        df_main[dt] = None
                    df_main[dt] = df_main[dt].astype(object)
                
                # Ищем нужные колонки
                brand_col = next((c for c in df_main.columns if 'БРЕНД' in str(c).upper()), 'Бренд')
                stock_col = next((c for c in df_main.columns if 'ОСТАТОК' in str(c).upper()), 'ОСТАТОК ')
                sebes_col = next((c for c in df_main.columns if 'СЕБЕСТ' in str(c).upper()), 'Себест-сть')
                dohod_col = next((c for c in df_main.columns if 'ДОХОД' in str(c).upper()), 'Доход')
                pribil_col = next((c for c in df_main.columns if 'ПРИБЫЛЬ' in str(c).upper()), 'Прибыль')
                cena_col = next((c for c in df_main.columns if 'ЦЕНА РЕАЛИЗАЦИИ' in str(c).upper()), 'Цена реализации')

                for idx, row in df_main.iterrows():
                    art = str(row['Артикул']).strip()
                    
                    # Пишем Бренд
                    if art in brand_data:
                        df_main.at[idx, brand_col] = brand_data[art]

                    # Пишем продажи
                    for dt, items in sales_data.items():
                        if art in items:
                            df_main.at[idx, dt] = items[art]
                    
                    # Пишем остатки
                    if art in stock_data:
                        df_main.at[idx, stock_col] = stock_data[art]
                        
                    # Пишем себестоимость
                    if art in cost_data:
                        df_main.at[idx, sebes_col] = cost_data[art].get('Себест-сть')
                        df_main.at[idx, dohod_col] = cost_data[art].get('Доход')
                        df_main.at[idx, pribil_col] = cost_data[art].get('Прибыль')
                        df_main.at[idx, cena_col] = cost_data[art].get('Цена реализации')
                
                # 7. СОХРАНЕНИЕ
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_main.to_excel(writer, index=False, sheet_name="Общий отчет")
                output.seek(0)
                
                # --- ВЫВОД РЕЗУЛЬТАТА ---
                st.success("✅ Отчет успешно пересобран и готов к скачиванию!")
                
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("🗓 Дней продаж", len(sales_data))
                m2.metric("🏷️ Обновлено брендов", len(brand_data))
                m3.metric("📦 Товаров на складе", len(stock_data))
                m4.metric("💳 С себестоимостью", len(cost_data))
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 СКАЧАТЬ ОБНОВЛЕННЫЙ ОТЧЕТ",
                    data=output,
                    file_name="Обновленный_loe_отчет.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ Произошла ошибка при сведении данных: {e}")

