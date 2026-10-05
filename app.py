from pathlib import Path
import runpy
import hashlib
import streamlit as st
import pandas as pd
from report_engine import build_report

st.set_page_config(page_title='LOE — отчет по категориям', page_icon='📊', layout='wide')
mode = st.sidebar.radio('Режим', ['По категориям — КЕТИК', 'Прежний режим — 5 файлов'])
if mode == 'Прежний режим — 5 файлов':
    runpy.run_path(str(Path(__file__).with_name('legacy_app.py')))
else:
    st.title('Отчет по категориям')
    st.write('Категория определяется по артикулу на первом листе КЕТИК. Количество, доход и валовая прибыль берутся из ежедневных выгрузок.')
    left, right = st.columns(2)
    with left:
        base = st.file_uploader('1. КЕТИК — категории и артикулы', type=['xlsx'])
    with right:
        sales = st.file_uploader('2. Валовая прибыль — одна полная выгрузка за каждый день', type=['xlsx'], accept_multiple_files=True)
    st.caption('Возвраты учитываются со знаком минус. Повторная обработка заменяет данные загруженных дней. Итоги на экране и отдельном листе относятся только к выбранным выгрузкам. История остальных дней сохраняется.')
    payload = [(f.name, f.getvalue()) for f in sales]
    fingerprint = hashlib.sha256((base.getvalue() if base else b'') + b''.join(n.encode()+d for n,d in payload)).hexdigest()
    if st.button('Сформировать отчет', type='primary', use_container_width=True):
        st.session_state.pop('category_result', None)
        if not base or not sales:
            st.error('Загрузите КЕТИК и хотя бы одну ежедневную выгрузку.')
        else:
            try:
                with st.spinner('Сопоставляю артикулы и рассчитываю итоги…'):
                    result = build_report(base.getvalue(), payload)
                st.session_state.category_result = (fingerprint, result)
            except Exception as error:
                st.error(f'Не удалось сформировать отчет: {error}')
    saved = st.session_state.get('category_result')
    if saved and saved[0] == fingerprint:
        data, records, totals = saved[1]
        st.success('Отчет сформирован. Категории и общие итоги заполнены.')
        for col,label,value in zip(st.columns(3), ['Общее количество','Общий доход','Общая прибыль'], totals):
            col.metric(label, f'{value:,.2f}'.replace(',', ' '))
        st.dataframe(pd.DataFrame(records,columns=['Категория','Общее количество','Общий доход','Общая прибыль']),hide_index=True,use_container_width=True)
        st.download_button('Скачать заполненный КЕТИК', data, 'КЕТИК_заполненный.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
