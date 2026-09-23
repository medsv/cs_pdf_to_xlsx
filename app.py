#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit-приложение для распознавания кабельных ведомостей из PDF-файлов
и формирования сводного Excel-файла (аналог main() из process_pdfs.py,
но с веб-интерфейсом: загрузка нескольких PDF и скачивание результата).

Запуск:
    streamlit run app.py

Использует process_pdfs.py в качестве библиотеки:
    - process_pdfs.process_pdf(pdf_path)  -> список записей о кабелях
    - process_pdfs.write_excel(cables, output_path) -> итоговый xlsx
"""

import os
import tempfile
from datetime import datetime

import streamlit as st
# process_pdfs - оригинальный файл
import process_pdfs_1


# ============================================================================
# Настройка страницы
# ============================================================================
st.set_page_config(
    page_title='Конвертер кабельного журнала из PDF в XLSX',
    page_icon='📘',
    layout='centered',
)


def render_header():
    """Заголовок и описание приложения."""
    st.title('Кабельный журнал: PDF → XLSX')

def reset_results_if_needed(uploaded_files):
    """
    Сбрасывает ранее полученные результаты, если изменился список
    загруженных файлов (чтобы старый результат не оставался от прежней сессии).
    """
    current_keys = [f.name for f in uploaded_files] if uploaded_files else []
    if st.session_state.get('uploaded_keys') != current_keys:
        st.session_state.pop('xlsx_bytes', None)
        st.session_state.pop('result', None)
        st.session_state['uploaded_keys'] = current_keys


def process_uploaded_files(uploaded_files):
    """
    Обрабатывает загруженные PDF-файлы:
      1. Сохраняет каждый файл во временную директорию
      2. Вызывает process_pdfs.process_pdf() для каждого файла
      3. Формирует итоговый xlsx через process_pdfs.write_excel()
      4. Возвращает байты xlsx-файла и сводку по обработке
    """
    all_cables = []
    per_file_stats = []
    errors = []

    progress_bar = st.progress(0.0, text='Подготовка…')
    status = st.empty()

    total = len(uploaded_files)

    with tempfile.TemporaryDirectory() as tmp_dir:
        # 1. Обработка каждого PDF
        for i, uploaded_file in enumerate(uploaded_files):
            file_label = f'{i + 1}/{total}: {uploaded_file.name}'
            status.info(f'⏳ Обработка {file_label}')
            progress_bar.progress((i) / total, text=f'Обработка {file_label}')

            # Сохраняем загруженный файл на диск (process_pdf принимает путь)
            pdf_path = os.path.join(tmp_dir, uploaded_file.name)
            with open(pdf_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            try:
                cables = process_pdfs_1.process_pdf(pdf_path)
                all_cables.extend(cables)
                per_file_stats.append((uploaded_file.name, len(cables)))
            except Exception as e:  # защита от непредвиденных ошибок
                errors.append(f'{uploaded_file.name}: {e}')
                per_file_stats.append((uploaded_file.name, 0))

            progress_bar.progress((i + 1) / total, text=f'Обработано {file_label}')

        # 2. Формирование итогового Excel-файла
        if all_cables:
            status.info('Формирование Excel-файла…')
            xlsx_path = os.path.join(tmp_dir, 'Output.xlsx')
            process_pdfs_1.write_excel(all_cables, xlsx_path)
            with open(xlsx_path, 'rb') as f:
                xlsx_bytes = f.read()
        else:
            xlsx_bytes = None

    status.empty()
    progress_bar.empty()

    return xlsx_bytes, per_file_stats, errors


def render_result_section(result):
    """Отображает сводку по обработке и кнопку скачивания xlsx."""
    xlsx_bytes, per_file_stats, errors = result

    st.subheader('Результат обработки')

    if errors:
        st.warning('⚠️ Часть файлов не удалось обработать:')
        for err in errors:
            st.error(err)

    if not xlsx_bytes:
        st.error(
            'Не удалось извлечь записи о кабелях ни из одного файла. '
            'Проверьте, что PDF содержат таблицы кабельных ведомостей.'
        )
        return

    total_cables = sum(count for _, count in per_file_stats)

    st.success(
        f'✅ Обработано файлов: {len(per_file_stats)}, '
        f'извлечено записей о кабелях: {total_cables}'
    )

    # Имя файла: yyyy-mm-dd_hh-mm_N_N_Кабели.xlsx
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    file_name = f'{timestamp}_{len(per_file_stats)}_{total_cables}_Кабели.xlsx'

    # Сводная таблица по файлам
    st.dataframe(
        {
            'Файл': [name for name, _ in per_file_stats],
            'Записей о кабелях': [count for _, count in per_file_stats],
        },
        use_container_width=True,
        hide_index=True,
    )

    # Кнопка скачивания итогового файла
    st.download_button(
        label='⬇️ Скачать итоговый Excel-файл',
        data=xlsx_bytes,
        file_name=file_name,
        mime=(
            'application/vnd.openxmlformats-officedocument'
            '.spreadsheetml.sheet'
        ),
        type='primary',
    )



# ============================================================================
# Основной поток приложения
# ============================================================================
def main():
    render_header()

    # Счётчик для key виджета file_uploader. При нажатии «Очистить» key меняется,
    # поэтому Streamlit пересоздаёт виджет заново — уже без загруженных файлов.
    if 'uploader_counter' not in st.session_state:
        st.session_state['uploader_counter'] = 0

    # Загрузка нескольких PDF-файлов с любыми именами
    uploaded_files = st.file_uploader(
        'Выберите один или несколько PDF файлов, нажав Upload, или перетащите их сюда мышкой из Проводника',
        type=['pdf'],
        accept_multiple_files=True,
        key=f'pdf_uploader_{st.session_state["uploader_counter"]}',
        help='Можно выбрать несколько файлов одновременно',
    )

    reset_results_if_needed(uploaded_files)

    # Кнопка запуска обработки
    col1, col2 = st.columns(2)
    with col1:
        if uploaded_files:
            if st.button('🚀 Обработать', type='primary'):
                with st.spinner('Обработка PDF-файлов…'):
                    st.session_state['result'] = process_uploaded_files(uploaded_files)
    with col2:
        if uploaded_files:
            if st.button('Очистить', type='primary'):
                # Меняем key виджета: file_uploader пересоздаётся пустым,
                # а результаты прошлой обработки сбрасываются.
                st.session_state['uploader_counter'] += 1
                st.session_state.pop('xlsx_bytes', None)
                st.session_state.pop('result', None)
                st.rerun()

    # Отображение последнего результата (переживает rerun'ы Streamlit)
    if st.session_state.get('result'):
        render_result_section(st.session_state['result'])

    st.markdown(
        """
        <hr>
        <p style="text-align: left; color: gray;">
        <small>
        2026, С.В. Медведев, engpython@yandex.ru
        </small>
        </p>
        """,
        unsafe_allow_html=True,
    )

if __name__ == '__main__':
    main()