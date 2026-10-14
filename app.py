"""Локальный отчёт: пользовательские данные живут только в текущей сессии."""

import asyncio
import hashlib
import logging
from pathlib import Path

import streamlit as st

from data import RATIO_NOTE, TABLE_HEADER, parse_csv, table_rows, validate_report, validate_report_field
from slides import build_presentation, slide_pages


LOGGER = logging.getLogger('report_slides')


st.set_page_config(page_title='Report Slides', layout='wide')
st.title('Report Slides')
st.caption('План и факт из CSV — в редактируемую презентацию PowerPoint. Модель и API-ключ не нужны.')
st.info('При обновлении страницы или остановке приложения рабочая сессия может исчезнуть. Скачайте нужный PPTX заранее.')

for key, value in (('input_profile', None), ('metrics', None), ('import_error', None),
                   ('pptx', None), ('export_fields', None), ('title', 'Еженедельный отчёт'),
                   ('period', ''), ('comment', ''), ('upload_revision', 0), ('source_file', None),
                   ('edited_fields', set())):
    if key not in st.session_state:
        st.session_state[key] = value


def reset_upload() -> None:
    """Clear the active upload and its report fields before selecting a replacement

    Returns:
        None
    """
    st.session_state['upload_revision'] += 1
    for key, value in (('input_profile', None), ('metrics', None), ('import_error', None),
                       ('pptx', None), ('export_fields', None), ('period', ''), ('comment', ''),
                       ('source_file', None)):
        st.session_state[key] = value
    st.session_state['edited_fields'].intersection_update(('title',))


def mark_report_edited(field: str) -> None:
    """Enable validation feedback only for the field the user has edited

    Args:
        field: Report field key

    Returns:
        None
    """
    st.session_state['edited_fields'].add(field)


st.subheader('1. Загрузите показатели')
st.download_button('Скачать пример CSV', (Path(__file__).parent / 'samples/weekly-report.csv').read_bytes(),
                   file_name='weekly-report.csv', mime='text/csv', key='sample', on_click='ignore')
upload_key = f"csv_file_{st.session_state['upload_revision']}"
source_file = st.session_state['source_file']
columns = st.columns([3, 1])
with columns[0]:
    if source_file is None:
        slot = st.empty()
        uploaded = slot.file_uploader('CSV UTF-8 с планом и фактом', type=['csv'], key=upload_key,
                                       accept_multiple_files=False, max_upload_size=1,
                                       help='До 256 КиБ, от 1 до 12 показателей. Колонки: показатель, план, факт, единица.')
        if uploaded is not None:
            source_file = (uploaded.name, uploaded.getvalue())
            st.session_state['source_file'] = source_file
            slot.text('Выбран CSV: ' + source_file[0])
    else:
        st.text('Выбран CSV: ' + source_file[0])
with columns[1]:
    delimiter = st.radio('Разделитель CSV', [';', ','], key='delimiter',
                         format_func=lambda value: '; — точка с запятой' if value == ';' else ', — запятая')

st.button('Загрузить новый CSV', key='new_csv', on_click=reset_upload, disabled=source_file is None,
           help='Очищает таблицу, PPTX, период и комментарий до выбора следующего файла. Название отчёта сохраняется.')

filename = source_file[0] if source_file is not None else None
raw = source_file[1] if source_file is not None else None
profile = (filename,
           hashlib.sha256(raw).hexdigest() if raw is not None else None, delimiter)
previous_profile = st.session_state['input_profile']
if previous_profile != profile:
    st.session_state['metrics'] = None
    st.session_state['import_error'] = None
    st.session_state['pptx'] = None
    source_changed = previous_profile is None or previous_profile[:2] != profile[:2]
    if source_changed:
        st.session_state['period'] = ''
        st.session_state['comment'] = ''
        st.session_state['edited_fields'].intersection_update(('title',))
    st.session_state['input_profile'] = profile
    if raw is not None:
        try:
            st.session_state['metrics'] = parse_csv(raw, filename, delimiter)
        except ValueError as error:
            st.session_state['import_error'] = str(error)

metrics = st.session_state['metrics']
if st.session_state['import_error']:
    st.error(st.session_state['import_error'])
elif metrics is None:
    st.caption('Скачайте образец или загрузите свой CSV, сохранённый в UTF-8.')

if metrics is not None:
    st.subheader('2. Проверьте данные и расчёты')
    rows = table_rows(metrics)
    preview = {header: [row[column] for row in rows] for column, header in enumerate(TABLE_HEADER)}
    st.dataframe(preview, hide_index=True, width='stretch',
                  column_config={header: st.column_config.TextColumn(header) for header in TABLE_HEADER})
    st.caption(RATIO_NOTE)
    st.caption('Цвета в презентации обозначают план и факт. Программа не оценивает, хорошо или плохо отклонение.')

st.subheader('3. Оформите отчёт')
columns = st.columns(2)
with columns[0]:
    title = st.text_input('Название отчёта', key='title', help='Обязательно, до 100 символов.',
                          on_change=mark_report_edited, args=('title',))
with columns[1]:
    period = st.text_input('Период', key='period', placeholder='5–11 октября 2026',
                           help='Обязательно, до 60 символов; свободный текст.', disabled=metrics is None,
                           on_change=mark_report_edited, args=('period',))
comment = st.text_area('Комментарий на титуле (необязательно)', key='comment',
                        help='Обычный текст, до 400 символов и 6 непустых абзацев.', disabled=metrics is None,
                        on_change=mark_report_edited, args=('comment',))

fields = (title, period, comment)
if st.session_state['export_fields'] != fields:
    st.session_state['pptx'] = None
    st.session_state['export_fields'] = fields
try:
    validated_fields = validate_report(*fields)
except ValueError:
    validated_fields = None
for field, value in zip(('title', 'period', 'comment'), fields):
    field_disabled = metrics is None and field != 'title'
    if field_disabled or field not in st.session_state['edited_fields']:
        continue
    try:
        validate_report_field(value, field)
    except ValueError as error:
        st.error(str(error))

if metrics is not None:
    st.subheader('4. Состав презентации')
    pages = slide_pages(metrics)
    tables = pages[0]
    charts = pages[1]
    st.text(f'{1 + len(tables) + len(charts)} слайдов: титул — 1, таблицы — {len(tables)}, графики — {len(charts)}.')
    for index, (unit, page) in enumerate(charts, 1):
        st.text(f'График {index}: единица «{unit}», показателей — {len(page)}.')

if st.button('Создать презентацию', key='generate', type='primary',
              disabled=metrics is None or validated_fields is None):
    try:
        with st.spinner('Собираю таблицы и диаграммы…'):
            result = build_presentation(metrics, *validated_fields)
        st.session_state['pptx'] = result
    except asyncio.CancelledError:
        raise
    except Exception as error:
        LOGGER.exception('Не удалось собрать PPTX для CSV %s', profile[1],
                         exc_info=(RuntimeError, RuntimeError('Ошибка сборки PPTX'), error.__traceback__))
        previous = ' Прежний успешный PPTX сохранён для этих же данных.' if st.session_state['pptx'] else ''
        st.error('Не удалось собрать файл PPTX. Введённые данные сохранены; подробности записаны в журнал сервера.'
                 + previous)

if st.session_state['pptx'] is not None:
    st.success('Презентация готова. Скачайте файл перед закрытием или обновлением страницы.')
    st.download_button('Скачать PPTX', st.session_state['pptx'], file_name='report-slides.pptx',
                        mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        key='download_pptx', on_click='ignore')

st.caption('Следующий отчёт: нажмите «Загрузить новый CSV», выберите файл и заполните период заново. '
           'Название сохранится, комментарий очистится. '
           'Правки из PowerPoint в следующий экспорт не переносятся.')
