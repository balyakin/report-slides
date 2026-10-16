"""Локальный отчёт: пользовательские данные живут только в текущей сессии."""

import asyncio
import hashlib
import logging
from decimal import Decimal
from pathlib import Path
import re
from typing import Dict, Sequence, Tuple

import streamlit as st

import data as report_data
from data import RATIO_NOTE, TABLE_HEADER, parse_csv, table_rows, validate_report, validate_report_field
from slides import build_presentation, slide_pages


LOGGER = logging.getLogger('report_slides')
AXIS_WIDGET_PREFIXES = ('axis_fixed_', 'axis_min_', 'axis_max_')
SLIDE_PREVIEW_WIDTH = 192
SLIDE_PREVIEW_HEIGHT = 108


st.set_page_config(page_title='Report Slides', page_icon=str(Path(__file__).parent / 'static/mark.svg'), layout='wide')
st.html('''<style>
[data-testid="stMainBlockContainer"] {max-width:1128px; padding:40px 24px 56px;}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] {gap:32px;}
.st-key-step_upload, .st-key-step_data, .st-key-step_style, .st-key-step_deck, .st-key-action {
    background:#FFFFFF; border-radius:12px; padding:24px; gap:20px;
    box-shadow:0 1px 2px rgba(15,23,42,.04);
}
[data-testid="stCaptionContainer"] {color:#475569; max-width:80ch;}
[data-testid="stCaptionContainer"] p {color:#475569;}
[data-testid="stMarkdownContainer"] {max-width:80ch;}
[data-testid="stText"], [data-testid="stMarkdownContainer"] {overflow-wrap:anywhere;}
[data-testid="stText"], [data-testid="stDataFrame"], [data-testid="stTextInput"] {
    font-variant-numeric:tabular-nums;
}
[data-testid="stFileUploaderDropzone"] {min-height:136px; border-radius:8px;}
[data-testid="stFileUploaderDropzoneInstructions"] {display:none;}
.st-key-brand_mark [data-testid="stElementToolbarButtonContainer"] {display:none;}
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-tertiary"], [data-testid="stBaseButton-segmented_control"],
[data-testid="stCheckbox"] {min-height:40px;}
[data-testid="stButtonGroup"] button, [data-testid="stExpander"] summary {min-height:40px;}
[data-testid="stTooltipHoverTarget"] button {min-width:40px; min-height:40px; border-radius:8px;}
[data-testid="stButton"] button, [data-testid="stDownloadButton"] button, [data-testid="stButtonGroup"] button {
    transition:background-color 120ms, border-color 120ms;
}
[data-testid="stTextInput"]:focus-within, [data-testid="stTextArea"]:focus-within,
[data-testid="stBaseButton-primary"]:focus-visible, [data-testid="stBaseButton-secondary"]:focus-visible,
[data-testid="stBaseButton-tertiary"]:focus-visible,
[data-testid="stButtonGroup"] button:focus-visible, [data-testid="stTooltipHoverTarget"] button:focus-visible,
[data-testid="stExpander"] summary:focus-visible {
    outline:2px solid #2563EB; outline-offset:3px; border-radius:8px;
}
.st-key-step_number_1, .st-key-step_number_2, .st-key-step_number_3, .st-key-step_number_4 {
    background:#F1F5F9; color:#475569; border-radius:50%; min-height:32px; justify-content:center; font-weight:600;
}
.st-key-slide_strip [data-testid="stVerticalBlock"] {gap:4px; padding:8px;}
.st-key-axis_groups [data-testid="stTextInput"] [data-testid="stWidgetLabel"] {display:none;}
.st-key-slide_strip [data-testid="stText"] {font-size:14px; line-height:1.35;}
.st-key-action:has([data-testid="stSpinner"]) .st-key-generate {display:none;}
.st-key-action [data-testid="stSpinner"] {
    min-height:40px; padding:8px 12px; background:#F1F5F9; border:1px solid #E2E8F0; border-radius:8px;
}
@media (max-width:640px) {
    [data-testid="stMainBlockContainer"] {padding:24px 16px 40px;}
    .st-key-step_upload, .st-key-step_data, .st-key-step_style, .st-key-step_deck, .st-key-action {padding:16px;}
    .st-key-axis_header {display:none;}
    .st-key-axis_groups [data-testid="stTextInput"] [data-testid="stWidgetLabel"] {display:flex;}
    [data-testid="stColumn"] {width:100%!important; flex:1 1 100%!important; min-width:0!important;}
}
@media (prefers-reduced-motion:reduce) {
    [data-testid="stButton"] *, [data-testid="stDownloadButton"] *, [data-testid="stSpinner"] * {
        animation:none!important; transition:none!important;
    }
}
</style>''')
with st.container(horizontal=True, vertical_alignment='center', gap='small'):
    with st.container(width=32, key='brand_mark'):
        st.image(str(Path(__file__).parent / 'static/mark.svg'), width=32)
    st.title('Report Slides', anchor=False)
st.text('План и факт из CSV — в редактируемую презентацию PowerPoint')
with st.container(horizontal=True, gap='small'):
    st.badge('Работает локально', color='gray')
    st.badge('Без ИИ и API-ключей', color='gray')
    st.badge('Редактируемый PPTX', color='gray')

for key, value in (('input_profile', None), ('metrics', None), ('import_error', None),
                   ('pptx', None), ('export_fields', None), ('title', 'Еженедельный отчёт'),
                   ('period', ''), ('comment', ''), ('upload_revision', 0), ('source_file', None),
                   ('edited_fields', set()), ('axis_settings', {}), ('axis_url_loaded', False),
                   ('axis_url_error', None), ('axis_saved_payload', ''), ('axis_save_error', None),
                   ('axis_export_settings', None)):
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


def update_axis_setting(unit: str, field: str, widget_key: str) -> None:
    """Copy a user edit into durable settings before widgets can disappear

    Args:
        unit: Normalized unit key
        field: Fixed mode, minimum or maximum field
        widget_key: Source widget state key

    Returns:
        None
    """
    st.session_state['axis_settings'][unit][field] = st.session_state[widget_key]
    st.session_state['axis_save_error'] = None
    st.session_state['pptx'] = None


def reset_axis_settings() -> None:
    """Clear only scale settings and their bookmark without resetting report input

    Returns:
        None
    """
    for key in tuple(st.session_state):
        if key.startswith(AXIS_WIDGET_PREFIXES):
            del st.session_state[key]
    st.session_state['axis_settings'] = {}
    st.session_state['axis_url_error'] = None
    st.session_state['axis_save_error'] = None
    st.session_state['axis_saved_payload'] = ''
    st.session_state['axis_url_loaded'] = True
    st.session_state['pptx'] = None
    remaining = {key: st.query_params.get_all(key) for key in st.query_params if key != report_data.AXIS_QUERY_KEY}
    st.query_params.from_dict(remaining)


def save_axis_settings() -> None:
    """Store a canonical bookmark while preserving unrelated query parameters

    Returns:
        None
    """
    try:
        payload = report_data.serialize_axis_settings(st.session_state['axis_settings'])
    except ValueError as error:
        st.session_state['axis_save_error'] = str(error)
        return
    if payload:
        if st.query_params.get_all(report_data.AXIS_QUERY_KEY) != [payload]:
            st.query_params[report_data.AXIS_QUERY_KEY] = payload
    else:
        remaining = {key: st.query_params.get_all(key) for key in st.query_params if key != report_data.AXIS_QUERY_KEY}
        st.query_params.from_dict(remaining)
    st.session_state['axis_saved_payload'] = payload
    st.session_state['axis_save_error'] = None


def _restore_axis_settings() -> None:
    if st.session_state['axis_url_loaded']:
        return
    st.session_state['axis_url_loaded'] = True
    values = st.query_params.get_all(report_data.AXIS_QUERY_KEY)
    if not values:
        return
    try:
        if len(values) != 1 or not values[0]:
            raise ValueError('В адресе нужен ровно один непустой параметр scales.')
        settings = report_data.parse_axis_settings(values[0])
        payload = report_data.serialize_axis_settings(settings)
    except ValueError as error:
        st.session_state['axis_url_error'] = str(error)
        return
    st.session_state['axis_settings'] = settings
    st.session_state['axis_saved_payload'] = payload


def _escape_axis_message(message: str) -> str:
    runs = re.findall('`+', message)
    delimiter = '`' * (max((len(run) for run in runs), default=0) + 1)
    return '  \n'.join(f'{delimiter} {line} {delimiter}' for line in message.splitlines())


def _get_current_axis_ranges(metrics: Sequence[report_data.Metric]) -> Dict[str, Tuple[Decimal, Decimal]]:
    ranges = {}
    for unit, group in report_data.group_metrics(metrics):
        setting = st.session_state['axis_settings'][unit.casefold()]
        if setting['fixed']:
            ranges[unit.casefold()] = report_data.parse_axis_range(setting['minimum'], setting['maximum'], unit)
    return ranges


def _render_step_header(number: int, title: str, status: str) -> None:
    with st.container(horizontal=True, vertical_alignment='center', gap='small'):
        with st.container(width=32, key=f'step_number_{number}'):
            st.text(str(number), text_alignment='center', width='stretch')
        st.subheader(title, anchor=False)
        st.badge(status, color='gray')


def _render_axis_group(unit: str) -> None:
    unit_key = unit.casefold()
    setting = st.session_state['axis_settings'][unit_key]
    keys = ('axis_fixed_' + unit_key, 'axis_min_' + unit_key, 'axis_max_' + unit_key)
    for field, widget_key in zip(('fixed', 'minimum', 'maximum'), keys):
        if widget_key not in st.session_state:
            st.session_state[widget_key] = setting[field]
    columns = st.columns([1.2, 1.7, 1, 1], vertical_alignment='center')
    with columns[0]:
        st.text(unit)
    with columns[1]:
        st.checkbox('Фиксированная шкала', key=keys[0], on_change=update_axis_setting,
                    args=(unit_key, 'fixed', keys[0]))
    with columns[2]:
        st.text_input('Минимум', key=keys[1], disabled=not setting['fixed'],
                      on_change=update_axis_setting, args=(unit_key, 'minimum', keys[1]))
    with columns[3]:
        st.text_input('Максимум', key=keys[2], disabled=not setting['fixed'],
                      on_change=update_axis_setting, args=(unit_key, 'maximum', keys[2]))


_restore_axis_settings()

with st.container(border=True, key='step_upload'):
    upload_header = st.container()
    upload_key = f"csv_file_{st.session_state['upload_revision']}"
    source_file = st.session_state['source_file']
    if source_file is None:
        slot = st.empty()
        uploaded = slot.file_uploader('CSV UTF-8 с планом и фактом', type=['csv'], key=upload_key,
                                      accept_multiple_files=False, max_upload_size=1,
                                      help='До 256 КиБ, от 1 до 12 показателей. '
                                           'Колонки: показатель, план, факт, единица.')
        if uploaded is not None:
            source_file = (uploaded.name, uploaded.getvalue())
            st.session_state['source_file'] = source_file
            slot.empty()
    if source_file is None:
        st.caption('CSV UTF-8, до 256 КиБ')
    with st.container(horizontal=True, vertical_alignment='bottom'):
        delimiter = st.segmented_control('Разделитель CSV', [';', ','], key='delimiter', default=';', required=True,
                                         format_func=lambda value: '; Точка с запятой' if value == ';' else ', Запятая',
                                         help='Посмотрите первую строку CSV в текстовом редакторе '
                                              'и выберите разделитель между названиями колонок.')
        st.download_button('Скачать пример CSV', (Path(__file__).parent / 'samples/weekly-report.csv').read_bytes(),
                           file_name='weekly-report.csv', mime='text/csv', key='sample', on_click='ignore',
                           type='tertiary')
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
    with upload_header:
        _render_step_header(1, 'Загрузите показатели', 'Готово' if metrics is not None else 'Нужно заполнить')
    with st.container(horizontal=True, vertical_alignment='center'):
        if source_file is not None:
            st.badge('CSV', icon=':material/description:', color='gray')
            count = f' · Показателей: {len(metrics)}' if metrics is not None else ''
            st.text(f'{filename} · {len(raw)} байт{count}')
            st.button('Загрузить новый CSV', key='new_csv', on_click=reset_upload,
                      type='tertiary', help='Очищает таблицу, PPTX, период и комментарий. Название сохраняется.')
    if st.session_state['import_error']:
        st.error(st.session_state['import_error'])
    elif metrics is None:
        st.caption('Скачайте образец или загрузите свой CSV, сохранённый в UTF-8.')

with st.container(border=True, key='step_data'):
    _render_step_header(2, 'Проверьте данные и расчёты', 'Готово' if metrics is not None else 'Ожидает CSV')
    if metrics is not None:
        st.text(f'Показателей: {len(metrics)} · Единиц: {len(report_data.group_metrics(metrics))}')
        rows = table_rows(metrics)
        preview = {header: [row[column] for row in rows] for column, header in enumerate(TABLE_HEADER)}
        st.dataframe(preview, hide_index=True, width='stretch',
                     column_config={header: st.column_config.TextColumn(header,
                                    alignment='left' if column < 2 else 'right')
                                    for column, header in enumerate(TABLE_HEADER)})
        st.caption(RATIO_NOTE + '. Цвета обозначают план и факт и не оценивают результат.')
    else:
        st.caption('После загрузки корректного CSV здесь появятся показатели, план, факт и отклонения.')

with st.container(border=True, key='step_style'):
    style_header = st.container()
    columns = st.columns(2)
    with columns[0]:
        title = st.text_input('Название отчёта', key='title', help='Обязательно, до 100 символов.',
                              on_change=mark_report_edited, args=('title',))
        title_feedback = st.empty()
    with columns[1]:
        period = st.text_input('Период', key='period', placeholder='5–11 октября 2026',
                               help='Обязательно, до 60 символов; свободный текст.', disabled=metrics is None,
                               on_change=mark_report_edited, args=('period',))
        period_feedback = st.empty()
    comment = st.text_area('Комментарий на титуле (необязательно)', key='comment',
                            help='Обычный текст, до 400 символов и 6 непустых абзацев.', disabled=metrics is None,
                            on_change=mark_report_edited, args=('comment',))
    comment_feedback = st.empty()

    fields = (title, period, comment)
    if st.session_state['export_fields'] != fields:
        st.session_state['pptx'] = None
        st.session_state['export_fields'] = fields
    try:
        validated_fields = validate_report(*fields)
    except ValueError:
        validated_fields = None
    feedback = {'title': title_feedback, 'period': period_feedback, 'comment': comment_feedback}
    for field, value in zip(('title', 'period', 'comment'), fields):
        field_disabled = metrics is None and field != 'title'
        if field_disabled or field not in st.session_state['edited_fields']:
            continue
        try:
            validate_report_field(value, field)
        except ValueError as error:
            feedback[field].error(str(error))

    axis_groups = report_data.group_metrics(metrics) if metrics is not None else ()
    for unit, group in axis_groups:
        st.session_state['axis_settings'].setdefault(unit.casefold(), {'fixed': False, 'minimum': '0', 'maximum': ''})
    axis_snapshot = tuple((unit.casefold(), tuple(st.session_state['axis_settings'][unit.casefold()].items()))
                          for unit, group in axis_groups)
    if st.session_state['axis_export_settings'] != axis_snapshot:
        st.session_state['pptx'] = None
        st.session_state['axis_export_settings'] = axis_snapshot
    axis_ranges = {}
    axis_error = None
    axis_warning = None
    try:
        axis_ranges = _get_current_axis_ranges(metrics) if metrics is not None else {}
    except ValueError as error:
        axis_error = str(error)
    if axis_error is None and metrics is not None:
        try:
            report_data.validate_axis_ranges(metrics, axis_ranges)
        except ValueError as error:
            axis_warning = str(error)
    axis_payload = None
    axis_save_error = st.session_state['axis_save_error']
    try:
        axis_payload = report_data.serialize_axis_settings(st.session_state['axis_settings'])
    except ValueError as error:
        if axis_error is None:
            axis_save_error = str(error)
    with style_header:
        status = 'Ожидает CSV' if metrics is None else 'Нужно заполнить'
        if metrics is not None and validated_fields is not None and not axis_error and not axis_warning:
            status = 'Готово'
        _render_step_header(3, 'Оформите отчёт', status)
    if axis_error or axis_warning or st.session_state['axis_url_error'] or axis_save_error:
        st.session_state['axis_expander'] = True
    with st.expander('Шкалы графиков', expanded=False, key='axis_expander', on_change='rerun'):
        st.caption('Границы общие для плана и факта. Для следующей недели сохраните шкалы в адресе '
                   'и добавьте страницу в закладки. CSV и поля отчёта в ссылку не входят.')
        if metrics is None:
            st.caption('Загрузите CSV, чтобы настроить шкалы его единиц. Сохранённая ссылка будет учтена.')
        with st.container(key='axis_groups'):
            if axis_groups:
                with st.container(key='axis_header'):
                    columns = st.columns([1.2, 1.7, 1, 1])
                    for column, label in zip(columns, ('Единица', 'Режим', 'Минимум', 'Максимум')):
                        column.caption(label)
            for unit, group in axis_groups:
                _render_axis_group(unit)
                values = tuple(value for metric in group for value in (metric.plan, metric.actual))
                st.caption(f'В текущих данных: от {report_data.format_number(min(values))} '
                           f'до {report_data.format_number(max(values))}')
        if axis_error:
            st.error(_escape_axis_message(axis_error))
        if axis_warning:
            st.warning(_escape_axis_message(axis_warning))
        if st.session_state['axis_url_error']:
            st.error(_escape_axis_message('Ошибка ссылки шкал: ' + st.session_state['axis_url_error']))
        if axis_save_error:
            st.error(_escape_axis_message('Не удалось сохранить шкалы в адресе: ' + axis_save_error))
        if axis_payload is None:
            current_units = {unit.casefold() for unit, group in axis_groups}
            for unit, setting in st.session_state['axis_settings'].items():
                if setting['fixed'] and unit not in current_units:
                    st.text(f'Сохранённая шкала «{unit}»: минимум {setting["minimum"]}, максимум {setting["maximum"]}')
        with st.container(horizontal=True, vertical_alignment='center'):
            st.button('Сохранить шкалы в адресе', key='axis_save', on_click=save_axis_settings,
                      disabled=axis_payload is None or st.session_state['axis_url_error'] is not None)
            st.button('Сбросить шкалы', key='axis_reset', on_click=reset_axis_settings, type='tertiary')
            if axis_payload != st.session_state['axis_saved_payload']:
                st.caption('Изменения шкал не сохранены в адресе')
        st.caption('Добавьте обновлённый адрес в закладки. Ссылка содержит единицы и границы шкал. '
                   'Несохранённые в адресе изменения не переживут новую сессию.')

with st.container(border=True, key='step_deck'):
    _render_step_header(4, 'Состав презентации', 'Готово' if metrics is not None else 'Ожидает CSV')
    if metrics is not None:
        pages = slide_pages(metrics)
        tables = pages[0]
        charts = pages[1]
        slide_count = 1 + len(tables) + len(charts)
        st.text(f'{slide_count} слайдов: титул — 1, таблицы — {len(tables)}, графики — {len(charts)}.')
        with st.container(horizontal=True, gap='small', key='slide_strip'):
            with st.container(border=True, width=SLIDE_PREVIEW_WIDTH, height=SLIDE_PREVIEW_HEIGHT):
                st.text('Титул')
                st.caption('Название, период, комментарий')
            for index, page in enumerate(tables, 1):
                with st.container(border=True, width=SLIDE_PREVIEW_WIDTH, height=SLIDE_PREVIEW_HEIGHT):
                    st.text(f'Таблица {index} из {len(tables)}')
                    st.caption('План, факт и отклонения')
            for unit, page in charts:
                with st.container(border=True, width=SLIDE_PREVIEW_WIDTH, height=SLIDE_PREVIEW_HEIGHT):
                    st.text('График · ' + unit)
                    bounds = axis_ranges.get(unit.casefold())
                    if bounds is not None:
                        st.text(f'Шкала {report_data.format_number(bounds[0])}…'
                                f'{report_data.format_number(bounds[1])}')
                    elif st.session_state['axis_settings'][unit.casefold()]['fixed']:
                        st.caption('Проверьте границы')
                    else:
                        st.caption('Авто')
    else:
        st.caption('После проверки CSV здесь появится состав презентации: титул, таблицы и графики.')

with st.container(border=True, key='action'):
    result_slot = st.container()
    reason = ''
    if metrics is None:
        reason = 'Исправьте CSV и загрузите его заново' if st.session_state['import_error'] else 'Загрузите CSV'
    elif not title.strip():
        reason = 'Заполните название отчёта'
    elif not period.strip():
        reason = 'Заполните период'
    elif validated_fields is None:
        reason = 'Исправьте поля оформления'
    elif st.session_state['axis_url_error']:
        reason = 'Сбросьте шкалы или откройте исправленную ссылку'
    elif axis_error:
        reason = 'Исправьте шкалы графиков'
    elif axis_warning:
        reason = 'Расширьте шкалу или выключите фиксацию'
    if st.button('Создать заново' if st.session_state['pptx'] else 'Создать презентацию',
                  key='generate', type='secondary' if st.session_state['pptx'] else 'primary',
                  disabled=metrics is None or validated_fields is None or axis_error is not None
                  or axis_warning is not None or st.session_state['axis_url_error'] is not None):
        try:
            with st.spinner('Создаю презентацию…'):
                result = build_presentation(metrics, *validated_fields, axis_ranges=axis_ranges)
            st.session_state['pptx'] = result
            st.rerun()
        except asyncio.CancelledError:
            raise
        except Exception as error:
            LOGGER.exception('Не удалось собрать PPTX для CSV %s', profile[1],
                             exc_info=(RuntimeError, RuntimeError('Ошибка сборки PPTX'), error.__traceback__))
            previous = ' Прежний успешный PPTX сохранён для этих же данных.' if st.session_state['pptx'] else ''
            st.error('Не удалось собрать файл PPTX. Введённые данные сохранены; подробности записаны в журнал сервера.'
                     + previous)

    if reason:
        st.caption(reason)
    if st.session_state['pptx'] is not None:
        with result_slot:
            st.badge('Готово', color='green')
            st.text(f'report-slides.pptx · {slide_count} слайдов · {len(st.session_state["pptx"]) // 1024} КиБ')
            st.download_button('Скачать PPTX', st.session_state['pptx'], file_name='report-slides.pptx',
                               mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                               key='download_pptx', on_click='ignore', type='primary')
    st.caption('При обновлении страницы или остановке приложения рабочая сессия может исчезнуть. '
               'Скачайте файл до обновления страницы.')

with st.container(key='next_steps'):
    st.subheader('Что дальше', anchor=False)
    st.markdown('1. Откройте PPTX и при необходимости отредактируйте текст, таблицы и данные диаграмм.\n'
                '2. Для следующей недели откройте закладку со шкалами, нажмите «Загрузить новый CSV» '
                'и заполните период. Название сохранится, комментарий очистится.\n'
                '3. Правки из PowerPoint в следующий экспорт не переносятся.')
