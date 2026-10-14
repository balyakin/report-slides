"""Строгий импорт CSV и расчёты без двоичной арифметики."""

import csv
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP, localcontext
from io import StringIO
import re
from typing import Dict, Sequence, Tuple


CSV_HEADER = ('показатель', 'план', 'факт', 'единица')
TABLE_HEADER = ('Показатель', 'Единица', 'План', 'Факт', 'Отклонение', 'Факт/план, %')
RATIO_NOTE = 'Отношение не рассчитывается при плане ≤ 0'
MAX_BYTES = 262144
LINE_BREAKS = '\r\n\x85\u2028\u2029'
NUMBER = re.compile(r'[+-]?[0-9]+([.,][0-9]{1,2})?\Z')
GROUP_SPACES = ' \u00a0\u202f'
REPORT_FIELDS: Dict[str, Tuple[str, int]] = {
    'title': ('Название отчёта', 100),
    'period': ('Период', 60),
    'comment': ('Комментарий', 400),
}


@dataclass(frozen=True)
class Metric:
    name: str
    plan: Decimal
    actual: Decimal
    unit: str

    @property
    def difference(self):
        with localcontext() as context:
            context.prec = 28
            return self.actual - self.plan

    @property
    def ratio(self):
        if self.plan <= 0:
            return None
        with localcontext() as context:
            context.prec = 28
            ratio = (self.actual / self.plan * 100).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            return ratio.copy_abs() if ratio == 0 else ratio


def format_number(value):
    if value == 0:
        return '0'
    text = format(value, 'f')
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    return text.replace('.', ',')


def table_rows(metrics):
    return [[metric.name, metric.unit, format_number(metric.plan), format_number(metric.actual),
             format_number(metric.difference),
             '—' if metric.ratio is None else format(metric.ratio, '.1f').replace('.', ',')]
            for metric in metrics]


def _check_xml(value, label):
    for char in value:
        code = ord(char)
        if not (char in '\t\n\r' or 0x20 <= code <= 0xD7FF or
                0xE000 <= code <= 0xFFFD or 0x10000 <= code <= 0x10FFFF):
            raise ValueError(f'{label}: удалите символ, недопустимый в XML.')


def _check_csv_text(value, label):
    _check_xml(value, label)
    if any(char in value for char in '\t' + LINE_BREAKS):
        raise ValueError(f'{label}: переводы строк и табуляция запрещены.')


def _csv_text(value, label, limit):
    value = value.strip()
    if not value or len(value) > limit:
        raise ValueError(f'{label}: нужен непустой текст длиной до {limit} символов.')
    return value


def _number(value: str, label: str) -> Decimal:
    value = value.strip()
    if any(char in value for char in GROUP_SPACES):
        raise ValueError(
            f'{label}: уберите разделитель разрядов в Excel и сохраните CSV заново; '
            'формат ячейки → Числовой → отключить разделитель групп разрядов.'
        )
    if value.endswith('%'):
        raise ValueError(f'{label}: запишите процент числом без знака %, например 25 вместо 25%.')
    if len(value) > 20 or not NUMBER.fullmatch(value):
        raise ValueError(f'{label}: введите число без пробелов, до двух дробных знаков; точка или запятая.')
    number = Decimal(value.replace(',', '.'))
    if not Decimal('-1000000000') <= number <= Decimal('1000000000'):
        raise ValueError(f'{label}: допустимы числа от −1 000 000 000 до 1 000 000 000.')
    return number if number else Decimal(0)


def _check_quotes(text: str, delimiter: str) -> None:
    state = 'start'
    line = 1
    previous = ''
    for char in text:
        if state == 'quoted':
            if char == '"':
                state = 'closed'
        elif state == 'closed':
            if char == '"':
                state = 'quoted'
            elif char == delimiter or char in '\r\n':
                state = 'start'
            else:
                raise ValueError(f'Строка {line}: повреждённый CSV; проверьте двойные кавычки.')
        elif char == '"':
            if state == 'padding':
                raise ValueError(f'Строка {line}: уберите пробелы перед открывающей кавычкой поля.')
            if state != 'start':
                raise ValueError(f'Строка {line}: кавычки внутри поля экранируйте двойными кавычками.')
            state = 'quoted'
        elif char == delimiter or char in '\r\n':
            state = 'start'
        elif char in GROUP_SPACES + '\t' and state in ('start', 'padding'):
            state = 'padding'
        else:
            state = 'bare'
        if char == '\r' or (char == '\n' and previous != '\r'):
            line += 1
        previous = char


def _check_field_count(row: Sequence[str], line: int, delimiter: str) -> None:
    count = len(row)
    if count == len(CSV_HEADER):
        return
    expected = delimiter.join(CSV_HEADER)
    message = (
        f'Строка {line}: число полей — {count} вместо {len(CSV_HEADER)}; '
        f'ожидаемый заголовок «{expected}», выбранный разделитель «{delimiter}».'
    )
    if count > len(CSV_HEADER):
        message += ' Удалите лишние колонки справа на листе Excel.'
    raise ValueError(message)


def parse_csv(content: bytes, filename: str, delimiter: str) -> Tuple[Metric, ...]:
    """Read a strictly validated UTF-8 report without rounding input numbers

    Args:
        content: Original CSV bytes
        filename: Uploaded file name
        delimiter: Semicolon or comma separator

    Returns:
        Metrics in their original order

    Raises:
        ValueError: If the CSV violates the report format or limits
    """
    if not filename.lower().endswith('.csv'):
        raise ValueError('Нужен файл с расширением .csv.')
    if len(content) > MAX_BYTES:
        raise ValueError('Размер CSV не должен превышать 256 КиБ (262 144 байт).')
    if delimiter not in (';', ','):
        raise ValueError('Выберите разделитель ; или ,.')
    try:
        text = content.decode('utf-8-sig')
    except UnicodeDecodeError as error:
        raise ValueError('Сохраните файл как CSV UTF-8: другая кодировка не поддерживается.') from error
    _check_quotes(text, delimiter)
    csv.field_size_limit(MAX_BYTES)
    reader = csv.reader(StringIO(text, newline=''), delimiter=delimiter, strict=True)
    expected = delimiter.join(CSV_HEADER)
    metrics = []
    names = set()
    try:
        header = next(reader, [])
        if header:
            _check_field_count(header, 1, delimiter)
        if tuple(value.strip().casefold() for value in header) != CSV_HEADER:
            raise ValueError(f'Строка 1: ожидаемый заголовок «{expected}»; выбранный разделитель «{delimiter}».')
        while True:
            line = reader.line_num + 1
            row = next(reader, None)
            if row is None:
                break
            if not row:
                continue
            for column, label in ((0, 'показатель'), (3, 'единица')):
                if len(row) > column:
                    _check_csv_text(row[column], f'Строка {line}, поле «{label}»')
            if all(not value.strip() for value in row):
                continue
            _check_field_count(row, line, delimiter)
            if len(metrics) == 12:
                raise ValueError(f'Строка {line}: допускается не больше 12 показателей.')
            name = _csv_text(row[0], f'Строка {line}, поле «показатель»', 60)
            unit = _csv_text(row[3], f'Строка {line}, поле «единица»', 16)
            if name.casefold() in names:
                raise ValueError(f'Строка {line}, поле «показатель»: название должно быть уникальным без учёта регистра.')
            plan = _number(row[1], f'Строка {line}, поле «план»')
            actual = _number(row[2], f'Строка {line}, поле «факт»')
            names.add(name.casefold())
            metrics.append(Metric(name, plan, actual, unit))
    except csv.Error as error:
        raise ValueError(
            f'Строка {max(1, reader.line_num)}: повреждённый CSV; проверьте двойные кавычки и поля.'
        ) from error
    if not metrics:
        raise ValueError('Добавьте от 1 до 12 показателей после заголовка.')
    return tuple(metrics)


def group_metrics(metrics):
    groups = {}
    for metric in metrics:
        groups.setdefault(metric.unit.casefold(), []).append(metric)
    return tuple((group[0].unit, tuple(group)) for group in groups.values())


def validate_report_field(value: str, field: str) -> str:
    """Validate and normalize one report field for independent UI feedback

    Args:
        value: Original field text
        field: Title, period or comment field key

    Returns:
        Text with surrounding whitespace removed

    Raises:
        ValueError: If the field key or text is invalid
    """
    try:
        details = REPORT_FIELDS[field]
    except KeyError as error:
        raise ValueError('Неизвестное поле отчёта.') from error
    label = details[0]
    limit = details[1]
    _check_xml(value, label)
    normalized = value.strip()
    if len(normalized) > limit:
        raise ValueError(f'{label}: допускается не больше {limit} символов.')
    if field != 'comment':
        if any(char in value for char in LINE_BREAKS + '\t'):
            raise ValueError(f'{label}: введите одну строку без переводов строк и табуляции.')
        if not normalized:
            raise ValueError(f'{label}: заполните поле.')
    if field == 'comment' and sum(bool(line.strip()) for line in value.splitlines()) > 6:
        raise ValueError('Комментарий: допускается не больше 6 непустых абзацев.')
    return normalized


def validate_report(title: str, period: str, comment: str) -> Tuple[str, str, str]:
    """Validate report fields and remove surrounding whitespace

    Args:
        title: Required single-line report title
        period: Required single-line reporting period
        comment: Optional title-slide paragraphs

    Returns:
        Normalized title, period and comment

    Raises:
        ValueError: If a field contains forbidden characters or exceeds its limits
    """
    return (validate_report_field(title, 'title'), validate_report_field(period, 'period'),
            validate_report_field(comment, 'comment'))
