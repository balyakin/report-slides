"""Строгий импорт CSV и расчёты без двоичной арифметики."""

import csv
from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING, ROUND_FLOOR, ROUND_HALF_UP, localcontext
from io import StringIO
import json
import re
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple
from urllib.parse import urlencode


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
AXIS_LIMIT = Decimal('1000000000')
AXIS_RESOLUTION = Decimal('0.01')
AXIS_TARGET_INTERVALS = Decimal(4)
AXIS_MIN_INTERVALS = Decimal(2)
AXIS_MAX_INTERVALS = Decimal(10)
AXIS_NUMBER_LENGTH = 20
AXIS_DECIMAL_PLACES = 2
AXIS_STEPS = tuple(Decimal(multiplier) * Decimal(10) ** exponent
                   for exponent in range(-2, 10) for multiplier in (1, 2, 5))
AXIS_VERSION = 'report-slides-axis-v1'
AXIS_QUERY_KEY = 'scales'
AXIS_MAX_SETTINGS = 32
AXIS_MAX_JSON_BYTES = 4096
AXIS_MAX_URL_BYTES = 8192
AXIS_PAYLOAD_SIZE = 2
AXIS_RECORD_SIZE = 3
AXIS_UNIT_LENGTH = 16
AXIS_NUMBER = re.compile(r'[+-]?[0-9]+([.,][0-9]+)?\Z')


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


def _parse_axis_number(text: str, unit: str, field: str) -> Decimal:
    value = text.strip()
    prefix = f'Единица «{unit}», поле «{field}»: '
    if not value:
        raise ValueError(prefix + 'введите число.')
    if len(value) > AXIS_NUMBER_LENGTH:
        raise ValueError(prefix + 'допускается не больше 20 символов.')
    if '%' in value:
        raise ValueError(prefix + 'введите число без знака %.')
    if any(char.isspace() for char in value):
        raise ValueError(prefix + 'введите число без внутренних пробелов и переводов строк.')
    if not AXIS_NUMBER.fullmatch(value):
        raise ValueError(prefix + 'введите обычное десятичное число; экспонента, NaN и бесконечность запрещены.')
    normalized = value.replace(',', '.')
    fraction = normalized.partition('.')[2]
    if len(fraction) > AXIS_DECIMAL_PLACES:
        raise ValueError(prefix + 'допускается не больше двух дробных знаков; округления нет.')
    number = Decimal(normalized)
    if not -AXIS_LIMIT <= number <= AXIS_LIMIT:
        raise ValueError(prefix + 'допустимы значения от −1000000000 до 1000000000.')
    return number if number else Decimal(0)


def _check_axis_bounds(minimum: Decimal, maximum: Decimal) -> None:
    for number in (minimum, maximum):
        if not number.is_finite():
            raise ValueError('границы должны быть конечными числами.')
        if not -AXIS_LIMIT <= number <= AXIS_LIMIT:
            raise ValueError('допустимы значения от −1000000000 до 1000000000.')
        if number.quantize(AXIS_RESOLUTION) != number:
            raise ValueError('границы должны быть представимы с точностью до сотых без округления.')
    if minimum >= maximum:
        raise ValueError('минимум должен быть меньше максимума.')
    if minimum > 0 or maximum < 0:
        raise ValueError('диапазон должен включать ноль.')


def _check_axis_intervals(minimum: Decimal, maximum: Decimal, step: Decimal) -> bool:
    intervals = (maximum - minimum) / step
    small = maximum - minimum == AXIS_RESOLUTION and step == AXIS_RESOLUTION
    return AXIS_MIN_INTERVALS <= intervals <= AXIS_MAX_INTERVALS or small


def calculate_fixed_axis_step(minimum: Decimal, maximum: Decimal) -> Decimal:
    """Choose aligned fixed ticks without changing the requested bounds

    Args:
        minimum: Inclusive lower axis bound
        maximum: Inclusive upper axis bound

    Returns:
        Decimal step shared by reports with these bounds

    Raises:
        ValueError: If bounds are invalid or no suitable ticks exist
    """
    with localcontext() as context:
        context.prec = 28
        _check_axis_bounds(minimum, maximum)
        candidates = []
        for step in AXIS_STEPS:
            aligned = minimum % step == 0 and maximum % step == 0
            if aligned and _check_axis_intervals(minimum, maximum, step):
                intervals = (maximum - minimum) / step
                candidates.append((abs(intervals - AXIS_TARGET_INTERVALS), -step, step))
        if candidates:
            return min(candidates)[2]
        suggested = suggest_axis_range(minimum, maximum)
        raise ValueError(
            f'для диапазона {format_number(minimum)}…{format_number(maximum)} нет подходящих делений. '
            f'Подойдёт диапазон {format_number(suggested[0])}…{format_number(suggested[1])}. Границы не изменены.'
        )


def suggest_axis_range(minimum: Decimal, maximum: Decimal) -> Tuple[Decimal, Decimal]:
    """Suggest the smallest outward expansion with readable aligned ticks

    Args:
        minimum: Required lower bound, including all data
        maximum: Required upper bound, including all data

    Returns:
        Suggested lower and upper bounds without applying them

    Raises:
        ValueError: If the requested bounds are invalid
    """
    with localcontext() as context:
        context.prec = 28
        _check_axis_bounds(minimum, maximum)
        candidates = []
        for step in AXIS_STEPS:
            lower = (minimum / step).to_integral_value(rounding=ROUND_FLOOR) * step
            upper = (maximum / step).to_integral_value(rounding=ROUND_CEILING) * step
            inside = lower >= -AXIS_LIMIT and upper <= AXIS_LIMIT
            if inside and _check_axis_intervals(lower, upper, step):
                expansion = minimum - lower + upper - maximum
                distance = abs((upper - lower) / step - AXIS_TARGET_INTERVALS)
                candidates.append((expansion, distance, -step, lower, upper))
        selected = min(candidates)
        return selected[3], selected[4]


def parse_axis_range(minimum_text: str, maximum_text: str, unit: str) -> Tuple[Decimal, Decimal]:
    """Parse both fixed bounds with field-specific diagnostics

    Args:
        minimum_text: Original minimum input
        maximum_text: Original maximum input
        unit: Display unit used in validation messages

    Returns:
        Validated bounds as Decimals

    Raises:
        ValueError: If either field or the resulting range is invalid
    """
    minimum = _parse_axis_number(minimum_text, unit, 'Минимум')
    maximum = _parse_axis_number(maximum_text, unit, 'Максимум')
    try:
        calculate_fixed_axis_step(minimum, maximum)
    except ValueError as error:
        raise ValueError(f'Единица «{unit}»: {error}') from error
    return minimum, maximum


def validate_axis_ranges(
        metrics: Sequence[Metric], axis_ranges: Optional[Mapping[str, Tuple[Decimal, Decimal]]]) -> None:
    """Validate fixed settings and report every out-of-range source value

    Args:
        metrics: Report metrics in CSV order
        axis_ranges: Fixed ranges keyed by normalized current units

    Returns:
        None

    Raises:
        ValueError: If settings are invalid or source values exceed them
    """
    if not axis_ranges:
        return
    groups = group_metrics(metrics)
    units = {unit.casefold() for unit, group in groups}
    for key in axis_ranges:
        if key not in units or key != key.casefold():
            raise ValueError(f'Неизвестная или ненормализованная единица шкалы: «{key}».')
    for unit, group in groups:
        bounds = axis_ranges.get(unit.casefold())
        if bounds is None:
            continue
        try:
            if len(bounds) != AXIS_PAYLOAD_SIZE:
                raise ValueError('нужны ровно две границы.')
            calculate_fixed_axis_step(bounds[0], bounds[1])
        except ValueError as error:
            raise ValueError(f'Единица «{unit}»: {error}') from error
    messages: List[str] = []
    for metric in metrics:
        bounds = axis_ranges.get(metric.unit.casefold())
        if bounds is not None:
            _collect_axis_violations(metric, bounds, messages)
    if messages:
        messages.extend(_get_axis_suggestions(groups, axis_ranges))
        messages.append('Измените границы или выключите фиксированную шкалу.')
        raise ValueError('\n'.join(messages))


def _collect_axis_violations(metric: Metric, bounds: Tuple[Decimal, Decimal], messages: List[str]) -> None:
    for field, value in (('План', metric.plan), ('Факт', metric.actual)):
        if not bounds[0] <= value <= bounds[1]:
            messages.append(
                f'Показатель «{metric.name}», единица «{metric.unit}», {field}: {format_number(value)} '
                f'вне диапазона {format_number(bounds[0])}…{format_number(bounds[1])}.'
            )


def _get_axis_suggestions(
        groups: Sequence[Tuple[str, Sequence[Metric]]],
        axis_ranges: Mapping[str, Tuple[Decimal, Decimal]]) -> List[str]:
    messages = []
    for unit, group in groups:
        bounds = axis_ranges.get(unit.casefold())
        if bounds is None:
            continue
        values = tuple(value for metric in group for value in (metric.plan, metric.actual))
        minimum = min(bounds[0], *values)
        maximum = max(bounds[1], *values)
        if minimum != bounds[0] or maximum != bounds[1]:
            suggested = suggest_axis_range(minimum, maximum)
            messages.append(
                f'Единица «{unit}»: подойдёт диапазон '
                f'{format_number(suggested[0])}…{format_number(suggested[1])}. Границы не изменены.'
            )
    return messages


def _check_axis_unit(unit: str) -> None:
    _check_csv_text(unit, 'Единица шкалы')
    if not unit or len(unit) > AXIS_UNIT_LENGTH or unit != unit.strip() or unit != unit.casefold():
        raise ValueError('Единица шкалы должна быть непустой, нормализованной и не длиннее 16 символов.')


def _check_axis_payload_size(payload: str) -> None:
    try:
        decoded_size = len(payload.encode('utf-8'))
        encoded_size = len(urlencode({AXIS_QUERY_KEY: payload}).encode('ascii'))
    except UnicodeError as error:
        raise ValueError('Ссылка шкал содержит недопустимый Unicode.') from error
    if decoded_size > AXIS_MAX_JSON_BYTES or encoded_size > AXIS_MAX_URL_BYTES:
        raise ValueError('Ссылка шкал слишком длинная. Сбросьте шкалы и настройте нужные группы заново.')


def serialize_axis_settings(settings: Mapping[str, Mapping[str, Any]]) -> str:
    """Create the canonical URL payload from all enabled settings

    Args:
        settings: Session settings, including temporarily hidden units

    Returns:
        Compact JSON or an empty string when no fixed scales are enabled

    Raises:
        ValueError: If an enabled setting or the payload size is invalid
    """
    records = []
    for unit in sorted(settings):
        setting = settings[unit]
        if setting['fixed']:
            _check_axis_unit(unit)
            bounds = parse_axis_range(setting['minimum'], setting['maximum'], unit)
            records.append([unit, format_number(bounds[0]), format_number(bounds[1])])
    if len(records) > AXIS_MAX_SETTINGS:
        raise ValueError('В ссылке допускается не больше 32 шкал. Сбросьте шкалы и настройте нужные группы заново.')
    if not records:
        return ''
    payload = json.dumps([AXIS_VERSION, records], ensure_ascii=False, separators=(',', ':'))
    _check_axis_payload_size(payload)
    return payload


def parse_axis_settings(payload: str) -> Dict[str, Dict[str, Any]]:
    """Restore a complete validated settings payload from an external URL

    Args:
        payload: Decoded scales query parameter

    Returns:
        Independent session settings for every saved unit

    Raises:
        ValueError: If the payload, structure, units or ranges are invalid
    """
    _check_axis_payload_size(payload)
    try:
        document = json.loads(payload)
    except (ValueError, RecursionError) as error:
        raise ValueError('Не удалось прочитать JSON шкал из ссылки.') from error
    if type(document) is not list or len(document) != AXIS_PAYLOAD_SIZE or document[0] != AXIS_VERSION:
        raise ValueError('Неподдерживаемый формат или версия ссылки шкал.')
    records = document[1]
    if type(records) is not list or len(records) > AXIS_MAX_SETTINGS:
        raise ValueError('В ссылке должен быть список не более 32 шкал.')
    settings: Dict[str, Dict[str, Any]] = {}
    for record in records:
        valid = type(record) is list and len(record) == AXIS_RECORD_SIZE
        if not valid or any(type(value) is not str for value in record):
            raise ValueError('Запись шкалы должна содержать три строки: единицу, минимум и максимум.')
        unit = record[0]
        _check_axis_unit(unit)
        if unit in settings:
            raise ValueError(f'Единица «{unit}»: шкала повторяется в ссылке.')
        bounds = parse_axis_range(record[1], record[2], unit)
        settings[unit] = {'fixed': True, 'minimum': format_number(bounds[0]), 'maximum': format_number(bounds[1])}
    return settings
