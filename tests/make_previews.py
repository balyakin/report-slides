"""Вымышленные макеты для ручной приёмки; пользовательские файлы не нужны."""

from decimal import Decimal
from io import BytesIO
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation
from data import Metric, parse_csv
from slides import build_presentation


def main(output: Path) -> None:
    """Write synthetic presentations for complete visual acceptance

    Args:
        output: External directory receiving generated PPTX files

    Returns:
        None

    Raises:
        OSError: If the output directory or files cannot be written
    """
    root = Path(__file__).resolve().parents[1]
    output.mkdir(parents=True, exist_ok=True)
    sample = parse_csv((root / 'samples/weekly-report.csv').read_bytes(), 'weekly-report.csv', ';')
    long_names = tuple(Metric(f'{n + 1:02}' + 'Ш' * 58, Decimal('-999999999.99'),
                              Decimal('999999999.99'), 'Щ' * 16) for n in range(12))
    unique = tuple(Metric(f'{n + 1:02}' + 'Ш' * 58, Decimal('0.01'),
                         Decimal('1000000000'), f'{n + 1:02}' + 'Щ' * 14) for n in range(12))
    mixed = (Metric('Отрицательные значения', Decimal('-1000000000'), Decimal('-999999999.99'), 'руб.'),
             Metric('Переход через ноль', Decimal('-0.01'), Decimal('0.29'), 'руб.'),
             Metric('Процент в исходной шкале', Decimal('25'), Decimal('30'), '%'),
             Metric('Нулевой показатель А', Decimal('0'), Decimal('0'), 'шт.'),
             Metric('Нулевой показатель Б', Decimal('0'), Decimal('0'), 'шт.'))
    cases = [('sample', sample, 'Еженедельный отчёт', '5–11 октября 2026', 'Данные вымышленные.', 5),
             ('long-labels', long_names, 'Ш' * 100, 'Щ' * 60, '\n'.join(['Ш' * 65] * 6), 7),
             ('many-units', unique, 'Отчёт: 12 разных единиц', 'Щ' * 60, '', 16),
             ('negative-zero-percent', mixed, 'Отрицательные, нулевые и проценты', 'Неделя 41', '', 6),
             ('blank-comment', sample[:1], 'Отчёт с пустыми абзацами', 'Неделя 41',
              'Начало' + '\n' * 30 + 'Конец комментария', 3)]
    cases = [case + (None,) for case in cases]
    week_one = (Metric('Общий показатель', Decimal(100), Decimal(135), 'шт.'),
                Metric('Другой показатель', Decimal(120), Decimal(180), 'шт.'))
    week_two = (Metric('Общий показатель', Decimal(100), Decimal(135), 'шт.'),
                Metric('Другой показатель', Decimal('50.25'), Decimal('160.25'), 'шт.'))
    extremes = (Metric('А' + 'Ш' * 59, Decimal('-1000000000'), Decimal('1000000000'), 'шт.'),
                Metric('Б' + 'Ш' * 59, Decimal('-500000000.01'), Decimal('500000000.01'), 'шт.'),
                Metric('В' + 'Ш' * 59, Decimal('-1000000000'), Decimal('1000000000'), 'шт.'),
                Metric('Г' + 'Ш' * 59, Decimal('-500000000'), Decimal('500000000'), 'шт.'))
    dense = (Metric('А' + 'Ш' * 59, Decimal(0), Decimal(900000000), 'шт.'),
             Metric('Б' + 'Ш' * 59, Decimal(100000000), Decimal(800000000), 'шт.'),
             Metric('В' + 'Ш' * 59, Decimal(0), Decimal(900000000), 'шт.'),
             Metric('Г' + 'Ш' * 59, Decimal(100000000), Decimal(800000000), 'шт.'))
    pages = (Metric('А', Decimal(100), Decimal(135), 'шт.'), Metric('Б', Decimal(0), Decimal(200), 'шт.'),
             Metric('В', Decimal(50), Decimal(150), 'шт.'), Metric('Г', Decimal(10), Decimal(20), 'шт.'),
             Metric('Д', Decimal(100), Decimal(135), 'шт.'))
    fixed = {'шт.': (Decimal(0), Decimal(200))}
    cases.extend([
        ('fixed-week-1', week_one, 'Фиксированный отчёт', 'Неделя 1', '', 3, fixed),
        ('fixed-week-2', week_two, 'Фиксированный отчёт', 'Неделя 2', '', 3, fixed),
        ('fixed-negative', (Metric('Границы', Decimal(-200), Decimal(0), 'шт.'),),
         'Отрицательные значения', 'Неделя', '', 3, {'шт.': (Decimal(-200), Decimal(0))}),
        ('fixed-mixed', (Metric('Границы', Decimal(-100), Decimal(100), 'шт.'),),
         'Переход через ноль', 'Неделя', '', 3, {'шт.': (Decimal(-100), Decimal(100))}),
        ('fixed-fractional', (Metric('Границы', Decimal('-0.03'), Decimal('0.02'), 'шт.'),),
         'Малые значения', 'Неделя', '', 3, {'шт.': (Decimal('-0.03'), Decimal('0.02'))}),
        ('fixed-zero', (Metric('Нули', Decimal(0), Decimal(0), 'шт.'),),
         'Нулевые данные', 'Неделя', '', 3, {'шт.': (Decimal(0), Decimal('0.02'))}),
        ('fixed-extremes', extremes, 'Крайние значения', 'Неделя', '', 3,
         {'шт.': (Decimal('-1000000000'), Decimal('1000000000'))}),
        ('fixed-dense', dense, 'Плотные деления и длинные названия', 'Неделя', '', 3,
         {'шт.': (Decimal(0), Decimal(900000000))}),
        ('fixed-asymmetric', tuple(Metric(metric.name, Decimal('-1000000000'), Decimal('800000000'), metric.unit)
                                   for metric in extremes), 'Длинные отрицательные деления', 'Неделя', '', 3,
         {'шт.': (Decimal('-1000000000'), Decimal('800000000'))}),
        ('fixed-pages', pages, 'Две страницы одной единицы', 'Неделя', '', 5, fixed),
    ])
    for name, source, title, period, comment, expected, ranges in cases:
        blob = build_presentation(source, title, period, comment, axis_ranges=ranges)
        assert len(Presentation(BytesIO(blob)).slides) == expected
        (output / f'{name}.pptx').write_bytes(blob)
        print(f'{name}.pptx: {expected} слайдов')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('Укажите каталог: python tests/make_previews.py /tmp/report-slides-previews')
    main(Path(sys.argv[1]))
