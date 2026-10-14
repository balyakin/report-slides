"""Вымышленные макеты для ручной приёмки; пользовательские файлы не нужны."""

from decimal import Decimal
from io import BytesIO
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation
from data import Metric, parse_csv
from slides import build_presentation


def main(output):
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
    for name, source, title, period, comment, expected in cases:
        blob = build_presentation(source, title, period, comment)
        assert len(Presentation(BytesIO(blob)).slides) == expected
        (output / f'{name}.pptx').write_bytes(blob)
        print(f'{name}.pptx: {expected} слайдов')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('Укажите каталог: python tests/make_previews.py /tmp/report-slides-previews')
    main(Path(sys.argv[1]))
