import csv
from decimal import Decimal, localcontext
from io import StringIO
import json
from pathlib import Path
import unittest

import data as report_data
from data import CSV_HEADER, group_metrics, parse_csv, table_rows, validate_report, validate_report_field



HEADER = 'показатель;план;факт;единица\n'
SAMPLE = Path(__file__).resolve().parents[1] / 'samples/weekly-report.csv'


def content(rows, delimiter=';', newline='\n'):
    stream = StringIO(newline='')
    writer = csv.writer(stream, delimiter=delimiter, lineterminator=newline)
    writer.writerow(['показатель', 'план', 'факт', 'единица'])
    writer.writerows(rows)
    return stream.getvalue().encode('utf-8')


class CSVTests(unittest.TestCase):
    def test_excel_grouping_errors_explain_how_to_resave_csv(self) -> None:
        """Reject grouping spaces with an actionable Excel message"""
        # ARRANGE
        amounts = ('250 000,50', '250\u00a0000,50', '250\u202f000,50')
        for amount in amounts:
            raw = content([['Выручка', amount, '275000', 'руб.']])
            # ACT
            with self.subTest(amount=amount), self.assertRaises(ValueError) as failure:
                parse_csv(raw, 'weekly.csv', ';')
            # ASSERT
            self.assertRegex(str(failure.exception), 'Excel.*разделитель групп')

    def test_percent_errors_explain_the_required_scale(self) -> None:
        """Explain the numeric percentage scale without silently converting input"""
        # ARRANGE
        raw = content([['Доля', '25%', '30', '%']])
        # ACT
        with self.assertRaisesRegex(ValueError, '25 вместо 25%'):
            parse_csv(raw, 'weekly.csv', ';')
        # ASSERT
        self.assertEqual(table_rows(parse_csv(content([['Доля', '25', '30', '%']]), 'weekly.csv', ';')),
                         [['Доля', '%', '25', '30', '5', '120,0']])

    def test_extra_column_errors_name_the_count_and_excel_columns(self) -> None:
        """Report trailing separators precisely without relaxing the four-column contract"""
        # ARRANGE
        documents = (
            ('показатель;план;факт;единица;\nА;1;2;шт.;\n', 'Строка 1'),
            (HEADER + 'А;1;2;шт.;\n', 'Строка 2'),
        )
        for document, location in documents:
            # ACT
            with self.subTest(location=location), self.assertRaises(ValueError) as failure:
                parse_csv(document.encode(), 'weekly.csv', ';')
            # ASSERT
            self.assertRegex(str(failure.exception), location + '.*5 вместо 4.*колонки справа')

    def test_missing_columns_and_wrong_separator_show_the_expected_header(self) -> None:
        """Explain missing columns without advising the user to delete more columns"""
        # ARRANGE
        cases = (
            (HEADER + 'А;1;2;шт.\n', ','),
            ('показатель;план;факт\nА;1;2\n', ';'),
            (HEADER + 'А;1;2\n', ';'),
        )
        for document, delimiter in cases:
            # ACT
            with self.subTest(delimiter=delimiter), self.assertRaises(ValueError) as failure:
                parse_csv(document.encode(), 'weekly.csv', delimiter)
            # ASSERT
            message = str(failure.exception)
            self.assertIn('ожидаемый заголовок «' + delimiter.join(CSV_HEADER) + '»', message)
            self.assertIn('выбранный разделитель «' + delimiter + '»', message)
            self.assertNotIn('колонки справа', message)

    def test_quote_error_lines_match_cr_crlf_and_lf(self) -> None:
        """Count physical lines consistently for all supported line endings"""
        # ARRANGE
        for newline in ('\r', '\r\n', '\n'):
            raw = newline.join(('показатель;план;факт;единица', 'А;1;2;шт.', 'Б; "1";2;шт.')).encode()
            # ACT
            with self.subTest(newline=newline), self.assertRaises(ValueError) as failure:
                parse_csv(raw, 'weekly.csv', ';')
            # ASSERT
            self.assertRegex(str(failure.exception), 'Строка 3: уберите пробелы перед открывающей кавычкой')

    def test_sample_values_and_all_calculations(self):
        metrics = parse_csv(SAMPLE.read_bytes(), SAMPLE.name, ';')
        self.assertEqual(table_rows(metrics), [
            ['Заявки', 'шт.', '120', '135', '15', '112,5'],
            ['Выручка', 'руб.', '250000,5', '275000,75', '25000,25', '110,0'],
            ['Время ответа', 'ч', '2,5', '2', '-0,5', '80,0'],
            ['Возвраты', 'шт.', '0', '3', '3', '—'],
        ])
        self.assertIsInstance(metrics[1].plan, Decimal)

    def test_bom_crlf_quotes_both_delimiters_and_outer_spaces(self):
        for delimiter in (';', ','):
            with self.subTest(delimiter=delimiter):
                raw = b'\xef\xbb\xbf' + content([
                    [f'  Заявки{delimiter} "новые"  ', ' +1,25 ', '-0.50', ' шт. '],
                ], delimiter, '\r\n')
                metric = parse_csv(raw, 'weekly.CSV', delimiter)[0]
                self.assertEqual((metric.name, metric.plan, metric.actual, metric.unit),
                                 (f'Заявки{delimiter} "новые"', Decimal('1.25'), Decimal('-0.5'), 'шт.'))

    def test_header_case_whitespace_and_empty_records(self):
        raw = ' ПОКАЗАТЕЛЬ ; План ; ФАКТ ; Единица \n\n;;;\nА;1;2;шт.\n ; ; ; \n'
        self.assertEqual(len(parse_csv(raw.encode(), 'a.csv', ';')), 1)

    def test_bad_headers_and_wrong_separator(self):
        for header in ('план;показатель;факт;единица', 'показатель;план;факт',
                       'показатель;план;факт;единица;лишнее', 'name;plan;actual;unit'):
            with self.subTest(header=header), self.assertRaisesRegex(ValueError, 'заголовок|колонк'):
                parse_csv((header + '\nА;1;2;шт.').encode(), 'a.csv', ';')
        with self.assertRaisesRegex(ValueError, 'разделитель'):
            parse_csv((HEADER + 'А;1;2;шт.').encode(), 'a.csv', ',')

    def test_invalid_encoding_extension_delimiter_and_empty_file(self):
        for raw, filename, delimiter, message in (
            ((HEADER + 'А;1;2;шт.').encode('cp1251'), 'a.csv', ';', 'UTF-8'),
            (b'', 'a.csv', ';', 'заголовок'),
            (HEADER.encode(), 'a.csv', ';', 'показател'),
            (HEADER.encode(), 'a.txt', ';', '\\.csv'),
            (HEADER.encode(), 'a.csv', '|', 'разделитель'),
        ):
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                parse_csv(raw, filename, delimiter)

    def test_malformed_csv_and_wrong_field_count(self):
        for record in ('"А;1;2;шт.', '"А"х;1;2;шт.', 'А"Б;1;2;шт.',
                       'А;1;2', 'А;1;2;шт.;5', 'А;1;;шт.', ';1;2;шт.'):
            with self.subTest(record=record), self.assertRaisesRegex(ValueError, 'Строка 2'):
                parse_csv((HEADER + record).encode(), 'a.csv', ';')

    def test_limits_one_twelve_thirteen_and_exact_byte_limit(self):
        for count in (1, 12):
            raw = content([[f'Показатель {n}', '1', '2', 'шт.'] for n in range(count)])
            self.assertEqual(len(parse_csv(raw, 'a.csv', ';')), count)
        with self.assertRaisesRegex(ValueError, 'Строка 14.*12'):
            parse_csv(content([[str(n), '1', '2', 'шт.'] for n in range(13)]), 'a.csv', ';')
        raw = content([['А', '1', '2', 'шт.']])
        padded = raw + b'\n' * (262144 - len(raw))
        self.assertEqual(len(parse_csv(padded, 'a.csv', ';')), 1)
        with self.assertRaisesRegex(ValueError, '262 144|256'):
            parse_csv(padded + b'\n', 'a.csv', ';')

    def test_byte_limit_allows_large_outer_whitespace_before_short_name(self):
        header, row = HEADER.encode(), 'А;1;2;шт.\n'.encode()
        raw = header + b' ' * (262144 - len(header) - len(row)) + row
        self.assertEqual(len(raw), 262144)
        self.assertEqual(table_rows(parse_csv(raw, 'a.csv', ';')), [['А', 'шт.', '1', '2', '1', '200,0']])

    def test_numeric_boundaries_and_negative_zero(self):
        metrics = parse_csv(content([
            ['А', '-1000000000', '1000000000', '%'],
            ['Б', '+0.00', '-0', 'ч'],
        ]), 'a.csv', ';')
        self.assertEqual(table_rows(metrics), [
            ['А', '%', '-1000000000', '1000000000', '2000000000', '—'],
            ['Б', 'ч', '0', '0', '0', '—'],
        ])

    def test_rejects_invalid_numbers_without_rounding(self):
        for number in ('1.234', '1e2', '1 000', '1\t0', 'NaN', 'Inf', '∞',
                       '1%', '₽1', '.5', '1.', '--1', '１２', '1000000000.01',
                       '-1000000000.01', '0' * 21, ''):
            for column in (1, 2):
                row = ['А', '1', '2', 'шт.']
                row[column] = number
                with self.subTest(number=number, column=column), self.assertRaisesRegex(
                    ValueError, f'Строка 2.*{("план", "факт")[column - 1]}'
                ):
                    parse_csv(content([row]), 'a.csv', ';')

    def test_text_lengths_duplicates_and_forbidden_characters_before_strip(self):
        accepted = [['Я' * 60, '1', '2', 'ю' * 16]]
        self.assertEqual(parse_csv(content(accepted), 'a.csv', ';')[0].name, 'Я' * 60)
        for name, unit in [('Я' * 61, 'шт.'), ('А', 'ю' * 17),
                           ('\tА', 'шт.'), ('А\n', 'шт.'), ('А', '\rшт.'),
                           ('А\x00', 'шт.'), ('А', 'шт.\uffff')]:
            with self.subTest(name=name, unit=unit), self.assertRaisesRegex(ValueError, 'Строка 2'):
                parse_csv(content([[name, '1', '2', unit]]), 'a.csv', ';')
        with self.assertRaisesRegex(ValueError, 'Строка 3.*показатель'):
            parse_csv(content([['А', '1', '2', 'шт.'], [' а ', '1', '2', 'шт.']]), 'a.csv', ';')

    def test_group_order_case_and_no_synonym_merging(self):
        metrics = parse_csv(content([
            ['А', '1', '2', ' Руб. '], ['Б', '1', '2', 'шт.'],
            ['В', '1', '2', 'руб.'], ['Г', '1', '2', '₽'], ['Д', '1', '2', 'рублей'],
        ]), 'a.csv', ';')
        self.assertEqual([(unit, [m.name for m in group]) for unit, group in group_metrics(metrics)],
                         [('Руб.', ['А', 'В']), ('шт.', ['Б']), ('₽', ['Г']), ('рублей', ['Д'])])

    def test_round_half_up_and_context_independent_calculation(self):
        with localcontext() as context:
            context.prec = 3
            metrics = parse_csv(content([
                ['А', '16', '1', 'шт.'], ['Б', '16', '-1', 'шт.'],
                ['В', '0.01', '1000000000', 'руб.'],
            ]), 'a.csv', ';')
            rows = table_rows(metrics)
        self.assertEqual([row[-1] for row in rows], ['6,3', '-6,3', '10000000000000,0'])
        self.assertEqual(rows[-1][-2], '999999999,99')

    def test_user_text_remains_literal(self):
        names = ['=1+1', 'https://example.test', '<А & "Б">']
        metrics = parse_csv(content([[name, '1', '2', 'шт.'] for name in names]), 'a.csv', ';')
        self.assertEqual([row[0] for row in table_rows(metrics)], names)

    def test_negative_zero_ratio_is_unsigned(self):
        metrics = parse_csv(content([['А', '1', '-0.00', 'шт.']]), 'a.csv', ';')
        self.assertEqual(table_rows(metrics)[0], ['А', 'шт.', '1', '0', '-1', '0,0'])

    def test_negative_ratio_rounded_to_zero_is_unsigned(self):
        metrics = parse_csv(content([['А', '1000000000', '-0.01', 'шт.']]), 'a.csv', ';')
        self.assertEqual(table_rows(metrics)[0][-1], '0,0')

    def test_empty_record_cannot_hide_forbidden_name_or_unit_characters(self):
        for row in (['\t', '', '', ''], ['', '', '', '\n']):
            with self.subTest(row=row), self.assertRaisesRegex(ValueError, 'Строка 2'):
                parse_csv(content([row, ['А', '1', '2', 'шт.']]), 'a.csv', ';')


    def test_unicode_line_breaks_are_rejected_before_trimming_csv_text(self):
        for line_break in ('\x85', '\u2028', '\u2029'):
            for column in (0, 3):
                for value in (line_break + 'А', 'А' + line_break, 'А' + line_break + 'Б'):
                    row = ['А', '1', '2', 'шт.']
                    row[column] = value
                    with self.subTest(char=repr(line_break), column=column, value=value), self.assertRaisesRegex(ValueError, 'Строка 2'):
                        parse_csv(content([row]), 'a.csv', ';')


class ReportFieldsTests(unittest.TestCase):
    def test_independent_validation_rejects_unknown_field_keys(self) -> None:
        """Report unsupported field keys as validation errors with their original cause"""
        # ARRANGE
        field = 'unknown'
        # ACT
        with self.assertRaises(ValueError) as failure:
            validate_report_field('Текст', field)
        # ASSERT
        self.assertEqual(str(failure.exception), 'Неизвестное поле отчёта.')
        self.assertIsInstance(failure.exception.__cause__, KeyError)

    def test_outer_whitespace_does_not_consume_report_limits(self) -> None:
        """Apply length limits to the text actually exported"""
        # ARRANGE
        title = 'Я' * 100
        period = 'ю' * 60
        comment = 'а' * 400
        # ACT
        fields = validate_report(' ' + title + ' ', ' ' + period + ' ', ' ' + comment + ' ')
        # ASSERT
        self.assertEqual(fields, (title, period, comment))
        self.assertEqual(validate_report('Отчёт', 'Неделя', ' ' * 401), ('Отчёт', 'Неделя', ''))

    def test_tabs_are_rejected_before_trimming_single_line_fields(self) -> None:
        """Keep tabs available only in multiline comments"""
        # ARRANGE
        fields = (('\tОтчёт', 'Неделя', ''), ('Отчёт', 'Неделя\t', ''))
        for report in fields:
            # ACT
            with self.subTest(report=report), self.assertRaisesRegex(ValueError, 'табуляции'):
                validate_report(*report)
            # ASSERT
            self.assertEqual(validate_report('Отчёт', 'Неделя', 'А\tБ'), ('Отчёт', 'Неделя', 'А\tБ'))

    def test_unicode_line_breaks_are_rejected_in_single_line_report_fields(self):
        for line_break in ('\x85', '\u2028', '\u2029'):
            for column, label in ((0, 'Название'), (1, 'Период')):
                fields = ['Отчёт', 'Неделя', '']
                fields[column] = 'А' + line_break + 'Б'
                with self.subTest(char=repr(line_break), column=column), self.assertRaisesRegex(ValueError, label):
                    validate_report(*fields)
            self.assertEqual(validate_report('А', 'Б', 'Начало' + line_break + 'Конец')[-1],
                             'Начало' + line_break + 'Конец')

    def test_valid_report_and_boundaries(self):
        self.assertEqual(validate_report(' Отчёт ', ' Неделя ', 'Текст\n\nЕщё'),
                         ('Отчёт', 'Неделя', 'Текст\n\nЕщё'))
        self.assertEqual(validate_report('Я' * 100, 'ю' * 60, 'а' * 400),
                         ('Я' * 100, 'ю' * 60, 'а' * 400))
        self.assertEqual(validate_report('А', 'Б', '')[-1], '')
        self.assertEqual(validate_report('А', 'Б', '\n'.join(['Абзац'] * 6))[-1].count('\n'), 5)

    def test_invalid_report_fields_and_paragraphs(self):
        for title, period, comment, field in (
            (' ', 'Б', '', 'Название'), ('А', '', '', 'Период'),
            ('Я' * 101, 'Б', '', 'Название'), ('А', 'ю' * 61, '', 'Период'),
            ('А', 'Б', 'а' * 401, 'Комментарий'), ('А', 'Б', 'а\n' * 7, 'Комментарий'),
            ('\nА', 'Б', '', 'Название'), ('А', 'Б\r', '', 'Период'),
            ('А', 'Б', 'текст\x0b', 'Комментарий'),
        ):
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, field):
                validate_report(title, period, comment)


class AxisRangeTests(unittest.TestCase):
    """Exercise fixed bounds and untrusted bookmark payloads"""

    def test_fixed_steps_and_outward_suggestions(self) -> None:
        """Keep endpoints and zero on ticks while suggesting the closest safe expansion"""
        # ARRANGE
        cases = (('0', '200', '50'), ('-200', '0', '50'), ('-100', '100', '50'),
                 ('0', '500', '100'), ('0', '250', '50'), ('0', '120', '20'),
                 ('0', '0.02', '0.01'), ('-0.03', '0.02', '0.01'),
                 ('0', '0.01', '0.01'), ('-0.01', '0', '0.01'), ('0', '0.3', '0.1'))
        # ACT
        for minimum, maximum, expected in cases:
            result = report_data.calculate_fixed_axis_step(Decimal(minimum), Decimal(maximum))
            # ASSERT
            with self.subTest(minimum=minimum, maximum=maximum):
                self.assertEqual(result, Decimal(expected))
                self.assertEqual(Decimal(minimum) % result, 0)
                self.assertEqual(Decimal(maximum) % result, 0)
        self.assertEqual(report_data.suggest_axis_range(Decimal(0), Decimal('0.11')), (Decimal(0), Decimal('0.12')))
        self.assertEqual(report_data.suggest_axis_range(Decimal(0), Decimal(117)), (Decimal(0), Decimal(120)))
        self.assertEqual(report_data.suggest_axis_range(Decimal(0), Decimal(230)), (Decimal(0), Decimal(250)))

    def test_range_parser_accepts_decimal_syntax_and_preserves_context(self) -> None:
        """Parse comma and dot bounds without changing caller precision"""
        # ARRANGE
        cases = ((' -0 ', '+200', (Decimal(0), Decimal(200))),
                 ('-200', '0', (Decimal(-200), Decimal(0))),
                 ('-0,03', '0.02', (Decimal('-0.03'), Decimal('0.02'))))
        # ACT
        for minimum, maximum, expected in cases:
            result = report_data.parse_axis_range(minimum, maximum, 'шт.')
            # ASSERT
            self.assertEqual(result, expected)
        with localcontext() as context:
            context.prec = 3
            result = report_data.calculate_fixed_axis_step(Decimal('-1000000000'), Decimal('1000000000'))
            self.assertEqual(result, Decimal('500000000'))
            self.assertEqual(context.prec, 3)

    def test_range_errors_follow_field_and_rule_order(self) -> None:
        """Give UI diagnostics without Excel-specific advice or silent correction"""
        # ARRANGE
        cases = (('', 'NaN', 'Минимум.*введите число'),
                 ('0', '', 'Максимум.*введите число'),
                 ('0', '1' * 21, 'Максимум.*20 символов'),
                 ('0', '1 %', 'Максимум.*без знака %'),
                 ('0', '1\t0', 'Максимум.*без внутренних пробелов'),
                 ('0', 'NaN', 'Максимум.*обычное десятичное число'),
                 ('0', '1e3', 'Максимум.*экспонента'),
                 ('0', '1.000', 'Максимум.*двух дробных'),
                 ('0', '1000000000.01', 'Максимум.*от −1000000000'),
                 ('0', '0', 'минимум должен быть меньше'),
                 ('200', '100', 'минимум должен быть меньше'),
                 ('10', '200', 'диапазон должен включать ноль'),
                 ('-200', '-10', 'диапазон должен включать ноль'),
                 ('0', '117', 'Подойдёт диапазон 0…120'))
        # ACT
        for minimum, maximum, expected in cases:
            # ASSERT
            with self.subTest(minimum=minimum, maximum=maximum):
                with self.assertRaisesRegex(ValueError, 'Единица «шт.».*' + expected):
                    report_data.parse_axis_range(minimum, maximum, 'шт.')

    def test_direct_api_rejects_bad_bounds_and_unknown_keys(self) -> None:
        """Reject nonfinite and imprecise bounds before exporting any chart"""
        # ARRANGE
        source = (report_data.Metric('Заявки', Decimal(0), Decimal(200), 'ШТ.'),)
        cases = ((Decimal(0),), (Decimal(0), Decimal('NaN')), (Decimal('sNaN'), Decimal(200)),
                 (Decimal(0), Decimal('Infinity')), (Decimal(0), Decimal('1e99999')),
                 (Decimal('0.001'), Decimal(200)), (Decimal(0), Decimal(117)))
        # ACT
        for bounds in cases:
            # ASSERT
            with self.subTest(bounds=bounds), self.assertRaises(ValueError):
                report_data.validate_axis_ranges(source, {'шт.': bounds})
        for key in ('ШТ.', 'руб.'):
            with self.assertRaisesRegex(ValueError, 'Неизвестная или ненормализованная'):
                report_data.validate_axis_ranges(source, {key: (Decimal(0), Decimal(200))})
        report_data.validate_axis_ranges(source, {'шт.': (Decimal('0.000'), Decimal('200.000'))})
        report_data.validate_axis_ranges(source, None)
        report_data.validate_axis_ranges(source, {})

    def test_collects_all_outliers_and_keeps_other_units_automatic(self) -> None:
        """Report plan then actual without confusing table calculations with chart values"""
        # ARRANGE
        source = (report_data.Metric('Заявки', Decimal(201), Decimal(230), 'шт.'),
                  report_data.Metric('Границы', Decimal(0), Decimal(200), 'ШТ.'),
                  report_data.Metric('Другая единица', Decimal(-1000), Decimal(1000), 'руб.'))
        expected = (
            'Показатель «Заявки», единица «шт.», План: 201 вне диапазона 0…200.\n'
            'Показатель «Заявки», единица «шт.», Факт: 230 вне диапазона 0…200.\n'
            'Единица «шт.»: подойдёт диапазон 0…250. Границы не изменены.\n'
            'Измените границы или выключите фиксированную шкалу.'
        )
        # ACT
        with self.assertRaises(ValueError) as caught:
            report_data.validate_axis_ranges(source, {'шт.': (Decimal(0), Decimal(200))})
        # ASSERT
        self.assertEqual(str(caught.exception), expected)

    def test_bookmarks_are_canonical_and_do_not_mutate_settings(self) -> None:
        """Round-trip special units without floats, source files or raw formatting"""
        # ARRANGE
        settings = {'шт.': {'fixed': True, 'minimum': '-0', 'maximum': '200.00'},
                    'ед:ч': {'fixed': True, 'minimum': '0', 'maximum': '0,02'},
                    '%': {'fixed': True, 'minimum': '0', 'maximum': '100'},
                    'a"b': {'fixed': True, 'minimum': '-200', 'maximum': '0'},
                    'авто': {'fixed': False, 'minimum': 'NaN', 'maximum': ''}}
        # ACT
        payload = report_data.serialize_axis_settings(settings)
        restored = report_data.parse_axis_settings(payload)
        # ASSERT
        self.assertEqual(restored, {
            '%': {'fixed': True, 'minimum': '0', 'maximum': '100'},
            'a"b': {'fixed': True, 'minimum': '-200', 'maximum': '0'},
            'ед:ч': {'fixed': True, 'minimum': '0', 'maximum': '0,02'},
            'шт.': {'fixed': True, 'minimum': '0', 'maximum': '200'},
        })
        self.assertEqual(report_data.serialize_axis_settings(restored), payload)
        self.assertEqual(settings.get('шт.'), {'fixed': True, 'minimum': '-0', 'maximum': '200.00'})
        disabled = {'авто': {'fixed': False, 'minimum': 'NaN', 'maximum': ''}}
        self.assertEqual(report_data.serialize_axis_settings(disabled), '')

    def test_untrusted_bookmark_structure_units_and_limits(self) -> None:
        """Reject an entire malformed bookmark rather than restoring a partial configuration"""
        # ARRANGE
        invalid = ('', '{}', 'null', '[]', '["wrong",[]]', '["report-slides-axis-v1",{}]',
                   '["report-slides-axis-v1",[["шт.",0,200]]]',
                   '["report-slides-axis-v1",[["шт.","0"]]]',
                   '["report-slides-axis-v1",[["ШТ.","0","200"]]]',
                   '["report-slides-axis-v1",[[" шт.","0","200"]]]',
                   '["report-slides-axis-v1",[["шт.","0","117"]]]',
                   '["report-slides-axis-v1",[["шт.","0","200"],["шт.","0","100"]]]',
                   '["report-slides-axis-v1",[["\\ud800","0","200"]]]',
                   '[' * 1500 + ']' * 1500, '\ud800', ' ' * 4097)
        # ACT
        for payload in invalid:
            # ASSERT
            with self.subTest(payload=repr(payload)[:80]), self.assertRaises(ValueError):
                report_data.parse_axis_settings(payload)

    def test_bookmark_record_and_encoded_length_boundaries(self) -> None:
        """Accept thirty-two settings and reject larger records or percent-encoded payloads"""
        # ARRANGE
        units = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'
        records = [[unit, '0', '200'] for unit in units]
        valid = json.dumps(['report-slides-axis-v1', records[:-1]])
        excessive = json.dumps(['report-slides-axis-v1', records])
        # ACT
        restored = report_data.parse_axis_settings(valid)
        # ASSERT
        self.assertEqual(len(restored), 32)
        with self.assertRaisesRegex(ValueError, '32 шкал'):
            report_data.parse_axis_settings(excessive)
        with self.assertRaisesRegex(ValueError, 'слишком длинная'):
            report_data.parse_axis_settings('?' * 3000)
        settings = {unit: {'fixed': True, 'minimum': '0', 'maximum': '200'} for unit in units}
        with self.assertRaisesRegex(ValueError, '32 шкал'):
            report_data.serialize_axis_settings(settings)


if __name__ == '__main__':
    unittest.main()
