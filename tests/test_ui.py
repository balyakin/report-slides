from io import BytesIO
from pathlib import Path
import socket
import unittest
from unittest.mock import patch

from pptx import Presentation
from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = (ROOT / 'samples/weekly-report.csv').read_bytes()


class InterfaceTests(unittest.TestCase):
    def test_field_feedback_waits_for_editing_and_uses_an_error_alert(self) -> None:
        """Avoid premature field errors while retaining feedback after user interaction"""
        # ARRANGE
        app = self.app()
        self.assertFalse(app.error)
        self.assertFalse(any('Период: заполните поле' in message.value for message in app.caption))
        self.upload(app)
        self.assertFalse(app.error)
        # ACT
        app.text_input(key='period').set_value('Неделя').run()
        app.text_input(key='period').set_value('').run()
        # ASSERT
        self.assertEqual([message.value for message in app.error], ['Период: заполните поле.'])
        self.assertTrue(app.button(key='generate').disabled)
        app.button(key='new_csv').click().run()
        self.assertFalse(app.error)
        self.assertFalse(app.session_state['edited_fields'])

    def test_feedback_does_not_mark_other_untouched_fields_as_invalid(self) -> None:
        """Validate edited fields independently even when the period is still empty"""
        # ARRANGE
        app = self.upload(self.app())
        # ACT
        app.text_input(key='title').set_value('Другой отчёт').run()
        # ASSERT
        self.assertFalse(app.error)
        self.assertTrue(app.button(key='generate').disabled)
        app.text_area(key='comment').set_value('а' * 401).run()
        self.assertEqual([message.value for message in app.error],
                         ['Комментарий: допускается не больше 400 символов.'])

    def test_title_feedback_survives_upload_and_reset(self) -> None:
        """Keep explaining an invalid title while its value survives CSV replacement"""
        # ARRANGE
        app = self.app()
        title_error = ['Название отчёта: заполните поле.']
        # ACT
        app.text_input(key='title').set_value('').run()
        # ASSERT
        self.assertEqual([message.value for message in app.error], title_error)
        self.upload(app)
        app.text_input(key='period').set_value('Неделя').run()
        self.assertEqual([message.value for message in app.error], title_error)
        self.assertTrue(app.button(key='generate').disabled)
        app.button(key='new_csv').click().run()
        self.assertEqual([message.value for message in app.error], title_error)
        self.upload(app)
        app.text_input(key='period').set_value('Неделя').run()
        self.assertEqual([message.value for message in app.error], title_error)
        app.text_input(key='title').set_value('Отчёт').run()
        self.assertFalse(app.error)
        self.assertFalse(app.button(key='generate').disabled)

    def test_generation_logs_a_traceback_without_user_content(self) -> None:
        """Retain a diagnosable stack while hiding sensitive exception text"""
        # ARRANGE
        app = self.generated()
        previous = app.session_state['pptx']
        # ACT
        with self.assertLogs('report_slides', level='ERROR') as messages:
            with patch('slides.build_presentation', side_effect=RuntimeError('SECRET XML <csv>')):
                app.button(key='generate').click().run()
        # ASSERT
        self.assertIn('Traceback', '\n'.join(messages.output))
        self.assertNotIn('SECRET', '\n'.join(messages.output))
        self.assertTrue(all(record.exc_info for record in messages.records))
        self.assertTrue(all('SECRET' not in message.value for message in app.error))
        self.assertEqual(app.session_state['pptx'], previous)

    def app(self):
        app = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=15).run()
        self.assertFalse(app.exception)
        return app

    def upload(self, app, raw=SAMPLE, filename='weekly.csv'):
        if not app.file_uploader:
            app.button(key='new_csv').click().run()
        app.file_uploader[0].set_value((filename, raw, 'text/csv')).run()
        self.assertFalse(app.exception)
        return app

    def ready(self):
        app = self.upload(self.app())
        app.text_input(key='period').set_value('5–11 октября 2026').run()
        app.text_area(key='comment').set_value('Комментарий').run()
        self.assertFalse(app.exception)
        return app

    def generated(self):
        app = self.ready()
        app.button(key='generate').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(Presentation(BytesIO(app.session_state['pptx'])).slides), 5)
        return app

    def assert_no_export(self, app):
        self.assertIsNone(app.session_state['pptx'])
        self.assertFalse(any(button.label == 'Скачать PPTX' for button in app.download_button))

    def test_empty_app_and_session_isolation(self):
        app = self.app()
        self.assertEqual(app.title[0].value, 'Report Slides')
        self.assertEqual(app.text_input(key='title').value, 'Еженедельный отчёт')
        self.assertEqual(app.text_input(key='period').value, '')
        self.assertTrue(app.button(key='generate').disabled)
        self.assertIsNone(app.session_state['metrics'])
        self.assert_no_export(app)
        self.assertTrue(any('сессия' in message.value for message in app.info))
        another = self.app()
        self.upload(app)
        self.assertIsNone(another.session_state['metrics'])

    def test_metadata_fields_wait_for_valid_csv_during_replacement(self):
        app = self.app()
        self.assertTrue(app.text_input(key='period').disabled)
        self.assertTrue(app.text_area(key='comment').disabled)
        self.assertFalse(app.text_input(key='title').disabled)
        self.upload(app)
        self.assertFalse(app.text_input(key='period').disabled)
        self.assertFalse(app.text_area(key='comment').disabled)
        app.button(key='new_csv').click().run()
        self.assertTrue(app.text_input(key='period').disabled)
        self.upload(app, SAMPLE.replace(b'120', b'NaN'))
        self.assertTrue(app.text_input(key='period').disabled)
        self.assertTrue(app.text_area(key='comment').disabled)

    def test_real_upload_preview_composition_and_export_only_by_button(self):
        app = self.ready()
        self.assert_no_export(app)
        preview = app.dataframe[0].value
        self.assertEqual(preview['Отклонение'].tolist(), ['15', '25000,25', '-0,5', '3'])
        self.assertEqual(preview['Факт/план, %'].tolist(), ['112,5', '110,0', '80,0', '—'])
        self.assertTrue(any('5 слайдов' in text.value for text in app.text))
        app.button(key='generate').click().run()
        blob = app.session_state['pptx']
        self.assertEqual(len(Presentation(BytesIO(blob)).slides), 5)
        app.run()
        self.assertEqual(app.session_state['pptx'], blob)
        self.assertTrue(app.download_button(key='download_pptx').proto.url.endswith('.pptx'))
        self.assertTrue(app.download_button(key='download_pptx').proto.ignore_rerun)

    def test_replacing_filename_or_bytes_clears_fields_and_export_but_preserves_title(self):
        for raw, filename in ((SAMPLE, 'new.csv'), (SAMPLE.replace(b'120', b'121'), 'weekly.csv')):
            with self.subTest(filename=filename):
                app = self.generated()
                self.upload(app, raw, filename)
                self.assert_no_export(app)
                self.assertEqual(app.text_input(key='title').value, 'Еженедельный отчёт')
                self.assertEqual(app.text_input(key='period').value, '')
                self.assertEqual(app.text_area(key='comment').value, '')
                self.assertEqual(len(app.session_state['metrics']), 4)
                self.assertTrue(app.button(key='generate').disabled)

    def test_invalid_replacement_removes_old_data_and_download(self):
        app = self.generated()
        self.upload(app, SAMPLE.replace(b'120', b'NaN'))
        self.assertIsNone(app.session_state['metrics'])
        self.assertEqual(len(app.dataframe), 0)
        self.assert_no_export(app)
        self.assertTrue(app.error)
        self.assertIn('Строка 2', app.error[0].value)
        self.assertEqual(app.text_input(key='period').value, '')

    def test_removal_and_delimiter_change_clear_old_result(self):
        app = self.generated()
        app.button(key='new_csv').click().run()
        self.assertFalse(app.exception)
        self.assertIsNone(app.session_state['metrics'])
        self.assert_no_export(app)
        self.assertEqual(app.text_input(key='period').value, '')
        app = self.generated()
        app.radio(key='delimiter').set_value(',').run()
        self.assertFalse(app.exception)
        self.assertIsNone(app.session_state['metrics'])
        self.assert_no_export(app)
        self.assertEqual(app.text_area(key='comment').value, 'Комментарий')
        self.assertEqual(app.text_input(key='period').value, '5–11 октября 2026')
        app.radio(key='delimiter').set_value(';').run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.session_state['metrics']), 4)
        self.assertEqual(app.text_area(key='comment').value, 'Комментарий')
        self.assertEqual(app.text_input(key='period').value, '5–11 октября 2026')
        self.assert_no_export(app)

    def test_report_field_changes_invalidate_result_keep_data(self):
        for key, value in (('title', 'Другой отчёт'), ('period', 'Новая неделя'),
                            ('comment', 'Новый комментарий')):
            with self.subTest(key=key):
                app = self.generated()
                before = app.session_state['metrics']
                widget = app.text_area(key=key) if key == 'comment' else app.text_input(key=key)
                widget.set_value(value).run()
                self.assertFalse(app.exception)
                self.assert_no_export(app)
                self.assertEqual(app.session_state['metrics'], before)

    def test_invalid_report_fields_disable_generation_without_losing_data(self):
        for key, value in (('title', ' '), ('period', ''), ('comment', 'а' * 401)):
            with self.subTest(key=key):
                app = self.ready()
                widget = app.text_area(key=key) if key == 'comment' else app.text_input(key=key)
                widget.set_value(value).run()
                self.assertFalse(app.exception)
                self.assertTrue(app.button(key='generate').disabled)
                self.assertEqual(len(app.session_state['metrics']), 4)
                self.assert_no_export(app)

    def test_generation_failure_preserves_inputs_and_previous_success_only_if_unchanged(self):
        app = self.generated()
        previous = app.session_state['pptx']
        with patch('slides.build_presentation', side_effect=RuntimeError('SECRET XML <csv>')):
            app.button(key='generate').click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['pptx'], previous)
            self.assertTrue(any(button.label == 'Скачать PPTX' for button in app.download_button))
            self.assertTrue(app.error)
            self.assertNotIn('SECRET', app.error[0].value)
            app.text_input(key='period').set_value('Новая неделя').run()
            app.button(key='generate').click().run()
            self.assert_no_export(app)
        self.assertEqual(app.text_input(key='period').value, 'Новая неделя')
        self.assertEqual(app.text_area(key='comment').value, 'Комментарий')
        self.assertEqual(len(app.session_state['metrics']), 4)

    def test_no_network_during_upload_preview_generation_and_rerun(self):
        app = self.app()
        with patch.object(socket.socket, 'connect', side_effect=AssertionError('External network')):
            self.upload(app)
            app.text_input(key='period').set_value('Неделя').run()
            app.button(key='generate').click().run()
            app.run()
        self.assertFalse(app.exception)
        self.assertEqual(len(Presentation(BytesIO(app.session_state['pptx'])).slides), 5)

    def test_preview_does_not_interpret_user_markdown_or_urls(self):
        app = self.upload(self.app(), 'показатель;план;факт;единица\n**А**;1;2;<b>ч</b>\nhttps://example.test;1;2;шт.\n'.encode())
        preview = app.dataframe[0].value
        self.assertEqual(preview['Показатель'].tolist(), ['**А**', 'https://example.test'])
        self.assertEqual(preview['Единица'].tolist(), ['<b>ч</b>', 'шт.'])
        self.assertEqual(len(app.get('link_button')), 0)


    def test_replacement_starts_by_clearing_report_before_new_file_is_selected(self):
        app = self.generated()
        self.assertEqual(len(app.file_uploader), 0)
        previous_revision = app.session_state['upload_revision']
        app.button(key='new_csv').click().run()
        self.assertFalse(app.exception)
        self.assertIsNone(app.session_state['metrics'])
        self.assertIsNone(app.session_state['source_file'])
        self.assert_no_export(app)
        self.assertIsNone(app.file_uploader[0].value)
        self.assertFalse(app.file_uploader[0].disabled)
        self.assertEqual(app.session_state['upload_revision'], previous_revision + 1)
        self.assertEqual(app.text_input(key='title').value, 'Еженедельный отчёт')
        self.assertEqual(app.text_input(key='period').value, '')
        self.assertEqual(app.text_area(key='comment').value, '')
        self.upload(app, SAMPLE.replace(b'120', b'NaN'))
        self.assertIsNone(app.session_state['metrics'])
        self.assert_no_export(app)


if __name__ == '__main__':
    unittest.main()
