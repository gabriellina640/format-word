import unittest
from dataclasses import asdict
from app.config import FormatSettings, SettingsError
from app import ui_fields


class UiValuesTest(unittest.TestCase):
    def test_every_supported_setting_has_a_form_field(self):
        covered = {field.name for field in ui_fields.FIELDS}
        self.assertEqual(covered, set(asdict(FormatSettings())) - {
            'template_path', 'migration_warnings', 'category_rules', 'style_categories'})

    def test_optional_image_width_and_category_state_roundtrip(self):
        self.assertIn('body_image_width_cm', ui_fields.FIELD_MAP)
        settings = FormatSettings(formatting_mode='by_category', category_rules={
            'title': {'mode': 'preserve', 'values': {}}}, style_categories={'Heading1': 'title'})
        values = ui_fields.form_values(settings)
        self.assertEqual(values['body_image_width_cm'], '')
        self.assertEqual(ui_fields.settings_from_values(values, base=settings), settings)
        values['body_image_width_cm'] = '8,5'
        result = ui_fields.settings_from_values(values, base=settings)
        self.assertEqual(result.body_image_width_cm, 8.5)
        result.category_rules['title']['values']['bold'] = True
        self.assertEqual(settings.category_rules['title']['values'], {})

    def test_summary_shows_categories_and_image_policy(self):
        self.assertIn('formatting_mode', ui_fields.FIELD_MAP)
        summary = ui_fields.settings_summary(FormatSettings(formatting_mode='by_category',
            category_rules={'title': {'mode': 'custom', 'values': {'font_size': 18}}},
            body_image_width_cm=8.5, body_image_alignment='center', complex_content_mode='reject'))
        for text in ('Por categoria', 'Título', 'Personalizar', '8,5', 'Centro', 'Recusar'):
            self.assertIn(text, summary)

    def test_roundtrip_preserves_custom_font_decimals_and_tristates(self):
        settings = FormatSettings(font_name='Minha fonte',font_size=12.5,bold=False,
            italic=True,keep_with_next=None,line_spacing=1.65,font_color='123456')
        values = ui_fields.form_values(settings)
        self.assertEqual(ui_fields.settings_from_values(values),settings)

    def test_comma_decimal_is_applied_without_clamping(self):
        values = ui_fields.form_values(FormatSettings())
        values['font_size'] = '36,5'
        values['line_spacing'] = '1,6'
        settings = ui_fields.settings_from_values(values)
        self.assertEqual(settings.font_size,36.5)
        self.assertEqual(settings.line_spacing,1.6)

    def test_invalid_or_empty_numeric_values_are_not_replaced(self):
        for value in ['', 'abc', 'nan', 'inf', '0', '401', '12,3']:
            with self.subTest(value=value):
                values = ui_fields.form_values(FormatSettings())
                values['font_size'] = value
                with self.assertRaises(SettingsError) as caught:
                    ui_fields.settings_from_values(values)
                self.assertEqual(caught.exception.field,'font_size')

    def test_integer_limits_do_not_truncate_decimals(self):
        values = ui_fields.form_values(FormatSettings())
        values['max_input_mb'] = '50,5'
        with self.assertRaises(SettingsError):
            ui_fields.settings_from_values(values)

    def test_summary_uses_effective_paper_and_line_units(self):
        settings = FormatSettings(paper_size='Letter',orientation='landscape',
            line_spacing_mode='exact',line_spacing=17.5,font_size=36)
        summary = ui_fields.settings_summary(settings)
        self.assertIn('36', summary)
        self.assertIn('17,5 pt', summary)
        self.assertIn('27,94', summary)
        self.assertIn('21,59', summary)

    def test_legacy_blocker_survives_collecting_form(self):
        settings = FormatSettings(template_path='legacy.docx')
        with self.assertRaisesRegex(SettingsError,'[Tt]emplate'):
            ui_fields.settings_from_values(ui_fields.form_values(settings),base=settings)

    def test_unique_profile_name_does_not_overwrite_existing(self):
        self.assertEqual(ui_fields.unique_profile_name('Perfil',{'Perfil','Perfil 2'}),'Perfil 3')

    def test_summary_includes_overrides_and_output_settings(self):
        settings = FormatSettings(bold=True, font_color='ABCDEF', header_distance_cm=.9, output_suffix='_novo')
        summary = ui_fields.settings_summary(settings)
        for value in ('Negrito: Ativar', '#ABCDEF', '0,9', '_novo'):
            self.assertIn(value, summary)

    def test_form_does_not_round_valid_precision_silently(self):
        settings = FormatSettings(line_spacing=1.23456789)
        self.assertEqual(ui_fields.settings_from_values(ui_fields.form_values(settings)), settings)


if __name__ == '__main__':
    unittest.main()
