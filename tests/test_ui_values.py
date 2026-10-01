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
        for text in ('Por tipo de texto', 'Título', 'Personalizar', '8,5', 'Centralizado', 'Recusar'):
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

    def test_word_labels_keep_existing_values_and_groups(self):
        self.assertEqual(ui_fields.FIELD_MAP['formatting_mode'].choices,
                         {'Todo o texto': 'uniform', 'Por tipo de texto': 'by_category'})
        self.assertEqual(ui_fields.FIELD_MAP['line_spacing_mode'].choices,
                         {'Múltiplo': 'multiple', 'Exatamente': 'exact', 'Pelo menos': 'at_least'})
        expected = {'font_color': 'Cor da fonte', 'line_spacing_mode': 'Espaçamento entre linhas',
                    'line_spacing': 'Em', 'paragraph_spacing_before': 'Espaçamento antes (pt)',
                    'paragraph_spacing_after': 'Espaçamento depois (pt)',
                    'first_line_indent_cm': 'Recuo especial (cm)',
                    'left_indent_cm': 'Recuo à esquerda (cm)', 'right_indent_cm': 'Recuo à direita (cm)',
                    'keep_with_next': 'Manter com o próximo', 'widow_control': 'Controle de viúvas/órfãs',
                    'paper_size': 'Tamanho do papel'}
        for name, label in expected.items():
            self.assertEqual(ui_fields.FIELD_MAP[name].label, label)
        self.assertEqual(ui_fields.FIELD_MAP['first_line_indent_cm'].group, 'Parágrafo')
        self.assertEqual(ui_fields.TRISTATE, {'Preservar': None, 'Ativar': True, 'Desativar': False})
        for name in ('alignment', 'header_alignment', 'footer_alignment', 'body_image_alignment'):
            self.assertEqual(ui_fields.FIELD_MAP[name].choices['Centralizado'], 'center')

    def test_every_choice_roundtrips_without_changing_serialized_values(self):
        for field in ui_fields.FIELDS:
            for label, value in (field.choices or {}).items():
                with self.subTest(field=field.name, choice=label):
                    settings = FormatSettings()
                    setattr(settings, field.name, value)
                    displayed = ui_fields.form_values(settings)
                    self.assertEqual(displayed[field.name], label)
                    self.assertEqual(asdict(ui_fields.settings_from_values(displayed)), asdict(settings))

    def test_summary_uses_word_line_spacing_and_special_indent(self):
        for mode, value, expected in (
                ('multiple', 1, 'Simples'), ('multiple', 1.5, '1,5 linhas'),
                ('multiple', 2, 'Duplo'), ('multiple', 1.25, 'Múltiplo 1,25'),
                ('exact', 18, 'Exatamente 18 pt'), ('at_least', 14.5, 'Pelo menos 14,5 pt')):
            with self.subTest(mode=mode, value=value):
                summary = ui_fields.settings_summary(FormatSettings(line_spacing_mode=mode, line_spacing=value))
                self.assertIn('Espaçamento entre linhas: ' + expected, summary)
        for indent, expected in ((0, 'Nenhum'), (1.25, 'Primeira linha 1,25 cm'), (-1.5, 'Deslocado 1,5 cm')):
            with self.subTest(indent=indent):
                summary = ui_fields.settings_summary(FormatSettings(first_line_indent_cm=indent))
                self.assertIn('Recuo especial: ' + expected, summary)
                self.assertNotIn('-1,5', summary)

    def test_custom_type_summary_explains_hanging_indent_and_line_spacing(self):
        settings = FormatSettings(formatting_mode='by_category', category_rules={
            'quote': {'mode': 'custom', 'values': {'first_line_indent_cm': -1.5,
                'line_spacing_mode': 'at_least', 'line_spacing': 15}}})
        summary = ui_fields.settings_summary(settings)
        self.assertIn('Recuo especial: Deslocado 1,5 cm', summary)
        self.assertIn('Espaçamento entre linhas: Pelo menos 15 pt', summary)
        self.assertNotIn('-1,5', summary)


if __name__ == '__main__':
    unittest.main()
