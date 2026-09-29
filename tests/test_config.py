from __future__ import annotations

import json
import random
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from app import config as module
from app.config import AppConfig, ConfigStore, FormatSettings


class ValidationTest(unittest.TestCase):
    def test_inherit_follows_effective_body_rule(self):
        settings = FormatSettings(formatting_mode='by_category', category_rules={
            'body': {'mode': 'custom', 'values': {'font_size': 18}},
            'quote': {'mode': 'inherit', 'values': {}}})
        self.assertEqual(module.settings_for_category(settings, 'quote').font_size, 18)
        settings.category_rules['body'] = {'mode': 'preserve', 'values': {}}
        self.assertIsNone(module.settings_for_category(settings, 'quote'))
        settings.category_rules['body'] = {'mode': 'inherit', 'values': {}}
        self.assertEqual(module.settings_for_category(settings, 'body').font_size, 12)
        self.assertEqual(module.settings_for_category(settings, 'quote').font_size, 12)

    def test_category_defaults_and_resolution(self):
        self.assertEqual(getattr(FormatSettings(), 'formatting_mode', None), 'uniform')
        self.assertEqual(set(module.CATEGORY_LABELS), {'body', 'title', 'subtitle', 'quote', 'signature', 'table', 'textbox'})
        settings = FormatSettings(formatting_mode='by_category', category_rules={
            'title': {'mode': 'custom', 'values': {'font_size': 18, 'font_color': '#ab12cd'}},
            'quote': {'mode': 'inherit', 'values': {}}})
        self.assertEqual(module.settings_for_category(settings, 'body').font_size, 12)
        self.assertIsNone(module.settings_for_category(settings, 'table'))
        self.assertEqual(module.settings_for_category(settings, 'quote').font_size, 12)
        self.assertEqual(module.settings_for_category(settings, 'title').font_color, 'AB12CD')
        self.assertEqual(module.settings_for_category(replace(settings, formatting_mode='uniform'), 'title').font_size, 12)
        with self.assertRaises(ValueError):
            module.settings_for_category(settings, 'unknown')

    def test_category_validation_deeply_isolates_snapshots(self):
        source = FormatSettings(category_rules={'title': {'mode': 'custom', 'values': {
            'font_name': ' Custom ', 'font_color': '#abcdef'}}}, style_categories={'Heading1': 'title'})
        result = self.validate(source)
        self.assertEqual(result.category_rules['title']['values']['font_color'], 'ABCDEF')
        self.assertEqual(result.category_rules['title']['values']['font_name'], 'Custom')
        result.category_rules['title']['values']['font_name'] = 'Changed'
        result.style_categories['Normal'] = 'body'
        self.assertEqual(source.category_rules['title']['values']['font_name'], ' Custom ')
        self.assertEqual(source.style_categories, {'Heading1': 'title'})
        effective = module.settings_for_category(source, 'body')
        effective.category_rules['title']['values']['font_name'] = 'Other'
        self.assertEqual(source.category_rules['title']['values']['font_name'], ' Custom ')

    def test_nested_category_rules_reject_invalid_structure_and_geometry(self):
        rules = [[], {'unknown': {}}, {'title': []}, {'title': {'mode': 'oops'}},
                 {'title': {'mode': 'custom', 'unknown': True}},
                 {'title': {'mode': 'custom', 'values': []}},
                 {'title': {'mode': 'custom', 'values': {'margin_left_cm': 1}}},
                 {'title': {'mode': 'custom', 'values': {'font_size': True}}},
                 {'title': {'mode': 'custom', 'values': {'font_color': 'red'}}},
                 {'title': {'mode': 'custom', 'values': {'left_indent_cm': 16}}}]
        for value in rules:
            with self.subTest(value=value), self.assertRaises(module.SettingsError):
                self.validate(FormatSettings(category_rules=value))
        for value in ([], {'': 'body'}, {'Normal': 'unknown'}, {1: 'title'}, {'Normal': []}):
            with self.subTest(value=value), self.assertRaises(module.SettingsError):
                self.validate(FormatSettings(style_categories=value))
        with self.assertRaises(module.SettingsError):
            self.validate(FormatSettings(margin_left_cm=10, category_rules={
                'quote': {'mode': 'custom', 'values': {'left_indent_cm': 9}}}))

    def test_body_images_and_complex_content_validate(self):
        self.validate(FormatSettings(body_image_width_cm=16, body_image_alignment='center'))
        for overrides in ({'body_image_width_cm': 17}, {'body_image_width_cm': True},
                          {'body_image_width_cm': float('nan')}, {'body_image_width_cm': 0},
                          {'body_image_alignment': 'justify'}, {'formatting_mode': 'bad'},
                          {'complex_content_mode': 'bad'}):
            with self.subTest(overrides=overrides), self.assertRaises(module.SettingsError):
                self.validate(replace(FormatSettings(), **overrides))

    def validate(self, settings, **kwargs):
        self.assertTrue(callable(getattr(module, 'validate_settings', None)), 'Shared validation must exist')
        return module.validate_settings(settings, **kwargs)

    def test_normalizes_a_copy_and_preserves_custom_font(self):
        settings = FormatSettings(font_name=' My Custom Font ', font_size=12.5)
        result = self.validate(settings)
        self.assertIsNot(result, settings)
        self.assertEqual(result.font_name, 'My Custom Font')
        self.assertEqual(settings.font_name, ' My Custom Font ')
        self.assertEqual(result.font_size, 12.5)

    def test_rejects_invalid_numbers_and_reports_field(self):
        for value in (True, '12', float('nan'), float('inf'), 0, 400.5, 12.2):
            with self.subTest(value=value):
                try:
                    self.validate(FormatSettings(font_size=value))
                except ValueError as exc:
                    self.assertEqual(exc.field, 'font_size')
                else:
                    self.fail('Invalid font size accepted')

    def test_defaults_include_new_fields(self):
        settings = FormatSettings()
        expected = {'alignment': 'justify', 'line_spacing_mode': 'multiple',
                    'paper_size': 'A4', 'orientation': 'portrait',
                    'header_mode': 'preserve', 'footer_mode': 'preserve',
                    'bold': None, 'font_color': '', 'paragraph_spacing_before': 0}
        for key, value in expected.items():
            self.assertEqual(getattr(settings, key, 'MISSING'), value)

    def test_enums_tristates_suffix_and_text_are_strict(self):
        self.validate(FormatSettings())
        cases = {'alignment': ['invalid', []], 'line_spacing_mode': ['x'],
                 'orientation': ['x'], 'paper_size': ['x'], 'header_mode': ['x'],
                 'header_alignment': ['justify'], 'bold': [1, 'true'],
                 'font_color': ['red', '#12345g'], 'font_name': ['', 'x\ny', 12],
                 'output_suffix': ['', '../x', 'with space', 'á'],
                 'max_input_mb': [True, 0, 1.5], 'header_image_path': [None]}
        for field, values in cases.items():
            for value in values:
                with self.subTest(field=field, value=value):
                    with self.assertRaises(ValueError) as caught:
                        self.validate(replace(FormatSettings(), **{field: value}))
                    self.assertEqual(caught.exception.field, field)

    def test_color_line_spacing_and_geometry(self):
        self.validate(FormatSettings())
        valid = replace(FormatSettings(), font_color='#ab12cd', first_line_indent_cm=-1,
                        left_indent_cm=1, line_spacing_mode='exact', line_spacing=12.5)
        self.assertEqual(self.validate(valid).font_color, 'AB12CD')
        for overrides in ({'line_spacing': .4}, {'line_spacing_mode': 'exact', 'line_spacing': .5},
                          {'paragraph_spacing_before': -1}, {'margin_left_cm': 20},
                          {'left_indent_cm': 15, 'right_indent_cm': 5},
                          {'header_distance_cm': 50}, {'page_width_cm': 0}):
            with self.subTest(overrides=overrides), self.assertRaises(ValueError):
                self.validate(replace(FormatSettings(), **overrides))
        self.assertEqual(module.paper_dimensions(replace(valid, orientation='landscape')), (29.7, 21))
        self.assertEqual(module.paper_dimensions(replace(valid, paper_size='Letter')), (21.59, 27.94))

    def test_template_blocks_execution(self):
        with self.assertRaisesRegex(ValueError, 'template|Template'):
            self.validate(FormatSettings(template_path='/missing/template.docx'))

    def test_assets_validated_only_when_requested(self):
        self.validate(FormatSettings())
        settings = replace(FormatSettings(), header_mode='image', header_image_path='/missing.png')
        self.validate(settings)
        with self.assertRaises(ValueError) as caught:
            self.validate(settings, check_assets=True)
        self.assertEqual(caught.exception.field, 'header_image_path')

    def test_image_geometry_reserves_paragraph_clearance(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'banner.png'
            Image.new('RGB', (1000, 100)).save(path)
            settings = replace(FormatSettings(), header_mode='image', header_image_path=str(path),
                               header_image_width_cm=10, header_distance_cm=1,
                               margin_top_cm=2.05)
            with self.assertRaises(ValueError) as caught:
                self.validate(settings, check_assets=True)
            self.assertEqual(caught.exception.field, 'header_image_width_cm')
            self.assertIn('Cabeçalho', str(caught.exception))
            self.validate(replace(settings, margin_top_cm=2.2), check_assets=True)

    def test_geometry_error_identifies_header_or_footer(self):
        for slot, label in (('header', 'Cabeçalho'), ('footer', 'Rodapé')):
            with self.subTest(slot=slot):
                with self.assertRaisesRegex(ValueError, label):
                    self.validate(replace(FormatSettings(), **{f'{slot}_distance_cm': 10}))


class ConfigStoreTest(unittest.TestCase):
    def test_category_profile_roundtrip_and_old_defaults(self):
        settings = FormatSettings(formatting_mode='by_category', category_rules={
            'signature': {'mode': 'custom', 'values': {'font_size': 10}}},
            style_categories={'Assinatura': 'signature'}, body_image_width_cm=8)
        self.store.save(AppConfig(settings, {'Promotor': settings}, 'Promotor'))
        loaded = self.store.load()
        self.assertEqual(loaded.settings, settings)
        loaded.settings.category_rules['signature']['values']['font_size'] = 20
        self.assertEqual(loaded.stacks['Promotor'].category_rules['signature']['values']['font_size'], 10)
        self.write({'settings': {'font_size': 11}})
        old = self.store.load().settings
        self.assertEqual(old.formatting_mode, 'uniform')
        self.assertEqual(old.complex_content_mode, 'preserve')
        self.assertEqual(old.category_rules, {})

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.store = ConfigStore(self.root / 'config')

    def write(self, payload):
        self.store.config_path.write_text(json.dumps(payload), encoding='utf-8')

    def test_roundtrip_custom_font_decimals_and_profiles(self):
        settings = FormatSettings(font_name='Fonte Personalizada', font_size=12.5, paragraph_spacing_after=6.5)
        self.store.save(AppConfig(settings, {'Meu perfil': settings}, 'Meu perfil'))
        loaded = self.store.load()
        self.assertEqual(loaded.settings, settings)
        self.assertEqual(loaded.stacks['Meu perfil'], settings)
        self.assertEqual(loaded.active_stack, 'Meu perfil')

    def test_corrupt_root_recovers_with_warning_and_backup(self):
        for raw in ('[]', 'null', '{invalid', '{"settings": []}'):
            with self.subTest(raw=raw):
                self.store.config_path.write_text(raw, encoding='utf-8')
                loaded = self.store.load()
                self.assertEqual(loaded.settings, FormatSettings())
                self.assertTrue(getattr(self.store, 'warnings', []))
                self.store.save(loaded)
                self.assertTrue(any(path.read_text(encoding='utf-8') == raw
                                    for path in self.store.config_dir.glob('settings*.bak')))

    def test_bad_profile_does_not_discard_good_profiles(self):
        self.write({'settings': {'font_size': 'bad'}, 'stacks': {
            'good': {'font_name': 'Custom'}, 'bad': {'line_spacing': False}}, 'active_stack': 'bad'})
        loaded = self.store.load()
        self.assertEqual(set(loaded.stacks), {'good'})
        self.assertEqual(loaded.settings, FormatSettings())
        self.assertEqual(loaded.active_stack, '')
        self.assertTrue(self.store.warnings)
        self.assertTrue(list(self.store.config_dir.glob('settings*.bak')))

    def test_rejects_invalid_settings_without_overwriting(self):
        self.store.save(AppConfig())
        original = self.store.config_path.read_bytes()
        for config in (AppConfig(settings=FormatSettings(font_size='x')),
                       AppConfig(stacks={'bad': FormatSettings(font_size=False)})):
            with self.assertRaises(ValueError):
                self.store.save(config)
        self.assertEqual(self.store.config_path.read_bytes(), original)

    def test_missing_asset_can_be_persisted_but_execution_requires_it(self):
        profile = replace(FormatSettings(), header_mode='image', header_image_path='/missing.png')
        self.store.save(AppConfig(stacks={'missing asset': profile}))
        self.assertEqual(self.store.load().stacks['missing asset'], profile)
        with self.assertRaises(ValueError):
            module.validate_settings(profile, check_assets=True)

    def test_valid_profile_edit_preserves_legacy_profile_for_later_review(self):
        legacy = replace(FormatSettings(), template_path='/saved/legacy.docx',
                         migration_warnings=('Revise o deslocamento antigo.',))
        self.store.save(AppConfig(stacks={'valid': FormatSettings(), 'legacy': legacy}))
        config = self.store.load()
        config.stacks['valid'].font_size = 14.5
        self.store.save(config)
        loaded = self.store.load()
        self.assertEqual(loaded.stacks['valid'].font_size, 14.5)
        self.assertEqual(loaded.stacks['legacy'], legacy)
        with self.assertRaises(ValueError):
            module.validate_settings(loaded.stacks['legacy'])

    def test_migrates_legacy_alignment_modes_retains_template_for_review(self):
        template = self.store.templates_dir / 'saved.docx'
        template.write_bytes(b'legacy document')
        self.write({'settings': {'justify_text': False, 'include_header': True,
                    'include_footer': False, 'template_path': str(template)}})
        loaded = self.store.load()
        self.assertEqual(getattr(loaded.settings, 'alignment', None), 'left')
        self.assertEqual(loaded.settings.header_mode, 'image')
        self.assertEqual(loaded.settings.footer_mode, 'preserve')
        self.assertEqual(loaded.settings.template_path, str(template))
        self.assertTrue(self.store.warnings)
        self.assertTrue(template.exists())
        self.store.save(loaded)
        self.assertEqual(self.store.load().settings.template_path, str(template))
        with self.assertRaises(ValueError):
            module.validate_settings(loaded.settings)

    def test_legacy_nonzero_offsets_require_review(self):
        self.write({'settings': {'header_offset_x_cm': 1.2}})
        loaded = self.store.load()
        self.assertTrue(getattr(self.store, 'warnings', []))
        with self.assertRaises(ValueError):
            module.validate_settings(loaded.settings)

    def test_failed_atomic_replace_retains_original_and_cleans_tempfile(self):
        self.store.save(AppConfig())
        original = self.store.config_path.read_bytes()
        before = set(self.store.config_dir.iterdir())
        with patch('pathlib.Path.replace', side_effect=OSError('disk failure')):
            with self.assertRaises(OSError):
                self.store.save(AppConfig(settings=FormatSettings(font_size=16)))
        self.assertEqual(self.store.config_path.read_bytes(), original)
        self.assertEqual(set(self.store.config_dir.iterdir()), before)

    def test_image_preserves_dimensions_and_content(self):
        source = self.root / 'logo.png'
        Image.new('RGBA', (160, 80), (20, 40, 60, 128)).save(source)
        first = self.store.store_image(source, 'header')
        second = self.store.store_image(source, 'header')
        self.assertNotEqual(first, second)
        with Image.open(first) as image:
            self.assertEqual(image.size, (160, 80))
            self.assertEqual(image.getpixel((0, 0)), (20, 40, 60, 128))

    def test_image_normalizes_exif_orientation(self):
        source = self.root / 'photo.jpg'
        image = Image.new('RGB', (80, 40), 'red')
        exif = image.getexif()
        exif[274] = 6
        image.save(source, exif=exif)
        with Image.open(self.store.store_image(source, 'footer')) as normalized:
            self.assertEqual(normalized.size, (40, 80))

    def test_imported_jpeg_expansion_remains_usable_without_losing_pixels(self):
        source = self.root / 'photo.jpg'
        noise = random.Random(2026).randbytes(100 * 100 * 3)
        Image.frombytes('RGB', (100, 100), noise).save(source, quality=90)
        import_limit = source.stat().st_size + 1
        with patch.object(module, 'MAX_IMAGE_BYTES', import_limit):
            stored = Path(self.store.store_image(source, 'header'))
            self.assertGreater(stored.stat().st_size, import_limit)
            settings = replace(FormatSettings(), header_mode='image',
                               header_image_path=str(stored), header_image_width_cm=1)
            module.validate_settings(settings, check_assets=True)
        with Image.open(source) as original, Image.open(stored) as normalized:
            self.assertEqual(original.convert('RGB').tobytes(), normalized.convert('RGB').tobytes())

    def test_reserved_legacy_profile_name_is_migrated_without_collision(self):
        self.store.config_path.write_text(json.dumps({
            'stacks': {'Configuração atual': {'font_size': 36},
                       'Configuração atual (perfil)': {'font_size': 18}},
            'active_stack': 'Configuração atual',
        }), encoding='utf-8')
        loaded = self.store.load()
        self.assertNotIn('Configuração atual', loaded.stacks)
        self.assertEqual(loaded.stacks[loaded.active_stack].font_size, 36)
        self.assertEqual(loaded.stacks['Configuração atual (perfil)'].font_size, 18)
        self.assertTrue(self.store.warnings)
        self.store.save(loaded)
        self.assertEqual(self.store.load(), loaded)

    def test_reserved_profile_name_cannot_be_saved_by_new_clients(self):
        with self.assertRaises(ValueError):
            self.store.save(AppConfig(stacks={'Configuração atual': FormatSettings()}))

    def test_image_rejects_unknown_slots_wrong_real_type_and_fake_bytes(self):
        source = self.root / 'fake.png'
        for data in (b'not png', b'\x89PNG\r\n\x1a\nbroken'):
            source.write_bytes(data)
            with self.assertRaises(ValueError):
                self.store.store_image(source, 'header')
        Image.new('RGB', (2, 2)).save(source, format='GIF')
        with self.assertRaises(ValueError):
            self.store.store_image(source, 'header')
        with self.assertRaises(ValueError):
            self.store.store_image(source, '../header')

    def test_image_byte_and_pixel_limits(self):
        source = self.root / 'large.png'
        source.write_bytes(b'\x00' * (8 * 1024 * 1024 + 1))
        with self.assertRaises(ValueError):
            self.store.store_image(source, 'header')
        Image.new('1', (6400, 6400)).save(source)
        with self.assertRaises(ValueError):
            self.store.store_image(source, 'header')


if __name__ == '__main__':
    unittest.main()
