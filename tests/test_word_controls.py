"""Desktop checks for familiar Word controls backed by existing profile values."""
import os
import tkinter as tk
import unittest


@unittest.skipUnless(os.environ.get('FORMATWORD_GUI_TESTS') == '1', 'Requires desktop Tk session')
class WordControlTests(unittest.TestCase):
    def setUp(self):
        from app.word_controls import SpecialIndentControl
        self.root = tk.Tk()
        self.root.withdraw()
        self.value = tk.StringVar(self.root, '1,25')
        self.control = SpecialIndentControl(self.root, self.value)

    def tearDown(self):
        self.root.destroy()

    def test_existing_values_show_familiar_special_indent(self):
        for raw, kind, magnitude in [('1,25', 'Primeira linha', '1,25'),
                                      ('-0,75', 'Deslocado', '0,75'), ('0', 'Nenhum', '0')]:
            self.value.set(raw)
            self.assertEqual(self.control.kind.get(), kind)
            self.assertEqual(self.control.amount.get(), magnitude)
            self.assertEqual(self.value.get(), raw)

    def test_user_selects_type_and_positive_amount(self):
        self.control.kind.set('Deslocado')
        self.control.selector.event_generate('<<ComboboxSelected>>')
        self.control.amount.set('0,7654321')
        self.assertEqual(float(self.value.get().replace(',', '.')), -.7654321)
        self.control.kind.set('Nenhum')
        self.control.selector.event_generate('<<ComboboxSelected>>')
        self.assertEqual(float(self.value.get()), 0)
        self.control.kind.set('Primeira linha')
        self.control.selector.event_generate('<<ComboboxSelected>>')
        self.control.amount.set('1,5')
        self.assertEqual(float(self.value.get().replace(',', '.')), 1.5)

    def test_invalid_entry_stays_invalid_in_canonical_value(self):
        for value in ('', 'abc', '-1'):
            self.control.amount.set(value)
            with self.assertRaises(ValueError):
                float(self.value.get())
            self.assertEqual(self.control.amount.get(), value)

    def test_external_invalid_value_stays_editable(self):
        self.value.set('0')
        self.value.set('abc')
        self.assertEqual(self.control.amount.get(), 'abc')
        self.assertEqual(str(self.control.entry.cget('state')), 'normal')
        self.control.amount.set('2')
        self.assertEqual(float(self.value.get()), 2)

    def test_disabled_and_destroy_remove_trace(self):
        self.control.configure(state='disabled')
        self.assertEqual(str(self.control.selector.cget('state')), 'disabled')
        self.assertEqual(str(self.control.entry.cget('state')), 'disabled')
        self.control.configure(state='normal')
        self.assertEqual(str(self.control.selector.cget('state')), 'readonly')
        self.control.destroy()
        self.assertFalse(self.value.trace_info())


if __name__ == '__main__':
    unittest.main()
