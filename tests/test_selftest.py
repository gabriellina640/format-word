import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from app.selftest import run_self_test


class SelfTestTests(unittest.TestCase):
    def test_reports_success_and_cleans_temporary_documents(self):
        with tempfile.TemporaryDirectory() as tmp:
            report = Path(tmp)/'report.json'
            self.assertEqual(run_self_test(report,gui=False),0)
            data = json.loads(report.read_text(encoding='utf-8'))
            self.assertTrue(data['ok'])
            self.assertTrue(any('Importação' in check for check in data['checks']))
            self.assertEqual(list(Path(tmp).iterdir()),[report])

    def test_reports_failure_without_false_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            report=Path(tmp)/'report.json'
            with patch('app.batch.format_documents_batch',side_effect=RuntimeError('smoke failure')):
                self.assertEqual(run_self_test(report,gui=False),1)
            data=json.loads(report.read_text(encoding='utf-8'))
            self.assertFalse(data['ok'])
            self.assertIn('smoke failure',data['error'])
