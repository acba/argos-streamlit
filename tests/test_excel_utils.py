import unittest
import pandas as pd
import io
from utils import detect_header, validar_schema, carregar_dados

class TestExcelUtils(unittest.TestCase):

    def setUp(self):
        # Create sample Excel files in memory
        self.valid_excel = io.BytesIO()
        df_valid = pd.DataFrame({
            'col1': [1, 2],
            'col2': ['a', 'b']
        })
        df_valid.to_excel(self.valid_excel, index=False)
        self.valid_excel.seek(0)

        self.offset_excel = io.BytesIO()
        # Simulate header on row 3 (index 2)
        with pd.ExcelWriter(self.offset_excel, engine='xlsxwriter') as writer:
            worksheet = writer.book.add_worksheet('Sheet1')
            worksheet.write(0, 0, 'Metadata 1')
            worksheet.write(1, 0, 'Metadata 2')
            # Write header
            worksheet.write(2, 0, 'col1')
            worksheet.write(2, 1, 'col2')
            # Write data
            worksheet.write(3, 0, 1)
            worksheet.write(3, 1, 'a')
        self.offset_excel.seek(0)

        self.missing_col_excel = io.BytesIO()
        df_missing = pd.DataFrame({
            'col1': [1, 2],
            'wrong_col': ['a', 'b']
        })
        df_missing.to_excel(self.missing_col_excel, index=False)
        self.missing_col_excel.seek(0)

    def test_detect_header_simple(self):
        idx = detect_header(self.valid_excel, 0, ['col1', 'col2'])
        self.assertEqual(idx, 0)

    def test_detect_header_offset(self):
        idx = detect_header(self.offset_excel, 0, ['col1', 'col2'])
        self.assertEqual(idx, 2)

    def test_detect_header_missing(self):
        # Should fail if columns are not found in sample
        with self.assertRaises(ValueError) as cm:
            detect_header(self.missing_col_excel, 0, ['col1', 'col2'])
        self.assertIn("Não foi possível detectar o cabeçalho", str(cm.exception))

    def test_validar_schema_success(self):
        df = pd.DataFrame({'Col1': [1], 'COL2': [2]})
        # Should be case-insensitive/stripped
        validar_schema(df, ['col1', 'col2'])

    def test_validar_schema_failure(self):
        df = pd.DataFrame({'Col1': [1]})
        with self.assertRaises(ValueError) as cm:
            validar_schema(df, ['col1', 'col2'])
        self.assertIn("Colunas obrigatórias ausentes", str(cm.exception))
        self.assertIn("col2", str(cm.exception))

    def test_carregar_dados_with_detection(self):
        # Test auto-detection via carregar_dados
        self.offset_excel.seek(0)
        df = carregar_dados(self.offset_excel, skiprows=None, required_columns=['col1', 'col2'])
        self.assertEqual(len(df), 1)
        self.assertIn('col1', df.columns)
        self.assertEqual(df.iloc[0]['col1'], 1)

    def test_carregar_dados_validation_failure(self):
        self.missing_col_excel.seek(0)
        with self.assertRaises(ValueError) as cm:
            carregar_dados(self.missing_col_excel, skiprows=None, required_columns=['col1', 'col2'])
        # The error usually comes from detect_header first if headers aren't found
        self.assertIn("Não foi possível detectar o cabeçalho", str(cm.exception))

if __name__ == '__main__':
    unittest.main()
