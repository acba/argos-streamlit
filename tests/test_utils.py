import unittest
import pandas as pd
import os
from unittest.mock import MagicMock, patch
from utils import (
    aplicar_variaveis_temporarias,
    carregar_dados,
    safe_compare,
    avalia_logica,
    avalia_expressao,
    processa_imagens_contexto,
    parse_expression,
    infix_to_rpn
)
from classes import FonteInformacao

class TestUtils(unittest.TestCase):

    def test_safe_compare(self):
        # Numeric comparisons
        self.assertTrue(safe_compare(10, "> 5"))
        self.assertTrue(safe_compare(5, ">= 5"))
        self.assertTrue(safe_compare(5, "== 5"))
        self.assertTrue(safe_compare(5, "< 10"))
        self.assertTrue(safe_compare(5, "<= 5"))
        self.assertTrue(safe_compare(5, "!= 6"))
        
        # String comparisons
        self.assertTrue(safe_compare("Sim", "== 'Sim'"))
        self.assertTrue(safe_compare("Sim", "Sim")) # Implicit equality
        self.assertTrue(safe_compare("Não", "!= 'Sim'"))
        
        # Mixed types (handling strings that look like numbers)
        self.assertTrue(safe_compare("10", "> 5"))
        self.assertTrue(safe_compare(10, "> '5'"))

        # Edge cases
        self.assertFalse(safe_compare(5, "> 10"))
        self.assertFalse(safe_compare("A", "== 'B'"))

    def test_avalia_logica(self):
        contexto = {'A': True, 'B': False, 'C': True}
        
        # Simple
        self.assertTrue(avalia_logica("A", contexto))
        self.assertFalse(avalia_logica("B", contexto))
        
        # OR
        self.assertTrue(avalia_logica("A | B", contexto))
        self.assertFalse(avalia_logica("B | B", contexto))
        
        # AND
        self.assertFalse(avalia_logica("A & B", contexto))
        self.assertTrue(avalia_logica("A & C", contexto))
        
        # NOT
        self.assertTrue(avalia_logica("~B", contexto))
        self.assertFalse(avalia_logica("~A", contexto))
        
        # Complex
        self.assertTrue(avalia_logica("(A & C) | B", contexto))
        self.assertFalse(avalia_logica("(A & B) | B", contexto))

    def test_avalia_expressao(self):
        # Test logic evaluation with safe_compare integration
        self.assertTrue(avalia_expressao("> 10", 15))
        self.assertFalse(avalia_expressao("> 10", 5))
        self.assertTrue(avalia_expressao("Sim", "Sim"))
        
        # Test compound expressions (not typical for this function but supported by logic)
        # Note: avalia_expressao splits by operators and compares 'situacao_encontrada' against each token
        # This logic in utils.py seems to assume the expression is a set of conditions ANDed/ORed together
        # applied to the SAME value.
        # E.g. val=10, expr="> 5 & < 15" -> (10 > 5) & (10 < 15) -> True
        self.assertTrue(avalia_expressao("> 5 & < 15", 10))
        self.assertFalse(avalia_expressao("> 5 & < 8", 10))

    def test_processa_imagens_contexto(self):
        contexto = {
            'img1': 'foto.png',
            'img2': 'missing.jpg',
            'text': 'texto normal'
        }
        files_map = {'foto.png': '/path/to/foto.png'}
        
        # Test Markdown mode
        new_ctx, warnings = processa_imagens_contexto(contexto.copy(), files_map, 'md')
        self.assertEqual(new_ctx['img1'], '/path/to/foto.png')
        self.assertIn("[Imagem 'missing.jpg' não encontrada]", new_ctx['img2'])
        self.assertEqual(len(warnings), 1)
        self.assertIn("Arquivo de imagem 'missing.jpg' para a variável 'img2' não encontrado", warnings[0])

    @patch('pandas.read_excel')
    def test_carregar_dados_success(self, mock_read_excel):
        mock_df = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        mock_read_excel.return_value = mock_df
        
        result = carregar_dados('dummy.xlsx')
        pd.testing.assert_frame_equal(result, mock_df)

    def test_carregar_dados_failure(self):
        # Test that it raises ValueError on file not found or other errors
        with self.assertRaises(ValueError) as cm:
            carregar_dados('non_existent_file.xlsx')
        self.assertIn("Erro ao carregar a planilha", str(cm.exception))

    def test_parser_helpers(self):
        # parse_expression
        tokens = parse_expression("A | (B & ~C)")
        self.assertEqual(tokens, ['A', '|', '(', 'B', '&', '~', 'C', ')'])
        
        # infix_to_rpn
        rpn = infix_to_rpn(tokens)
        # Expected RPN for A | (B & ~C) is A B C ~ & |
        self.assertEqual(rpn, ['A', 'B', 'C', '~', '&', '|'])

    def test_aplicar_variaveis_temporarias_em_ordem(self):
        fonte = FonteInformacao("Questionário", "dummy.xlsx", id="questionario")
        fonte.info = pd.DataFrame({
            'q0105[TI_efetivos]': [2, 0],
            'q0105[TI_terceirizados]': [1, 3],
        }, index=['A', 'B'])
        variaveis = pd.DataFrame([
            {
                'id': 'VT01',
                'id_fonte_informacao': 'questionario',
                'nome': 'total_TI_interno',
                'expressao': 'q0105[TI_efetivos]',
                'descricao': 'Total interno',
            },
            {
                'id': 'VT02',
                'id_fonte_informacao': 'questionario',
                'nome': 'predominio_terceiros',
                'expressao': 'q0105[TI_terceirizados] > total_TI_interno',
                'descricao': 'Predomínio de terceiros',
            },
        ])

        aplicar_variaveis_temporarias({'questionario': fonte}, variaveis)

        self.assertEqual(fonte.info['total_TI_interno'].tolist(), [2, 0])
        self.assertEqual(fonte.info['predominio_terceiros'].tolist(), [False, True])

if __name__ == '__main__':
    unittest.main()
