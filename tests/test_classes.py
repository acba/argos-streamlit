import unittest
import pandas as pd
import numpy as np
from classes import (
    FonteInformacao,
    Achado,
    AcaoVerificacao,
    ProcedimentoAuditoria,
    Auditado
)

class TestClasses(unittest.TestCase):

    def test_fonte_informacao_serialization(self):
        fi = FonteInformacao("Fonte Teste", "path/to/file.xlsx", "Sigla")
        data = fi.to_dict()
        fi2 = FonteInformacao.from_dict(data)
        
        self.assertEqual(fi.descricao, fi2.descricao)
        self.assertEqual(fi.filepath, fi2.filepath)
        self.assertEqual(fi.id, fi2.id)

    def test_achado_serialization(self):
        achado = Achado("1", "Achado Teste", ["Situação 1"], ["Evidência 1"], [{"tipo": "Rec", "encaminhamento": "Fazer X"}])
        data = achado.to_dict()
        achado2 = Achado.from_dict(data)
        
        self.assertEqual(achado.nome, achado2.nome)
        self.assertEqual(achado.situacoes_encontradas, achado2.situacoes_encontradas)
        # Check for list/dict equality
        self.assertEqual(len(achado.encaminhamentos), len(achado2.encaminhamentos))

    def test_acao_verificacao_serialization(self):
        fi = FonteInformacao("Fonte", "file.xlsx")
        av = AcaoVerificacao(
            fonte_informacao=fi,
            informacao_requerida="Campo",
            descricao_evidencia="Evidencia base",
            situacao_inconforme="> 0",
            tipo_encaminhamento="Rec",
            encaminhamento="Ajustar",
            pre_encaminhamento=None,
            criterio="Crit",
            descricao_situacao_inconforme="Desc"
        )
        av.resultado = True
        av.situacao_encontrada = 10
        
        data = av.to_dict()
        av2 = AcaoVerificacao.from_dict(data)
        
        self.assertEqual(av.id, av2.id)
        self.assertEqual(av.situacao_inconforme, av2.situacao_inconforme)
        self.assertEqual(av.fonte_informacao.descricao, av2.fonte_informacao.descricao)
        self.assertEqual(av.situacao_encontrada, av2.situacao_encontrada)

    def test_procedimento_auditoria_executar(self):
        # Setup
        fi = FonteInformacao("Fonte", "dummy.xlsx")
        # Mock dataframe in FonteInformacao
        fi.info = pd.DataFrame({
            'Campo1': [10],
            'Campo2': ['Sim']
        }, index=['ORG'])
        
        av1 = AcaoVerificacao(fi, 'Campo1', 'Evid1', '> 5', 'Rec', 'Enc1', None, 'Crit', 'Desc1', id='AV01')
        av2 = AcaoVerificacao(fi, 'Campo2', 'Evid2', "== 'Sim'", 'Rec', 'Enc2', None, 'Crit', 'Desc2', id='AV02')
        
        # Procedimento: (AV01 AND AV02)
        pa = ProcedimentoAuditoria("Proc 1", "AV01 & AV02", "1.1", "Achado Teste")
        pa.adicionar_acao(av1)
        pa.adicionar_acao(av2)
        
        # Execute
        auditado = Auditado("Orgao Teste", "ORG") # ID 'ORG' matches index in fi.info
        pa.executar('ORG') # Pass the sigla/index key
        
        self.assertTrue(pa.executado)
        self.assertTrue(pa.achado_ocorreu)
        self.assertIsNotNone(pa.achado)
        self.assertEqual(pa.achado.nome, "Achado Teste")
        
        # Check that DataFrame reference was removed (optimization)
        self.assertIsNone(fi.info)

    def test_auditado_serialization(self):
        auditado = Auditado("Nome", "SIG")
        auditado.tem_achados = True
        auditado.foi_auditado = True
        
        # Create a dummy procedure result to attach
        pa = ProcedimentoAuditoria("Desc", "True", "1", "Achado")
        pa.executado = True
        pa.achado_ocorreu = True
        pa.achado = Achado("1", "Achado")
        
        auditado.procedimentos_executados.append(pa)
        
        data = auditado.to_dict()
        auditado2 = Auditado.from_dict(data)
        
        self.assertEqual(auditado.sigla, auditado2.sigla)
        self.assertEqual(len(auditado2.procedimentos_executados), 1)
        self.assertEqual(auditado2.procedimentos_executados[0].achado.nome, "Achado")

if __name__ == '__main__':
    unittest.main()
