"""Testes dos textos de sucesso da interface (sem instanciar a janela Tk).

Cobrem apenas os ``@staticmethod`` puros de ``app`` que recebem o dict de
resultado dos pipelines e devolvem a string de feedback. Não importam nem
instanciam ``ConversorPlanilhasApp`` e não dependem de widgets/messagebox.
"""

import unittest
from pathlib import Path

import app


class TestTextoSucessoPadrao(unittest.TestCase):
    def test_contem_mensagem_metricas_e_arquivo(self) -> None:
        resultado = {
            "registros_lidos": 120,
            "removidos_sem_telefone": 5,
            "duplicados_removidos": 3,
            "registros_finais": 112,
            "arquivos_gerados": [Path(r"C:\saida\Contatos_Tratados.xlsx")],
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_padrao(resultado, None)
        self.assertIn("Processamento concluído com sucesso!", texto)
        self.assertIn("112", texto)
        self.assertIn("Contatos_Tratados.xlsx", texto)
        self.assertIn("C:\\saida", texto)


class TestTextoSucessoBlipPreparacao(unittest.TestCase):
    def test_contem_contatos_csvs_invalidos_e_excel(self) -> None:
        resumo = {
            "contatos_exportados": 250,
            "lotes": 3,
            "total_invalidos": 2,
            "arquivos_gerados": [
                Path(r"C:\blip\Blip_001.csv"),
                Path(r"C:\blip\Blip_002.csv"),
                Path(r"C:\blip\Blip_003.csv"),
            ],
            "excel_revisao": Path(r"C:\blip\Blip_Contatos_Revisao.xlsx"),
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_preparacao(resumo)
        self.assertIn("Preparação para o Blip concluída com sucesso!", texto)
        self.assertIn("250", texto)
        self.assertIn("3", texto)
        self.assertIn("2", texto)
        self.assertIn("Blip_Contatos_Revisao.xlsx", texto)


class TestTextoSucessoBlipRevisao(unittest.TestCase):
    def test_mensagem_diferente_da_preparacao(self) -> None:
        resumo = {
            "contatos_exportados": 150,
            "lotes": 2,
            "total_invalidos": 0,
            "arquivos_gerados": [
                Path(r"C:\final\Blip_001.csv"),
                Path(r"C:\final\Blip_002.csv"),
            ],
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao(resumo)
        self.assertIn("CSVs finais gerados com sucesso!", texto)
        self.assertNotIn("Preparação para o Blip concluída com sucesso!", texto)
        self.assertIn("150", texto)
        self.assertIn("C:\\final", texto)

    def test_arquivos_gerados_vazio_nao_lanca(self) -> None:
        resumo = {
            "contatos_exportados": 0,
            "lotes": 0,
            "total_invalidos": 4,
            "arquivos_gerados": [],
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao(resumo)
        self.assertIn("CSVs finais gerados com sucesso!", texto)
        self.assertIn("—", texto)


class TestTextoSucessoBlipPreparacaoDDD(unittest.TestCase):
    def test_contem_total_grupos_csvs_sem_ddd_e_pasta(self) -> None:
        resumo = {
            "total_contatos": 300,
            "contatos_exportados": 300,
            "lotes": 5,
            "total_sem_ddd": 12,
            "pasta_raiz": Path(r"C:\saida\Blip"),
            "totais_por_grupo": {
                "DDD_47": 100,
                "DDD_49": 80,
                "DDD_48_E_OUTROS": 108,
                "SEM_DDD": 12,
            },
            "por_grupo": {
                "DDD_47": {"lotes": 1},
                "DDD_49": {"lotes": 1},
                "DDD_48_E_OUTROS": {"lotes": 2},
                "SEM_DDD": {"lotes": 1},
            },
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_preparacao_ddd(resumo)
        self.assertIn("Preparação para o Blip concluída com sucesso!", texto)
        self.assertIn("300", texto)  # total de contatos
        self.assertIn("DDD_47", texto)
        self.assertIn("DDD_49", texto)
        self.assertIn("DDD_48_E_OUTROS", texto)
        self.assertIn("SEM_DDD", texto)
        self.assertIn("CSV(s)", texto)  # CSVs por grupo
        self.assertIn("12", texto)  # SEM_DDD
        self.assertIn("C:\\saida\\Blip", texto)  # pasta raiz Blip

    def test_sem_pasta_nao_lanca(self) -> None:
        resumo = {
            "total_contatos": 0,
            "contatos_exportados": 0,
            "lotes": 0,
            "total_sem_ddd": 0,
            "totais_por_grupo": {},
            "por_grupo": {},
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_preparacao_ddd(resumo)
        self.assertIn("Preparação para o Blip concluída com sucesso!", texto)
        self.assertIn("—", texto)


class TestTextoSucessoBlipRevisaoDDD(unittest.TestCase):
    def test_contem_total_mudaram_grupos_sem_ddd_e_pasta(self) -> None:
        resumo = {
            "contatos_exportados": 150,
            "mudaram_de_grupo": 7,
            "lotes": 3,
            "total_sem_ddd": 4,
            "pasta_raiz": Path(r"C:\final\Blip"),
            "totais_por_grupo": {
                "DDD_47": 60,
                "DDD_49": 50,
                "DDD_48_E_OUTROS": 36,
                "SEM_DDD": 4,
            },
            "por_grupo": {
                "DDD_47": {"lotes": 1},
                "DDD_49": {"lotes": 1},
                "DDD_48_E_OUTROS": {"lotes": 1},
            },
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao_ddd(resumo)
        self.assertIn("CSVs finais gerados com sucesso!", texto)
        self.assertNotIn("Preparação para o Blip concluída com sucesso!", texto)
        self.assertIn("150", texto)  # total exportado
        self.assertIn("mudaram de grupo", texto.lower())
        self.assertIn("7", texto)  # mudaram de grupo
        self.assertIn("DDD_47", texto)
        self.assertIn("DDD_48_E_OUTROS", texto)
        self.assertIn("SEM_DDD", texto)
        self.assertIn("CSV(s)", texto)
        self.assertIn("4", texto)  # SEM_DDD
        self.assertIn("C:\\final\\Blip", texto)

    def test_sem_mudaram_de_grupo_usa_zero(self) -> None:
        resumo = {
            "contatos_exportados": 0,
            "lotes": 0,
            "total_sem_ddd": 0,
            "totais_por_grupo": {},
            "por_grupo": {},
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao_ddd(resumo)
        self.assertIn("CSVs finais gerados com sucesso!", texto)
        self.assertIn("Contatos que mudaram de grupo: 0", texto)
        self.assertIn("—", texto)


if __name__ == "__main__":
    unittest.main()
