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


class TestTextoSucessoBlipPreparacaoCursoDDD(unittest.TestCase):
    def test_contem_total_cursos_ddd_nao_identificado_e_pasta(self) -> None:
        resumo = {
            "total_contatos": 300,
            "contatos_exportados": 300,
            "total_cursos_identificados": 2,
            "total_curso_nao_identificado": 20,
            "total_csvs": 6,
            "lotes": 6,
            "pasta_raiz": Path(r"C:\saida\Blip"),
            "contatos_por_curso": {
                "Análise e Desenvolvimento de Sistemas": 180,
                "Logística": 100,
                "CURSO_NAO_IDENTIFICADO": 20,
            },
            "contatos_por_ddd": {
                "DDD_47": 120,
                "DDD_49": 90,
                "DDD_48_E_OUTROS": 80,
                "SEM_DDD": 10,
            },
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_preparacao_curso_ddd(resumo)
        self.assertIn(
            "Preparação para o Blip (curso e DDD) concluída com sucesso!", texto
        )
        self.assertIn("300", texto)  # total de contatos
        self.assertIn("Análise e Desenvolvimento de Sistemas", texto)
        self.assertIn("Logística", texto)
        self.assertIn("CURSO_NAO_IDENTIFICADO", texto)
        self.assertIn("20", texto)  # curso não identificado
        self.assertIn("DDD_47", texto)
        self.assertIn("DDD_49", texto)
        self.assertIn("DDD_48_E_OUTROS", texto)
        self.assertIn("SEM_DDD", texto)
        self.assertIn("contato(s)", texto)
        self.assertIn("C:\\saida\\Blip", texto)  # pasta raiz Blip

    def test_sem_pasta_nao_lanca(self) -> None:
        resumo = {
            "total_contatos": 0,
            "contatos_exportados": 0,
            "total_cursos_identificados": 0,
            "total_curso_nao_identificado": 0,
            "total_csvs": 0,
            "contatos_por_curso": {},
            "contatos_por_ddd": {},
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_preparacao_curso_ddd(resumo)
        self.assertIn(
            "Preparação para o Blip (curso e DDD) concluída com sucesso!", texto
        )
        self.assertIn("—", texto)


class TestTextoSucessoBlipRevisaoCursoDDD(unittest.TestCase):
    def test_contem_mudancas_curso_ddd_e_pasta(self) -> None:
        resumo = {
            "contatos_exportados": 150,
            "mudaram_de_curso": 9,
            "mudaram_de_ddd": 5,
            "foram_para_curso_nao_identificado": 3,
            "foram_para_sem_ddd": 2,
            "total_csvs": 4,
            "lotes": 4,
            "pasta_raiz": Path(r"C:\final\Blip"),
            "contatos_por_curso": {
                "Análise e Desenvolvimento de Sistemas": 150,
            },
            "contatos_por_ddd": {
                "DDD_47": 60,
                "DDD_49": 40,
                "DDD_48_E_OUTROS": 48,
                "SEM_DDD": 2,
            },
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao_curso_ddd(resumo)
        self.assertIn("CSVs finais por curso e DDD gerados com sucesso!", texto)
        self.assertNotIn(
            "Preparação para o Blip (curso e DDD) concluída com sucesso!", texto
        )
        self.assertIn("150", texto)  # total exportado
        self.assertIn("Mudaram de curso: 9", texto)
        self.assertIn("Mudaram de DDD: 5", texto)
        self.assertIn("Foram p/ CURSO_NAO_IDENTIFICADO: 3", texto)
        self.assertIn("Foram p/ SEM_DDD: 2", texto)
        self.assertIn("Análise e Desenvolvimento de Sistemas", texto)
        self.assertIn("DDD_47", texto)
        self.assertIn("C:\\final\\Blip", texto)

    def test_sem_metricas_usa_zero(self) -> None:
        resumo = {
            "contatos_exportados": 0,
            "total_csvs": 0,
            "contatos_por_curso": {},
            "contatos_por_ddd": {},
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_blip_revisao_curso_ddd(resumo)
        self.assertIn("CSVs finais por curso e DDD gerados com sucesso!", texto)
        self.assertIn("Mudaram de curso: 0", texto)
        self.assertIn("Mudaram de DDD: 0", texto)
        self.assertIn("Foram p/ CURSO_NAO_IDENTIFICADO: 0", texto)
        self.assertIn("Foram p/ SEM_DDD: 0", texto)
        self.assertIn("—", texto)


class TestTextoSucessoUnificar(unittest.TestCase):
    def test_contem_total_lido_unico_grupos_sem_ddd_invalidos_e_pasta(self) -> None:
        resumo = {
            "total_lido": 500,
            "total_unico": 420,
            "por_grupo": {
                "DDD_47": 150,
                "DDD_49": 120,
                "DDD_48_E_OUTROS": 140,
                "SEM_DDD": 10,
            },
            "total_sem_ddd": 10,
            "total_invalidos": 8,
            "invalidos": [
                {
                    "valor_original": "sem numeros",
                    "arquivo": "a.csv",
                    "linha": 3,
                    "motivo": "Telefone inválido ou não classificável.",
                }
            ],
            "relatorio_invalidos": Path(
                r"C:\saida\LISTA_UNIFICADA\_telefones_invalidos.csv"
            ),
            "pasta_raiz": Path(r"C:\saida\LISTA_UNIFICADA"),
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_unificar(resumo)
        self.assertIn("Lista unificada gerada com sucesso!", texto)
        self.assertIn("500", texto)  # total lido
        self.assertIn("420", texto)  # total único
        self.assertIn("DDD_47", texto)
        self.assertIn("DDD_49", texto)
        self.assertIn("DDD_48_E_OUTROS", texto)
        self.assertIn("SEM_DDD", texto)
        self.assertIn("8", texto)  # inválidos
        self.assertIn("relatório", texto.lower())
        # O caminho REAL do relatório é apontado, não um "relatório no resumo".
        self.assertIn("_telefones_invalidos.csv", texto)
        # Os registros identificados aparecem para conferência direta.
        self.assertIn("sem numeros", texto)
        self.assertIn("a.csv", texto)
        self.assertIn("C:\\saida\\LISTA_UNIFICADA", texto)

    def test_sem_pasta_nao_lanca(self) -> None:
        resumo = {
            "total_lido": 0,
            "total_unico": 0,
            "por_grupo": {},
            "total_sem_ddd": 0,
            "total_invalidos": 0,
        }
        texto = app.ConversorPlanilhasApp._texto_sucesso_unificar(resumo)
        self.assertIn("Lista unificada gerada com sucesso!", texto)
        self.assertIn("—", texto)


class TestMensagemAmigavelTelefoneColunaError(unittest.TestCase):
    def test_retorna_a_propria_mensagem(self) -> None:
        import unificar

        mensagem = (
            "Não foi possível identificar a coluna de telefone nos arquivos a seguir:\n"
            "- planilha.xlsx: nenhuma coluna de telefone reconhecível"
        )
        erro = unificar.TelefoneColunaError(mensagem)
        self.assertEqual(
            app.ConversorPlanilhasApp._mensagem_amigavel(erro), mensagem
        )


if __name__ == "__main__":
    unittest.main()
