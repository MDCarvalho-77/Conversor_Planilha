"""Testes do pipeline V2: consolidação e deduplicação global."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from adapters import (
    CANON_CPF,
    CANON_NOME_COMPLETO,
    CANON_REGIONAL,
    CANON_TELEFONE,
)
from pipeline_v2 import (
    FORMATO_AMBAS_SAIDAS,
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_PLANILHA_ESTENDIDA,
    FORMATO_SOMENTE_NUMEROS,
    LIMITE_ARQUIVOS,
    LimiteArquivosExcedidoError,
    gerar_saidas_v2,
    processar_consolidado,
)


def _salvar(df: pd.DataFrame, caminho: Path) -> Path:
    df.to_excel(caminho, index=False, engine="openpyxl")
    return caminho


def _df_crm1(linhas: list[dict[str, object]]) -> pd.DataFrame:
    colunas = [
        "Razão do Status",
        "Nome Completo",
        "CPF_CNPJ",
        "E-mail",
        "Telefone da Empresa/Celular",
        "Nome do produto",
        "Primário (Nome do produto) (Produto)",
    ]
    return pd.DataFrame(linhas, columns=colunas)


def _df_crm2(linhas: list[dict[str, object]]) -> pd.DataFrame:
    colunas = [
        "Regional/Unidade do Produto",
        "Razão do Status",
        "Nome Completo",
        "Nome",
        "Segundo Nome",
        "Sobrenome",
        "CPF",
        "E-mail Comercial",
        "Telefone Celular",
        "Telefone da Empresa",
        "Telefone Residencial",
        "Nome do Produto",
        "Produto Primário (Nome do Produto) (Produto)",
    ]
    return pd.DataFrame(linhas, columns=colunas)


def _df_generico(linhas: list[dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(linhas, columns=["Cliente", "CPF", "Telefone"])


class ConsolidacaoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_remove_registro_sem_telefone(self) -> None:
        arq = _salvar(
            _df_crm1(
                [
                    {"Telefone da Empresa/Celular": "48999990001", "CPF_CNPJ": "1"},
                    {"Telefone da Empresa/Celular": "", "CPF_CNPJ": "2"},
                ]
            ),
            self.dir / "a.xlsx",
        )
        final, metricas = processar_consolidado([arq])
        self.assertEqual(metricas["registros_lidos"], 2)
        self.assertEqual(metricas["removidos_sem_telefone"], 1)
        self.assertEqual(metricas["registros_finais"], 1)

    def test_normaliza_telefone_como_v1(self) -> None:
        arq = _salvar(
            _df_crm1([{"Telefone da Empresa/Celular": "(48) 99999-0001", "CPF_CNPJ": "1"}]),
            self.dir / "a.xlsx",
        )
        final, _ = processar_consolidado([arq])
        self.assertEqual(final.iloc[0][CANON_TELEFONE], "5548999990001")

    def test_dedup_dentro_do_mesmo_arquivo(self) -> None:
        linha = {
            "Razão do Status": "Em contato",
            "Nome Completo": "Ana",
            "CPF_CNPJ": "1",
            "E-mail": "ana@x.com",
            "Telefone da Empresa/Celular": "48999990001",
            "Nome do produto": "P",
            "Primário (Nome do produto) (Produto)": "Q",
        }
        arq = _salvar(_df_crm1([linha, dict(linha)]), self.dir / "a.xlsx")
        _, metricas = processar_consolidado([arq])
        self.assertEqual(metricas["duplicados_removidos"], 1)
        self.assertEqual(metricas["registros_finais"], 1)

    def test_dedup_entre_arquivos_diferentes(self) -> None:
        linha = {
            "Razão do Status": "Em contato",
            "Nome Completo": "Ana",
            "CPF_CNPJ": "1",
            "E-mail": "ana@x.com",
            "Telefone da Empresa/Celular": "48999990001",
            "Nome do produto": "P",
            "Primário (Nome do produto) (Produto)": "Q",
        }
        a = _salvar(_df_crm1([linha]), self.dir / "a.xlsx")
        b = _salvar(_df_crm1([dict(linha)]), self.dir / "b.xlsx")
        _, metricas = processar_consolidado([a, b])
        self.assertEqual(metricas["registros_lidos"], 2)
        self.assertEqual(metricas["duplicados_removidos"], 1)
        self.assertEqual(metricas["registros_finais"], 1)

    def test_dedup_entre_formatos_diferentes(self) -> None:
        # CRM 1 e genérico que resultam no MESMO registro canônico devem
        # colapsar em uma única linha (dedup global sobre o canônico).
        crm1 = _salvar(
            _df_crm1(
                [
                    {
                        "Razão do Status": "",
                        "Nome Completo": "Ana",
                        "CPF_CNPJ": "1",
                        "E-mail": "",
                        "Telefone da Empresa/Celular": "48999990001",
                        "Nome do produto": "",
                        "Primário (Nome do produto) (Produto)": "",
                    }
                ]
            ),
            self.dir / "crm1.xlsx",
        )
        gen = _salvar(
            _df_generico([{"Cliente": "Ana", "CPF": "1", "Telefone": "48999990001"}]),
            self.dir / "gen.xlsx",
        )
        final, metricas = processar_consolidado([crm1, gen])
        self.assertEqual(metricas["registros_lidos"], 2)
        self.assertEqual(metricas["duplicados_removidos"], 1)
        self.assertEqual(metricas["registros_finais"], 1)
        linha = final.iloc[0]
        self.assertEqual(linha[CANON_NOME_COMPLETO], "Ana")
        self.assertEqual(linha[CANON_CPF], "1")

    def test_diferenca_em_campo_canonico_nao_deduplica(self) -> None:
        a = _salvar(
            _df_crm1(
                [
                    {
                        "Nome Completo": "Ana",
                        "CPF_CNPJ": "1",
                        "Telefone da Empresa/Celular": "48999990001",
                        "Nome do produto": "Produto A",
                        "Primário (Nome do produto) (Produto)": "Q",
                    },
                    {
                        "Nome Completo": "Ana",
                        "CPF_CNPJ": "1",
                        "Telefone da Empresa/Celular": "48999990001",
                        "Nome do produto": "Produto B",
                        "Primário (Nome do produto) (Produto)": "Q",
                    },
                ]
            ),
            self.dir / "a.xlsx",
        )
        _, metricas = processar_consolidado([a])
        self.assertEqual(metricas["duplicados_removidos"], 0)
        self.assertEqual(metricas["registros_finais"], 2)

    def test_regional_do_crm2_preservada(self) -> None:
        arq = _salvar(
            _df_crm2(
                [
                    {
                        "Regional/Unidade do Produto": "SESI SC",
                        "Nome Completo": "Maria",
                        "CPF": "9",
                        "E-mail Comercial": "m@x.com",
                        "Telefone Celular": "48988880002",
                        "Nome do Produto": "Curso",
                        "Produto Primário (Nome do Produto) (Produto)": "Linha",
                    }
                ]
            ),
            self.dir / "crm2.xlsx",
        )
        final, _ = processar_consolidado([arq])
        self.assertEqual(final.iloc[0][CANON_REGIONAL], "SESI SC")

    def test_limite_de_arquivos(self) -> None:
        arquivos = []
        for i in range(LIMITE_ARQUIVOS + 1):
            arquivos.append(
                _salvar(
                    _df_crm1([{"Telefone da Empresa/Celular": f"4899999{i:04d}", "CPF_CNPJ": str(i)}]),
                    self.dir / f"f{i}.xlsx",
                )
            )
        with self.assertRaises(LimiteArquivosExcedidoError):
            processar_consolidado(arquivos)

    def test_ate_o_limite_processa(self) -> None:
        arquivos = []
        for i in range(LIMITE_ARQUIVOS):
            arquivos.append(
                _salvar(
                    _df_crm1([{"Telefone da Empresa/Celular": f"4899999{i:04d}", "CPF_CNPJ": str(i)}]),
                    self.dir / f"f{i}.xlsx",
                )
            )
        _, metricas = processar_consolidado(arquivos)
        self.assertEqual(metricas["registros_lidos"], LIMITE_ARQUIVOS)
        self.assertEqual(metricas["registros_finais"], LIMITE_ARQUIVOS)


class SaidaEstendidaTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _arquivo_crm2(self) -> Path:
        return _salvar(
            _df_crm2(
                [
                    {
                        "Regional/Unidade do Produto": "SESI SC",
                        "Razão do Status": "Em contato",
                        "Nome Completo": "Maria",
                        "CPF": "9",
                        "E-mail Comercial": "m@x.com",
                        "Telefone Celular": "(48) 98888-0002",
                        "Nome do Produto": "Curso",
                        "Produto Primário (Nome do Produto) (Produto)": "Linha",
                    }
                ]
            ),
            self.dir / "crm2.xlsx",
        )

    def test_estendida_tem_sete_colunas_na_ordem(self) -> None:
        arq = self._arquivo_crm2()
        gerar_saidas_v2([arq], FORMATO_PLANILHA_ESTENDIDA)
        saida = pd.read_excel(arq.with_name("crm2_estendido.xlsx"))
        self.assertEqual(
            saida.columns.tolist(),
            [
                "Razão do Status",
                "Nome Completo",
                "CPF",
                "E-mail",
                "Telefone da Empresa/Celular",
                "Regional/Unidade do Produto",
                "Nome do Produto",
            ],
        )

    def test_estendida_valores_e_telefone_normalizado(self) -> None:
        arq = self._arquivo_crm2()
        gerar_saidas_v2([arq], FORMATO_PLANILHA_ESTENDIDA)
        linha = pd.read_excel(arq.with_name("crm2_estendido.xlsx")).iloc[0]
        self.assertEqual(linha["Regional/Unidade do Produto"], "SESI SC")
        self.assertEqual(linha["Nome do Produto"], "Curso")
        self.assertEqual(str(linha["Telefone da Empresa/Celular"]), "5548988880002")

    def test_estendida_regional_vazia_para_crm1(self) -> None:
        arq = _salvar(
            _df_crm1(
                [
                    {
                        "Razão do Status": "Em contato",
                        "Nome Completo": "Ana",
                        "CPF_CNPJ": "1",
                        "E-mail": "ana@x.com",
                        "Telefone da Empresa/Celular": "48999990001",
                        "Nome do produto": "Produto X",
                        "Primário (Nome do produto) (Produto)": "Q",
                    }
                ]
            ),
            self.dir / "crm1.xlsx",
        )
        gerar_saidas_v2([arq], FORMATO_PLANILHA_ESTENDIDA)
        linha = pd.read_excel(arq.with_name("crm1_estendido.xlsx")).iloc[0]
        # Célula vazia é lida como NaN pelo pandas.
        self.assertTrue(pd.isna(linha["Regional/Unidade do Produto"]))
        self.assertEqual(linha["Nome do Produto"], "Produto X")

    def test_saida_completa_v2_tem_cinco_colunas(self) -> None:
        arq = self._arquivo_crm2()
        gerar_saidas_v2([arq], FORMATO_PLANILHA_COMPLETA)
        saida = pd.read_excel(arq.with_name("crm2_tratado.xlsx"))
        self.assertEqual(
            saida.columns.tolist(),
            ["Razão do Status", "Nome Completo", "CPF", "E-mail", "Telefone da Empresa/Celular"],
        )

    def test_ambas_saidas_v2_geram_dois_arquivos(self) -> None:
        arq = self._arquivo_crm2()
        resultado = gerar_saidas_v2([arq], FORMATO_AMBAS_SAIDAS)
        self.assertEqual(len(resultado["arquivos_gerados"]), 2)

    def test_nome_consolidado_para_multiplos_arquivos(self) -> None:
        a = _salvar(
            _df_crm1([{"Telefone da Empresa/Celular": "48999990001", "CPF_CNPJ": "1"}]),
            self.dir / "primeiro.xlsx",
        )
        b = _salvar(
            _df_crm1([{"Telefone da Empresa/Celular": "48999990002", "CPF_CNPJ": "2"}]),
            self.dir / "segundo.xlsx",
        )
        resultado = gerar_saidas_v2([a, b], FORMATO_PLANILHA_ESTENDIDA)
        gerado = Path(resultado["arquivos_gerados"][0]).name
        self.assertEqual(gerado, "Consolidado_primeiro_estendido.xlsx")


if __name__ == "__main__":
    unittest.main()
