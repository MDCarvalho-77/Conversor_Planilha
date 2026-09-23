"""Testes de regressão da V2 usando a planilha de cenários validada."""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from pandas.testing import assert_frame_equal

from processamento import (
    COLUNA_TELEFONE,
    COLUNA_TELEFONE_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    gerar_saidas,
)


PASTA_PROJETO = Path(__file__).parent
ARQUIVO_CENARIOS = PASTA_PROJETO / "Testes.xlsx"
SAIDA_V1 = PASTA_PROJETO / "Testes_tratado.xlsx"


class ProcessamentoV2Test(unittest.TestCase):
    def setUp(self) -> None:
        self.diretorio_temporario = tempfile.TemporaryDirectory()
        self.entrada = Path(self.diretorio_temporario.name) / "cenarios.xlsx"
        shutil.copy2(ARQUIVO_CENARIOS, self.entrada)

    def tearDown(self) -> None:
        self.diretorio_temporario.cleanup()

    def test_planilha_completa_mantem_saida_e_metricas_da_v1(self) -> None:
        resultado = gerar_saidas(self.entrada, FORMATO_PLANILHA_COMPLETA)
        completo = self.entrada.with_name("cenarios_tratado.xlsx")

        self.assertEqual(
            {chave: resultado[chave] for chave in (
                "registros_lidos", "removidos_sem_telefone", "duplicados_removidos", "registros_finais"
            )},
            {
                "registros_lidos": 16,
                "removidos_sem_telefone": 1,
                "duplicados_removidos": 1,
                "registros_finais": 14,
            },
        )
        assert_frame_equal(pd.read_excel(completo), pd.read_excel(SAIDA_V1))

    def test_somente_numeros_tem_apenas_coluna_telefone(self) -> None:
        gerar_saidas(self.entrada, FORMATO_SOMENTE_NUMEROS)
        telefones = pd.read_excel(self.entrada.with_name("cenarios_telefones.xlsx"))

        self.assertEqual(telefones.columns.tolist(), [COLUNA_TELEFONE_SOMENTE_NUMEROS])
        self.assertEqual(len(telefones), 14)

    def test_ambas_saidas_geram_telefones_identicos(self) -> None:
        resultado = gerar_saidas(self.entrada, FORMATO_AMBAS_SAIDAS)
        completo = pd.read_excel(self.entrada.with_name("cenarios_tratado.xlsx"))
        telefones = pd.read_excel(self.entrada.with_name("cenarios_telefones.xlsx"))

        self.assertEqual(len(resultado["arquivos_gerados"]), 2)
        self.assertEqual(
            completo[COLUNA_TELEFONE].tolist(),
            telefones[COLUNA_TELEFONE_SOMENTE_NUMEROS].tolist(),
        )


if __name__ == "__main__":
    unittest.main()
