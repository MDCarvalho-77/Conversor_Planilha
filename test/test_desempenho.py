"""Teste de carga do processamento com planilhas CRM temporárias."""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
import unittest
from pathlib import Path

import pandas as pd

PROJETO = Path(__file__).resolve().parents[1]
if str(PROJETO) not in sys.path:
    sys.path.insert(0, str(PROJETO))

from processamento import (  # noqa: E402
    COLUNAS_PRODUTO,
    COLUNAS_SAIDA,
    COLUNA_TELEFONE,
    processar_planilha,
)


TAMANHOS = (5_000, 10_000, 25_000)


def _telefone_realista(indice: int) -> str:
    """Gera telefones nos formatos tratados pela regra de negócio atual."""
    tipo = indice % 5
    base_10 = f"48{indice:08d}"
    base_11 = f"48{indice:09d}"
    if tipo == 0:
        return f"({base_10[:2]}) {base_10[2:6]}-{base_10[6:]}"
    if tipo == 1:
        return f"({base_11[:2]}) {base_11[2:7]}-{base_11[7:]}"
    if tipo == 2:
        return f"55 {base_10}"
    if tipo == 3:
        return f"+55 ({base_11[:2]}) {base_11[2:7]}-{base_11[7:]}"
    return base_11


def _registro_crm(indice: int) -> dict[str, object]:
    """Cria um registro completo com variações de produto e telefone."""
    registro = {
        COLUNAS_SAIDA[0]: "Em contato",
        COLUNAS_SAIDA[1]: f"Cliente de carga {indice:05d}",
        COLUNAS_SAIDA[2]: f"CPF-{indice:011d}",
        COLUNAS_SAIDA[3]: f"cliente{indice:05d}@exemplo.local",
        COLUNA_TELEFONE: None if indice % 41 == 0 else _telefone_realista(indice),
        COLUNAS_PRODUTO[0]: f"Produto {indice % 7}",
        COLUNAS_PRODUTO[1]: f"Linha de produto {indice % 5}",
    }
    return registro


def criar_planilha_crm(caminho: Path, quantidade: int) -> None:
    """Gera mistura determinística de únicos, duplicatas e telefones variados."""
    registros: list[dict[str, object]] = []
    for indice in range(quantidade):
        if indice > 0 and indice % 29 == 0:
            # Cópia integral: deve ser removida pela deduplicação já validada.
            registros.append(registros[-1].copy())
        else:
            registros.append(_registro_crm(indice))
    pd.DataFrame(registros).to_excel(caminho, index=False, engine="openpyxl")


def _tamanho_formatado(tamanho_bytes: int) -> str:
    return f"{tamanho_bytes / 1024 / 1024:.2f} MB"


class TesteDesempenhoProcessamento(unittest.TestCase):
    """Mede a execução ponta a ponta sem deixar arquivos no projeto."""

    def test_cargas_de_planilha_crm(self) -> None:
        resultados: list[dict[str, object]] = []
        with tempfile.TemporaryDirectory(prefix="conversor_crm_desempenho_") as pasta:
            diretorio = Path(pasta)
            for quantidade in TAMANHOS:
                entrada = diretorio / f"crm_{quantidade}.xlsx"
                saida = diretorio / f"crm_{quantidade}_tratado.xlsx"
                criar_planilha_crm(entrada, quantidade)

                inicio = time.perf_counter()
                metricas = processar_planilha(entrada, saida)
                tempo = time.perf_counter() - inicio

                self.assertEqual(metricas["registros_lidos"], quantidade)
                self.assertTrue(saida.is_file())
                self.assertLess(metricas["registros_finais"], quantidade)
                resultados.append(
                    {
                        "registros": quantidade,
                        "entrada": entrada.stat().st_size,
                        "tempo": tempo,
                        "finais": metricas["registros_finais"],
                        "saida": saida.stat().st_size,
                    }
                )

        print("\nResultados do teste de desempenho")
        print("| Registros | Tamanho entrada | Tempo processamento | Registros finais | Saída |")
        print("| ---: | ---: | ---: | ---: | ---: |")
        for resultado in resultados:
            print(
                "| {registros:,} | {entrada} | {tempo:.2f} s | {finais:,} | {saida} |".format(
                    registros=resultado["registros"],
                    entrada=_tamanho_formatado(resultado["entrada"]),
                    tempo=resultado["tempo"],
                    finais=resultado["finais"],
                    saida=_tamanho_formatado(resultado["saida"]),
                )
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
