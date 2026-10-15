"""Testes do modo Blip (V3): preparação, divisão, CSV e Excel de revisão."""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

import blip
from adapters import CANON_NOME_COMPLETO, CANON_TELEFONE, COLUNAS_CANONICAS
from blip import (
    COLUNA_ID,
    COLUNA_NOME,
    MAX_BYTES_POR_ARQUIVO,
    MAX_CONTATOS_POR_ARQUIVO,
    NOME_EXCEL_REVISAO,
    BlipError,
    dividir_em_lotes,
    formatar_identificador_telefone,
    gerar_csvs_do_excel_revisado,
    gerar_excel_revisao,
    gravar_csvs,
    ler_excel_revisado,
    preparar_blip,
    preparar_contatos,
    validar_csv_gravado,
    validar_identificador,
)


def _canonico(telefones: list[str], nomes: list[str] | None = None) -> pd.DataFrame:
    """Monta um DataFrame canônico mínimo com telefone e nome."""
    tamanho = len(telefones)
    nomes = nomes if nomes is not None else [f"Contato {i}" for i in range(tamanho)]
    dados = {coluna: [""] * tamanho for coluna in COLUNAS_CANONICAS}
    dados[CANON_TELEFONE] = telefones
    dados[CANON_NOME_COMPLETO] = nomes
    return pd.DataFrame(dados, columns=list(COLUNAS_CANONICAS))


def _telefone(i: int) -> str:
    """Telefone já normalizado pela V2 (dígitos com DDI 55), sem '+'."""
    return f"5548{i:09d}"


class FormatacaoIdentificadorTest(unittest.TestCase):
    def test_adiciona_mais_a_telefone_sem_mais(self) -> None:
        self.assertEqual(formatar_identificador_telefone("5548999990001"), "+5548999990001")

    def test_preserva_mais_existente(self) -> None:
        self.assertEqual(formatar_identificador_telefone("+5548999990001"), "+5548999990001")

    def test_remove_caracteres_nao_numericos_mantendo_mais(self) -> None:
        self.assertEqual(formatar_identificador_telefone("+55 (48) 99999-0001"), "+5548999990001")

    def test_vazio_resulta_vazio(self) -> None:
        self.assertEqual(formatar_identificador_telefone(""), "")
        self.assertEqual(formatar_identificador_telefone(None), "")
        self.assertEqual(formatar_identificador_telefone("nan"), "")


class ValidarIdentificadorTest(unittest.TestCase):
    def test_telefone_com_mais_valido(self) -> None:
        self.assertEqual(validar_identificador("+5548999990001"), (True, "+5548999990001"))

    def test_digitos_viram_telefone(self) -> None:
        self.assertEqual(validar_identificador("5548999990001"), (True, "+5548999990001"))

    def test_bsuid_valido(self) -> None:
        ok, ident = validar_identificador("abc123@blip")
        self.assertTrue(ok)
        self.assertEqual(ident, "abc123@blip")

    def test_vazio_invalido(self) -> None:
        self.assertEqual(validar_identificador(""), (False, ""))

    def test_token_com_espaco_invalido(self) -> None:
        ok, _ = validar_identificador("com espaco")
        self.assertFalse(ok)


class PrepararContatosTest(unittest.TestCase):
    def test_mapeia_telefone_e_nome(self) -> None:
        df = _canonico(["5548999990001"], ["João da Silva"])
        contatos, invalidos = preparar_contatos(df)
        self.assertEqual(contatos, [("+5548999990001", "João da Silva")])
        self.assertEqual(invalidos, [])

    def test_registro_sem_telefone_vira_invalido(self) -> None:
        df = _canonico(["", "5548999990002"], ["Sem Tel", "Com Tel"])
        contatos, invalidos = preparar_contatos(df)
        self.assertEqual(len(contatos), 1)
        self.assertEqual(len(invalidos), 1)
        self.assertEqual(invalidos[0].nome, "Sem Tel")

    def test_nome_vazio_permitido(self) -> None:
        df = _canonico(["5548999990001"], [""])
        contatos, invalidos = preparar_contatos(df)
        self.assertEqual(contatos, [("+5548999990001", "")])
        self.assertEqual(invalidos, [])


class DivisaoLotesTest(unittest.TestCase):
    def _contatos(self, n: int) -> list[tuple[str, str]]:
        return [(f"+{_telefone(i)}", f"Contato {i}") for i in range(n)]

    def test_lista_vazia(self) -> None:
        lotes, invalidos = dividir_em_lotes([])
        self.assertEqual(lotes, [])
        self.assertEqual(invalidos, [])

    def test_um_contato(self) -> None:
        lotes, _ = dividir_em_lotes(self._contatos(1))
        self.assertEqual(len(lotes), 1)
        self.assertEqual(len(lotes[0]), 1)

    def test_exatamente_100(self) -> None:
        lotes, _ = dividir_em_lotes(self._contatos(100))
        self.assertEqual(len(lotes), 1)
        self.assertEqual(len(lotes[0]), 100)

    def test_101(self) -> None:
        lotes, _ = dividir_em_lotes(self._contatos(101))
        self.assertEqual(len(lotes), 2)
        self.assertEqual(len(lotes[0]), 100)
        self.assertEqual(len(lotes[1]), 1)

    def test_200(self) -> None:
        lotes, _ = dividir_em_lotes(self._contatos(200))
        self.assertEqual([len(l) for l in lotes], [100, 100])

    def test_201(self) -> None:
        lotes, _ = dividir_em_lotes(self._contatos(201))
        self.assertEqual([len(l) for l in lotes], [100, 100, 1])

    def test_ordem_preservada(self) -> None:
        contatos = self._contatos(150)
        lotes, _ = dividir_em_lotes(contatos)
        achatado = [c for lote in lotes for c in lote]
        self.assertEqual(achatado, contatos)

    def test_limite_de_tamanho_antes_de_100(self) -> None:
        # Nomes muito grandes forçam o corte por 100 KB antes de 100 contatos.
        nome_grande = "N" * 1500
        contatos = [(f"+{_telefone(i)}", nome_grande) for i in range(100)]
        lotes, invalidos = dividir_em_lotes(contatos)
        self.assertEqual(invalidos, [])
        self.assertGreater(len(lotes), 1)  # cortou por tamanho, não por contagem
        for lote in lotes:
            self.assertLessEqual(len(lote), MAX_CONTATOS_POR_ARQUIVO)

    def test_registro_que_nao_cabe_sozinho_vira_invalido(self) -> None:
        enorme = "X" * (MAX_BYTES_POR_ARQUIVO + 10)
        contatos = [("+5548999990001", "ok"), ("+5548999990002", enorme)]
        lotes, invalidos = dividir_em_lotes(contatos)
        self.assertEqual(len(invalidos), 1)
        achatado = [c for lote in lotes for c in lote]
        self.assertEqual(achatado, [("+5548999990001", "ok")])


class CsvGravacaoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_cabecalho_contado_no_tamanho(self) -> None:
        # Com nomes grandes, o cabeçalho afeta onde o lote é cortado.
        nome_grande = "N" * 1000
        contatos = [(f"+{_telefone(i)}", nome_grande) for i in range(130)]
        lotes, _ = dividir_em_lotes(contatos)
        resultados = gravar_csvs(lotes, self.dir)
        for r in resultados:
            self.assertLessEqual(r["bytes"], MAX_BYTES_POR_ARQUIVO)

    def test_mais_preservado_no_csv_gravado(self) -> None:
        lotes = [[("+5548999990001", "Fulano")]]
        gravar_csvs(lotes, self.dir)
        caminho = self.dir / "Blip_001.csv"
        with open(caminho, "r", encoding="utf-8", newline="") as f:
            linhas = list(csv.reader(f))
        self.assertEqual(linhas[0], [COLUNA_ID, COLUNA_NOME])
        self.assertEqual(linhas[1][0], "+5548999990001")

    def test_acentos_virgulas_aspas(self) -> None:
        nome = 'Antônio "Tonho", o Júnior; çãé'
        lotes = [[("+5548999990001", nome)]]
        gravar_csvs(lotes, self.dir)
        caminho = self.dir / "Blip_001.csv"
        with open(caminho, "r", encoding="utf-8", newline="") as f:
            linhas = list(csv.reader(f))
        self.assertEqual(linhas[1], ["+5548999990001", nome])

    def test_sem_coluna_indice_ou_extra(self) -> None:
        lotes = [[("+5548999990001", "A"), ("+5548999990002", "B")]]
        gravar_csvs(lotes, self.dir)
        caminho = self.dir / "Blip_001.csv"
        with open(caminho, "r", encoding="utf-8", newline="") as f:
            linhas = list(csv.reader(f))
        for linha in linhas:
            self.assertEqual(len(linha), 2)

    def test_nenhum_csv_ultrapassa_limites(self) -> None:
        contatos = [(f"+{_telefone(i)}", f"Contato {i}") for i in range(250)]
        lotes, _ = dividir_em_lotes(contatos)
        resultados = gravar_csvs(lotes, self.dir)
        self.assertEqual(len(resultados), 3)  # 100,100,50
        for r in resultados:
            self.assertLessEqual(r["bytes"], MAX_BYTES_POR_ARQUIVO)
            self.assertLessEqual(r["contatos"], MAX_CONTATOS_POR_ARQUIVO)

    def test_nao_sobrescreve_silenciosamente(self) -> None:
        gravar_csvs([[("+5548999990001", "A")]], self.dir)
        with self.assertRaises(BlipError):
            gravar_csvs([[("+5548999990002", "B")]], self.dir)

    def test_validar_csv_detecta_excesso_de_bytes(self) -> None:
        caminho = self.dir / "Blip_001.csv"
        with open(caminho, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f, lineterminator="\r\n")
            w.writerow([COLUNA_ID, COLUNA_NOME])
            w.writerow(["+5548999990001", "X" * (MAX_BYTES_POR_ARQUIVO + 100)])
        with self.assertRaises(BlipError):
            validar_csv_gravado(caminho)


class ExcelRevisaoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_varias_abas_uma_por_lote(self) -> None:
        contatos = [(f"+{_telefone(i)}", f"Contato {i}") for i in range(250)]
        lotes, _ = dividir_em_lotes(contatos)
        caminho = gerar_excel_revisao(lotes, self.dir)
        wb = load_workbook(caminho)
        self.assertEqual(wb.sheetnames, ["Lote_001", "Lote_002", "Lote_003"])
        aba = wb["Lote_001"]
        self.assertEqual([aba.cell(1, 1).value, aba.cell(1, 2).value], [COLUNA_ID, COLUNA_NOME])
        # Identificador como texto, com '+'.
        self.assertEqual(aba.cell(2, 1).value, "+" + _telefone(0))
        self.assertEqual(aba.cell(2, 1).number_format, "@")

    def test_nome_do_arquivo(self) -> None:
        caminho = gerar_excel_revisao([[("+5548999990001", "A")]], self.dir)
        self.assertEqual(caminho.name, NOME_EXCEL_REVISAO)


class ReimportacaoExcelTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _excel_manual(self, abas: dict[str, list[tuple[str, str]]]) -> Path:
        wb = Workbook()
        wb.remove(wb.active)
        for nome_aba, linhas in abas.items():
            aba = wb.create_sheet(title=nome_aba)
            aba.append([COLUNA_ID, COLUNA_NOME])
            for ident, nome in linhas:
                aba.append([ident, nome])
        caminho = self.dir / "revisado.xlsx"
        wb.save(caminho)
        return caminho

    def test_le_varias_abas_em_ordem(self) -> None:
        caminho = self._excel_manual(
            {
                "Lote_001": [("+5548999990001", "A"), ("+5548999990002", "B")],
                "Lote_002": [("+5548999990003", "C")],
            }
        )
        contatos = ler_excel_revisado(caminho)
        self.assertEqual(
            contatos,
            [("+5548999990001", "A"), ("+5548999990002", "B"), ("+5548999990003", "C")],
        )

    def test_aba_com_colunas_erradas_falha(self) -> None:
        wb = Workbook()
        wb.remove(wb.active)
        aba = wb.create_sheet(title="Lote_001")
        aba.append(["Coluna Errada", "Nome"])
        aba.append(["+5548999990001", "A"])
        caminho = self.dir / "ruim.xlsx"
        wb.save(caminho)
        with self.assertRaises(BlipError):
            ler_excel_revisado(caminho)

    def test_reimportacao_com_alteracoes_manuais(self) -> None:
        # Usuário inseriu, moveu e removeu contatos manualmente.
        caminho = self._excel_manual(
            {
                "Lote_001": [("+5548900000001", "Novo"), ("abc@blip", "BSUID")],
                "Lote_002": [("+5548900000002", "Outro"), ("", ""), ("telefone invalido", "X")],
            }
        )
        pasta = self.dir / "saida"
        resumo = gerar_csvs_do_excel_revisado(caminho, pasta)
        # 3 válidos (incluindo BSUID), 1 inválido (telefone invalido), vazio ignorado.
        self.assertEqual(resumo["contatos_exportados"], 3)
        self.assertEqual(resumo["total_invalidos"], 1)
        # Excel original não foi modificado.
        wb = load_workbook(caminho)
        self.assertIn("Lote_001", wb.sheetnames)

    def test_mais_de_100_em_uma_aba_redistribui(self) -> None:
        # Uma única aba com 150 contatos deve virar 2 CSVs (100 + 50).
        linhas = [(f"+{_telefone(i)}", f"C{i}") for i in range(150)]
        caminho = self._excel_manual({"Lote_001": linhas})
        pasta = self.dir / "saida"
        resumo = gerar_csvs_do_excel_revisado(caminho, pasta)
        self.assertEqual(resumo["lotes"], 2)
        for detalhe in resumo["detalhes_csv"]:
            self.assertLessEqual(detalhe["contatos"], MAX_CONTATOS_POR_ARQUIVO)
            self.assertLessEqual(detalhe["bytes"], MAX_BYTES_POR_ARQUIVO)

    def test_nao_modifica_excel_revisado(self) -> None:
        caminho = self._excel_manual({"Lote_001": [("+5548999990001", "A")]})
        antes = caminho.stat().st_mtime_ns
        gerar_csvs_do_excel_revisado(caminho, self.dir / "saida")
        depois = caminho.stat().st_mtime_ns
        self.assertEqual(antes, depois)


class PrepararBlipFluxoCompletoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_fluxo_gera_csvs_e_excel(self) -> None:
        df = _canonico([_telefone(i) for i in range(201)])
        resumo = preparar_blip(df, self.dir)
        self.assertEqual(resumo["lotes"], 3)
        self.assertEqual(resumo["contatos_exportados"], 201)
        self.assertTrue(Path(resumo["excel_revisao"]).exists())
        for detalhe in resumo["detalhes_csv"]:
            self.assertLessEqual(detalhe["bytes"], MAX_BYTES_POR_ARQUIVO)
            self.assertLessEqual(detalhe["contatos"], MAX_CONTATOS_POR_ARQUIVO)

    def test_lista_vazia_gera_resumo_sem_csv(self) -> None:
        df = _canonico([])
        resumo = preparar_blip(df, self.dir)
        self.assertEqual(resumo["contatos_exportados"], 0)
        self.assertEqual(resumo["lotes"], 0)
        self.assertEqual(resumo["arquivos_gerados"], [])
        # Excel de revisão é criado mesmo vazio (com uma aba vazia).
        self.assertTrue(Path(resumo["excel_revisao"]).exists())


if __name__ == "__main__":
    unittest.main()
