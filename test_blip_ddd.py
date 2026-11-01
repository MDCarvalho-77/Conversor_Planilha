"""Testes da V3.1: classificação por DDD, exportação Blip por grupo e
reimportação de revisão com reclassificação pelo telefone atual."""

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
    GRUPO_DDD_47,
    GRUPO_DDD_48_OUTROS,
    GRUPO_DDD_49,
    GRUPO_SEM_DDD,
    GRUPOS_DDD,
    MAX_BYTES_POR_ARQUIVO,
    MAX_CONTATOS_POR_ARQUIVO,
    NOME_EXCEL_REVISAO,
    PREFIXO_CSV,
    SUBPASTA_BLIP,
    classificar_ddd,
    extrair_ddd_brasil,
    gerar_csvs_por_ddd_de_revisao,
    preparar_blip_por_ddd,
    separar_por_ddd,
)


def _canonico(telefones: list[str], nomes: list[str] | None = None) -> pd.DataFrame:
    """Monta um DataFrame canônico mínimo com telefone e nome."""
    tamanho = len(telefones)
    nomes = nomes if nomes is not None else [f"Contato {i}" for i in range(tamanho)]
    dados = {coluna: [""] * tamanho for coluna in COLUNAS_CANONICAS}
    dados[CANON_TELEFONE] = telefones
    dados[CANON_NOME_COMPLETO] = nomes
    return pd.DataFrame(dados, columns=list(COLUNAS_CANONICAS))


# Telefones brasileiros já normalizados (dígitos com DDI 55), por DDD.
def _tel_br(ddd: str, i: int) -> str:
    """Telefone BR móvel (13 dígitos: 55 + DDD + 9 dígitos), sem '+'."""
    return f"55{ddd}9{i:08d}"


class ExtrairDddBrasilTest(unittest.TestCase):
    def test_ddd_47(self) -> None:
        self.assertEqual(extrair_ddd_brasil("+5547999990001"), "47")

    def test_ddd_49(self) -> None:
        self.assertEqual(extrair_ddd_brasil("+5549999990001"), "49")

    def test_ddd_48(self) -> None:
        self.assertEqual(extrair_ddd_brasil("+5548999990001"), "48")

    def test_ddd_sem_mais(self) -> None:
        self.assertEqual(extrair_ddd_brasil("5547999990001"), "47")

    def test_fixo_br_12_digitos(self) -> None:
        # 55 + DDD(11) + 8 dígitos = 12 dígitos.
        self.assertEqual(extrair_ddd_brasil("+551133334444"), "11")

    def test_internacional_nao_br_eua(self) -> None:
        # +1 (EUA): não começa com 55 -> None (não confundir '1x' com DDD).
        self.assertIsNone(extrair_ddd_brasil("+14155550123"))

    def test_internacional_nao_br_portugal(self) -> None:
        self.assertIsNone(extrair_ddd_brasil("+351912345678"))

    def test_internacional_nao_br_reino_unido(self) -> None:
        self.assertIsNone(extrair_ddd_brasil("+447911123456"))

    def test_bsuid_retorna_none(self) -> None:
        self.assertIsNone(extrair_ddd_brasil("abc123@blip"))

    def test_vazio_retorna_none(self) -> None:
        self.assertIsNone(extrair_ddd_brasil(""))
        self.assertIsNone(extrair_ddd_brasil(None))
        self.assertIsNone(extrair_ddd_brasil("nan"))

    def test_55_comprimento_incompativel_retorna_none(self) -> None:
        # Começa com 55 mas tem comprimento fora de 12/13 -> None.
        self.assertIsNone(extrair_ddd_brasil("+5547"))
        self.assertIsNone(extrair_ddd_brasil("+554799999000112345"))

    def test_ddd_fora_da_faixa_retorna_none(self) -> None:
        # 55 + '09' + ... : DDD 09 fora da faixa 11-99.
        self.assertIsNone(extrair_ddd_brasil("+550933334444"))


class ClassificarDddTest(unittest.TestCase):
    def test_47(self) -> None:
        self.assertEqual(classificar_ddd("+5547999990001"), GRUPO_DDD_47)

    def test_49(self) -> None:
        self.assertEqual(classificar_ddd("+5549999990001"), GRUPO_DDD_49)

    def test_48_vai_para_outros(self) -> None:
        self.assertEqual(classificar_ddd("+5548999990001"), GRUPO_DDD_48_OUTROS)

    def test_outro_ddd_br_11(self) -> None:
        self.assertEqual(classificar_ddd("+5511999990001"), GRUPO_DDD_48_OUTROS)

    def test_outro_ddd_br_85(self) -> None:
        self.assertEqual(classificar_ddd("+5585999990001"), GRUPO_DDD_48_OUTROS)

    def test_internacional_nao_br_vai_sem_ddd(self) -> None:
        self.assertEqual(classificar_ddd("+14155550123"), GRUPO_SEM_DDD)
        self.assertEqual(classificar_ddd("+351912345678"), GRUPO_SEM_DDD)

    def test_bsuid_vai_sem_ddd(self) -> None:
        self.assertEqual(classificar_ddd("abc123@blip"), GRUPO_SEM_DDD)

    def test_vazio_vai_sem_ddd(self) -> None:
        self.assertEqual(classificar_ddd(""), GRUPO_SEM_DDD)

    def test_55_malformado_vai_sem_ddd(self) -> None:
        self.assertEqual(classificar_ddd("+5547"), GRUPO_SEM_DDD)

    def test_cada_entrada_em_exatamente_um_grupo(self) -> None:
        amostras = [
            "+5547999990001",
            "+5549999990001",
            "+5548999990001",
            "+5511999990001",
            "+14155550123",
            "+351912345678",
            "abc123@blip",
            "",
            "+5547",
        ]
        for amostra in amostras:
            grupo = classificar_ddd(amostra)
            # Exatamente 1 dos 4 grupos.
            self.assertIn(grupo, GRUPOS_DDD)
            self.assertEqual(sum(1 for g in GRUPOS_DDD if g == grupo), 1)


class SepararPorDddTest(unittest.TestCase):
    def test_sempre_quatro_chaves(self) -> None:
        grupos = separar_por_ddd([])
        self.assertEqual(list(grupos.keys()), list(GRUPOS_DDD))
        for lista in grupos.values():
            self.assertEqual(lista, [])

    def test_preserva_ordem_dentro_do_grupo(self) -> None:
        contatos = [
            ("+5547999990001", "A47"),
            ("+5547999990002", "B47"),
            ("+5549999990001", "A49"),
            ("+5548999990001", "A48"),
            ("abc@blip", "SemDDD"),
        ]
        grupos = separar_por_ddd(contatos)
        self.assertEqual(
            grupos[GRUPO_DDD_47],
            [("+5547999990001", "A47"), ("+5547999990002", "B47")],
        )
        self.assertEqual(grupos[GRUPO_DDD_49], [("+5549999990001", "A49")])
        self.assertEqual(grupos[GRUPO_DDD_48_OUTROS], [("+5548999990001", "A48")])
        self.assertEqual(grupos[GRUPO_SEM_DDD], [("abc@blip", "SemDDD")])


class PrepararBlipPorDddTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_quatro_grupos_geram_quatro_subpastas(self) -> None:
        df = _canonico(
            [
                _tel_br("47", 1),
                _tel_br("49", 1),
                _tel_br("48", 1),
                "abc123@blip",  # SEM_DDD
            ],
            ["N47", "N49", "N48", "NSem"],
        )
        resumo = preparar_blip_por_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        for grupo in GRUPOS_DDD:
            with self.subTest(grupo=grupo):
                self.assertTrue((raiz / grupo).is_dir())
                self.assertTrue((raiz / grupo / f"{PREFIXO_CSV}001.csv").exists())
                self.assertTrue((raiz / grupo / NOME_EXCEL_REVISAO).exists())
        self.assertEqual(resumo["pasta_raiz"], raiz)
        self.assertEqual(resumo["contatos_exportados"], 4)
        self.assertEqual(resumo["total_sem_ddd"], 1)

    def test_grupo_vazio_nao_cria_pasta(self) -> None:
        # Só DDD 47; os demais grupos NÃO devem ter pasta.
        df = _canonico([_tel_br("47", 1), _tel_br("47", 2)])
        preparar_blip_por_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        self.assertTrue((raiz / GRUPO_DDD_47).is_dir())
        self.assertFalse((raiz / GRUPO_DDD_49).exists())
        self.assertFalse((raiz / GRUPO_DDD_48_OUTROS).exists())
        self.assertFalse((raiz / GRUPO_SEM_DDD).exists())

    def test_mais_de_100_num_grupo_gera_csvs_numerados(self) -> None:
        df = _canonico([_tel_br("47", i) for i in range(150)])
        resumo = preparar_blip_por_ddd(df, self.dir)
        pasta = self.dir / SUBPASTA_BLIP / GRUPO_DDD_47
        self.assertTrue((pasta / f"{PREFIXO_CSV}001.csv").exists())
        self.assertTrue((pasta / f"{PREFIXO_CSV}002.csv").exists())
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_47]["lotes"], 2)

    def test_mais_de_100kb_num_grupo_divide(self) -> None:
        nome_grande = "N" * 1500
        df = _canonico(
            [_tel_br("49", i) for i in range(100)],
            [nome_grande] * 100,
        )
        resumo = preparar_blip_por_ddd(df, self.dir)
        detalhe = resumo["por_grupo"][GRUPO_DDD_49]
        self.assertGreater(detalhe["lotes"], 1)  # dividiu por tamanho
        for csv_detalhe in detalhe["detalhes_csv"]:
            self.assertLessEqual(csv_detalhe["bytes"], MAX_BYTES_POR_ARQUIVO)

    def test_numeracao_independente_por_grupo(self) -> None:
        # Cada grupo recomeça em Blip_001.csv.
        df = _canonico(
            [_tel_br("47", 1), _tel_br("49", 1), _tel_br("48", 1)],
        )
        preparar_blip_por_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        for grupo in (GRUPO_DDD_47, GRUPO_DDD_49, GRUPO_DDD_48_OUTROS):
            self.assertTrue((raiz / grupo / f"{PREFIXO_CSV}001.csv").exists())

    def test_csvs_nao_misturam_grupos(self) -> None:
        df = _canonico(
            [_tel_br("47", 1), _tel_br("49", 1), _tel_br("48", 1), "abc@blip"],
        )
        preparar_blip_por_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        esperado = {
            GRUPO_DDD_47: "5547",
            GRUPO_DDD_49: "5549",
            GRUPO_DDD_48_OUTROS: "5548",
        }
        for grupo, prefixo_digitos in esperado.items():
            caminho = raiz / grupo / f"{PREFIXO_CSV}001.csv"
            with open(caminho, "r", encoding="utf-8", newline="") as arquivo:
                linhas = list(csv.reader(arquivo))
            for linha in linhas[1:]:
                self.assertTrue(linha[0].lstrip("+").startswith(prefixo_digitos))

    def test_um_excel_de_revisao_por_subpasta(self) -> None:
        df = _canonico([_tel_br("47", 1), _tel_br("49", 1)])
        preparar_blip_por_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        for grupo in (GRUPO_DDD_47, GRUPO_DDD_49):
            excel = raiz / grupo / NOME_EXCEL_REVISAO
            self.assertTrue(excel.exists())
            wb = load_workbook(excel)
            # Nome da aba prefixado pelo grupo, 1a linha intacta.
            aba = wb[wb.sheetnames[0]]
            self.assertTrue(wb.sheetnames[0].startswith(grupo))
            self.assertEqual(
                [aba.cell(1, 1).value, aba.cell(1, 2).value],
                [COLUNA_ID, COLUNA_NOME],
            )


class ReimportacaoPorDddTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _excel_grupo(self, grupo: str, linhas: list[tuple[str, str]]) -> Path:
        """Cria <dir>/<grupo>/Blip_Contatos_Revisao.xlsx com os contatos."""
        pasta = self.dir / grupo
        pasta.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        wb.remove(wb.active)
        aba = wb.create_sheet(title=f"{grupo}_Lote_001"[:31])
        aba.append([COLUNA_ID, COLUNA_NOME])
        for ident, nome in linhas:
            aba.append([ident, nome])
        caminho = pasta / NOME_EXCEL_REVISAO
        wb.save(caminho)
        return caminho

    def test_telefone_alterado_muda_de_grupo(self) -> None:
        # Estava em DDD_47 mas o telefone agora é DDD 49.
        origem = self._excel_grupo(GRUPO_DDD_47, [("+5549999990001", "Mudou")])
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(resumo["mudaram_de_grupo"], 1)
        self.assertTrue((saida / SUBPASTA_BLIP / GRUPO_DDD_49 / f"{PREFIXO_CSV}001.csv").exists())
        self.assertFalse((saida / SUBPASTA_BLIP / GRUPO_DDD_47).exists())

    def test_passou_a_inidentificavel_vai_sem_ddd(self) -> None:
        origem = self._excel_grupo(GRUPO_DDD_47, [("abc123@blip", "AgoraBSUID")])
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(resumo["mudaram_de_grupo"], 1)
        self.assertTrue((saida / SUBPASTA_BLIP / GRUPO_SEM_DDD).is_dir())

    def test_recebeu_ddd_valido_vai_para_grupo_certo(self) -> None:
        # Estava em SEM_DDD, agora tem DDD 48.
        origem = self._excel_grupo(GRUPO_SEM_DDD, [("+5548999990001", "GanhouDDD")])
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(resumo["mudaram_de_grupo"], 1)
        self.assertTrue((saida / SUBPASTA_BLIP / GRUPO_DDD_48_OUTROS).is_dir())

    def test_sem_mudanca_nao_conta(self) -> None:
        origem = self._excel_grupo(GRUPO_DDD_47, [("+5547999990001", "Fica")])
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(resumo["mudaram_de_grupo"], 0)

    def test_redistribuicao_de_varios_arquivos(self) -> None:
        origem_47 = self._excel_grupo(
            GRUPO_DDD_47,
            [("+5547999990001", "Fica47"), ("+5549999990002", "VaiPro49")],
        )
        origem_49 = self._excel_grupo(
            GRUPO_DDD_49,
            [("+5548999990003", "VaiProOutros")],
        )
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem_47, origem_49], saida)
        self.assertEqual(resumo["mudaram_de_grupo"], 2)
        self.assertEqual(resumo["totais_por_grupo"][GRUPO_DDD_47], 1)
        self.assertEqual(resumo["totais_por_grupo"][GRUPO_DDD_49], 1)
        self.assertEqual(resumo["totais_por_grupo"][GRUPO_DDD_48_OUTROS], 1)

    def test_invalido_vai_para_relatorio(self) -> None:
        origem = self._excel_grupo(
            GRUPO_DDD_47,
            [("+5547999990001", "OK"), ("telefone invalido", "Ruim")],
        )
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(resumo["total_invalidos"], 1)
        self.assertEqual(resumo["contatos_exportados"], 1)

    def test_nao_modifica_arquivos_de_origem(self) -> None:
        origem = self._excel_grupo(GRUPO_DDD_47, [("+5549999990001", "Mudou")])
        antes_mtime = origem.stat().st_mtime_ns
        antes_bytes = origem.read_bytes()
        saida = self.dir / "saida"
        gerar_csvs_por_ddd_de_revisao([origem], saida)
        self.assertEqual(origem.stat().st_mtime_ns, antes_mtime)
        self.assertEqual(origem.read_bytes(), antes_bytes)


class FuncoesLegadasPreservadasTest(unittest.TestCase):
    """Garante que as funções internas da V3 continuam existindo."""

    def test_funcoes_v3_existem(self) -> None:
        for nome in (
            "preparar_blip",
            "gerar_blip_de_planilhas",
            "gerar_csvs_do_excel_revisado",
            "dividir_em_lotes",
            "gravar_csvs",
            "gerar_excel_revisao",
            "ler_excel_revisado",
            "validar_identificador",
            "formatar_identificador_telefone",
            "preparar_contatos",
            "_resumo",
        ):
            self.assertTrue(hasattr(blip, nome), nome)


if __name__ == "__main__":
    unittest.main()
