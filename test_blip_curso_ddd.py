"""Testes da V3.2: fluxo Curso -> DDD, validação do Requisito D (coluna de
curso), exportação em Blip/<curso>/<grupo_ddd>/, Excel de revisão de 4 colunas
e reimportação v3.2 com autodetecção de formato.

Os testes reabrem os arquivos gravados (CSV e XLSX) para conferir bytes e
estrutura reais, e preservam a V3/V3.1 (não alteram as funções existentes)."""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

import cursos
import blip
from adapters import (
    CANON_CURSO,
    CANON_EMAIL,
    CANON_NOME_COMPLETO,
    CANON_TELEFONE,
    COLUNAS_CANONICAS,
)
from blip import (
    COLUNA_CURSO,
    COLUNA_EMAIL,
    COLUNA_ID,
    COLUNA_NOME,
    COLUNAS_BLIP,
    COLUNAS_BLIP_V32,
    CursoColunaError,
    GRUPO_DDD_47,
    GRUPO_DDD_49,
    GRUPO_SEM_DDD,
    GRUPOS_DDD,
    MAX_BYTES_POR_ARQUIVO,
    MAX_CONTATOS_POR_ARQUIVO,
    NOME_EXCEL_REVISAO,
    PREFIXO_CSV,
    SUBPASTA_BLIP,
    colunas_visiveis,
    detectar_formato_revisao,
    gerar_blip_por_curso_e_ddd_de_planilhas,
    gerar_csvs_por_curso_e_ddd_de_revisao,
    gerar_excel_revisao_v32,
    ler_excel_revisao_v32,
    preparar_blip_por_curso_e_ddd,
    preparar_contatos_v32,
    separar_por_curso_e_ddd,
    validar_coluna_curso_planilhas,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _canonico_v32(
    telefones: list[str],
    nomes: list[str] | None = None,
    emails: list[str] | None = None,
    cursos_col: list[str] | None = None,
) -> pd.DataFrame:
    """Monta um DataFrame canônico com telefone, nome, e-mail e curso."""
    tamanho = len(telefones)
    nomes = nomes if nomes is not None else [f"Contato {i}" for i in range(tamanho)]
    emails = emails if emails is not None else [""] * tamanho
    cursos_col = cursos_col if cursos_col is not None else [""] * tamanho
    dados = {coluna: [""] * tamanho for coluna in COLUNAS_CANONICAS}
    dados[CANON_TELEFONE] = telefones
    dados[CANON_NOME_COMPLETO] = nomes
    dados[CANON_EMAIL] = emails
    dados[CANON_CURSO] = cursos_col
    return pd.DataFrame(dados, columns=list(COLUNAS_CANONICAS))


def _escrever_planilha(caminho: Path, cabecalho: list[str], linhas: list[list]) -> None:
    """Grava uma planilha .xlsx simples (uma aba) com cabeçalho e linhas."""
    wb = Workbook()
    aba = wb.active
    aba.append(cabecalho)
    for linha in linhas:
        aba.append(linha)
    wb.save(caminho)


# ---------------------------------------------------------------------------
# preparar_contatos_v32
# ---------------------------------------------------------------------------
class PrepararContatosV32Test(unittest.TestCase):
    def test_tupla_quatro_campos(self) -> None:
        df = _canonico_v32(
            ["5547999990001"],
            nomes=["Ana"],
            emails=["ana@x.com"],
            cursos_col=["Analise e Desenvolvimento de Sistemas"],
        )
        contatos, invalidos = preparar_contatos_v32(df)
        self.assertEqual(invalidos, [])
        self.assertEqual(len(contatos), 1)
        ident, nome, email, curso = contatos[0]
        self.assertEqual(ident, "+5547999990001")
        self.assertEqual(nome, "Ana")
        self.assertEqual(email, "ana@x.com")
        # Padronizado para a denominação cadastrada.
        self.assertEqual(curso, "Analise e Desenvolvimento de Sistemas")

    def test_curso_vazio_vira_nao_identificado(self) -> None:
        df = _canonico_v32(["5547999990001"], cursos_col=[""])
        contatos, _ = preparar_contatos_v32(df)
        self.assertEqual(contatos[0][3], cursos.CURSO_NAO_IDENTIFICADO)

    def test_email_ausente_fica_vazio(self) -> None:
        df = _canonico_v32(["5547999990001"], emails=[""])
        contatos, _ = preparar_contatos_v32(df)
        self.assertEqual(contatos[0][2], "")

    def test_sem_telefone_vai_para_invalidos(self) -> None:
        df = _canonico_v32(["", "5547999990001"], cursos_col=["X", "Y"])
        contatos, invalidos = preparar_contatos_v32(df)
        self.assertEqual(len(contatos), 1)
        self.assertEqual(len(invalidos), 1)
        self.assertEqual(invalidos[0].indice, 0)

    def test_padroniza_tres_variacoes_ads(self) -> None:
        df = _canonico_v32(
            ["5547999990001", "5547999990002", "5547999990003"],
            cursos_col=[
                "Superior de Tecnologia em Analise e Desenvolvimento de Sistemas",
                "Tecnologia em Análise e Desenvolvimento de Sistemas",
                "análise e desenvolvimento de sistemas",
            ],
        )
        contatos, _ = preparar_contatos_v32(df)
        cursos_resultantes = {c[3] for c in contatos}
        self.assertEqual(
            cursos_resultantes, {"Analise e Desenvolvimento de Sistemas"}
        )


# ---------------------------------------------------------------------------
# separar_por_curso_e_ddd
# ---------------------------------------------------------------------------
class SepararPorCursoEDddTest(unittest.TestCase):
    def test_cada_curso_recebe_quatro_grupos(self) -> None:
        contatos = [
            ("+5547999990001", "A", "", "ADS"),
            ("+5549999990001", "B", "", "ADS"),
        ]
        por_curso = separar_por_curso_e_ddd(contatos)
        self.assertEqual(list(por_curso["ADS"].keys()), list(GRUPOS_DDD))
        self.assertEqual(len(por_curso["ADS"][GRUPO_DDD_47]), 1)
        self.assertEqual(len(por_curso["ADS"][GRUPO_DDD_49]), 1)

    def test_curso_nao_identificado_subdividido(self) -> None:
        contatos = [
            ("+5547999990001", "A", "", cursos.CURSO_NAO_IDENTIFICADO),
            ("abc@blip", "B", "", cursos.CURSO_NAO_IDENTIFICADO),
        ]
        por_curso = separar_por_curso_e_ddd(contatos)
        self.assertIn(cursos.CURSO_NAO_IDENTIFICADO, por_curso)
        grupos = por_curso[cursos.CURSO_NAO_IDENTIFICADO]
        self.assertEqual(list(grupos.keys()), list(GRUPOS_DDD))
        self.assertEqual(len(grupos[GRUPO_DDD_47]), 1)
        self.assertEqual(len(grupos[GRUPO_SEM_DDD]), 1)

    def test_cada_contato_em_exatamente_uma_combinacao(self) -> None:
        contatos = [
            ("+5547999990001", "A", "", "ADS"),
            ("+5549999990001", "B", "", "Direito"),
            ("abc@blip", "C", "", "Direito"),
        ]
        por_curso = separar_por_curso_e_ddd(contatos)
        total = sum(
            len(lst)
            for grupos in por_curso.values()
            for lst in grupos.values()
        )
        self.assertEqual(total, 3)


# ---------------------------------------------------------------------------
# gerar_excel_revisao_v32
# ---------------------------------------------------------------------------
class GerarExcelRevisaoV32Test(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_quatro_colunas_e_dados(self) -> None:
        lotes = [[("+5547999990001", "Ana", "ana@x.com", "ADS")]]
        caminho = gerar_excel_revisao_v32(lotes, self.dir, prefixo_aba="ADS_DDD_47")
        wb = load_workbook(caminho)
        aba = wb.worksheets[0]
        cabecalho = [c.value for c in aba[1]]
        self.assertEqual(cabecalho, list(COLUNAS_BLIP_V32))
        self.assertEqual([c.value for c in aba[2]],
                         ["+5547999990001", "Ana", "ana@x.com", "ADS"])
        # Coluna A como texto.
        self.assertEqual(aba.cell(row=2, column=1).number_format, "@")
        # Autofiltro cobre A1:D...
        self.assertTrue(aba.auto_filter.ref.startswith("A1:D"))
        wb.close()

    def test_prefixo_respeita_limite_31(self) -> None:
        lotes = [[("+5547999990001", "Ana", "", "ADS")]]
        prefixo = "Curso_Muito_Longo_Que_Ultrapassa_Limite_DDD_47"
        caminho = gerar_excel_revisao_v32(lotes, self.dir, prefixo_aba=prefixo)
        wb = load_workbook(caminho)
        for aba in wb.worksheets:
            self.assertLessEqual(len(aba.title), 31)
            # A 1ª linha continua sendo o cabeçalho de colunas.
            self.assertEqual([c.value for c in aba[1]], list(COLUNAS_BLIP_V32))
        wb.close()


# ---------------------------------------------------------------------------
# colunas_visiveis + validar_coluna_curso_planilhas (Requisito D)
# ---------------------------------------------------------------------------
class ColunasVisiveisTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_le_cabecalhos_visiveis(self) -> None:
        caminho = self.dir / "p.xlsx"
        _escrever_planilha(caminho, ["Nome", "Telefone", "Curso"], [["A", "1", "X"]])
        colunas = colunas_visiveis(caminho)
        self.assertIn("Curso", colunas)
        self.assertIn("Nome", colunas)

    def test_ignora_abas_ocultas(self) -> None:
        caminho = self.dir / "p.xlsx"
        wb = Workbook()
        visivel = wb.active
        visivel.title = "Visivel"
        visivel.append(["Nome", "Telefone"])
        oculta = wb.create_sheet("Oculta")
        oculta.append(["Curso"])
        oculta.sheet_state = "hidden"
        wb.save(caminho)
        colunas = colunas_visiveis(caminho)
        self.assertNotIn("Curso", colunas)
        self.assertIn("Nome", colunas)


class ValidarColunaCursoPlanilhasTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_todas_com_coluna_passa(self) -> None:
        p1 = self.dir / "a.xlsx"
        p2 = self.dir / "b.xlsx"
        _escrever_planilha(p1, ["Nome", "Curso"], [["A", "X"]])
        _escrever_planilha(p2, ["Nome", "Nome do Produto"], [["B", "Y"]])
        # Não deve levantar.
        validar_coluna_curso_planilhas([p1, p2])

    def test_uma_sem_coluna_interrompe(self) -> None:
        p1 = self.dir / "a.xlsx"
        p2 = self.dir / "b.xlsx"
        _escrever_planilha(p1, ["Nome", "Curso"], [["A", "X"]])
        _escrever_planilha(p2, ["Nome", "Telefone"], [["B", "1"]])
        with self.assertRaises(CursoColunaError) as ctx:
            validar_coluna_curso_planilhas([p1, p2])
        self.assertIn("b.xlsx", str(ctx.exception))

    def test_varias_problematicas_mensagem_unica(self) -> None:
        p1 = self.dir / "a.xlsx"
        p2 = self.dir / "b.xlsx"
        _escrever_planilha(p1, ["Nome", "Telefone"], [["A", "1"]])
        _escrever_planilha(p2, ["Nome", "Idade"], [["B", "20"]])
        with self.assertRaises(CursoColunaError) as ctx:
            validar_coluna_curso_planilhas([p1, p2])
        mensagem = str(ctx.exception)
        self.assertIn("a.xlsx", mensagem)
        self.assertIn("b.xlsx", mensagem)

    def test_ambiguidade_lista_candidatas(self) -> None:
        p1 = self.dir / "a.xlsx"
        _escrever_planilha(p1, ["Curso", "Curso de Interesse"], [["X", "Y"]])
        with self.assertRaises(CursoColunaError) as ctx:
            validar_coluna_curso_planilhas([p1])
        self.assertIn("a.xlsx", str(ctx.exception))


# ---------------------------------------------------------------------------
# Exportação completa curso -> DDD
# ---------------------------------------------------------------------------
class PrepararBlipPorCursoEDddTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_estrutura_pastas_curso_ddd(self) -> None:
        df = _canonico_v32(
            ["5547999990001", "5549999990001"],
            nomes=["Ana", "Bruno"],
            emails=["ana@x.com", "bruno@x.com"],
            cursos_col=["ADS", "ADS"],
        )
        resumo = preparar_blip_por_curso_e_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        self.assertTrue((raiz / "ADS" / GRUPO_DDD_47).is_dir())
        self.assertTrue((raiz / "ADS" / GRUPO_DDD_49).is_dir())
        # Grupos sem contatos não criam pasta.
        self.assertFalse((raiz / "ADS" / GRUPO_SEM_DDD).exists())
        self.assertEqual(resumo["total_contatos"], 2)
        self.assertEqual(resumo["total_cursos_identificados"], 1)

    def test_csv_tem_exatamente_duas_colunas(self) -> None:
        df = _canonico_v32(
            ["5547999990001"],
            nomes=["Ana"],
            emails=["ana@x.com"],
            cursos_col=["ADS"],
        )
        preparar_blip_por_curso_e_ddd(df, self.dir)
        csv_path = next(
            (self.dir / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47).glob(
                f"{PREFIXO_CSV}*.csv"
            )
        )
        with open(csv_path, "r", encoding="utf-8", newline="") as arq:
            linhas = list(csv.reader(arq))
        self.assertEqual(linhas[0], list(COLUNAS_BLIP))
        self.assertEqual(linhas[1], ["+5547999990001", "Ana"])
        for linha in linhas:
            self.assertEqual(len(linha), 2)

    def test_excel_revisao_quatro_colunas(self) -> None:
        df = _canonico_v32(
            ["5547999990001"],
            nomes=["Ana"],
            emails=["ana@x.com"],
            cursos_col=["Direito"],
        )
        preparar_blip_por_curso_e_ddd(df, self.dir)
        excel = (
            self.dir / SUBPASTA_BLIP / "Direito" / GRUPO_DDD_47 / NOME_EXCEL_REVISAO
        )
        wb = load_workbook(excel)
        aba = wb.worksheets[0]
        self.assertEqual([c.value for c in aba[1]], list(COLUNAS_BLIP_V32))
        self.assertEqual(aba.cell(row=2, column=3).value, "ana@x.com")
        self.assertEqual(aba.cell(row=2, column=4).value, "Direito")
        wb.close()

    def test_nao_mistura_cursos(self) -> None:
        df = _canonico_v32(
            ["5547999990001", "5547999990002"],
            nomes=["Ana", "Bruno"],
            cursos_col=["ADS", "Direito"],
        )
        preparar_blip_por_curso_e_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP
        csv_ads = next((raiz / "ADS" / GRUPO_DDD_47).glob(f"{PREFIXO_CSV}*.csv"))
        csv_dir = next((raiz / "Direito" / GRUPO_DDD_47).glob(f"{PREFIXO_CSV}*.csv"))
        with open(csv_ads, encoding="utf-8", newline="") as a:
            nomes_ads = [l[1] for l in list(csv.reader(a))[1:]]
        with open(csv_dir, encoding="utf-8", newline="") as a:
            nomes_dir = [l[1] for l in list(csv.reader(a))[1:]]
        self.assertEqual(nomes_ads, ["Ana"])
        self.assertEqual(nomes_dir, ["Bruno"])

    def test_numeracao_reinicia_por_combinacao(self) -> None:
        df = _canonico_v32(
            ["5547999990001", "5549999990001"],
            cursos_col=["ADS", "ADS"],
        )
        preparar_blip_por_curso_e_ddd(df, self.dir)
        raiz = self.dir / SUBPASTA_BLIP / "ADS"
        self.assertTrue((raiz / GRUPO_DDD_47 / f"{PREFIXO_CSV}001.csv").exists())
        self.assertTrue((raiz / GRUPO_DDD_49 / f"{PREFIXO_CSV}001.csv").exists())

    def test_limite_bytes_respeitado(self) -> None:
        telefones = [f"55479999{i:05d}" for i in range(120)]
        df = _canonico_v32(
            telefones,
            nomes=[f"Contato {i}" for i in range(120)],
            cursos_col=["ADS"] * 120,
        )
        preparar_blip_por_curso_e_ddd(df, self.dir)
        pasta = self.dir / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47
        for csv_path in pasta.glob(f"{PREFIXO_CSV}*.csv"):
            self.assertLessEqual(csv_path.stat().st_size, MAX_BYTES_POR_ARQUIVO)
            with open(csv_path, encoding="utf-8", newline="") as arq:
                contatos = list(csv.reader(arq))[1:]
            self.assertLessEqual(len(contatos), MAX_CONTATOS_POR_ARQUIVO)

    def test_colisao_de_pasta_desambigua(self) -> None:
        # Dois cursos diferentes que sanitizam para o mesmo nome.
        df = _canonico_v32(
            ["5547999990001", "5547999990002"],
            cursos_col=["Curso: X", 'Curso" X'],
        )
        resumo = preparar_blip_por_curso_e_ddd(df, self.dir)
        nomes = set(resumo["nomes_pasta_curso"].values())
        self.assertEqual(len(nomes), 2)


class GerarBlipPorCursoEDddDePlanilhasTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_validacao_interrompe_sem_gerar_arquivo(self) -> None:
        bom = self.dir / "bom.xlsx"
        ruim = self.dir / "ruim.xlsx"
        _escrever_planilha(
            bom,
            ["Nome", "Telefone Celular", "Curso"],
            [["Ana", "5547999990001", "ADS"]],
        )
        _escrever_planilha(
            ruim, ["Nome", "Telefone Celular"], [["Bruno", "5549999990001"]]
        )
        saida = self.dir / "saida"
        with self.assertRaises(CursoColunaError):
            gerar_blip_por_curso_e_ddd_de_planilhas([bom, ruim], saida)
        # Nenhum arquivo parcial gerado.
        self.assertFalse((saida / SUBPASTA_BLIP).exists())

    def test_contato_sem_valor_de_curso_nao_interrompe(self) -> None:
        planilha = self.dir / "p.xlsx"
        _escrever_planilha(
            planilha,
            ["Nome", "Telefone Celular", "Curso"],
            [["Ana", "5547999990001", ""], ["Bruno", "5549999990001", "ADS"]],
        )
        saida = self.dir / "saida"
        resumo = gerar_blip_por_curso_e_ddd_de_planilhas([planilha], saida)
        self.assertIn(cursos.CURSO_NAO_IDENTIFICADO, resumo["contatos_por_curso"])
        self.assertIn("metricas_processamento", resumo)


# ---------------------------------------------------------------------------
# Reimportação v3.2: leitura + detecção de formato
# ---------------------------------------------------------------------------
class DetectarFormatoRevisaoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_detecta_v31(self) -> None:
        caminho = self.dir / "v31.xlsx"
        _escrever_planilha(
            caminho, [COLUNA_ID, COLUNA_NOME], [["+5547999990001", "Ana"]]
        )
        self.assertEqual(detectar_formato_revisao(caminho), "v31")

    def test_detecta_v32(self) -> None:
        caminho = self.dir / "v32.xlsx"
        _escrever_planilha(
            caminho,
            list(COLUNAS_BLIP_V32),
            [["+5547999990001", "Ana", "a@x.com", "ADS"]],
        )
        self.assertEqual(detectar_formato_revisao(caminho), "v32")

    def test_formato_desconhecido_levanta(self) -> None:
        caminho = self.dir / "x.xlsx"
        _escrever_planilha(caminho, ["Foo", "Bar"], [["1", "2"]])
        with self.assertRaises(blip.BlipError):
            detectar_formato_revisao(caminho)


class LerExcelRevisaoV32Test(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_le_quatro_colunas(self) -> None:
        caminho = self.dir / "v32.xlsx"
        _escrever_planilha(
            caminho,
            list(COLUNAS_BLIP_V32),
            [["+5547999990001", "Ana", "a@x.com", "ADS"]],
        )
        registros = ler_excel_revisao_v32(caminho)
        self.assertEqual(registros, [("+5547999990001", "Ana", "a@x.com", "ADS")])

    def test_cabecalho_invalido_levanta(self) -> None:
        caminho = self.dir / "ruim.xlsx"
        _escrever_planilha(
            caminho, [COLUNA_ID, COLUNA_NOME], [["+5547999990001", "Ana"]]
        )
        with self.assertRaises(blip.BlipError):
            ler_excel_revisao_v32(caminho)


class GerarCsvsPorCursoEDddDeRevisaoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _revisao(self, pasta: Path, linhas: list[list[str]]) -> Path:
        pasta.mkdir(parents=True, exist_ok=True)
        caminho = pasta / NOME_EXCEL_REVISAO
        _escrever_planilha(caminho, list(COLUNAS_BLIP_V32), linhas)
        return caminho

    def test_telefone_alterado_muda_ddd(self) -> None:
        # Origem: Blip/ADS/DDD_47, mas telefone é de DDD 49 agora.
        fonte_pasta = self.dir / "orig" / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47
        caminho = self._revisao(
            fonte_pasta, [["+5549999990001", "Ana", "a@x.com", "ADS"]]
        )
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_curso_e_ddd_de_revisao([caminho], saida)
        self.assertTrue(
            (saida / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_49 / f"{PREFIXO_CSV}001.csv").exists()
        )
        self.assertEqual(resumo["mudaram_de_ddd"], 1)

    def test_curso_apagado_vai_para_nao_identificado(self) -> None:
        fonte_pasta = self.dir / "orig" / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47
        caminho = self._revisao(
            fonte_pasta, [["+5547999990001", "Ana", "a@x.com", ""]]
        )
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_curso_e_ddd_de_revisao([caminho], saida)
        self.assertTrue(
            (
                saida
                / SUBPASTA_BLIP
                / cursos.CURSO_NAO_IDENTIFICADO
                / GRUPO_DDD_47
                / f"{PREFIXO_CSV}001.csv"
            ).exists()
        )
        self.assertEqual(resumo["foram_para_curso_nao_identificado"], 1)

    def test_identificador_invalido_vai_ao_relatorio(self) -> None:
        fonte_pasta = self.dir / "orig" / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47
        caminho = self._revisao(
            fonte_pasta, [["", "Ana", "a@x.com", "ADS"]]
        )
        saida = self.dir / "saida"
        resumo = gerar_csvs_por_curso_e_ddd_de_revisao([caminho], saida)
        self.assertEqual(resumo["total_invalidos"], 1)

    def test_originais_preservados(self) -> None:
        fonte_pasta = self.dir / "orig" / SUBPASTA_BLIP / "ADS" / GRUPO_DDD_47
        caminho = self._revisao(
            fonte_pasta, [["+5547999990001", "Ana", "a@x.com", "ADS"]]
        )
        mtime_antes = caminho.stat().st_mtime
        bytes_antes = caminho.read_bytes()
        saida = self.dir / "saida"
        gerar_csvs_por_curso_e_ddd_de_revisao([caminho], saida)
        self.assertEqual(caminho.stat().st_mtime, mtime_antes)
        self.assertEqual(caminho.read_bytes(), bytes_antes)


if __name__ == "__main__":
    unittest.main()
