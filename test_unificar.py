"""Testes da V3.3 (REQUISITO A): modo Unificar lista — Somente telefones.

Cobre leitura multi-formato (.csv/.xlsx/misto), reconhecimento por NOME de
coluna (lista controlada, tolerante a acento/espaço/caixa), ausência e
ambiguidade -> TelefoneColunaError, dedup global preservando ordem, os 4 grupos
de DDD, internacional/BSUID, preservação do '+', Excel number_format '@', CSV
sem BOM, grupos vazios sem pasta, ausência de limites 100/100KB, relatório de
inválidos com identificação, interrupção total sem parcial, anti-sobrescrita
(False/True) e CSV em latin-1/utf-8-sig/separador ';'.
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from blip import (
    GRUPO_DDD_47,
    GRUPO_DDD_48_OUTROS,
    GRUPO_DDD_49,
    GRUPO_SEM_DDD,
    GRUPOS_DDD,
)
from unificar import (
    CABECALHO_RELATORIO_INVALIDOS,
    NOME_ARQUIVO_SAIDA,
    NOME_RELATORIO_INVALIDOS,
    ROTULOS_TELEFONE_UNIFICAR,
    SUBPASTA_UNIFICADA,
    SobrescritaError,
    TelefoneColunaError,
    unificar_telefones,
    validar_fontes,
)


def _tel_br(ddd: str, i: int) -> str:
    """Telefone BR móvel já com DDI 55 (13 dígitos), sem '+'."""
    return f"55{ddd}9{i:08d}"


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    # ---- helpers de montagem de arquivos ---------------------------------
    def _xlsx(self, nome: str, cabecalho: list[str], linhas: list[list]) -> Path:
        caminho = self.dir / nome
        wb = Workbook()
        aba = wb.active
        aba.append(cabecalho)
        for linha in linhas:
            aba.append(linha)
        wb.save(caminho)
        return caminho

    def _csv_bytes(self, nome: str, conteudo: bytes) -> Path:
        caminho = self.dir / nome
        caminho.write_bytes(conteudo)
        return caminho

    def _csv(
        self,
        nome: str,
        cabecalho: list[str],
        linhas: list[list[str]],
        sep: str = ",",
        encoding: str = "utf-8",
    ) -> Path:
        partes = [sep.join(cabecalho)]
        partes += [sep.join(linha) for linha in linhas]
        texto = "\r\n".join(partes) + "\r\n"
        return self._csv_bytes(nome, texto.encode(encoding))


class ReconhecimentoColunaTest(_Base):
    def test_cada_cabecalho_da_lista_controlada(self) -> None:
        for rotulo in ROTULOS_TELEFONE_UNIFICAR:
            with self.subTest(rotulo=rotulo):
                arq = self._xlsx("a.xlsx", [rotulo], [[_tel_br("47", 1)]])
                resumo = unificar_telefones([arq], self.dir / rotulo)
                self.assertEqual(resumo["total_unico"], 1)

    def test_variacoes_acento_espaco_caixa(self) -> None:
        # 'Número de Telefone' com acento/caixa diferentes deve ser reconhecido.
        arq = self._xlsx(
            "a.xlsx", ["  NÚMERO   DE  telefone "], [[_tel_br("47", 1)]]
        )
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)

    def test_ausencia_levanta_erro(self) -> None:
        arq = self._xlsx("semtel.xlsx", ["Nome", "Email"], [["João", "j@x.com"]])
        with self.assertRaises(TelefoneColunaError) as ctx:
            validar_fontes([arq])
        self.assertIn("semtel.xlsx", str(ctx.exception))

    def test_nao_reconhece_apenas_por_conter_numero(self) -> None:
        # 'Numero do Documento' NÃO está na lista controlada.
        arq = self._xlsx("a.xlsx", ["Numero do Documento"], [["12345678"]])
        with self.assertRaises(TelefoneColunaError):
            validar_fontes([arq])


class AmbiguidadeTest(_Base):
    def test_duas_candidatas_distintas_preenchidas_ambiguo(self) -> None:
        arq = self._xlsx(
            "ambi.xlsx",
            ["Celular", "Telefone Residencial"],
            [[_tel_br("47", 1), _tel_br("49", 1)]],
        )
        with self.assertRaises(TelefoneColunaError) as ctx:
            validar_fontes([arq])
        msg = str(ctx.exception)
        self.assertIn("ambígu", msg.lower())
        self.assertIn("Celular", msg)
        self.assertIn("Telefone Residencial", msg)

    def test_uma_candidata_preenchida_outra_vazia_ok(self) -> None:
        # Duas colunas candidatas, mas só uma preenchida -> não é ambíguo.
        arq = self._xlsx(
            "ok.xlsx",
            ["Celular", "Telefone Residencial"],
            [[_tel_br("47", 1), ""]],
        )
        validar_fontes([arq])  # não levanta
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)

    def test_erro_unico_lista_todos_os_arquivos(self) -> None:
        a1 = self._xlsx("sem1.xlsx", ["Nome"], [["x"]])
        a2 = self._xlsx(
            "ambi2.xlsx",
            ["Celular", "Telefone da Empresa"],
            [[_tel_br("47", 1), _tel_br("48", 1)]],
        )
        with self.assertRaises(TelefoneColunaError) as ctx:
            validar_fontes([a1, a2])
        msg = str(ctx.exception)
        self.assertIn("sem1.xlsx", msg)
        self.assertIn("ambi2.xlsx", msg)


class ConsolidacaoDedupTest(_Base):
    def test_dedup_global_entre_arquivos_preserva_ordem(self) -> None:
        a1 = self._xlsx(
            "a.xlsx",
            ["Telefone"],
            [[_tel_br("47", 1)], [_tel_br("49", 1)]],
        )
        a2 = self._csv(
            "b.csv",
            ["Telefone"],
            [[_tel_br("47", 1)], [_tel_br("48", 1)]],  # 47 repetido
        )
        resumo = unificar_telefones([a1, a2], self.dir / "out")
        # 3 únicos: 47, 49, 48 (o 47 do segundo arquivo é dedup).
        self.assertEqual(resumo["total_lido"], 4)
        self.assertEqual(resumo["total_unico"], 3)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_47], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_49], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_48_OUTROS], 1)

    def test_quatro_grupos(self) -> None:
        arq = self._xlsx(
            "a.xlsx",
            ["Telefone"],
            [
                [_tel_br("47", 1)],
                [_tel_br("49", 1)],
                [_tel_br("48", 1)],
                ["+351912345678"],  # internacional (12 díg.) -> SEM_DDD
            ],
        )
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_47], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_49], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_48_OUTROS], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_SEM_DDD], 1)
        self.assertEqual(resumo["total_sem_ddd"], 1)

    def test_internacional_vai_sem_ddd(self) -> None:
        arq = self._csv("a.csv", ["Telefone"], [["+351912345678"]])
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["por_grupo"][GRUPO_SEM_DDD], 1)

    def test_numero_curto_preserva_sem_prefixo(self) -> None:
        # 8 dígitos: _formatar_telefone não prefixa 55; vira +<digitos>.
        arq = self._csv("a.csv", ["Telefone"], [["12345678"]])
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_SEM_DDD], 1)


class SaidaArquivosTest(_Base):
    def test_preserva_mais_no_xlsx_como_texto(self) -> None:
        arq = self._xlsx("a.xlsx", ["Telefone"], [["47999990001"]])
        resumo = unificar_telefones([arq], self.dir / "out")
        xlsx = self.dir / "out" / SUBPASTA_UNIFICADA / GRUPO_DDD_47 / f"{NOME_ARQUIVO_SAIDA}.xlsx"
        self.assertIn(xlsx, resumo["caminhos_gerados"])
        wb = load_workbook(xlsx)
        aba = wb.active
        self.assertEqual(aba.cell(1, 1).value, "Telefone")
        self.assertEqual(aba.cell(2, 1).value, "+5547999990001")
        self.assertEqual(aba.cell(2, 1).number_format, "@")

    def test_csv_sem_bom_uma_coluna(self) -> None:
        arq = self._xlsx("a.xlsx", ["Telefone"], [[_tel_br("49", 1)]])
        unificar_telefones([arq], self.dir / "out")
        csv_path = self.dir / "out" / SUBPASTA_UNIFICADA / GRUPO_DDD_49 / f"{NOME_ARQUIVO_SAIDA}.csv"
        bruto = csv_path.read_bytes()
        self.assertFalse(bruto.startswith(b"\xef\xbb\xbf"))  # sem BOM
        with open(csv_path, "r", encoding="utf-8", newline="") as arquivo:
            linhas = list(csv.reader(arquivo))
        self.assertEqual(linhas[0], ["Telefone"])
        self.assertEqual(linhas[1], ["+5549900000001"])
        self.assertTrue(all(len(linha) == 1 for linha in linhas))

    def test_grupos_vazios_nao_criam_pasta(self) -> None:
        arq = self._xlsx("a.xlsx", ["Telefone"], [[_tel_br("47", 1)]])
        unificar_telefones([arq], self.dir / "out")
        raiz = self.dir / "out" / SUBPASTA_UNIFICADA
        self.assertTrue((raiz / GRUPO_DDD_47).is_dir())
        self.assertFalse((raiz / GRUPO_DDD_49).exists())
        self.assertFalse((raiz / GRUPO_DDD_48_OUTROS).exists())
        self.assertFalse((raiz / GRUPO_SEM_DDD).exists())

    def test_sem_limite_100_gera_arquivo_unico(self) -> None:
        linhas = [[_tel_br("47", i)] for i in range(150)]
        arq = self._xlsx("a.xlsx", ["Telefone"], linhas)
        unificar_telefones([arq], self.dir / "out")
        pasta = self.dir / "out" / SUBPASTA_UNIFICADA / GRUPO_DDD_47
        # Um único arquivo, sem Blip_002 nem divisão.
        self.assertTrue((pasta / f"{NOME_ARQUIVO_SAIDA}.csv").exists())
        with open(pasta / f"{NOME_ARQUIVO_SAIDA}.csv", "r", encoding="utf-8", newline="") as arquivo:
            linhas_lidas = list(csv.reader(arquivo))
        self.assertEqual(len(linhas_lidas), 151)  # cabeçalho + 150


class InvalidosTest(_Base):
    def test_invalido_vai_ao_relatorio_com_identificacao(self) -> None:
        arq = self._csv(
            "a.csv",
            ["Telefone"],
            [[_tel_br("47", 1)], ["sem numeros aqui"]],
        )
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_invalidos"], 1)
        invalido = resumo["invalidos"][0]
        self.assertEqual(invalido["valor_original"], "sem numeros aqui")
        self.assertEqual(invalido["arquivo"], "a.csv")
        self.assertEqual(invalido["linha"], 3)  # cabeçalho=1, 47=2, inválido=3
        self.assertEqual(resumo["total_unico"], 1)

    def test_relatorio_invalidos_persistido_com_identificacao(self) -> None:
        arq = self._csv(
            "a.csv",
            ["Telefone"],
            [[_tel_br("47", 1)], ["sem numeros aqui"]],
        )
        saida = self.dir / "out"
        resumo = unificar_telefones([arq], saida)
        relatorio = saida / SUBPASTA_UNIFICADA / f"{NOME_RELATORIO_INVALIDOS}.csv"
        # O resumo aponta o caminho REAL e o arquivo existe e está nos gerados.
        self.assertEqual(resumo["relatorio_invalidos"], relatorio)
        self.assertIn(relatorio, resumo["caminhos_gerados"])
        self.assertTrue(relatorio.is_file())
        self.assertFalse(relatorio.read_bytes().startswith(b"\xef\xbb\xbf"))  # sem BOM
        with open(relatorio, "r", encoding="utf-8", newline="") as arquivo:
            linhas = list(csv.reader(arquivo))
        self.assertEqual(linhas[0], list(CABECALHO_RELATORIO_INVALIDOS))
        self.assertEqual(linhas[1][0], "sem numeros aqui")  # valor_original
        self.assertEqual(linhas[1][1], "a.csv")  # arquivo
        self.assertEqual(linhas[1][2], "3")  # linha

    def test_sem_invalidos_nao_cria_relatorio(self) -> None:
        arq = self._csv("a.csv", ["Telefone"], [[_tel_br("47", 1)]])
        saida = self.dir / "out"
        resumo = unificar_telefones([arq], saida)
        self.assertIsNone(resumo["relatorio_invalidos"])
        relatorio = saida / SUBPASTA_UNIFICADA / f"{NOME_RELATORIO_INVALIDOS}.csv"
        self.assertFalse(relatorio.exists())


class InterrupcaoTotalTest(_Base):
    def test_validacao_nao_grava_nada_quando_falha(self) -> None:
        ok = self._xlsx("ok.xlsx", ["Telefone"], [[_tel_br("47", 1)]])
        ruim = self._xlsx("ruim.xlsx", ["Nome"], [["x"]])
        saida = self.dir / "out"
        with self.assertRaises(TelefoneColunaError):
            unificar_telefones([ok, ruim], saida)
        self.assertFalse((saida / SUBPASTA_UNIFICADA).exists())


class AntiSobrescritaTest(_Base):
    def test_sobrescrever_false_levanta_sem_parcial(self) -> None:
        arq = self._xlsx("a.xlsx", ["Telefone"], [[_tel_br("47", 1)], [_tel_br("49", 1)]])
        saida = self.dir / "out"
        unificar_telefones([arq], saida)  # primeira gravação
        # Segunda tentativa sem autorização: erro e nada regravado.
        with self.assertRaises(SobrescritaError) as ctx:
            unificar_telefones([arq], saida, sobrescrever=False)
        self.assertIn(NOME_ARQUIVO_SAIDA, str(ctx.exception))

    def test_sobrescrever_true_regrava(self) -> None:
        a1 = self._xlsx("a.xlsx", ["Telefone"], [[_tel_br("47", 1)]])
        saida = self.dir / "out"
        unificar_telefones([a1], saida)
        # Novo conjunto com 2 telefones 47, regravando.
        a2 = self._xlsx("b.xlsx", ["Telefone"], [[_tel_br("47", 1)], [_tel_br("47", 2)]])
        resumo = unificar_telefones([a2], saida, sobrescrever=True)
        csv_path = saida / SUBPASTA_UNIFICADA / GRUPO_DDD_47 / f"{NOME_ARQUIVO_SAIDA}.csv"
        with open(csv_path, "r", encoding="utf-8", newline="") as arquivo:
            linhas = list(csv.reader(arquivo))
        self.assertEqual(len(linhas), 3)  # cabeçalho + 2
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_47], 2)


class EncodingSeparadorTest(_Base):
    def test_csv_utf8_sig(self) -> None:
        arq = self._csv("a.csv", ["Telefone"], [[_tel_br("47", 1)]], encoding="utf-8-sig")
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)

    def test_csv_latin1(self) -> None:
        # Cabeçalho com acento gravado em latin-1.
        arq = self._csv(
            "a.csv", ["Número de Telefone"], [[_tel_br("49", 1)]], encoding="latin-1"
        )
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_49], 1)

    def test_csv_separador_ponto_e_virgula(self) -> None:
        arq = self._csv(
            "a.csv",
            ["Telefone", "Nome"],
            [[_tel_br("48", 1), "Fulano"]],
            sep=";",
        )
        resumo = unificar_telefones([arq], self.dir / "out")
        self.assertEqual(resumo["total_unico"], 1)
        self.assertEqual(resumo["por_grupo"][GRUPO_DDD_48_OUTROS], 1)


if __name__ == "__main__":
    unittest.main()
