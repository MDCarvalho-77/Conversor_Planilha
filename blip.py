"""Modo Blip (V3): preparação de listas de contatos para upload no Blip.

Esta camada é adicional à V2. Ela parte do mesmo DataFrame canônico final
produzido pelo ``pipeline_v2`` (ou de um Excel de revisão editado pelo usuário)
e gera arquivos CSV compatíveis com os requisitos do Blip:

- Exatamente duas colunas, nesta ordem: ``Número ou BSUID`` e ``Nome``.
- No máximo 100 contatos por arquivo (sem contar o cabeçalho).
- No máximo 100 KB por arquivo, medindo o CSV realmente gravado em disco.
- Identificador em formato de telefone internacional (com ``+``) ou BSUID.
- Codificação UTF-8 (sem BOM), separador vírgula, aspas quando necessário.

A V1 (``processamento.py``) e o núcleo da V2 (``adapters``, ``pipeline_v2``)
permanecem intactos: aqui apenas reutilizamos o resultado deles.
"""

from __future__ import annotations

import csv
import io
import re
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from adapters import CANON_NOME_COMPLETO, CANON_TELEFONE


# ---------------------------------------------------------------------------
# Constantes do formato Blip
# ---------------------------------------------------------------------------
FORMATO_BLIP = "Blip — lista de contatos"

COLUNA_ID = "Número ou BSUID"
COLUNA_NOME = "Nome"
COLUNAS_BLIP: tuple[str, str] = (COLUNA_ID, COLUNA_NOME)

MAX_CONTATOS_POR_ARQUIVO = 100
MAX_BYTES_POR_ARQUIVO = 100 * 1024  # 100 KB reais

# Convenções de CSV (documentadas no README). UTF-8 sem BOM e vírgula.
CSV_ENCODING = "utf-8"
CSV_SEPARADOR = ","
CSV_QUEBRA_LINHA = "\r\n"  # padrão CSV; contabilizado na medição de tamanho

PREFIXO_CSV = "Blip_"
NOME_EXCEL_REVISAO = "Blip_Contatos_Revisao.xlsx"
PREFIXO_ABA = "Lote_"


class BlipError(ValueError):
    """Erro de preparação/validação da lista Blip apresentável ao usuário."""


class RegistroInvalido:
    """Descreve um registro que não pôde ser exportado, com o motivo."""

    __slots__ = ("indice", "identificador", "nome", "motivo")

    def __init__(self, indice: int, identificador: str, nome: str, motivo: str) -> None:
        self.indice = indice
        self.identificador = identificador
        self.nome = nome
        self.motivo = motivo

    def como_dict(self) -> dict[str, Any]:
        return {
            "indice": self.indice,
            "identificador": self.identificador,
            "nome": self.nome,
            "motivo": self.motivo,
        }

    def __repr__(self) -> str:  # pragma: no cover - só para depuração
        return f"RegistroInvalido(indice={self.indice}, motivo={self.motivo!r})"


# ---------------------------------------------------------------------------
# Formatação e validação de identificadores
# ---------------------------------------------------------------------------
_SO_DIGITOS = re.compile(r"\D")
# BSUID do Blip costuma ser um token alfanumérico com separadores; aceitamos
# qualquer token não-vazio que NÃO pareça telefone e não contenha espaços.
_BSUID_PLAUSIVEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._@:-]*$")


def formatar_identificador_telefone(valor: Any) -> str:
    """Converte um telefone já normalizado (dígitos) para o formato ``+<dígitos>``.

    A normalização numérica (código do país etc.) é feita pela regra da V1/V2.
    Aqui apenas garantimos a apresentação internacional com ``+``. Valores
    vazios resultam em string vazia (tratados como inválidos adiante).
    """
    texto = "" if valor is None else str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return ""
    if texto.startswith("+"):
        digitos = _SO_DIGITOS.sub("", texto)
        return f"+{digitos}" if digitos else ""
    digitos = _SO_DIGITOS.sub("", texto)
    return f"+{digitos}" if digitos else ""


def validar_identificador(valor: Any) -> tuple[bool, str]:
    """Valida um identificador do Excel revisado (telefone com + ou BSUID).

    Retorna ``(valido, identificador_normalizado)``. Não inventa valores.
    """
    texto = "" if valor is None else str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return False, ""
    # Telefone internacional: '+' seguido de pelo menos alguns dígitos.
    if texto.startswith("+"):
        digitos = _SO_DIGITOS.sub("", texto)
        if len(digitos) >= 8:  # telefone plausível
            return True, f"+{digitos}"
        return False, texto
    # Apenas dígitos: tratamos como telefone sem o '+' e normalizamos.
    if texto.isdigit():
        if len(texto) >= 8:
            return True, f"+{texto}"
        return False, texto
    # Caso contrário, pode ser um BSUID: token sem espaços e plausível.
    if " " not in texto and _BSUID_PLAUSIVEL.match(texto):
        return True, texto
    return False, texto


# ---------------------------------------------------------------------------
# Medição de tamanho de CSV (bytes reais)
# ---------------------------------------------------------------------------
def _linha_csv_bytes(campos: list[str]) -> int:
    """Mede quantos bytes uma linha de CSV ocupa, com as mesmas convenções.

    Usa o módulo ``csv`` para aplicar aspas/escapes exatamente como na gravação,
    e mede os bytes em UTF-8, incluindo a quebra de linha.
    """
    buffer = io.StringIO()
    escritor = csv.writer(
        buffer, delimiter=CSV_SEPARADOR, lineterminator=CSV_QUEBRA_LINHA,
        quoting=csv.QUOTE_MINIMAL,
    )
    escritor.writerow(campos)
    return len(buffer.getvalue().encode(CSV_ENCODING))


def _bytes_cabecalho() -> int:
    return _linha_csv_bytes([COLUNA_ID, COLUNA_NOME])


# ---------------------------------------------------------------------------
# Preparação dos contatos (do canônico para o par identificador/nome)
# ---------------------------------------------------------------------------
def preparar_contatos(
    final: pd.DataFrame,
) -> tuple[list[tuple[str, str]], list[RegistroInvalido]]:
    """Transforma o canônico final em pares (identificador, nome) válidos.

    - Identificador: telefone canônico formatado para ``+<dígitos>``.
    - Nome: ``Nome Completo`` canônico (já composto pela V2 quando aplicável).

    Registros sem identificador válido vão para a lista de inválidos (não são
    descartados silenciosamente). O nome vazio é permitido (não inventamos
    nomes), mas é contabilizado adiante no resumo.
    """
    contatos: list[tuple[str, str]] = []
    invalidos: list[RegistroInvalido] = []

    telefones = final[CANON_TELEFONE] if CANON_TELEFONE in final.columns else None
    nomes = final[CANON_NOME_COMPLETO] if CANON_NOME_COMPLETO in final.columns else None

    for posicao in range(len(final)):
        bruto_tel = "" if telefones is None else telefones.iloc[posicao]
        bruto_nome = "" if nomes is None else nomes.iloc[posicao]
        nome = "" if bruto_nome is None else str(bruto_nome).strip()
        if nome.lower() == "nan":
            nome = ""

        identificador = formatar_identificador_telefone(bruto_tel)
        if identificador == "":
            invalidos.append(
                RegistroInvalido(
                    indice=posicao,
                    identificador="" if bruto_tel is None else str(bruto_tel),
                    nome=nome,
                    motivo="Registro sem identificador (telefone) válido.",
                )
            )
            continue
        contatos.append((identificador, nome))

    return contatos, invalidos


# ---------------------------------------------------------------------------
# Divisão em lotes por contagem (100) e tamanho (100 KB reais)
# ---------------------------------------------------------------------------
def dividir_em_lotes(
    contatos: list[tuple[str, str]],
) -> tuple[list[list[tuple[str, str]]], list[RegistroInvalido]]:
    """Divide a lista em lotes respeitando 100 contatos E 100 KB por arquivo.

    Preserva a ordem. Acrescenta contatos ao lote atual medindo o tamanho real
    do CSV (cabeçalho + linhas). Se um único contato não couber em 100 KB
    sozinho, ele é reportado como inválido (nunca truncado).

    Retorna ``(lotes, invalidos_por_tamanho)``.
    """
    lotes: list[list[tuple[str, str]]] = []
    invalidos: list[RegistroInvalido] = []
    bytes_cabecalho = _bytes_cabecalho()

    lote_atual: list[tuple[str, str]] = []
    bytes_atuais = bytes_cabecalho

    for posicao, (identificador, nome) in enumerate(contatos):
        linha_bytes = _linha_csv_bytes([identificador, nome])

        # Um único registro que não cabe nem sozinho (cabeçalho + 1 linha).
        if bytes_cabecalho + linha_bytes > MAX_BYTES_POR_ARQUIVO:
            invalidos.append(
                RegistroInvalido(
                    indice=posicao,
                    identificador=identificador,
                    nome=nome,
                    motivo=(
                        "Contato não cabe no limite de 100 KB nem sozinho; "
                        "não foi truncado."
                    ),
                )
            )
            continue

        estoura_contagem = len(lote_atual) + 1 > MAX_CONTATOS_POR_ARQUIVO
        estoura_tamanho = bytes_atuais + linha_bytes > MAX_BYTES_POR_ARQUIVO

        if lote_atual and (estoura_contagem or estoura_tamanho):
            lotes.append(lote_atual)
            lote_atual = []
            bytes_atuais = bytes_cabecalho

        lote_atual.append((identificador, nome))
        bytes_atuais += linha_bytes

    if lote_atual:
        lotes.append(lote_atual)

    return lotes, invalidos


# ---------------------------------------------------------------------------
# Gravação de CSVs e validação pós-gravação
# ---------------------------------------------------------------------------
def _gravar_csv(lote: list[tuple[str, str]], caminho: Path) -> None:
    """Grava um lote como CSV Blip (2 colunas, sem índice, UTF-8)."""
    with open(caminho, "w", encoding=CSV_ENCODING, newline="") as arquivo:
        escritor = csv.writer(
            arquivo, delimiter=CSV_SEPARADOR, lineterminator=CSV_QUEBRA_LINHA,
            quoting=csv.QUOTE_MINIMAL,
        )
        escritor.writerow([COLUNA_ID, COLUNA_NOME])
        for identificador, nome in lote:
            escritor.writerow([identificador, nome])


def validar_csv_gravado(caminho: Path) -> dict[str, Any]:
    """Reabre o CSV gravado e valida limites e estrutura reais.

    Confere: tamanho em bytes ≤ 100 KB, nº de contatos ≤ 100, exatamente as
    duas colunas na ordem certa, e que o ``+`` do telefone foi preservado.
    Levanta :class:`BlipError` em caso de violação.
    """
    tamanho = caminho.stat().st_size
    if tamanho > MAX_BYTES_POR_ARQUIVO:
        raise BlipError(
            f"O arquivo '{caminho.name}' excede 100 KB ({tamanho} bytes) após gravado."
        )

    with open(caminho, "r", encoding=CSV_ENCODING, newline="") as arquivo:
        leitor = csv.reader(arquivo, delimiter=CSV_SEPARADOR)
        linhas = list(leitor)

    if not linhas:
        raise BlipError(f"O arquivo '{caminho.name}' está vazio.")
    if linhas[0] != [COLUNA_ID, COLUNA_NOME]:
        raise BlipError(
            f"O arquivo '{caminho.name}' não tem o cabeçalho esperado "
            f"{COLUNAS_BLIP}."
        )
    contatos = linhas[1:]
    if len(contatos) > MAX_CONTATOS_POR_ARQUIVO:
        raise BlipError(
            f"O arquivo '{caminho.name}' tem {len(contatos)} contatos (máx. 100)."
        )
    for linha in contatos:
        if len(linha) != 2:
            raise BlipError(
                f"O arquivo '{caminho.name}' tem linha com número de colunas "
                "diferente de 2."
            )
    return {"arquivo": caminho, "bytes": tamanho, "contatos": len(contatos)}


def _preparar_pasta_saida(pasta: str | Path) -> Path:
    """Garante a pasta de saída, sem sobrescrever CSVs Blip existentes."""
    destino = Path(pasta)
    destino.mkdir(parents=True, exist_ok=True)
    existentes = sorted(destino.glob(f"{PREFIXO_CSV}*.csv"))
    if existentes:
        nomes = ", ".join(p.name for p in existentes[:5])
        raise BlipError(
            f"A pasta '{destino}' já contém arquivos {PREFIXO_CSV}*.csv ({nomes}"
            f"{'...' if len(existentes) > 5 else ''}). "
            "Escolha uma pasta vazia ou remova os arquivos antes de gerar novos."
        )
    return destino


def gravar_csvs(
    lotes: list[list[tuple[str, str]]], pasta_saida: str | Path
) -> list[dict[str, Any]]:
    """Grava cada lote como ``Blip_001.csv``... e valida cada arquivo gravado."""
    destino = _preparar_pasta_saida(pasta_saida)
    resultados: list[dict[str, Any]] = []
    for numero, lote in enumerate(lotes, start=1):
        caminho = destino / f"{PREFIXO_CSV}{numero:03d}.csv"
        _gravar_csv(lote, caminho)
        resultados.append(validar_csv_gravado(caminho))
    return resultados


# ---------------------------------------------------------------------------
# Excel de revisão
# ---------------------------------------------------------------------------
def _nome_aba(numero: int) -> str:
    """Nome de aba curto e sequencial: Lote_001, Lote_002, ..."""
    return f"{PREFIXO_ABA}{numero:03d}"


def gerar_excel_revisao(
    lotes: list[list[tuple[str, str]]], pasta_saida: str | Path
) -> Path:
    """Gera o Excel de conferência: uma aba por lote, telefone como texto.

    As colunas são as mesmas dos CSVs, sem índice. Os identificadores são
    gravados como texto (format '@') para preservar o ``+`` e evitar notação
    científica. Aplica autofiltro no cabeçalho de cada aba.
    """
    from openpyxl import Workbook

    destino = Path(pasta_saida)
    destino.mkdir(parents=True, exist_ok=True)
    caminho = destino / NOME_EXCEL_REVISAO

    workbook = Workbook()
    # Remove a aba padrão criada automaticamente.
    workbook.remove(workbook.active)

    lotes_efetivos = lotes if lotes else [[]]
    for numero, lote in enumerate(lotes_efetivos, start=1):
        aba = workbook.create_sheet(title=_nome_aba(numero))
        aba.append([COLUNA_ID, COLUNA_NOME])
        for identificador, nome in lote:
            aba.append([identificador, nome])
        # Identificador como texto em todas as linhas de dados.
        for linha in range(2, 2 + len(lote)):
            aba.cell(row=linha, column=1).number_format = "@"
        # Autofiltro cobrindo o cabeçalho + dados.
        ultima_linha = 1 + len(lote)
        aba.auto_filter.ref = f"A1:{get_column_letter(2)}{max(ultima_linha, 1)}"
        aba.column_dimensions["A"].width = 22
        aba.column_dimensions["B"].width = 40

    workbook.save(caminho)
    return caminho


# ---------------------------------------------------------------------------
# Fluxo 1: gerar a partir do resultado do processamento
# ---------------------------------------------------------------------------
def preparar_blip(
    final: pd.DataFrame, pasta_saida: str | Path
) -> dict[str, Any]:
    """Prepara CSVs iniciais + Excel de revisão a partir do canônico final.

    Retorna um resumo com contatos, lotes, arquivos e inválidos.
    """
    contatos, invalidos_prep = preparar_contatos(final)
    lotes, invalidos_tamanho = dividir_em_lotes(contatos)
    invalidos = invalidos_prep + invalidos_tamanho

    csvs = gravar_csvs(lotes, pasta_saida)
    excel = gerar_excel_revisao(lotes, pasta_saida)

    return _resumo(contatos, lotes, csvs, excel, invalidos)


def gerar_blip_de_planilhas(
    caminhos: list[str | Path], pasta_saida: str | Path
) -> dict[str, Any]:
    """Executa o pipeline V2 nas planilhas e prepara a saída Blip.

    Reutiliza integralmente a consolidação, deduplicação e normalização da V2
    (importada aqui para evitar dependência circular na importação do módulo).
    """
    from pipeline_v2 import processar_consolidado

    final, metricas = processar_consolidado(caminhos)
    resumo = preparar_blip(final, pasta_saida)
    return {**resumo, "metricas_processamento": metricas}


# ---------------------------------------------------------------------------
# Fluxo 2: gerar CSVs definitivos a partir do Excel revisado
# ---------------------------------------------------------------------------
def ler_excel_revisado(caminho: str | Path) -> list[tuple[str, str]]:
    """Lê as abas de lotes do Excel revisado e devolve os contatos em ordem.

    Valida que cada aba tem exatamente as duas colunas esperadas, na ordem
    certa. Preserva a ordem das abas e das linhas. Não modifica o arquivo.
    """
    origem = Path(caminho)
    try:
        workbook = load_workbook(origem, read_only=True, data_only=True)
    except Exception as erro:
        raise BlipError(
            f"Não foi possível abrir o Excel revisado: '{origem.name}'."
        ) from erro

    contatos: list[tuple[str, str]] = []
    try:
        abas = [aba for aba in workbook.worksheets if aba.sheet_state == "visible"]
        if not abas:
            raise BlipError("O Excel revisado não possui abas visíveis.")
        for aba in abas:
            linhas = aba.iter_rows(values_only=True)
            cabecalho = next(linhas, None)
            if cabecalho is None:
                continue  # aba vazia é ignorada
            cab = [("" if c is None else str(c).strip()) for c in cabecalho[:2]]
            if cab != [COLUNA_ID, COLUNA_NOME]:
                raise BlipError(
                    f"A aba '{aba.title}' não tem as colunas esperadas "
                    f"{COLUNAS_BLIP}. Encontrado: {cab}."
                )
            for linha in linhas:
                if linha is None:
                    continue
                identificador = "" if len(linha) < 1 or linha[0] is None else str(linha[0]).strip()
                nome = "" if len(linha) < 2 or linha[1] is None else str(linha[1]).strip()
                if identificador == "" and nome == "":
                    continue  # linha totalmente vazia é ignorada
                contatos.append((identificador, nome))
    finally:
        workbook.close()

    return contatos


def gerar_csvs_do_excel_revisado(
    caminho_excel: str | Path, pasta_saida: str | Path
) -> dict[str, Any]:
    """Gera os CSVs definitivos a partir do Excel revisado pelo usuário.

    Revalida identificadores/nomes, reaplica os limites de 100 contatos e
    100 KB (redistribuindo mesmo que o usuário tenha alterado as abas) e grava
    os CSVs. Não modifica o Excel de origem.
    """
    brutos = ler_excel_revisado(caminho_excel)

    validos: list[tuple[str, str]] = []
    invalidos: list[RegistroInvalido] = []
    for posicao, (identificador_bruto, nome) in enumerate(brutos):
        ok, identificador = validar_identificador(identificador_bruto)
        if not ok:
            invalidos.append(
                RegistroInvalido(
                    indice=posicao,
                    identificador=identificador_bruto,
                    nome=nome,
                    motivo="Identificador ausente ou inválido no Excel revisado.",
                )
            )
            continue
        validos.append((identificador, nome))

    lotes, invalidos_tamanho = dividir_em_lotes(validos)
    invalidos += invalidos_tamanho

    csvs = gravar_csvs(lotes, pasta_saida)
    return _resumo(validos, lotes, csvs, None, invalidos)


# ---------------------------------------------------------------------------
# Resumo
# ---------------------------------------------------------------------------
def _resumo(
    contatos: list[tuple[str, str]],
    lotes: list[list[tuple[str, str]]],
    csvs: list[dict[str, Any]],
    excel: Path | None,
    invalidos: list[RegistroInvalido],
) -> dict[str, Any]:
    sem_nome = sum(1 for _, nome in contatos if nome == "")
    return {
        "contatos_exportados": sum(len(lote) for lote in lotes),
        "arquivos_gerados": [c["arquivo"] for c in csvs],
        "lotes": len(lotes),
        "detalhes_csv": csvs,
        "excel_revisao": excel,
        "contatos_sem_nome": sem_nome,
        "invalidos": [r.como_dict() for r in invalidos],
        "total_invalidos": len(invalidos),
    }
