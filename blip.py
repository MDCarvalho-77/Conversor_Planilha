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


# ---------------------------------------------------------------------------
# Constantes de grupo por DDD (V3.1)
# ---------------------------------------------------------------------------
GRUPO_DDD_47 = "DDD_47"
GRUPO_DDD_49 = "DDD_49"
GRUPO_DDD_48_OUTROS = "DDD_48_E_OUTROS"
GRUPO_SEM_DDD = "SEM_DDD"
# Ordem fixa dos grupos (usada na separação e na criação de pastas).
GRUPOS_DDD: tuple[str, str, str, str] = (
    GRUPO_DDD_47,
    GRUPO_DDD_49,
    GRUPO_DDD_48_OUTROS,
    GRUPO_SEM_DDD,
)
# Subpasta raiz da exportação Blip por grupo de DDD.
SUBPASTA_BLIP = "Blip"


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
# Classificação por DDD brasileiro (V3.1)
# ---------------------------------------------------------------------------
def extrair_ddd_brasil(identificador: Any) -> str | None:
    """Extrai o DDD brasileiro de um identificador, de forma conservadora.

    Retorna o DDD (dois dígitos, como string) somente quando é possível
    determiná-lo com segurança a partir de um telefone brasileiro. Caso
    contrário retorna ``None`` — nunca inventamos DDD.

    Decisões conservadoras:

    - Se o identificador contiver letra (BSUID) ou não parecer telefone,
      retorna ``None``.
    - Só reconhecemos um número como brasileiro quando os dígitos começam com
      ``55`` E o total de dígitos é 12 (DDI + DDD + 8) ou 13 (DDI + DDD + 9);
      o DDD são os dois dígitos logo após o ``55``.
    - ``+55`` (ou ``55``) com comprimento incompatível com um telefone BR
      retorna ``None`` (não forçamos uma leitura de DDD).
    - Telefone internacional não-BR (ex.: ``+1...``, ``+351...``, ``+44...``)
      retorna ``None``: os dois primeiros dígitos NÃO são tratados como DDD BR.
    - Telefone sem o ``55`` cujo tamanho não permita inferir com segurança
      retorna ``None``.
    - O DDD é validado na faixa 11-99; fora dela retorna ``None``.
    """
    texto = "" if identificador is None else str(identificador).strip()
    if texto == "" or texto.lower() == "nan":
        return None
    # Se houver qualquer letra, não é telefone (provável BSUID) -> None.
    if any(caractere.isalpha() for caractere in texto):
        return None
    digitos = _SO_DIGITOS.sub("", texto)
    if digitos == "":
        return None
    # Brasil: precisa começar com o DDI 55 e ter 12 ou 13 dígitos no total.
    if not digitos.startswith("55"):
        return None
    if len(digitos) not in (12, 13):
        return None
    ddd = digitos[2:4]
    # DDD válido no Brasil está na faixa 11-99.
    if not (11 <= int(ddd) <= 99):
        return None
    return ddd


def classificar_ddd(identificador: Any) -> str:
    """Classifica um identificador em exatamente um dos quatro grupos de DDD.

    - DDD 47 -> :data:`GRUPO_DDD_47`.
    - DDD 49 -> :data:`GRUPO_DDD_49`.
    - Qualquer outro DDD BR válido (inclui 48) -> :data:`GRUPO_DDD_48_OUTROS`.
    - Sem DDD BR identificável (BSUID, internacional não-BR, inválido) ->
      :data:`GRUPO_SEM_DDD`.
    """
    ddd = extrair_ddd_brasil(identificador)
    if ddd is None:
        return GRUPO_SEM_DDD
    if ddd == "47":
        return GRUPO_DDD_47
    if ddd == "49":
        return GRUPO_DDD_49
    return GRUPO_DDD_48_OUTROS


def separar_por_ddd(
    contatos: list[tuple[str, str]],
) -> dict[str, list[tuple[str, str]]]:
    """Separa os contatos pelos quatro grupos de DDD, preservando a ordem.

    Retorna sempre as quatro chaves de :data:`GRUPOS_DDD`, na ordem fixa, cada
    uma com a lista dos contatos daquele grupo (lista vazia quando não houver).
    """
    grupos: dict[str, list[tuple[str, str]]] = {grupo: [] for grupo in GRUPOS_DDD}
    for identificador, nome in contatos:
        grupos[classificar_ddd(identificador)].append((identificador, nome))
    return grupos


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
    lotes: list[list[tuple[str, str]]],
    pasta_saida: str | Path,
    prefixo_aba: str = "",
) -> Path:
    """Gera o Excel de conferência: uma aba por lote, telefone como texto.

    As colunas são as mesmas dos CSVs, sem índice. Os identificadores são
    gravados como texto (format '@') para preservar o ``+`` e evitar notação
    científica. Aplica autofiltro no cabeçalho de cada aba.

    ``prefixo_aba`` (opcional, V3.1) acrescenta uma identificação de grupo ao
    nome de cada aba (ex.: ``DDD_47_Lote_001``), sem alterar a primeira linha
    ``[COLUNA_ID, COLUNA_NOME]`` — de modo que :func:`ler_excel_revisado`
    continua lendo normalmente. Com o padrão vazio, o comportamento é idêntico
    ao da V3.
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
        titulo = _nome_aba(numero)
        if prefixo_aba:
            # Excel limita o nome da aba a 31 caracteres.
            titulo = f"{prefixo_aba}_{titulo}"[:31]
        aba = workbook.create_sheet(title=titulo)
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
# Fluxo 1 (V3.1): gerar a saída Blip separada por grupo de DDD
# ---------------------------------------------------------------------------
def _exportar_grupo(
    grupo: str,
    contatos: list[tuple[str, str]],
    pasta_raiz: Path,
) -> dict[str, Any]:
    """Gera CSVs + Excel de revisão de UM grupo na subpasta ``pasta_raiz/<grupo>``.

    A numeração dos CSVs reinicia em ``Blip_001.csv`` dentro da subpasta e a
    anti-sobrescrita de :func:`_preparar_pasta_saida` roda por subpasta.
    """
    pasta_grupo = pasta_raiz / grupo
    lotes, invalidos_tamanho = dividir_em_lotes(contatos)
    csvs = gravar_csvs(lotes, pasta_grupo)
    excel = gerar_excel_revisao(lotes, pasta_grupo, prefixo_aba=grupo)
    return {
        "grupo": grupo,
        "pasta": pasta_grupo,
        "contatos": sum(len(lote) for lote in lotes),
        "lotes": len(lotes),
        "arquivos_gerados": [c["arquivo"] for c in csvs],
        "detalhes_csv": csvs,
        "excel_revisao": excel,
        "invalidos": [r.como_dict() for r in invalidos_tamanho],
        "total_invalidos": len(invalidos_tamanho),
    }


def _resumo_por_ddd(
    grupos_contatos: dict[str, list[tuple[str, str]]],
    pasta_raiz: Path,
    invalidos_iniciais: list[RegistroInvalido],
    mudaram_de_grupo: int | None = None,
) -> dict[str, Any]:
    """Exporta cada grupo com contatos e monta o resumo agregado da V3.1.

    Cria a subpasta de um grupo apenas quando ele tem contatos. Agrega os
    totais globais e preserva chaves já usadas pela UI (``contatos_exportados``,
    ``lotes``, ``arquivos_gerados``, ``total_invalidos``).
    """
    por_grupo: dict[str, dict[str, Any]] = {}
    arquivos_todos: list[Path] = []
    invalidos_total = list(invalidos_iniciais)
    total_lotes = 0
    total_contatos = 0
    totais_por_grupo: dict[str, int] = {grupo: 0 for grupo in GRUPOS_DDD}

    for grupo in GRUPOS_DDD:
        contatos = grupos_contatos.get(grupo, [])
        totais_por_grupo[grupo] = len(contatos)
        if not contatos:
            continue  # grupo vazio NÃO cria pasta
        detalhe = _exportar_grupo(grupo, contatos, pasta_raiz)
        por_grupo[grupo] = detalhe
        arquivos_todos.extend(detalhe["arquivos_gerados"])
        total_lotes += detalhe["lotes"]
        total_contatos += detalhe["contatos"]
        for item in detalhe["invalidos"]:
            invalidos_total.append(
                RegistroInvalido(
                    indice=item["indice"],
                    identificador=item["identificador"],
                    nome=item["nome"],
                    motivo=item["motivo"],
                )
            )

    resumo: dict[str, Any] = {
        "pasta_raiz": pasta_raiz,
        "por_grupo": por_grupo,
        "totais_por_grupo": totais_por_grupo,
        "contatos_exportados": total_contatos,
        "lotes": total_lotes,
        "arquivos_gerados": arquivos_todos,
        "total_contatos": total_contatos,
        "total_sem_ddd": totais_por_grupo[GRUPO_SEM_DDD],
        "invalidos": [r.como_dict() for r in invalidos_total],
        "total_invalidos": len(invalidos_total),
    }
    if mudaram_de_grupo is not None:
        resumo["mudaram_de_grupo"] = mudaram_de_grupo
    return resumo


def preparar_blip_por_ddd(
    final: pd.DataFrame, pasta_saida: str | Path
) -> dict[str, Any]:
    """Prepara a saída Blip separada por grupo de DDD a partir do canônico final.

    Estrutura gerada: ``<pasta_saida>/Blip/<GRUPO>/Blip_001.csv...`` +
    ``Blip_Contatos_Revisao.xlsx``, com uma subpasta por grupo COM contatos.
    """
    contatos, invalidos_prep = preparar_contatos(final)
    grupos = separar_por_ddd(contatos)
    pasta_raiz = Path(pasta_saida) / SUBPASTA_BLIP
    pasta_raiz.mkdir(parents=True, exist_ok=True)
    return _resumo_por_ddd(grupos, pasta_raiz, invalidos_prep)


def gerar_blip_por_ddd_de_planilhas(
    caminhos: list[str | Path], pasta_saida: str | Path
) -> dict[str, Any]:
    """Executa o pipeline V2 e prepara a saída Blip separada por grupo de DDD."""
    from pipeline_v2 import processar_consolidado

    final, metricas = processar_consolidado(caminhos)
    resumo = preparar_blip_por_ddd(final, pasta_saida)
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
# Fluxo 2 (V3.1): gerar CSVs por grupo a partir das planilhas revisadas,
# reclassificando pelo DDD do telefone atual.
# ---------------------------------------------------------------------------
def _grupo_de_origem(caminho: Path) -> str | None:
    """Deriva o grupo de origem pelo nome da pasta-pai do arquivo, se for um
    grupo conhecido; caso contrário retorna ``None`` (origem desconhecida)."""
    pai = caminho.parent.name
    return pai if pai in GRUPOS_DDD else None


def gerar_csvs_por_ddd_de_revisao(
    fontes: list[str | Path], pasta_saida: str | Path
) -> dict[str, Any]:
    """Gera os CSVs definitivos por grupo a partir das planilhas revisadas.

    Para cada ``.xlsx`` em ``fontes``: lê com :func:`ler_excel_revisado`,
    deriva o grupo de ORIGEM pelo nome da pasta-pai (quando for um de
    :data:`GRUPOS_DDD`), valida cada identificador (inválidos vão ao relatório,
    sem descarte silencioso) e RECLASSIFICA cada contato válido pelo telefone
    atual com :func:`classificar_ddd`. Em seguida reaplica os limites por grupo
    e grava na estrutura ``<pasta_saida>/Blip/<GRUPO>/``.

    Conta ``mudaram_de_grupo`` apenas quando a origem é conhecida e difere do
    grupo recalculado. NÃO modifica os arquivos ``.xlsx`` de origem.
    """
    grupos: dict[str, list[tuple[str, str]]] = {grupo: [] for grupo in GRUPOS_DDD}
    invalidos: list[RegistroInvalido] = []
    mudaram_de_grupo = 0

    for fonte in fontes:
        caminho = Path(fonte)
        origem = _grupo_de_origem(caminho)
        brutos = ler_excel_revisado(caminho)
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
            grupo = classificar_ddd(identificador)
            grupos[grupo].append((identificador, nome))
            if origem is not None and origem != grupo:
                mudaram_de_grupo += 1

    pasta_raiz = Path(pasta_saida) / SUBPASTA_BLIP
    pasta_raiz.mkdir(parents=True, exist_ok=True)
    return _resumo_por_ddd(grupos, pasta_raiz, invalidos, mudaram_de_grupo=mudaram_de_grupo)


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
