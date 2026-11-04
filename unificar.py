"""Modo Unificar lista — Somente telefones (V3.3, REQUISITO A).

Esta modalidade é adicional às três da V3.2 e NÃO as altera. Ela lê telefones
de múltiplos arquivos ``.csv`` e ``.xlsx`` (sem limite de 7 arquivos), consolida
todos os números em uma única lista deduplicada globalmente, classifica por DDD
reutilizando ``blip`` e grava ``LISTA_UNIFICADA/<GRUPO>/Lista_Unificada.{xlsx,csv}``
apenas para os grupos COM contatos.

Reutiliza, SEM alterá-los:

- ``processamento._formatar_telefone``: normaliza o telefone (prefixa 55 em
  números de 10/11 dígitos).
- ``adapters.normalizar_rotulo``: reconhece colunas pelo NOME (NFKD, sem acento,
  colapsa espaços, casefold), nunca pelo conteúdo.
- ``blip.*``: ``formatar_identificador_telefone`` (``+<dígitos>``),
  ``classificar_ddd`` e os 4 grupos de ``GRUPOS_DDD``; convenções de CSV
  (``CSV_ENCODING='utf-8'`` sem BOM, ``CSV_SEPARADOR=','``, ``CSV_QUEBRA_LINHA``).

Decisões de UX importantes (brief V3.3):

- Telefones inválidos ou não classificáveis NUNCA são descartados em silêncio:
  vão para um relatório com valor original, arquivo e linha.
- A gravação NÃO sobrescreve arquivos existentes sem autorização explícita
  (``sobrescrever=True``); caso contrário levanta erro listando os conflitos,
  sem gerar nenhuma saída parcial.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from adapters import normalizar_rotulo
from blip import (
    CSV_ENCODING,
    CSV_QUEBRA_LINHA,
    CSV_SEPARADOR,
    GRUPO_SEM_DDD,
    GRUPOS_DDD,
    classificar_ddd,
    formatar_identificador_telefone,
)
from processamento import _formatar_telefone


# ---------------------------------------------------------------------------
# Constantes do modo Unificar
# ---------------------------------------------------------------------------
SUBPASTA_UNIFICADA = "LISTA_UNIFICADA"
NOME_ARQUIVO_SAIDA = "Lista_Unificada"
COLUNA_TELEFONE = "Telefone"
NOME_RELATORIO_INVALIDOS = "_telefones_invalidos"
CABECALHO_RELATORIO_INVALIDOS: tuple[str, ...] = (
    "valor_original",
    "arquivo",
    "linha",
    "motivo",
)


class TelefoneColunaError(ValueError):
    """Requisito A (V3.3): coluna de telefone ausente ou ambígua em arquivo(s).

    Levantada ANTES de gravar qualquer arquivo quando ao menos um dos arquivos
    selecionados não tem uma coluna de telefone reconhecível ou tem colunas de
    telefone ambíguas. A mensagem lista TODOS os arquivos problemáticos de uma
    vez (nome + motivo) e os cabeçalhos aceitos. Deriva de ``ValueError`` para
    que a UI apresente ``str(erro)`` diretamente (mesmo padrão de
    ``blip.CursoColunaError``).
    """


class SobrescritaError(ValueError):
    """Gravação bloqueada porque já existe uma Lista_Unificada no destino.

    Levantada por :func:`unificar_telefones` quando ``sobrescrever`` é ``False``
    e algum ``LISTA_UNIFICADA/<grupo>/Lista_Unificada.{xlsx,csv}`` já existe.
    Lista os arquivos em conflito e NÃO grava nada (sem saída parcial).
    """


# ---------------------------------------------------------------------------
# Lista CONTROLADA e EXTENSÍVEL de rótulos aceitos como telefone
# ---------------------------------------------------------------------------
# O reconhecimento olha SOMENTE o NOME da coluna (via normalizar_rotulo), nunca
# o conteúdo das células. A lista é deliberadamente controlada: NÃO reconhecemos
# uma coluna só por ela conter a palavra "numero". Para estender a modalidade,
# acrescente o rótulo exato aqui.
ROTULOS_TELEFONE_UNIFICAR: tuple[str, ...] = (
    "Telefone",
    "Numero ou BSUID",
    "Telefone Celular",
    "Telefone da Empresa",
    "Telefone Residencial",
    "Telefone da Empresa/Celular",
    "Celular",
    "Numero de Telefone",
)

# Conjunto normalizado pré-computado para comparação rápida por NOME de coluna.
_ROTULOS_TELEFONE_NORMALIZADO: frozenset[str] = frozenset(
    normalizar_rotulo(r) for r in ROTULOS_TELEFONE_UNIFICAR
)

# Encodings tentados ao ler .csv, NA ORDEM: o primeiro que decodificar vence.
_ENCODINGS_CSV: tuple[str, ...] = ("utf-8-sig", "utf-8", "latin-1")


# ---------------------------------------------------------------------------
# Leitura de cabeçalhos/colunas por arquivo
# ---------------------------------------------------------------------------
# Representamos cada arquivo lido como um dict de colunas: nome_original ->
# lista de valores (como texto). Isso permite reconhecer a coluna pelo NOME e,
# depois, extrair os valores preservando o índice de linha (1-based de dados).
def _ler_colunas_xlsx(caminho: Path) -> dict[str, list[Any]]:
    """Lê cabeçalhos (linha 1) e valores das abas VISÍVEIS de um .xlsx.

    Usa ``openpyxl.load_workbook(read_only=True, data_only=True)`` e considera
    apenas abas ``sheet_state == 'visible'`` (mesmo padrão de
    ``processamento._abas_visiveis`` / ``app._colunas_visiveis``). Quando há mais
    de uma aba visível, as colunas de mesmo nome são concatenadas na ordem das
    abas. Colunas sem cabeçalho (vazio) são ignoradas.
    """
    try:
        workbook = load_workbook(caminho, read_only=True, data_only=True)
    except FileNotFoundError as erro:
        raise FileNotFoundError(
            f"Arquivo de entrada não encontrado: '{caminho}'."
        ) from erro
    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível acessar o arquivo de entrada: '{caminho}'."
        ) from erro
    except Exception as erro:
        raise ValueError(
            f"Não foi possível abrir o arquivo Excel '{caminho.name}'. "
            "Verifique se é um .xlsx válido."
        ) from erro

    colunas: dict[str, list[Any]] = {}
    try:
        for aba in workbook.worksheets:
            if aba.sheet_state != "visible":
                continue
            linhas = aba.iter_rows(values_only=True)
            cabecalho = next(linhas, None)
            if cabecalho is None:
                continue  # aba vazia
            # Mapeia índice de coluna -> nome original (com strip), ignorando
            # cabeçalhos vazios.
            indices: list[tuple[int, str]] = []
            for indice, valor in enumerate(cabecalho):
                nome = "" if valor is None else str(valor).strip()
                if nome == "":
                    continue
                indices.append((indice, nome))
                colunas.setdefault(nome, [])
            for linha in linhas:
                for indice, nome in indices:
                    valor = linha[indice] if indice < len(linha) else None
                    colunas[nome].append(valor)
    finally:
        workbook.close()
    return colunas


def _decodificar_csv(caminho: Path) -> str:
    """Decodifica o .csv tentando os encodings em ordem (primeiro que vence).

    Ordem: ``utf-8-sig`` (consome BOM se houver), ``utf-8``, ``latin-1``
    (``latin-1`` nunca falha, então é o fallback final).
    """
    bruto = caminho.read_bytes()
    for encoding in _ENCODINGS_CSV:
        try:
            return bruto.decode(encoding)
        except UnicodeDecodeError:
            continue
    # Fallback defensivo (latin-1 acima já não falha, mas garantimos retorno).
    return bruto.decode("latin-1")


def _detectar_separador(texto: str) -> str:
    """Detecta o separador ',' vs ';' de um CSV.

    Usa ``csv.Sniffer`` sobre uma amostra. Fallback documentado: quando o
    Sniffer não decide, escolhe o separador com MAIS ocorrências na linha de
    cabeçalho (primeira linha não vazia); empate/ausência -> ','.
    """
    amostra = texto[:4096]
    try:
        dialeto = csv.Sniffer().sniff(amostra, delimiters=";,")
        if dialeto.delimiter in (",", ";"):
            return dialeto.delimiter
    except csv.Error:
        pass
    # Fallback: contagem na linha de cabeçalho.
    cabecalho = ""
    for linha in texto.splitlines():
        if linha.strip() != "":
            cabecalho = linha
            break
    if cabecalho.count(";") > cabecalho.count(","):
        return ";"
    return ","


def _ler_colunas_csv(caminho: Path) -> dict[str, list[Any]]:
    """Lê cabeçalho e valores de um .csv como dict nome_original -> valores."""
    texto = _decodificar_csv(caminho)
    separador = _detectar_separador(texto)
    leitor = csv.reader(texto.splitlines(), delimiter=separador)
    linhas = list(leitor)

    colunas: dict[str, list[Any]] = {}
    if not linhas:
        return colunas
    cabecalho = linhas[0]
    indices: list[tuple[int, str]] = []
    for indice, valor in enumerate(cabecalho):
        nome = "" if valor is None else str(valor).strip()
        if nome == "":
            continue
        indices.append((indice, nome))
        colunas.setdefault(nome, [])
    for linha in linhas[1:]:
        for indice, nome in indices:
            valor = linha[indice] if indice < len(linha) else None
            colunas[nome].append(valor)
    return colunas


def _ler_colunas(caminho: Path) -> dict[str, list[Any]]:
    """Lê as colunas de um arquivo .xlsx ou .csv conforme a extensão."""
    sufixo = caminho.suffix.lower()
    if sufixo == ".csv":
        return _ler_colunas_csv(caminho)
    if sufixo in (".xlsx", ".xlsm"):
        return _ler_colunas_xlsx(caminho)
    raise ValueError(
        f"Formato de arquivo não suportado para '{caminho.name}'. "
        "Use .csv ou .xlsx."
    )


# ---------------------------------------------------------------------------
# Identificação da coluna de telefone por arquivo
# ---------------------------------------------------------------------------
def _coluna_preenchida(valores: list[Any]) -> bool:
    """Indica se a coluna tem ao menos um valor não-vazio."""
    for valor in valores:
        if valor is None:
            continue
        if str(valor).strip() != "":
            return True
    return False


def _identificar_coluna_telefone(
    colunas: dict[str, list[Any]],
) -> tuple[str | None, list[str]]:
    """Identifica a coluna de telefone de um arquivo já lido.

    Retorna ``(coluna_escolhida, candidatas_ambiguas)``.

    Reconhecimento olha SOMENTE o NOME (via normalizar_rotulo) contra a lista
    controlada ``ROTULOS_TELEFONE_UNIFICAR``.

    Heurística de ambiguidade CONSERVADORA (documentada): só há ambiguidade
    quando existem 2+ colunas candidatas DISTINTAS (nomes normalizados
    diferentes) que estejam PREENCHIDAS (ao menos um valor não-vazio). Uma única
    candidata resolve. Duas candidatas de mesmo nome normalizado NÃO geram
    ambiguidade. Candidatas vazias são ignoradas na contagem de ambiguidade.
    """
    # Preserva a ordem de entrada das colunas.
    candidatas: list[str] = [
        nome
        for nome in colunas
        if normalizar_rotulo(nome) in _ROTULOS_TELEFONE_NORMALIZADO
    ]
    if not candidatas:
        return None, []

    # Candidatas DISTINTAS (por nome normalizado) e PREENCHIDAS.
    preenchidas_por_norma: dict[str, str] = {}
    for nome in candidatas:
        if not _coluna_preenchida(colunas[nome]):
            continue
        chave = normalizar_rotulo(nome)
        preenchidas_por_norma.setdefault(chave, nome)

    if len(preenchidas_por_norma) >= 2:
        # Ambíguo: 2+ candidatas distintas preenchidas.
        return None, list(preenchidas_por_norma.values())

    if preenchidas_por_norma:
        # Exatamente uma categoria preenchida -> escolhe a primeira candidata
        # (ordem de entrada) daquela categoria.
        (escolhida,) = preenchidas_por_norma.values()
        return escolhida, []

    # Nenhuma candidata preenchida: escolhe a primeira candidata reconhecida
    # (ainda que vazia) — o arquivo tem a coluna, apenas sem valores.
    return candidatas[0], []


# ---------------------------------------------------------------------------
# Validação das fontes (REQUISITO A: mensagem única com todos os problemas)
# ---------------------------------------------------------------------------
def _analisar_arquivo(caminho: Path) -> tuple[str | None, str | None]:
    """Analisa um arquivo e retorna ``(coluna_escolhida, motivo_problema)``.

    ``motivo_problema`` é ``None`` quando o arquivo é válido. Caso contrário
    descreve o problema (ausência ou ambiguidade).
    """
    colunas = _ler_colunas(caminho)
    escolhida, candidatas = _identificar_coluna_telefone(colunas)
    if escolhida is not None:
        return escolhida, None
    if candidatas:
        lista = ", ".join(f"'{c}'" for c in candidatas)
        return None, f"colunas de telefone ambíguas ({lista})"
    return None, "nenhuma coluna de telefone reconhecível"


def validar_fontes(caminhos: list[str | Path]) -> None:
    """Valida TODOS os arquivos e levanta UM erro com todos os problemas.

    Percorre todos os ``caminhos``, coleta os problemáticos (sem coluna
    reconhecível OU ambíguos) e, havendo qualquer um, levanta um único
    :class:`TelefoneColunaError` com a lista completa (nome + motivo) e os
    cabeçalhos aceitos. NENHUMA saída é gravada aqui. A UI deve chamar esta
    função ANTES de gerar a lista.
    """
    problematicos: list[str] = []
    for bruto in caminhos:
        caminho = Path(bruto)
        _coluna, motivo = _analisar_arquivo(caminho)
        if motivo is not None:
            problematicos.append(f"- {caminho.name}: {motivo}")

    if problematicos:
        aceitos = ", ".join(f"'{r}'" for r in ROTULOS_TELEFONE_UNIFICAR)
        raise TelefoneColunaError(
            "Não foi possível identificar a coluna de telefone nos arquivos a "
            "seguir:\n"
            + "\n".join(problematicos)
            + "\n\nCabeçalhos de telefone aceitos: "
            + aceitos
            + "."
        )


# ---------------------------------------------------------------------------
# Consolidação + dedup global + classificação
# ---------------------------------------------------------------------------
def _consolidar(
    caminhos: list[str | Path],
) -> tuple[list[tuple[str, str]], list[dict[str, Any]], int]:
    """Lê a coluna de telefone escolhida de cada arquivo e consolida.

    Retorna ``(unicos, invalidos, total_lido)`` onde:

    - ``unicos``: lista ``(identificador, grupo)`` já deduplicada GLOBALMENTE
      pelo identificador final, preservando a ordem da PRIMEIRA ocorrência.
    - ``invalidos``: registros sem identificador utilizável, cada um com
      ``valor_original``, ``arquivo``, ``linha`` e ``motivo``.
    - ``total_lido``: total de valores lidos (não vazios após normalização de
      texto) considerados antes do dedup.
    """
    vistos: set[str] = set()
    unicos: list[tuple[str, str]] = []
    invalidos: list[dict[str, Any]] = []
    total_lido = 0

    for bruto in caminhos:
        caminho = Path(bruto)
        colunas = _ler_colunas(caminho)
        escolhida, _candidatas = _identificar_coluna_telefone(colunas)
        if escolhida is None:
            # Defensivo: validar_fontes já deveria ter interrompido antes.
            continue
        valores = colunas[escolhida]
        for posicao, valor_bruto in enumerate(valores):
            valor_original = "" if valor_bruto is None else str(valor_bruto).strip()
            if valor_original == "" or valor_original.lower() == "nan":
                continue  # célula vazia não é lida nem reportada como inválida
            total_lido += 1
            # Linha de dados 1-based (cabeçalho é a linha 1 no arquivo).
            linha = posicao + 2
            # Normaliza (regra V1) e formata para o identificador internacional.
            normalizado = _formatar_telefone(valor_original)
            identificador = formatar_identificador_telefone(normalizado)
            if identificador == "":
                invalidos.append(
                    {
                        "valor_original": valor_original,
                        "arquivo": caminho.name,
                        "linha": linha,
                        "motivo": "Telefone inválido ou não classificável.",
                    }
                )
                continue
            if identificador in vistos:
                continue  # dedup global pós-normalização
            vistos.add(identificador)
            unicos.append((identificador, classificar_ddd(identificador)))

    return unicos, invalidos, total_lido


# ---------------------------------------------------------------------------
# Gravação por grupo (sem limites 100/100KB)
# ---------------------------------------------------------------------------
def _gravar_xlsx_grupo(identificadores: list[str], caminho: Path) -> None:
    """Grava uma lista de telefones em um .xlsx (1 coluna 'Telefone').

    Os telefones são gravados como TEXTO (``number_format='@'``) para preservar
    o ``+`` e evitar notação científica.
    """
    workbook = Workbook()
    aba = workbook.active
    aba.title = "Telefones"
    aba.append([COLUNA_TELEFONE])
    for identificador in identificadores:
        aba.append([identificador])
    for linha in range(2, 2 + len(identificadores)):
        aba.cell(row=linha, column=1).number_format = "@"
    aba.column_dimensions["A"].width = 22
    workbook.save(caminho)


def _gravar_csv_grupo(identificadores: list[str], caminho: Path) -> None:
    """Grava uma lista de telefones em um .csv (UTF-8 sem BOM, 1 coluna)."""
    with open(caminho, "w", encoding=CSV_ENCODING, newline="") as arquivo:
        escritor = csv.writer(
            arquivo,
            delimiter=CSV_SEPARADOR,
            lineterminator=CSV_QUEBRA_LINHA,
            quoting=csv.QUOTE_MINIMAL,
        )
        escritor.writerow([COLUNA_TELEFONE])
        for identificador in identificadores:
            escritor.writerow([identificador])


def _gravar_relatorio_invalidos(invalidos: list[dict[str, Any]], caminho: Path) -> None:
    """Grava o relatório de telefones inválidos em um .csv (UTF-8 sem BOM).

    Uma linha por registro inválido, com ``valor_original``, ``arquivo``,
    ``linha`` e ``motivo``, para que o usuário possa conferir exatamente quais
    registros não foram classificados (brief V3.3: nunca descartar em silêncio).
    """
    with open(caminho, "w", encoding=CSV_ENCODING, newline="") as arquivo:
        escritor = csv.writer(
            arquivo,
            delimiter=CSV_SEPARADOR,
            lineterminator=CSV_QUEBRA_LINHA,
            quoting=csv.QUOTE_MINIMAL,
        )
        escritor.writerow(list(CABECALHO_RELATORIO_INVALIDOS))
        for registro in invalidos:
            escritor.writerow(
                [registro.get(coluna, "") for coluna in CABECALHO_RELATORIO_INVALIDOS]
            )


def _caminhos_de_saida(pasta_raiz: Path, grupo: str) -> tuple[Path, Path]:
    """Retorna ``(xlsx, csv)`` para um grupo dentro de ``LISTA_UNIFICADA``."""
    pasta_grupo = pasta_raiz / grupo
    xlsx = pasta_grupo / f"{NOME_ARQUIVO_SAIDA}.xlsx"
    csv_path = pasta_grupo / f"{NOME_ARQUIVO_SAIDA}.csv"
    return xlsx, csv_path


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------
def unificar_telefones(
    caminhos: list[str | Path],
    pasta_saida: str | Path,
    sobrescrever: bool = False,
) -> dict[str, Any]:
    """Unifica telefones de vários arquivos em listas por grupo de DDD.

    Executa, nesta ordem: validação defensiva (:func:`validar_fontes` também
    existe separada para a UI chamar ANTES de solicitar o destino), consolidação
    com normalização (regra V1) e formatação internacional, dedup GLOBAL pelo
    identificador final, classificação com :func:`blip.classificar_ddd` e
    gravação por grupo em ``<pasta_saida>/LISTA_UNIFICADA/<GRUPO>/``. Sem
    limites de 100 contatos nem 100 KB. Grupos vazios NÃO criam pasta.

    Anti-sobrescrita: com ``sobrescrever=False`` (padrão), se algum
    ``Lista_Unificada.{xlsx,csv}`` já existir, levanta :class:`SobrescritaError`
    listando os conflitos SEM gravar nada. Com ``sobrescrever=True``, regrava.

    Retorna um resumo com: ``total_lido``, ``total_unico`` (pós-dedup),
    ``por_grupo`` (grupo -> quantidade), ``total_sem_ddd``, ``total_invalidos``,
    ``invalidos`` (lista com ``valor_original``/``arquivo``/``linha``),
    ``relatorio_invalidos`` (``Path`` do CSV itemizado de inválidos, ou ``None``
    quando não há inválidos), ``caminhos_gerados`` (lista de ``Path``) e
    ``pasta_raiz`` (``Path`` de ``LISTA_UNIFICADA``).
    """
    # Validação defensiva: interrompe com UMA mensagem antes de qualquer saída.
    validar_fontes(caminhos)

    unicos, invalidos, total_lido = _consolidar(caminhos)

    # Agrupa os identificadores únicos preservando a ordem por grupo.
    por_grupo_listas: dict[str, list[str]] = {grupo: [] for grupo in GRUPOS_DDD}
    for identificador, grupo in unicos:
        por_grupo_listas[grupo].append(identificador)

    pasta_raiz = Path(pasta_saida) / SUBPASTA_UNIFICADA

    relatorio_invalidos = (
        pasta_raiz / f"{NOME_RELATORIO_INVALIDOS}.csv" if invalidos else None
    )

    # Anti-sobrescrita: confere TODOS os conflitos antes de gravar qualquer um.
    if not sobrescrever:
        conflitos: list[Path] = []
        for grupo in GRUPOS_DDD:
            if not por_grupo_listas[grupo]:
                continue
            xlsx, csv_path = _caminhos_de_saida(pasta_raiz, grupo)
            if xlsx.exists():
                conflitos.append(xlsx)
            if csv_path.exists():
                conflitos.append(csv_path)
        if relatorio_invalidos is not None and relatorio_invalidos.exists():
            conflitos.append(relatorio_invalidos)
        if conflitos:
            lista = "\n".join(f"- {c}" for c in conflitos)
            raise SobrescritaError(
                "Já existem arquivos de lista unificada no destino escolhido:\n"
                + lista
                + "\n\nEscolha outro destino ou autorize a substituição."
            )

    # Gravação por grupo COM contatos.
    caminhos_gerados: list[Path] = []
    for grupo in GRUPOS_DDD:
        identificadores = por_grupo_listas[grupo]
        if not identificadores:
            continue  # grupo vazio NÃO cria pasta
        xlsx, csv_path = _caminhos_de_saida(pasta_raiz, grupo)
        xlsx.parent.mkdir(parents=True, exist_ok=True)
        _gravar_xlsx_grupo(identificadores, xlsx)
        _gravar_csv_grupo(identificadores, csv_path)
        caminhos_gerados.extend([xlsx, csv_path])

    # Relatório de inválidos: grava o CSV itemizado (valor/arquivo/linha/motivo)
    # para que o usuário confira os registros não classificados. Nunca é
    # descartado em silêncio (brief V3.3).
    if relatorio_invalidos is not None:
        relatorio_invalidos.parent.mkdir(parents=True, exist_ok=True)
        _gravar_relatorio_invalidos(invalidos, relatorio_invalidos)
        caminhos_gerados.append(relatorio_invalidos)

    por_grupo = {grupo: len(por_grupo_listas[grupo]) for grupo in GRUPOS_DDD}
    return {
        "total_lido": total_lido,
        "total_unico": len(unicos),
        "por_grupo": por_grupo,
        "total_sem_ddd": por_grupo[GRUPO_SEM_DDD],
        "total_invalidos": len(invalidos),
        "invalidos": invalidos,
        "relatorio_invalidos": relatorio_invalidos,
        "caminhos_gerados": caminhos_gerados,
        "pasta_raiz": pasta_raiz,
    }
