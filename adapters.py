"""Camada de identificação e mapeamento de formatos de planilha (V2).

Esta camada converte planilhas de diferentes origens (CRM 1, CRM 2 e
planilhas genéricas) para um modelo canônico interno comum. O núcleo de
negócio da V2 opera apenas sobre esse modelo canônico, sem condicionais de
formato espalhadas pelo código.

A V1 (pipeline clássico em ``processamento.py``) permanece intacta: esta
camada é usada exclusivamente pelo pipeline V2.
"""

from __future__ import annotations

import unicodedata
from typing import Callable

import pandas as pd


# ---------------------------------------------------------------------------
# Modelo canônico interno
# ---------------------------------------------------------------------------
# Estes são os campos internos comuns a todas as origens. Campos ausentes na
# origem ficam vazios (string vazia). O telefone é o único obrigatório para o
# registro permanecer no resultado — mas essa regra é aplicada pelo pipeline de
# consolidação, não aqui.
CANON_RAZAO_STATUS = "Razão do Status"
CANON_NOME_COMPLETO = "Nome Completo"
CANON_CPF = "CPF"
CANON_EMAIL = "E-mail"
CANON_TELEFONE = "Telefone"
CANON_REGIONAL = "Regional/Unidade do Produto"
CANON_NOME_PRODUTO = "Nome do Produto"
CANON_PRODUTO_PRIMARIO = "Produto Primário"

# Ordem canônica estável das colunas do modelo interno.
COLUNAS_CANONICAS: tuple[str, ...] = (
    CANON_RAZAO_STATUS,
    CANON_NOME_COMPLETO,
    CANON_CPF,
    CANON_EMAIL,
    CANON_TELEFONE,
    CANON_REGIONAL,
    CANON_NOME_PRODUTO,
    CANON_PRODUTO_PRIMARIO,
)


# ---------------------------------------------------------------------------
# Identificadores de formato
# ---------------------------------------------------------------------------
FORMATO_CRM1 = "CRM1"
FORMATO_CRM2 = "CRM2"
FORMATO_GENERICO = "GENERICO"


class FormatoDesconhecidoError(ValueError):
    """A planilha não corresponde a CRM 1, CRM 2 nem a um genérico válido."""


# ---------------------------------------------------------------------------
# Assinaturas de colunas por formato (nomes exatos, já com strip aplicado)
# ---------------------------------------------------------------------------
# CRM 1: colunas que identificam inequivocamente o formato atual da V1.
CRM1_TELEFONE = "Telefone da Empresa/Celular"
CRM1_CPF = "CPF_CNPJ"
CRM1_NOME = "Nome Completo"
CRM1_EMAIL = "E-mail"
CRM1_RAZAO = "Razão do Status"
CRM1_PRODUTO = "Nome do produto"
CRM1_PRODUTO_PRIMARIO = "Primário (Nome do produto) (Produto)"

# Assinatura mínima do CRM 1: a presença de CPF_CNPJ junto do telefone da V1
# distingue o CRM 1 do CRM 2 (que usa "CPF" e telefones separados).
ASSINATURA_CRM1: frozenset[str] = frozenset({CRM1_CPF, CRM1_TELEFONE})

# CRM 2: colunas de assinatura do segundo formato.
CRM2_REGIONAL = "Regional/Unidade do Produto"
CRM2_CPF = "CPF"
CRM2_NOME = "Nome Completo"
CRM2_PRIMEIRO_NOME = "Nome"
CRM2_SEGUNDO_NOME = "Segundo Nome"
CRM2_SOBRENOME = "Sobrenome"
CRM2_EMAIL = "E-mail Comercial"
CRM2_TEL_CELULAR = "Telefone Celular"
CRM2_TEL_EMPRESA = "Telefone da Empresa"
CRM2_TEL_RESIDENCIAL = "Telefone Residencial"
CRM2_RAZAO = "Razão do Status"
CRM2_PRODUTO = "Nome do Produto"
CRM2_PRODUTO_PRIMARIO = "Produto Primário (Nome do Produto) (Produto)"

# Assinatura mínima do CRM 2: a Regional (exclusiva deste formato) junto de ao
# menos um telefone típico do CRM 2.
ASSINATURA_CRM2: frozenset[str] = frozenset({CRM2_REGIONAL, CRM2_TEL_CELULAR})


# ---------------------------------------------------------------------------
# Normalização de nomes de coluna (para reconhecimento tolerante)
# ---------------------------------------------------------------------------
def normalizar_rotulo(rotulo: object) -> str:
    """Normaliza um rótulo de coluna para comparação tolerante.

    Torna a comparação insensível a caixa, acentos e espaços extras. Não é
    usada para renomear colunas — apenas para reconhecê-las.
    """
    texto = "" if rotulo is None else str(rotulo)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = " ".join(texto.split())  # colapsa espaços internos e apara extremidades
    return texto.casefold()


# ---------------------------------------------------------------------------
# Reconhecimento genérico de colunas (lista controlada e extensível)
# ---------------------------------------------------------------------------
# Nomes de coluna aceitos como "telefone" em planilhas genéricas. A lista é
# deliberadamente controlada: o reconhecimento olha apenas o NOME da coluna,
# nunca o conteúdo das células. Comparação tolerante a caixa, acento e espaços
# via normalizar_rotulo. Para estender, basta acrescentar rótulos aqui.
ROTULOS_TELEFONE_GENERICO: tuple[str, ...] = (
    "Telefone",
    "Telefone Celular",
    "Celular",
    "CEL",
    "Telefone Residencial",
    "Telefone Comercial",
    "Telefone da Empresa",
    "Telefone Empresa",
    "Número do Telefone",
    "Numero Telefone",
    "Número Celular",
    "Numero Celular",
    "Fone",
    "Tel",
    "WhatsApp",
)

# Rótulos genéricos aceitos para os demais campos canônicos (opcionais).
ROTULOS_NOME_GENERICO: tuple[str, ...] = ("Nome Completo", "Nome", "Cliente")
ROTULOS_CPF_GENERICO: tuple[str, ...] = ("CPF",)
ROTULOS_EMAIL_GENERICO: tuple[str, ...] = ("E-mail", "Email", "E-mail Comercial")

# Conjuntos normalizados para comparação rápida.
_TELEFONE_NORMALIZADO = frozenset(normalizar_rotulo(r) for r in ROTULOS_TELEFONE_GENERICO)


def localizar_coluna_telefone(colunas: object) -> str | None:
    """Retorna o nome original da primeira coluna reconhecida como telefone.

    A comparação usa a lista controlada ``ROTULOS_TELEFONE_GENERICO`` de forma
    tolerante (caixa/acento/espaço). Preserva a ordem das colunas de entrada.
    Retorna ``None`` quando nenhuma coluna é reconhecida.
    """
    for coluna in colunas:
        if normalizar_rotulo(coluna) in _TELEFONE_NORMALIZADO:
            return str(coluna).strip()
    return None


def _localizar_primeira(colunas: object, rotulos: tuple[str, ...]) -> str | None:
    """Retorna a primeira coluna cujo nome normalizado casa com ``rotulos``."""
    alvos = frozenset(normalizar_rotulo(r) for r in rotulos)
    for coluna in colunas:
        if normalizar_rotulo(coluna) in alvos:
            return str(coluna).strip()
    return None


# ---------------------------------------------------------------------------
# Detecção de formato
# ---------------------------------------------------------------------------
def detectar_formato(colunas: object) -> str:
    """Identifica o formato a partir dos nomes de coluna (já com strip).

    Ordem de prioridade: CRM 1 → CRM 2 → genérico. Retorna um dos
    identificadores de formato. Levanta :class:`FormatoDesconhecidoError`
    quando nem sequer há uma coluna de telefone reconhecível (genérico).
    """
    presentes = {str(coluna).strip() for coluna in colunas}

    if ASSINATURA_CRM1.issubset(presentes):
        return FORMATO_CRM1
    if ASSINATURA_CRM2.issubset(presentes):
        return FORMATO_CRM2
    if _tem_coluna_telefone_generica(presentes):
        return FORMATO_GENERICO

    raise FormatoDesconhecidoError(
        "A planilha não corresponde a CRM 1, CRM 2 nem a um formato genérico "
        "com coluna de telefone reconhecível."
    )


def _tem_coluna_telefone_generica(colunas: set[str]) -> bool:
    """Indica se há ao menos uma coluna reconhecível como telefone."""
    return localizar_coluna_telefone(colunas) is not None


# ---------------------------------------------------------------------------
# Utilitários de montagem do DataFrame canônico
# ---------------------------------------------------------------------------
def _coluna_texto(df: pd.DataFrame, nome: str) -> pd.Series:
    """Retorna a coluna como texto limpo, ou uma coluna vazia se ausente."""
    if nome in df.columns:
        serie = df[nome].astype("string").fillna("").str.strip()
        # Normaliza o literal 'nan' que pode surgir de origens mistas.
        serie = serie.mask(serie.str.lower() == "nan", "")
        return serie
    return pd.Series([""] * len(df), index=df.index, dtype="string")


def _dataframe_canonico_vazio(indice: pd.Index) -> dict[str, pd.Series]:
    """Cria o esqueleto de todas as colunas canônicas vazias."""
    return {
        coluna: pd.Series([""] * len(indice), index=indice, dtype="string")
        for coluna in COLUNAS_CANONICAS
    }


def _montar_canonico(dados: dict[str, pd.Series], indice: pd.Index) -> pd.DataFrame:
    """Monta o DataFrame canônico na ordem estável, preenchendo ausentes."""
    base = _dataframe_canonico_vazio(indice)
    base.update(dados)
    canonico = pd.DataFrame(base, columns=list(COLUNAS_CANONICAS))
    return canonico


# ---------------------------------------------------------------------------
# Adapter: CRM 1
# ---------------------------------------------------------------------------
def adaptar_crm1(df: pd.DataFrame) -> pd.DataFrame:
    """Mapeia um DataFrame do CRM 1 para o modelo canônico.

    Espera colunas já com espaços das extremidades removidos. A normalização
    final do telefone (regra da V1) é aplicada pelo pipeline, não aqui.
    """
    dados = {
        CANON_RAZAO_STATUS: _coluna_texto(df, CRM1_RAZAO),
        CANON_NOME_COMPLETO: _coluna_texto(df, CRM1_NOME),
        CANON_CPF: _coluna_texto(df, CRM1_CPF),
        CANON_EMAIL: _coluna_texto(df, CRM1_EMAIL),
        CANON_TELEFONE: _coluna_texto(df, CRM1_TELEFONE),
        CANON_NOME_PRODUTO: _coluna_texto(df, CRM1_PRODUTO),
        CANON_PRODUTO_PRIMARIO: _coluna_texto(df, CRM1_PRODUTO_PRIMARIO),
        # CRM 1 não possui Regional/Unidade → permanece vazio (P3).
    }
    return _montar_canonico(dados, df.index)


# ---------------------------------------------------------------------------
# Adapter: CRM 2
# ---------------------------------------------------------------------------
# Ordem de prioridade dos telefones do CRM 2: usa o primeiro campo preenchido.
CRM2_PRIORIDADE_TELEFONE: tuple[str, ...] = (
    CRM2_TEL_CELULAR,
    CRM2_TEL_EMPRESA,
    CRM2_TEL_RESIDENCIAL,
)


def _telefone_por_prioridade(df: pd.DataFrame) -> pd.Series:
    """Escolhe, por linha, o primeiro telefone preenchido na ordem de prioridade.

    Celular → Empresa → Residencial. Se nenhum estiver preenchido, resulta em
    string vazia (o registro será removido pela regra de telefone obrigatório).
    """
    escolhido = pd.Series([""] * len(df), index=df.index, dtype="string")
    for coluna in CRM2_PRIORIDADE_TELEFONE:
        valores = _coluna_texto(df, coluna)
        # Preenche apenas onde ainda não há telefone escolhido.
        vazio = escolhido.eq("")
        escolhido = escolhido.mask(vazio, valores)
    return escolhido


def _nome_completo_com_fallback(df: pd.DataFrame) -> pd.Series:
    """Usa 'Nome Completo'; se vazio, compõe de Nome + Segundo Nome + Sobrenome.

    Campos vazios são simplesmente ignorados na composição, sem impedir o
    processamento.
    """
    nome_completo = _coluna_texto(df, CRM2_NOME)
    primeiro = _coluna_texto(df, CRM2_PRIMEIRO_NOME)
    segundo = _coluna_texto(df, CRM2_SEGUNDO_NOME)
    sobrenome = _coluna_texto(df, CRM2_SOBRENOME)

    def _compor(partes: tuple[str, str, str]) -> str:
        return " ".join(parte for parte in partes if parte)

    composto = pd.Series(
        [
            _compor((p, s, u))
            for p, s, u in zip(primeiro.tolist(), segundo.tolist(), sobrenome.tolist())
        ],
        index=df.index,
        dtype="string",
    )
    # Onde 'Nome Completo' está vazio, cai para a composição.
    return nome_completo.mask(nome_completo.eq(""), composto)


def adaptar_crm2(df: pd.DataFrame) -> pd.DataFrame:
    """Mapeia um DataFrame do CRM 2 para o modelo canônico."""
    dados = {
        CANON_RAZAO_STATUS: _coluna_texto(df, CRM2_RAZAO),
        CANON_NOME_COMPLETO: _nome_completo_com_fallback(df),
        CANON_CPF: _coluna_texto(df, CRM2_CPF),
        CANON_EMAIL: _coluna_texto(df, CRM2_EMAIL),
        CANON_TELEFONE: _telefone_por_prioridade(df),
        CANON_REGIONAL: _coluna_texto(df, CRM2_REGIONAL),
        CANON_NOME_PRODUTO: _coluna_texto(df, CRM2_PRODUTO),
        CANON_PRODUTO_PRIMARIO: _coluna_texto(df, CRM2_PRODUTO_PRIMARIO),
    }
    return _montar_canonico(dados, df.index)


# ---------------------------------------------------------------------------
# Adapter: planilha genérica
# ---------------------------------------------------------------------------
def adaptar_generico(df: pd.DataFrame) -> pd.DataFrame:
    """Mapeia uma planilha genérica para o modelo canônico.

    O telefone é obrigatório: a primeira coluna reconhecida como telefone
    (via lista controlada) vira o telefone canônico. Os demais campos são
    opcionais e ficam vazios quando não houver coluna reconhecível.
    """
    coluna_telefone = localizar_coluna_telefone(df.columns)
    if coluna_telefone is None:
        raise FormatoDesconhecidoError(
            "A planilha genérica não possui coluna de telefone reconhecível."
        )

    coluna_nome = _localizar_primeira(df.columns, ROTULOS_NOME_GENERICO)
    coluna_cpf = _localizar_primeira(df.columns, ROTULOS_CPF_GENERICO)
    coluna_email = _localizar_primeira(df.columns, ROTULOS_EMAIL_GENERICO)

    dados = {CANON_TELEFONE: _coluna_texto(df, coluna_telefone)}
    if coluna_nome is not None:
        dados[CANON_NOME_COMPLETO] = _coluna_texto(df, coluna_nome)
    if coluna_cpf is not None:
        dados[CANON_CPF] = _coluna_texto(df, coluna_cpf)
    if coluna_email is not None:
        dados[CANON_EMAIL] = _coluna_texto(df, coluna_email)
    # Razão do Status, Regional, Produto e Produto Primário ficam vazios.
    return _montar_canonico(dados, df.index)


# Registro de adapters por formato.
ADAPTERS: dict[str, Callable[[pd.DataFrame], pd.DataFrame]] = {
    FORMATO_CRM1: adaptar_crm1,
    FORMATO_CRM2: adaptar_crm2,
    FORMATO_GENERICO: adaptar_generico,
}
