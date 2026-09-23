"""Tratamento de planilhas de clientes potenciais."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook


COLUNA_TELEFONE = "Telefone da Empresa/Celular"
COLUNA_TELEFONE_SOMENTE_NUMEROS = "Telefone"
COLUNAS_PRODUTO = (
    "Nome do produto",
    "Primário (Nome do produto) (Produto)",
)
COLUNAS_SAIDA = (
    "Razão do Status",
    "Nome Completo",
    "CPF_CNPJ",
    "E-mail",
    COLUNA_TELEFONE,
)
NOMES_SAIDA = {"CPF_CNPJ": "CPF"}
FORMATO_PLANILHA_COMPLETA = "Planilha completa"
FORMATO_SOMENTE_NUMEROS = "Somente números"
FORMATO_AMBAS_SAIDAS = "As duas saídas"
FORMATOS_SAIDA = (
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
)


def _abas_visiveis(caminho_entrada: Path) -> list[str]:
    """Obtém os nomes das abas visíveis sem ler as abas técnicas ocultas."""
    try:
        workbook = load_workbook(caminho_entrada, read_only=True, data_only=False)
    except FileNotFoundError as erro:
        raise FileNotFoundError(
            f"Arquivo de entrada não encontrado: '{caminho_entrada}'."
        ) from erro
    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível acessar o arquivo de entrada: '{caminho_entrada}'."
        ) from erro
    except Exception as erro:
        raise ValueError(
            "Não foi possível abrir o arquivo Excel de entrada. "
            "Verifique se o arquivo é um .xlsx válido."
        ) from erro

    try:
        return [aba.title for aba in workbook.worksheets if aba.sheet_state == "visible"]
    finally:
        workbook.close()


def _limpar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Remove espaços no início e no fim dos nomes de colunas."""
    nomes_limpos = [str(coluna).strip() for coluna in df.columns]
    repetidas = sorted(
        {coluna for coluna in nomes_limpos if nomes_limpos.count(coluna) > 1}
    )
    if repetidas:
        nomes = ", ".join(repr(coluna) for coluna in repetidas)
        raise ValueError(
            "A limpeza dos nomes das colunas gerou nomes duplicados: " f"{nomes}."
        )

    resultado = df.copy()
    resultado.columns = nomes_limpos
    return resultado


def _validar_colunas(df: pd.DataFrame, colunas: tuple[str, ...], contexto: str) -> None:
    """Gera uma mensagem clara quando colunas obrigatórias não estão presentes."""
    ausentes = [coluna for coluna in colunas if coluna not in df.columns]
    if ausentes:
        nomes = ", ".join(repr(coluna) for coluna in ausentes)
        raise ValueError(f"Coluna(s) obrigatória(s) ausente(s) para {contexto}: {nomes}.")


def _formatar_telefone(telefone: Any) -> str:
    """Conserva só dígitos e acrescenta o código 55 quando aplicável."""
    numeros = re.sub(r"\D", "", str(telefone))
    if len(numeros) in (10, 11):
        return f"55{numeros}"
    return numeros


def processar_dados(caminho_entrada: str | Path) -> tuple[pd.DataFrame, dict[str, int]]:
    """Processa as abas visíveis e grava uma planilha pronta para uso.

    A deduplicação compara todas as colunas originais, incluindo as duas
    colunas de produto exigidas, antes de qualquer normalização de telefone.
    """
    entrada = Path(caminho_entrada)
    abas_visiveis = _abas_visiveis(entrada)

    if not abas_visiveis:
        raise ValueError("O arquivo Excel não possui abas visíveis para processar.")

    try:
        tabelas = [
            pd.read_excel(entrada, sheet_name=aba, engine="openpyxl")
            for aba in abas_visiveis
        ]
    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível ler o arquivo de entrada: '{entrada}'."
        ) from erro
    except Exception as erro:
        raise ValueError(
            "Não foi possível ler as abas visíveis do arquivo Excel de entrada."
        ) from erro

    df = _limpar_nomes_colunas(pd.concat(tabelas, ignore_index=True, sort=False))
    registros_lidos = len(df)

    _validar_colunas(df, (COLUNA_TELEFONE,), "identificar o telefone")

    # O filtro é calculado em uma série temporária: os dados originais não são
    # alterados antes da deduplicação.
    telefone_texto = df[COLUNA_TELEFONE].astype("string").str.strip()
    tem_telefone = (
        telefone_texto.notna()
        & telefone_texto.ne("")
        & telefone_texto.str.lower().ne("nan")
    ).fillna(False)
    df_com_telefone = df.loc[tem_telefone].copy()
    removidos_sem_telefone = registros_lidos - len(df_com_telefone)

    _validar_colunas(df_com_telefone, COLUNAS_PRODUTO, "deduplicação")

    # Sem subset: somente linhas idênticas em todas as colunas são removidas.
    df_sem_duplicatas = df_com_telefone.drop_duplicates().copy()
    duplicados_removidos = len(df_com_telefone) - len(df_sem_duplicatas)

    _validar_colunas(df_sem_duplicatas, COLUNAS_SAIDA, "gerar a saída")

    resultado = df_sem_duplicatas.loc[:, COLUNAS_SAIDA].copy()
    resultado[COLUNA_TELEFONE] = resultado[COLUNA_TELEFONE].map(_formatar_telefone)
    resultado = resultado.rename(columns=NOMES_SAIDA)

    metricas = {
        "registros_lidos": registros_lidos,
        "removidos_sem_telefone": removidos_sem_telefone,
        "duplicados_removidos": duplicados_removidos,
        "registros_finais": len(resultado),
    }
    return resultado, metricas

def _gravar_planilha(df: pd.DataFrame, caminho_saida: str | Path) -> None:
    """Grava uma saída Excel com mensagens de erro apropriadas."""
    saida = Path(caminho_saida)
    try:
        df.to_excel(saida, index=False, engine="openpyxl")
    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível gravar o arquivo de saída: '{saida}'. "
            "Feche-o caso esteja aberto no Excel e tente novamente."
        ) from erro
    except Exception as erro:
        raise ValueError(f"Não foi possível gerar o arquivo de saída: '{saida}'.") from erro


def gerar_saidas(
    caminho_entrada: str | Path, formato_saida: str = FORMATO_PLANILHA_COMPLETA
) -> dict[str, Any]:
    """Gera uma ou ambas as saídas a partir do mesmo DataFrame final."""
    if formato_saida not in FORMATOS_SAIDA:
        raise ValueError(f"Formato de saída inválido: {formato_saida!r}.")

    entrada = Path(caminho_entrada)
    resultado, metricas = processar_dados(entrada)
    arquivos_gerados: list[Path] = []

    if formato_saida in (FORMATO_PLANILHA_COMPLETA, FORMATO_AMBAS_SAIDAS):
        completo = entrada.with_name(f"{entrada.stem}_tratado.xlsx")
        _gravar_planilha(resultado, completo)
        arquivos_gerados.append(completo)

    if formato_saida in (FORMATO_SOMENTE_NUMEROS, FORMATO_AMBAS_SAIDAS):
        telefones = entrada.with_name(f"{entrada.stem}_telefones.xlsx")
        somente_telefones = resultado.loc[:, [COLUNA_TELEFONE]].rename(
            columns={COLUNA_TELEFONE: COLUNA_TELEFONE_SOMENTE_NUMEROS}
        )
        _gravar_planilha(somente_telefones, telefones)
        arquivos_gerados.append(telefones)

    return {**metricas, "arquivos_gerados": arquivos_gerados}


def processar_planilha(
    caminho_entrada: str | Path, caminho_saida: str | Path
) -> dict[str, int]:
    """Compatibilidade com a V1: gera a planilha completa no caminho indicado."""
    resultado, metricas = processar_dados(caminho_entrada)
    _gravar_planilha(resultado, caminho_saida)
    return metricas
