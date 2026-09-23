"""Pipeline V2: consolidação de múltiplos arquivos e deduplicação global.

Este módulo implementa o caminho adicional da V2, que opera sobre o modelo
canônico produzido por ``adapters``. A V1 (``processamento.py``) permanece
intacta; aqui apenas reutilizamos funções puras dela (normalização de telefone,
leitura de abas visíveis e limpeza de nomes de coluna).

Fluxo:
    arquivos → (por arquivo) leitura de abas visíveis → limpeza de nomes →
    detecção de formato → mapeamento canônico → consolidação →
    remoção de registros sem telefone → deduplicação GLOBAL (sobre os campos
    canônicos) → normalização dos telefones → DataFrame canônico final.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

import adapters
from adapters import (
    CANON_CPF,
    CANON_EMAIL,
    CANON_NOME_COMPLETO,
    CANON_NOME_PRODUTO,
    CANON_RAZAO_STATUS,
    CANON_REGIONAL,
    CANON_TELEFONE,
    COLUNAS_CANONICAS,
    FormatoDesconhecidoError,
)
from processamento import (
    COLUNA_TELEFONE,
    COLUNA_TELEFONE_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    _abas_visiveis,
    _formatar_telefone,
    _gravar_planilha,
    _limpar_nomes_colunas,
)


# Limite de arquivos por execução (regra de negócio da V2).
LIMITE_ARQUIVOS = 7

# Nova saída da V2 (adicional às três clássicas).
FORMATO_PLANILHA_ESTENDIDA = "Planilha completa estendida"

# Formatos aceitos pelo pipeline V2 (as três clássicas continuam válidas e a
# estendida é a novidade).
FORMATOS_SAIDA_V2 = (
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
    FORMATO_PLANILHA_ESTENDIDA,
)

# Colunas da saída COMPLETA clássica (5), mapeadas do canônico. A ordem e os
# nomes são idênticos aos da V1.
_MAPA_SAIDA_COMPLETA: tuple[tuple[str, str], ...] = (
    (CANON_RAZAO_STATUS, "Razão do Status"),
    (CANON_NOME_COMPLETO, "Nome Completo"),
    (CANON_CPF, "CPF"),
    (CANON_EMAIL, "E-mail"),
    (CANON_TELEFONE, COLUNA_TELEFONE),
)

# Colunas da saída ESTENDIDA (7): as cinco da completa + Regional + Produto.
# Produto Primário NÃO entra na saída estendida (P5).
_MAPA_SAIDA_ESTENDIDA: tuple[tuple[str, str], ...] = (
    (CANON_RAZAO_STATUS, "Razão do Status"),
    (CANON_NOME_COMPLETO, "Nome Completo"),
    (CANON_CPF, "CPF"),
    (CANON_EMAIL, "E-mail"),
    (CANON_TELEFONE, COLUNA_TELEFONE),
    (CANON_REGIONAL, "Regional/Unidade do Produto"),
    (CANON_NOME_PRODUTO, "Nome do Produto"),
)


class LimiteArquivosExcedidoError(ValueError):
    """Mais arquivos do que o permitido em uma única execução."""


def _ler_arquivo_canonico(caminho: str | Path) -> pd.DataFrame:
    """Lê as abas visíveis de um arquivo e devolve o DataFrame canônico.

    Reutiliza a regra da V1 de ignorar abas ``hidden``/``veryHidden`` e de
    limpar espaços nos nomes de coluna, antes de detectar o formato e aplicar
    o adapter correspondente.
    """
    entrada = Path(caminho)
    abas_visiveis = _abas_visiveis(entrada)
    if not abas_visiveis:
        raise ValueError(
            f"O arquivo '{entrada.name}' não possui abas visíveis para processar."
        )

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
            f"Não foi possível ler as abas visíveis do arquivo '{entrada.name}'."
        ) from erro

    df = _limpar_nomes_colunas(pd.concat(tabelas, ignore_index=True, sort=False))

    # Detecção por arquivo (P: identificação individual). Pode levantar
    # FormatoDesconhecidoError, tratada de forma amigável pela camada de UI.
    formato = adapters.detectar_formato(df.columns)
    adapter = adapters.ADAPTERS[formato]
    return adapter(df)


def consolidar_arquivos(caminhos: list[str | Path]) -> tuple[pd.DataFrame, int]:
    """Lê e consolida vários arquivos em um único DataFrame canônico.

    Retorna o DataFrame consolidado (ainda sem limpeza/dedup) e o total de
    registros lidos. Valida o limite de arquivos.
    """
    if not caminhos:
        raise ValueError("Nenhum arquivo selecionado para processar.")
    if len(caminhos) > LIMITE_ARQUIVOS:
        raise LimiteArquivosExcedidoError(
            f"Selecione no máximo {LIMITE_ARQUIVOS} arquivos por vez."
        )

    canonicos = [_ler_arquivo_canonico(caminho) for caminho in caminhos]
    consolidado = pd.concat(canonicos, ignore_index=True, sort=False)
    # Garante a ordem/o conjunto canônico estável mesmo após o concat.
    consolidado = consolidado.reindex(columns=list(COLUNAS_CANONICAS))
    registros_lidos = len(consolidado)
    return consolidado, registros_lidos


def processar_consolidado(
    caminhos: list[str | Path],
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Executa o pipeline V2 completo e devolve o canônico final + métricas.

    Ordem (espelha a V1): remover sem telefone → deduplicar → normalizar
    telefone. A deduplicação é GLOBAL sobre os campos canônicos.
    """
    consolidado, registros_lidos = consolidar_arquivos(caminhos)

    # Remoção de registros sem telefone (telefone é o único campo obrigatório).
    telefone = consolidado[CANON_TELEFONE].astype("string").fillna("").str.strip()
    tem_telefone = telefone.ne("") & telefone.str.lower().ne("nan")
    com_telefone = consolidado.loc[tem_telefone].copy()
    removidos_sem_telefone = registros_lidos - len(com_telefone)

    # Deduplicação GLOBAL sobre todos os campos canônicos (P1), antes de
    # normalizar o telefone — mesma ordem relativa da V1.
    sem_duplicatas = com_telefone.drop_duplicates().copy()
    duplicados_removidos = len(com_telefone) - len(sem_duplicatas)

    # Normalização de telefone reutilizando exatamente a regra da V1.
    sem_duplicatas[CANON_TELEFONE] = sem_duplicatas[CANON_TELEFONE].map(_formatar_telefone)

    final = sem_duplicatas.reset_index(drop=True)
    metricas = {
        "registros_lidos": registros_lidos,
        "removidos_sem_telefone": removidos_sem_telefone,
        "duplicados_removidos": duplicados_removidos,
        "registros_finais": len(final),
    }
    return final, metricas


def _projetar(final: pd.DataFrame, mapa: tuple[tuple[str, str], ...]) -> pd.DataFrame:
    """Projeta o canônico final para uma saída, renomeando as colunas.

    ``mapa`` associa (coluna_canonica, nome_de_saida) na ordem desejada.
    """
    origem = [canon for canon, _ in mapa]
    destino = [nome for _, nome in mapa]
    projetado = final.loc[:, origem].copy()
    projetado.columns = destino
    return projetado


def saida_completa(final: pd.DataFrame) -> pd.DataFrame:
    """Saída completa clássica (5 colunas) a partir do canônico final."""
    return _projetar(final, _MAPA_SAIDA_COMPLETA)


def saida_estendida(final: pd.DataFrame) -> pd.DataFrame:
    """Saída estendida (7 colunas) a partir do canônico final."""
    return _projetar(final, _MAPA_SAIDA_ESTENDIDA)


def saida_somente_numeros(final: pd.DataFrame) -> pd.DataFrame:
    """Saída somente com a coluna de telefone (nome 'Telefone')."""
    somente = final.loc[:, [CANON_TELEFONE]].copy()
    somente.columns = [COLUNA_TELEFONE_SOMENTE_NUMEROS]
    return somente


def _nome_base_saida(caminhos: list[str | Path]) -> Path:
    """Define o nome-base dos arquivos de saída.

    Um único arquivo → usa o nome do próprio arquivo. Vários arquivos →
    prefixa 'Consolidado_' ao nome do primeiro, para deixar claro que a saída
    reúne várias entradas. A saída é gravada na pasta do primeiro arquivo.
    """
    primeiro = Path(caminhos[0])
    if len(caminhos) == 1:
        return primeiro
    return primeiro.with_name(f"Consolidado_{primeiro.stem}{primeiro.suffix}")


def gerar_saidas_v2(
    caminhos: list[str | Path],
    formato_saida: str = FORMATO_PLANILHA_ESTENDIDA,
) -> dict[str, Any]:
    """Executa o pipeline V2 e grava a(s) saída(s) escolhida(s).

    Todas as saídas partem do MESMO DataFrame canônico final: não há segunda
    deduplicação por saída.
    """
    if formato_saida not in FORMATOS_SAIDA_V2:
        raise ValueError(f"Formato de saída inválido: {formato_saida!r}.")

    final, metricas = processar_consolidado(caminhos)
    base = _nome_base_saida(caminhos)
    arquivos_gerados: list[Path] = []

    if formato_saida in (FORMATO_PLANILHA_COMPLETA, FORMATO_AMBAS_SAIDAS):
        completo = base.with_name(f"{base.stem}_tratado.xlsx")
        _gravar_planilha(saida_completa(final), completo)
        arquivos_gerados.append(completo)

    if formato_saida in (FORMATO_SOMENTE_NUMEROS, FORMATO_AMBAS_SAIDAS):
        telefones = base.with_name(f"{base.stem}_telefones.xlsx")
        _gravar_planilha(saida_somente_numeros(final), telefones)
        arquivos_gerados.append(telefones)

    if formato_saida == FORMATO_PLANILHA_ESTENDIDA:
        estendido = base.with_name(f"{base.stem}_estendido.xlsx")
        _gravar_planilha(saida_estendida(final), estendido)
        arquivos_gerados.append(estendido)

    return {**metricas, "arquivos_gerados": arquivos_gerados}
