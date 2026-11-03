"""Padronização de curso e geração de nomes de pasta seguros (V3.2).

Este módulo é a base da modalidade de saída por curso + DDD. Ele faz três
coisas, todas puras e sem efeitos colaterais:

1. Padroniza o nome de um curso usando uma tabela de equivalências
   *cadastrada manualmente* (:data:`EQUIVALENCIAS_CURSO`). Nunca agrupa cursos
   por similaridade automática — apenas o que estiver explicitamente cadastrado
   é unificado. Cursos sem equivalência preservam o nome original.
2. Sanitiza um nome para uso como nome de pasta no Windows
   (:func:`sanitizar_nome_pasta`).
3. Resolve colisões de nomes de pasta de forma determinística
   (:func:`nomes_de_pasta_unicos`).

Reusa ``adapters.normalizar_rotulo`` (NFKD, remove acentos, colapsa espaços,
casefold) para comparar nomes de curso de forma tolerante a caixa, acento e
espaços — sem reescrever o nome original que será exibido.
"""

from __future__ import annotations

import unicodedata

from adapters import normalizar_rotulo


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
# Marcador usado quando o curso não pôde ser identificado (célula vazia, None
# ou o literal 'nan' vindo de planilhas lidas pelo pandas).
CURSO_NAO_IDENTIFICADO = "CURSO_NAO_IDENTIFICADO"

# Denominação padrão única desta semente (ADS).
_PADRAO_ADS = "Analise e Desenvolvimento de Sistemas"

# Tabela de equivalências de curso — revisável e extensível.
#
# CHAVE   = forma normalizada do nome (resultado de ``normalizar_rotulo``),
#           ou seja, sem acento, sem caixa e com espaços colapsados.
# VALOR   = denominação padrão exibida/usada para agrupar.
#
# Para ESTENDER: acrescente um par ``normalizar_rotulo("Nome Variante"):
# "Denominação Padrão"``. Só variações explicitamente cadastradas são
# unificadas; nenhuma similaridade é inferida automaticamente.
#
# Semente ÚNICA: as três variações de Análise e Desenvolvimento de Sistemas,
# todas apontando para a mesma denominação padrão.
EQUIVALENCIAS_CURSO: dict[str, str] = {
    normalizar_rotulo(
        "Superior de Tecnologia em Analise e Desenvolvimento de Sistemas"
    ): _PADRAO_ADS,
    normalizar_rotulo(
        "Tecnologia em Analise e Desenvolvimento de Sistemas"
    ): _PADRAO_ADS,
    normalizar_rotulo("Analise e Desenvolvimento de Sistemas"): _PADRAO_ADS,
}


# ---------------------------------------------------------------------------
# Padronização de curso
# ---------------------------------------------------------------------------
def padronizar_curso(nome: object) -> str:
    """Retorna a denominação padronizada de um curso.

    Regras:

    * Entrada vazia, ``None`` ou o literal ``'nan'`` -> ``CURSO_NAO_IDENTIFICADO``.
    * Se a forma normalizada constar em :data:`EQUIVALENCIAS_CURSO` -> retorna a
      denominação padrão cadastrada.
    * Caso contrário -> retorna o nome ORIGINAL preservado (apenas com
      ``strip``), sem reescrever.

    Nunca agrupa cursos por similaridade automática; só equivalência cadastrada.
    """
    if nome is None:
        return CURSO_NAO_IDENTIFICADO

    original = str(nome).strip()
    if not original or original.casefold() == "nan":
        return CURSO_NAO_IDENTIFICADO

    chave = normalizar_rotulo(original)
    if chave in EQUIVALENCIAS_CURSO:
        return EQUIVALENCIAS_CURSO[chave]

    return original


# ---------------------------------------------------------------------------
# Sanitização de nome de pasta (Windows)
# ---------------------------------------------------------------------------
# Caracteres proibidos em nomes de arquivo/pasta no Windows.
_CARACTERES_INVALIDOS_WINDOWS = set('\\/:*?"<>|')

# Comprimento máximo razoável do nome de pasta (bem abaixo do limite de 255).
_LIMITE_NOME_PASTA = 80

# Fallback estável quando o nome sanitizado ficaria vazio.
_FALLBACK_PASTA = "Curso"


def sanitizar_nome_pasta(nome: object) -> str:
    """Converte ``nome`` em um nome de pasta seguro para Windows.

    Remove caracteres inválidos do Windows (``\\ / : * ? " < > |``) e de
    controle, remove pontos e espaços finais, trunca a um comprimento razoável
    e nunca retorna string vazia (usa :data:`_FALLBACK_PASTA`).
    """
    texto = "" if nome is None else str(nome)

    caracteres: list[str] = []
    for caractere in texto:
        if caractere in _CARACTERES_INVALIDOS_WINDOWS:
            continue
        # Remove caracteres de controle (categoria Unicode "Cc").
        if unicodedata.category(caractere) == "Cc":
            continue
        caracteres.append(caractere)

    limpo = "".join(caracteres)
    # Colapsa espaços internos e apara extremidades.
    limpo = " ".join(limpo.split())
    # Trunca antes de remover pontos/espaços finais, para não deixar resíduo.
    limpo = limpo[:_LIMITE_NOME_PASTA]
    # Windows não aceita pontos nem espaços finais em nomes de pasta.
    limpo = limpo.rstrip(" .")

    if not limpo:
        return _FALLBACK_PASTA
    return limpo


# ---------------------------------------------------------------------------
# Desambiguação de colisão de nomes de pasta
# ---------------------------------------------------------------------------
def nomes_de_pasta_unicos(cursos: list[str]) -> dict[str, str]:
    """Mapeia cada curso ORIGINAL para um nome de pasta sanitizado distinto.

    Quando dois cursos DIFERENTES sanitizam para o mesmo nome, anexa um sufixo
    numérico determinístico (``" (2)"``, ``" (3)"``, ...) respeitando a ordem de
    entrada, garantindo unicidade. Cursos repetidos (mesmo texto) reaproveitam o
    mesmo nome de pasta, sem gerar novo sufixo.
    """
    resultado: dict[str, str] = {}
    usados: set[str] = set()

    for curso in cursos:
        if curso in resultado:
            # Mesmo curso já mapeado: reaproveita o nome, sem novo sufixo.
            continue

        base = sanitizar_nome_pasta(curso)
        candidato = base
        contador = 2
        while candidato in usados:
            sufixo = f" ({contador})"
            # Garante que o nome com sufixo também respeite o limite.
            recorte = base[: _LIMITE_NOME_PASTA - len(sufixo)].rstrip(" .")
            if not recorte:
                recorte = _FALLBACK_PASTA
            candidato = f"{recorte}{sufixo}"
            contador += 1

        resultado[curso] = candidato
        usados.add(candidato)

    return resultado
