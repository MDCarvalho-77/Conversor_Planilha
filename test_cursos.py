"""Testes do módulo cursos.py (padronização, sanitização e colisão) — V3.2."""

from __future__ import annotations

import unittest

from cursos import (
    CURSO_NAO_IDENTIFICADO,
    EQUIVALENCIAS_CURSO,
    nomes_de_pasta_unicos,
    padronizar_curso,
    sanitizar_nome_pasta,
)

_PADRAO_ADS = "Analise e Desenvolvimento de Sistemas"


class PadronizarCursoTest(unittest.TestCase):
    def test_tres_variacoes_ads_mesmo_padrao(self) -> None:
        self.assertEqual(
            padronizar_curso(
                "Superior de Tecnologia em Analise e Desenvolvimento de Sistemas"
            ),
            _PADRAO_ADS,
        )
        self.assertEqual(
            padronizar_curso("Tecnologia em Analise e Desenvolvimento de Sistemas"),
            _PADRAO_ADS,
        )
        self.assertEqual(
            padronizar_curso("Analise e Desenvolvimento de Sistemas"),
            _PADRAO_ADS,
        )

    def test_ads_tolera_acento_caixa_e_espacos(self) -> None:
        self.assertEqual(
            padronizar_curso("ANÁLISE E DESENVOLVIMENTO DE SISTEMAS"),
            _PADRAO_ADS,
        )
        self.assertEqual(
            padronizar_curso("  Análise  e  Desenvolvimento  de  Sistemas  "),
            _PADRAO_ADS,
        )
        self.assertEqual(
            padronizar_curso("tecnologia em análise e desenvolvimento de sistemas"),
            _PADRAO_ADS,
        )

    def test_curso_sem_equivalencia_preserva_original(self) -> None:
        self.assertEqual(padronizar_curso("Engenharia Civil"), "Engenharia Civil")
        # strip aplicado, mas texto não reescrito.
        self.assertEqual(padronizar_curso("  Direito  "), "Direito")

    def test_nao_agrupa_cursos_parecidos_mas_diferentes(self) -> None:
        # "Análise de Sistemas" NÃO é ADS cadastrado -> preserva original.
        self.assertEqual(padronizar_curso("Analise de Sistemas"), "Analise de Sistemas")
        self.assertEqual(
            padronizar_curso("Desenvolvimento de Sistemas"),
            "Desenvolvimento de Sistemas",
        )

    def test_entrada_vazia_none_ou_nan(self) -> None:
        self.assertEqual(padronizar_curso(""), CURSO_NAO_IDENTIFICADO)
        self.assertEqual(padronizar_curso("   "), CURSO_NAO_IDENTIFICADO)
        self.assertEqual(padronizar_curso(None), CURSO_NAO_IDENTIFICADO)
        self.assertEqual(padronizar_curso("nan"), CURSO_NAO_IDENTIFICADO)
        self.assertEqual(padronizar_curso("NaN"), CURSO_NAO_IDENTIFICADO)

    def test_semente_tem_apenas_as_tres_variacoes_de_ads(self) -> None:
        self.assertEqual(len(EQUIVALENCIAS_CURSO), 3)
        self.assertTrue(all(valor == _PADRAO_ADS for valor in EQUIVALENCIAS_CURSO.values()))


class SanitizarNomePastaTest(unittest.TestCase):
    def test_remove_caracteres_invalidos_windows(self) -> None:
        resultado = sanitizar_nome_pasta('Curso: A/B\\C*D?E"F<G>H|I')
        for proibido in '\\/:*?"<>|':
            self.assertNotIn(proibido, resultado)

    def test_remove_caracteres_de_controle(self) -> None:
        resultado = sanitizar_nome_pasta("Curso\tde\nTeste")
        self.assertNotIn("\t", resultado)
        self.assertNotIn("\n", resultado)

    def test_remove_pontos_e_espacos_finais(self) -> None:
        self.assertFalse(sanitizar_nome_pasta("Curso de Teste...  ").endswith("."))
        self.assertFalse(sanitizar_nome_pasta("Curso de Teste...  ").endswith(" "))

    def test_trunca_comprimento(self) -> None:
        nome = "A" * 200
        self.assertLessEqual(len(sanitizar_nome_pasta(nome)), 80)

    def test_nunca_retorna_vazio(self) -> None:
        self.assertTrue(sanitizar_nome_pasta(""))
        self.assertTrue(sanitizar_nome_pasta("   "))
        self.assertTrue(sanitizar_nome_pasta(None))
        self.assertTrue(sanitizar_nome_pasta("///:::"))


class NomesDePastaUnicosTest(unittest.TestCase):
    def test_colisao_gera_nomes_distintos_e_deterministicos(self) -> None:
        # Dois cursos DIFERENTES que sanitizam para o mesmo nome base.
        cursos = ["Curso: A", "Curso/ A"]
        mapa = nomes_de_pasta_unicos(cursos)
        self.assertEqual(len(set(mapa.values())), 2)
        self.assertNotEqual(mapa["Curso: A"], mapa["Curso/ A"])
        # Ordem de entrada estável: o primeiro mantém a base.
        self.assertEqual(mapa["Curso: A"], "Curso A")
        self.assertEqual(mapa["Curso/ A"], "Curso A (2)")

    def test_determinismo_entre_execucoes(self) -> None:
        cursos = ["X*", "X?", "X|"]
        self.assertEqual(nomes_de_pasta_unicos(cursos), nomes_de_pasta_unicos(cursos))

    def test_curso_repetido_reaproveita_mesmo_nome(self) -> None:
        cursos = ["Engenharia", "Engenharia", "Direito"]
        mapa = nomes_de_pasta_unicos(cursos)
        self.assertEqual(mapa["Engenharia"], "Engenharia")
        self.assertEqual(mapa["Direito"], "Direito")
        self.assertEqual(len(mapa), 2)

    def test_cursos_distintos_sem_colisao_preservam_base(self) -> None:
        cursos = ["Engenharia Civil", "Direito"]
        mapa = nomes_de_pasta_unicos(cursos)
        self.assertEqual(mapa["Engenharia Civil"], "Engenharia Civil")
        self.assertEqual(mapa["Direito"], "Direito")


if __name__ == "__main__":
    unittest.main()
