"""Testes do módulo ``tutorial`` (sem instanciar nenhuma janela Tk).

Cobrem a persistência à prova de falha da preferência "não mostrar novamente",
a lógica de ``deve_mostrar_tutorial`` e o controlador de navegação das 4 etapas.
Usam um diretório temporário para o ``config.json`` via o override de
``_caminho_config`` (parâmetro ``override``), sem tocar no ``%APPDATA%`` real.
"""

import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import tutorial


class TestPreferencias(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.config = Path(self._tmp.name) / "ConversorCRM" / "config.json"

    def test_deve_mostrar_true_sem_preferencia(self) -> None:
        self.assertTrue(tutorial.deve_mostrar_tutorial(override=self.config))

    def test_salvar_persiste_e_altera_deve_mostrar(self) -> None:
        tutorial.salvar_preferencia_nao_mostrar(True, override=self.config)
        self.assertTrue(self.config.exists())
        self.assertFalse(tutorial.deve_mostrar_tutorial(override=self.config))

    def test_releitura_reflete_valor_salvo(self) -> None:
        tutorial.salvar_preferencia_nao_mostrar(True, override=self.config)
        preferencias = tutorial.carregar_preferencias(override=self.config)
        self.assertTrue(preferencias.get("nao_mostrar"))

        # Reabertura manual continua possível: voltar a False permite mostrar.
        tutorial.salvar_preferencia_nao_mostrar(False, override=self.config)
        self.assertTrue(tutorial.deve_mostrar_tutorial(override=self.config))

    def test_salvar_com_destino_invalido_nao_lanca(self) -> None:
        # Um arquivo ocupa o lugar do diretório pai -> mkdir/escrita falham.
        arquivo = Path(self._tmp.name) / "bloqueio"
        arquivo.write_text("x", encoding="utf-8")
        destino_invalido = arquivo / "ConversorCRM" / "config.json"
        try:
            tutorial.salvar_preferencia_nao_mostrar(True, override=destino_invalido)
        except Exception as exc:  # pragma: no cover - não deve ocorrer
            self.fail(f"salvar não deveria propagar exceção: {exc!r}")
        # Leitura também é à prova de falha e retorna defaults.
        self.assertTrue(tutorial.deve_mostrar_tutorial(override=destino_invalido))

    def test_carregar_json_corrompido_retorna_defaults(self) -> None:
        self.config.parent.mkdir(parents=True, exist_ok=True)
        self.config.write_text("{isso não é json", encoding="utf-8")
        preferencias = tutorial.carregar_preferencias(override=self.config)
        self.assertEqual(preferencias, tutorial.PREFERENCIAS_PADRAO)
        self.assertTrue(tutorial.deve_mostrar_tutorial(override=self.config))

    def test_appdata_inexistente_usa_fallback_home(self) -> None:
        with mock.patch.object(tutorial.os, "getenv", return_value=None):
            caminho = tutorial._caminho_config()
        esperado = Path.home() / "ConversorCRM" / "config.json"
        self.assertEqual(caminho, esperado)

    def test_appdata_presente_usa_appdata(self) -> None:
        def fake_getenv(nome, default=None):
            if nome == "APPDATA":
                return self._tmp.name
            return None

        with mock.patch.object(tutorial.os, "getenv", side_effect=fake_getenv):
            caminho = tutorial._caminho_config()
        esperado = Path(self._tmp.name) / "ConversorCRM" / "config.json"
        self.assertEqual(caminho, esperado)


class TestEtapas(unittest.TestCase):
    def test_quatro_etapas_com_titulo_e_texto(self) -> None:
        self.assertEqual(len(tutorial.ETAPAS), 4)
        for etapa in tutorial.ETAPAS:
            self.assertTrue(etapa.titulo.strip())
            self.assertTrue(etapa.texto.strip())
            self.assertNotIn("...", etapa.texto)
            self.assertNotIn("…", etapa.texto)

    def test_etapa_importacao_menciona_limite_e_unificacao_sem_limite(self) -> None:
        texto = tutorial.ETAPAS[1].texto
        self.assertIn("7", texto)
        self.assertIn("NÃO tem limite", texto)


class TestControladorNavegacao(unittest.TestCase):
    def test_inicio_no_primeiro(self) -> None:
        c = tutorial.ControladorTutorial()
        self.assertEqual(c.indice, 0)
        self.assertTrue(c.primeiro)
        self.assertFalse(c.ultimo)

    def test_anterior_nao_estoura_limite_inferior(self) -> None:
        c = tutorial.ControladorTutorial()
        c.anterior()
        self.assertEqual(c.indice, 0)

    def test_proximo_percorre_e_para_no_ultimo(self) -> None:
        c = tutorial.ControladorTutorial()
        for esperado in (1, 2, 3, 3, 3):
            c.proximo()
            self.assertEqual(c.indice, esperado)
        self.assertTrue(c.ultimo)
        self.assertFalse(c.primeiro)

    def test_navega_para_frente_e_para_tras(self) -> None:
        c = tutorial.ControladorTutorial()
        c.proximo()
        c.proximo()
        self.assertEqual(c.indice, 2)
        c.anterior()
        self.assertEqual(c.indice, 1)

    def test_total_corresponde_as_etapas(self) -> None:
        c = tutorial.ControladorTutorial(len(tutorial.ETAPAS))
        self.assertEqual(c.total, 4)


if __name__ == "__main__":
    unittest.main()
