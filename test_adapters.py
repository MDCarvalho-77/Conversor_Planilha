"""Testes da camada de adapters/mapeamento da V2."""

from __future__ import annotations

import unittest

import pandas as pd

import adapters
from adapters import (
    COLUNAS_CANONICAS,
    CANON_CPF,
    CANON_EMAIL,
    CANON_NOME_COMPLETO,
    CANON_NOME_PRODUTO,
    CANON_PRODUTO_PRIMARIO,
    CANON_RAZAO_STATUS,
    CANON_REGIONAL,
    CANON_TELEFONE,
    FORMATO_CRM1,
    FORMATO_CRM2,
    FORMATO_GENERICO,
    FormatoDesconhecidoError,
    adaptar_crm1,
    adaptar_crm2,
    adaptar_generico,
    detectar_formato,
    localizar_coluna_telefone,
    normalizar_rotulo,
)


class NormalizacaoRotuloTest(unittest.TestCase):
    def test_tolera_caixa_acento_e_espacos(self) -> None:
        self.assertEqual(normalizar_rotulo("  Telefone  "), "telefone")
        self.assertEqual(normalizar_rotulo("TELEFONE"), "telefone")
        self.assertEqual(normalizar_rotulo("Número  Celular"), "numero celular")
        self.assertEqual(normalizar_rotulo(None), "")


class DeteccaoFormatoTest(unittest.TestCase):
    def test_detecta_crm1_pela_assinatura(self) -> None:
        colunas = [
            "Razão do Status",
            "Nome Completo",
            "CPF_CNPJ",
            "E-mail",
            "Telefone da Empresa/Celular",
            "Nome do produto",
            "Primário (Nome do produto) (Produto)",
        ]
        self.assertEqual(detectar_formato(colunas), FORMATO_CRM1)

    def test_detecta_crm1_com_espacos_nas_extremidades(self) -> None:
        colunas = [" CPF_CNPJ ", " Telefone da Empresa/Celular "]
        self.assertEqual(detectar_formato(colunas), FORMATO_CRM1)

    def test_formato_desconhecido_levanta_erro(self) -> None:
        colunas = ["Coluna A", "Coluna B", "Alguma Coisa"]
        with self.assertRaises(FormatoDesconhecidoError):
            detectar_formato(colunas)


class AdaptarCrm1Test(unittest.TestCase):
    def _df_crm1(self) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "Razão do Status": ["Em contato"],
                "Nome Completo": ["João da Silva"],
                "CPF_CNPJ": ["12345678900"],
                "E-mail": ["joao@exemplo.com"],
                "Telefone da Empresa/Celular": ["(48) 99999-0001"],
                "Nome do produto": ["Produto X"],
                "Primário (Nome do produto) (Produto)": ["Linha Y"],
                # Coluna fora do canônico deve ser simplesmente ignorada.
                "Cidade": ["Florianópolis"],
            }
        )

    def test_mapeia_todas_as_colunas_canonicas(self) -> None:
        canonico = adaptar_crm1(self._df_crm1())
        self.assertEqual(list(canonico.columns), list(COLUNAS_CANONICAS))

    def test_valores_mapeados_corretamente(self) -> None:
        canonico = adaptar_crm1(self._df_crm1())
        linha = canonico.iloc[0]
        self.assertEqual(linha[CANON_RAZAO_STATUS], "Em contato")
        self.assertEqual(linha[CANON_NOME_COMPLETO], "João da Silva")
        self.assertEqual(linha[CANON_CPF], "12345678900")
        self.assertEqual(linha[CANON_EMAIL], "joao@exemplo.com")
        self.assertEqual(linha[CANON_TELEFONE], "(48) 99999-0001")
        self.assertEqual(linha[CANON_NOME_PRODUTO], "Produto X")
        self.assertEqual(linha[CANON_PRODUTO_PRIMARIO], "Linha Y")

    def test_regional_fica_vazia_no_crm1(self) -> None:
        canonico = adaptar_crm1(self._df_crm1())
        self.assertEqual(canonico.iloc[0][CANON_REGIONAL], "")

    def test_coluna_ausente_vira_vazia(self) -> None:
        df = pd.DataFrame({"Telefone da Empresa/Celular": ["48999990001"], "CPF_CNPJ": ["1"]})
        canonico = adaptar_crm1(df)
        self.assertEqual(canonico.iloc[0][CANON_NOME_COMPLETO], "")
        self.assertEqual(canonico.iloc[0][CANON_EMAIL], "")

    def test_adapter_registrado(self) -> None:
        self.assertIn(FORMATO_CRM1, adapters.ADAPTERS)


class DeteccaoCrm2Test(unittest.TestCase):
    def test_detecta_crm2_pela_assinatura(self) -> None:
        colunas = [
            "Regional/Unidade do Produto",
            "Razão do Status",
            "Nome Completo",
            "CPF",
            "E-mail Comercial",
            "Telefone Celular",
            "Telefone da Empresa",
            "Telefone Residencial",
        ]
        self.assertEqual(detectar_formato(colunas), FORMATO_CRM2)

    def test_crm1_tem_prioridade_sobre_crm2(self) -> None:
        # Planilha que contém colunas dos dois; CRM 1 deve vencer.
        colunas = [
            "CPF_CNPJ",
            "Telefone da Empresa/Celular",
            "Regional/Unidade do Produto",
            "Telefone Celular",
        ]
        self.assertEqual(detectar_formato(colunas), FORMATO_CRM1)


class AdaptarCrm2Test(unittest.TestCase):
    def _df_crm2(self, **overrides: object) -> pd.DataFrame:
        base = {
            "Regional/Unidade do Produto": ["SESI SC"],
            "Razão do Status": ["Em contato"],
            "Nome Completo": ["Maria Souza"],
            "Nome": ["Maria"],
            "Segundo Nome": [""],
            "Sobrenome": ["Souza"],
            "CPF": ["98765432100"],
            "E-mail Comercial": ["maria@empresa.com"],
            "Telefone Celular": ["(48) 98888-0002"],
            "Telefone da Empresa": ["(48) 3333-0003"],
            "Telefone Residencial": ["(48) 2222-0004"],
            "Nome do Produto": ["Curso A"],
            "Produto Primário (Nome do Produto) (Produto)": ["Linha Z"],
            "Cidade": ["Joinville"],
        }
        base.update({chave: [valor] for chave, valor in overrides.items()})
        return pd.DataFrame(base)

    def test_mapeia_todas_as_colunas_canonicas(self) -> None:
        canonico = adaptar_crm2(self._df_crm2())
        self.assertEqual(list(canonico.columns), list(COLUNAS_CANONICAS))

    def test_valores_e_regional_mapeados(self) -> None:
        linha = adaptar_crm2(self._df_crm2()).iloc[0]
        self.assertEqual(linha[CANON_RAZAO_STATUS], "Em contato")
        self.assertEqual(linha[CANON_CPF], "98765432100")
        self.assertEqual(linha[CANON_EMAIL], "maria@empresa.com")
        self.assertEqual(linha[CANON_REGIONAL], "SESI SC")
        self.assertEqual(linha[CANON_NOME_PRODUTO], "Curso A")
        self.assertEqual(linha[CANON_PRODUTO_PRIMARIO], "Linha Z")

    # --- Prioridade de telefone: Celular > Empresa > Residencial ---
    def test_prioridade_usa_celular_quando_presente(self) -> None:
        linha = adaptar_crm2(self._df_crm2()).iloc[0]
        self.assertEqual(linha[CANON_TELEFONE], "(48) 98888-0002")

    def test_prioridade_cai_para_empresa_sem_celular(self) -> None:
        df = self._df_crm2(**{"Telefone Celular": ""})
        self.assertEqual(adaptar_crm2(df).iloc[0][CANON_TELEFONE], "(48) 3333-0003")

    def test_prioridade_cai_para_residencial_sem_celular_e_empresa(self) -> None:
        df = self._df_crm2(**{"Telefone Celular": "", "Telefone da Empresa": ""})
        self.assertEqual(adaptar_crm2(df).iloc[0][CANON_TELEFONE], "(48) 2222-0004")

    def test_sem_nenhum_telefone_fica_vazio(self) -> None:
        df = self._df_crm2(
            **{"Telefone Celular": "", "Telefone da Empresa": "", "Telefone Residencial": ""}
        )
        self.assertEqual(adaptar_crm2(df).iloc[0][CANON_TELEFONE], "")

    # --- Fallback de Nome Completo ---
    def test_nome_completo_direto_quando_preenchido(self) -> None:
        linha = adaptar_crm2(self._df_crm2()).iloc[0]
        self.assertEqual(linha[CANON_NOME_COMPLETO], "Maria Souza")

    def test_nome_composto_quando_completo_vazio(self) -> None:
        df = self._df_crm2(
            **{"Nome Completo": "", "Nome": "JOÃO", "Segundo Nome": "DA", "Sobrenome": "SILVA"}
        )
        self.assertEqual(adaptar_crm2(df).iloc[0][CANON_NOME_COMPLETO], "JOÃO DA SILVA")

    def test_nome_composto_ignora_campos_vazios(self) -> None:
        df = self._df_crm2(
            **{"Nome Completo": "", "Nome": "JOÃO", "Segundo Nome": "", "Sobrenome": ""}
        )
        self.assertEqual(adaptar_crm2(df).iloc[0][CANON_NOME_COMPLETO], "JOÃO")

    def test_adapter_crm2_registrado(self) -> None:
        self.assertIn(FORMATO_CRM2, adapters.ADAPTERS)


class LocalizarTelefoneTest(unittest.TestCase):
    def test_reconhece_diversos_nomes(self) -> None:
        for nome in (
            "Telefone",
            "telefone",
            "TELEFONE",
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
        ):
            with self.subTest(nome=nome):
                self.assertEqual(localizar_coluna_telefone([nome]), nome)

    def test_tolera_acento_caixa_e_espacos(self) -> None:
        self.assertEqual(localizar_coluna_telefone(["  NÚMERO  CELULAR "]), "NÚMERO  CELULAR")
        self.assertEqual(localizar_coluna_telefone(["whatsapp"]), "whatsapp")

    def test_nao_reconhece_coluna_arbitraria(self) -> None:
        self.assertIsNone(localizar_coluna_telefone(["Cliente", "Documento", "Endereço"]))

    def test_retorna_a_primeira_reconhecida(self) -> None:
        self.assertEqual(
            localizar_coluna_telefone(["Nome", "Celular", "Telefone"]), "Celular"
        )


class DeteccaoGenericaTest(unittest.TestCase):
    def test_detecta_generico_por_telefone(self) -> None:
        self.assertEqual(detectar_formato(["Cliente", "CPF", "E-mail", "Telefone"]), FORMATO_GENERICO)

    def test_detecta_generico_apenas_com_celular(self) -> None:
        self.assertEqual(detectar_formato(["Nome", "Celular"]), FORMATO_GENERICO)

    def test_sem_coluna_telefone_levanta_erro(self) -> None:
        with self.assertRaises(FormatoDesconhecidoError):
            detectar_formato(["Cliente", "Documento", "Endereço"])


class AdaptarGenericoTest(unittest.TestCase):
    def test_mapeia_campos_reconhecidos(self) -> None:
        df = pd.DataFrame(
            {
                "Cliente": ["Ana"],
                "CPF": ["111"],
                "E-mail": ["ana@x.com"],
                "Telefone": ["(48) 99999-1234"],
            }
        )
        linha = adaptar_generico(df).iloc[0]
        self.assertEqual(list(adaptar_generico(df).columns), list(COLUNAS_CANONICAS))
        self.assertEqual(linha[CANON_NOME_COMPLETO], "Ana")
        self.assertEqual(linha[CANON_CPF], "111")
        self.assertEqual(linha[CANON_EMAIL], "ana@x.com")
        self.assertEqual(linha[CANON_TELEFONE], "(48) 99999-1234")

    def test_campos_opcionais_ausentes_ficam_vazios(self) -> None:
        df = pd.DataFrame({"Nome": ["Bia"], "Celular": ["48988887777"]})
        linha = adaptar_generico(df).iloc[0]
        self.assertEqual(linha[CANON_NOME_COMPLETO], "Bia")
        self.assertEqual(linha[CANON_TELEFONE], "48988887777")
        self.assertEqual(linha[CANON_CPF], "")
        self.assertEqual(linha[CANON_EMAIL], "")
        self.assertEqual(linha[CANON_RAZAO_STATUS], "")
        self.assertEqual(linha[CANON_REGIONAL], "")

    def test_apenas_coluna_telefone(self) -> None:
        df = pd.DataFrame({"Telefone": ["48999990000"]})
        self.assertEqual(adaptar_generico(df).iloc[0][CANON_TELEFONE], "48999990000")

    def test_sem_telefone_levanta_erro(self) -> None:
        df = pd.DataFrame({"Cliente": ["X"], "Documento": ["Y"]})
        with self.assertRaises(FormatoDesconhecidoError):
            adaptar_generico(df)

    def test_adapter_generico_registrado(self) -> None:
        self.assertIn(FORMATO_GENERICO, adapters.ADAPTERS)


if __name__ == "__main__":
    unittest.main()
