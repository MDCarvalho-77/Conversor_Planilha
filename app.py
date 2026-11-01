"""Interface desktop do Conversor CRM."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from threading import Thread
from tkinter import StringVar, filedialog, messagebox

import customtkinter as ctk
from openpyxl import load_workbook
from PIL import Image

import adapters
from adapters import FormatoDesconhecidoError
from processamento import (
    COLUNAS_PRODUTO,
    COLUNAS_SAIDA,
    COLUNA_TELEFONE,
    FORMATO_AMBAS_SAIDAS,
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    gerar_saidas,
)
from pipeline_v2 import (
    FORMATO_PLANILHA_ESTENDIDA,
    LIMITE_ARQUIVOS,
    LimiteArquivosExcedidoError,
    gerar_saidas_v2,
)
import blip
from blip import (
    FORMATO_BLIP,
    GRUPO_DDD_47,
    GRUPO_DDD_48_OUTROS,
    GRUPO_DDD_49,
    GRUPO_SEM_DDD,
    GRUPOS_DDD,
    MAX_BYTES_POR_ARQUIVO,
    MAX_CONTATOS_POR_ARQUIVO,
    BlipError,
    gerar_blip_de_planilhas,
    gerar_blip_por_ddd_de_planilhas,
    gerar_csvs_do_excel_revisado,
    gerar_csvs_por_ddd_de_revisao,
)


MENSAGEM_ARQUIVO_INCOMPATIVEL = (
    "Arquivo incompatível. Selecione uma planilha exportada diretamente do CRM."
)
MENSAGEM_SEM_TELEFONE = (
    "Não foi possível identificar uma coluna de telefone nesta planilha.\n"
    "Para processar uma planilha genérica, a coluna que contém os números "
    "precisa ter um nome que indique telefone, como: Telefone, Celular, "
    "Telefone Celular, Telefone Residencial ou Telefone da Empresa."
)
MENSAGEM_LIMITE_ARQUIVOS = (
    f"Selecione no máximo {LIMITE_ARQUIVOS} arquivos por vez. "
    "Ajuste a seleção e tente novamente."
)
NOTA_BLIP = (
    f"Blip: até {MAX_CONTATOS_POR_ARQUIVO} contatos e "
    f"{MAX_BYTES_POR_ARQUIVO // 1024} KB por CSV. "
    "São gerados os CSVs e um Excel de revisão para edição manual."
)
PASTA_RECURSOS = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
CAMINHO_LOGOMARCA = PASTA_RECURSOS / "assets" / "marca-SESI-SENAI_reduzida_cor.png"
AZUL_PRINCIPAL = "#1479B8"
AZUL_DESTAQUE = "#0D639A"
FUNDO_SECAO = ("#E9F2F8", "#1C2730")
FUNDO_CARTAO = ("#F7FAFC", "#26343E")
TEXTO_SECUNDARIO = ("#52616B", "#BAC7D0")
VERDE_SUCESSO = ("#1E7E34", "#4ADE80")

# Formatos clássicos da V1 (para um único arquivo CRM 1 usam o pipeline V1).
FORMATOS_CLASSICOS = (
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
)

# Identificadores estáveis das duas opções de saída apresentadas ao usuário
# (V3.1). O roteamento interno continua usando as constantes FORMATO_* e as
# funções subjacentes, que permanecem disponíveis para compatibilidade/testes.
OPCAO_COMPLETA_NUMEROS = "Planilha completa + números convertidos"
OPCAO_BLIP_DDD = "Blip — listas de contatos (separadas por DDD)"


def _colunas_visiveis(caminho: Path) -> set[str]:
    """Lê os cabeçalhos das abas visíveis (ignora hidden/veryHidden)."""
    try:
        workbook = load_workbook(caminho, read_only=True, data_only=False)
    except Exception as erro:
        raise ValueError(MENSAGEM_ARQUIVO_INCOMPATIVEL) from erro
    try:
        colunas: set[str] = set()
        for aba in workbook.worksheets:
            if aba.sheet_state == "visible":
                cabecalho = next(aba.iter_rows(min_row=1, max_row=1, values_only=True), ())
                colunas.update(str(valor).strip() for valor in cabecalho if valor is not None)
    finally:
        workbook.close()
    return colunas


def _validar_estrutura_crm(caminho: Path) -> None:
    """Valida a estrutura do CRM 1 sem modificar dados ou regras (pipeline V1)."""
    colunas = _colunas_visiveis(caminho)
    obrigatorias = set(COLUNAS_SAIDA) | set(COLUNAS_PRODUTO) | {COLUNA_TELEFONE}
    if not obrigatorias.issubset(colunas):
        raise ValueError(MENSAGEM_ARQUIVO_INCOMPATIVEL)


def _e_crm1(caminho: Path) -> bool:
    """Indica se o arquivo é detectado como CRM 1 (para roteamento V1)."""
    try:
        return adapters.detectar_formato(_colunas_visiveis(caminho)) == adapters.FORMATO_CRM1
    except FormatoDesconhecidoError:
        return False


class ConversorPlanilhasApp(ctk.CTk):
    """Janela principal do Conversor CRM."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Conversor CRM")
        self.geometry("720x760")
        self.minsize(720, 560)
        self.caminhos_entrada: list[Path] = []
        self.pasta_saida: Path | None = None
        self.processando = False
        self._operacao_atual = ""
        self.botoes_formato: list[ctk.CTkRadioButton] = []
        self.valores_metricas: dict[str, ctk.CTkLabel] = {}
        self.grid_columnconfigure(0, weight=1)
        # Cabeçalho fixo (row 0); o corpo rolável (row 1) cresce com a janela.
        self.grid_rowconfigure(1, weight=1)

        self._criar_cabecalho()
        self.corpo = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.corpo.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.corpo.grid_columnconfigure(0, weight=1)
        self._criar_area_planilha()
        self._criar_area_formato()
        self._criar_area_revisao_blip()
        self._criar_area_resultado()

    def _criar_cabecalho(self) -> None:
        """Monta a identificação institucional do aplicativo."""
        cabecalho = ctk.CTkFrame(self, fg_color="transparent")
        cabecalho.grid(row=0, column=0, sticky="ew")
        cabecalho.grid_columnconfigure(0, weight=1)
        self.logo_original = Image.open(CAMINHO_LOGOMARCA)
        self.logo_exibicao = ctk.CTkImage(
            light_image=self.logo_original,
            dark_image=self.logo_original,
            size=(345, 50),
        )
        ctk.CTkLabel(cabecalho, text="", image=self.logo_exibicao).grid(
            row=0, column=0, pady=(22, 8)
        )
        ctk.CTkLabel(
            cabecalho, text="Conversor CRM", font=ctk.CTkFont(size=27, weight="bold")
        ).grid(row=1, column=0, pady=(0, 1))
        ctk.CTkLabel(
            cabecalho,
            text="Tratamento de dados do CRM",
            text_color=TEXTO_SECUNDARIO,
            font=ctk.CTkFont(size=14),
        ).grid(row=2, column=0, pady=(0, 16))

    def _criar_area_planilha(self) -> None:
        area = self._nova_secao(f"PLANILHAS (até {LIMITE_ARQUIVOS})", 0)
        self.arquivo_selecionado = ctk.CTkLabel(
            area,
            text="Nenhuma planilha selecionada",
            text_color=TEXTO_SECUNDARIO,
            font=ctk.CTkFont(size=14),
            justify="left",
            wraplength=560,
        )
        self.arquivo_selecionado.grid(row=1, column=0, padx=18, pady=(0, 10))
        self.botao_selecionar = ctk.CTkButton(
            area,
            text="SELECIONAR PLANILHAS",
            command=self.selecionar_arquivo,
            width=230,
            fg_color=AZUL_PRINCIPAL,
            hover_color=AZUL_DESTAQUE,
            font=ctk.CTkFont(weight="bold"),
        )
        self.botao_selecionar.grid(row=2, column=0, padx=18, pady=(0, 15))

    def _criar_area_formato(self) -> None:
        area = self._nova_secao("FORMATO DE SAÍDA", 1)
        self.formato_saida = StringVar(value=OPCAO_COMPLETA_NUMEROS)
        opcoes = (
            (
                OPCAO_COMPLETA_NUMEROS,
                "Gera a planilha completa consolidada e a lista de "
                "telefones convertidos.",
            ),
            (
                OPCAO_BLIP_DDD,
                "Separa automaticamente por grupo de DDD (47, 49, 48/outros "
                "e SEM_DDD): CSVs e Excel de revisão por grupo.",
            ),
        )
        for indice, (formato, descricao) in enumerate(opcoes):
            coluna = indice % 2
            linha = 1 + indice // 2
            area.grid_columnconfigure(coluna, weight=1)
            opcao = ctk.CTkFrame(area, fg_color="transparent")
            opcao.grid(row=linha, column=coluna, padx=8, pady=(0, 13), sticky="nsew")
            botao = ctk.CTkRadioButton(
                opcao,
                text=formato,
                variable=self.formato_saida,
                value=formato,
                font=ctk.CTkFont(size=13, weight="bold"),
                fg_color=AZUL_PRINCIPAL,
                hover_color=AZUL_DESTAQUE,
            )
            botao.grid(row=0, column=0, sticky="w")
            ctk.CTkLabel(
                opcao,
                text=descricao,
                text_color=TEXTO_SECUNDARIO,
                justify="left",
                wraplength=250,
                font=ctk.CTkFont(size=12),
            ).grid(row=1, column=0, padx=(25, 0), pady=(3, 0), sticky="w")
            self.botoes_formato.append(botao)

        self.botao_processar = ctk.CTkButton(
            self.corpo,
            text="PROCESSAR PLANILHA",
            command=self.processar,
            state="disabled",
            width=290,
            height=42,
            fg_color=AZUL_PRINCIPAL,
            hover_color=AZUL_DESTAQUE,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.botao_processar.grid(row=2, column=0, pady=(0, 13))

    def _criar_area_revisao_blip(self) -> None:
        area = self._nova_secao("BLIP — GERAR CSVs DA PLANILHA REVISADA", 3)
        ctk.CTkLabel(
            area,
            text=(
                "Depois de revisar o Excel de contatos, gere os CSVs definitivos. "
                "Os limites de 100 contatos e 100 KB são reaplicados."
            ),
            text_color=TEXTO_SECUNDARIO,
            justify="left",
            wraplength=560,
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, padx=18, pady=(0, 10))
        self.botao_revisar = ctk.CTkButton(
            area,
            text="GERAR CSVs A PARTIR DA PLANILHA REVISADA",
            command=self.gerar_csvs_revisados,
            width=320,
            fg_color="transparent",
            border_width=1,
            border_color=AZUL_PRINCIPAL,
            text_color=(AZUL_DESTAQUE, "#54B7ED"),
            hover_color=("#D9EAF5", "#253B49"),
            font=ctk.CTkFont(weight="bold"),
        )
        self.botao_revisar.grid(row=2, column=0, padx=18, pady=(0, 14))

    def _criar_area_resultado(self) -> None:
        area = self._nova_secao("RESULTADO", 4)
        nomes = (
            ("Registros lidos", "registros_lidos"),
            ("Sem telefone", "removidos_sem_telefone"),
            ("Duplicados", "duplicados_removidos"),
            ("Registros finais", "registros_finais"),
        )
        metricas_frame = ctk.CTkFrame(area, fg_color="transparent")
        metricas_frame.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")
        for coluna, (titulo, chave) in enumerate(nomes):
            metricas_frame.grid_columnconfigure(coluna, weight=1)
            cartao = ctk.CTkFrame(metricas_frame, fg_color=FUNDO_CARTAO, corner_radius=8)
            cartao.grid(row=0, column=coluna, padx=4, sticky="ew")
            ctk.CTkLabel(
                cartao, text=titulo, text_color=TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11)
            ).pack(padx=8, pady=(7, 0))
            valor = ctk.CTkLabel(cartao, text="—", font=ctk.CTkFont(size=20, weight="bold"))
            valor.pack(padx=8, pady=(0, 7))
            self.valores_metricas[chave] = valor
        self.arquivos_gerados = ctk.CTkLabel(
            area,
            text="Os arquivos gerados aparecerão aqui.",
            text_color=TEXTO_SECUNDARIO,
            justify="left",
            anchor="w",
            wraplength=560,
        )
        self.arquivos_gerados.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.botao_abrir_pasta = ctk.CTkButton(
            area,
            text="ABRIR PASTA",
            command=self.abrir_pasta,
            state="disabled",
            width=180,
            fg_color="transparent",
            border_width=1,
            border_color=AZUL_PRINCIPAL,
            text_color=(AZUL_DESTAQUE, "#54B7ED"),
            hover_color=("#D9EAF5", "#253B49"),
            font=ctk.CTkFont(weight="bold"),
        )
        self.botao_abrir_pasta.grid(row=3, column=0, pady=(0, 14))
        self.status = ctk.CTkLabel(
            self.corpo, text="", text_color=TEXTO_SECUNDARIO, wraplength=600
        )
        self.status.grid(row=5, column=0, pady=(0, 15))

    def _nova_secao(self, titulo: str, linha: int) -> ctk.CTkFrame:
        """Cria um painel visual discreto para agrupar uma etapa do fluxo."""
        area = ctk.CTkFrame(self.corpo, fg_color=FUNDO_SECAO, corner_radius=10)
        area.grid(row=linha, column=0, padx=42, pady=(0, 12), sticky="ew")
        area.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            area,
            text=titulo,
            text_color=(AZUL_DESTAQUE, "#58BCEB"),
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, padx=18, pady=(12, 8), sticky="w")
        return area

    def selecionar_arquivo(self) -> None:
        """Seleciona uma ou mais planilhas e remove resultados anteriores."""
        caminhos = filedialog.askopenfilenames(
            title="Selecionar planilhas do CRM", filetypes=[("Planilhas Excel", "*.xlsx")]
        )
        if not caminhos:
            return
        if len(caminhos) > LIMITE_ARQUIVOS:
            messagebox.showwarning("Limite de arquivos", MENSAGEM_LIMITE_ARQUIVOS, parent=self)
            return

        self.caminhos_entrada = [Path(caminho) for caminho in caminhos]
        self.pasta_saida = None
        self.arquivo_selecionado.configure(text=self._texto_selecao())
        self.botao_processar.configure(state="normal")
        self.botao_abrir_pasta.configure(state="disabled")
        for valor in self.valores_metricas.values():
            valor.configure(text="—")
        self.arquivos_gerados.configure(
            text="Os arquivos gerados aparecerão aqui.", text_color=TEXTO_SECUNDARIO
        )
        self.status.configure(text="")

    def _texto_selecao(self) -> str:
        """Descreve a seleção atual: quantidade e nomes dos arquivos."""
        total = len(self.caminhos_entrada)
        cabecalho = f"{total} planilha(s) selecionada(s):"
        nomes = "\n".join(f"• {caminho.name}" for caminho in self.caminhos_entrada)
        return f"{cabecalho}\n{nomes}"

    def processar(self) -> None:
        """Inicia o worker e impede processamento duplicado."""
        if not self.caminhos_entrada or self.processando:
            return
        formato = self.formato_saida.get()
        self._operacao_atual = (
            "blip_preparacao" if formato == OPCAO_BLIP_DDD else "padrao"
        )

        # No modo Blip, a pasta de saída é escolhida pelo usuário (na thread
        # principal) antes de iniciar o processamento pesado.
        pasta_blip: Path | None = None
        if formato == OPCAO_BLIP_DDD:
            escolha = filedialog.askdirectory(
                title="Escolha a pasta de saída para os arquivos do Blip"
            )
            if not escolha:
                return
            pasta_blip = Path(escolha)

        self._bloquear_interface("PROCESSANDO...")
        self.status.configure(text="Processando suas planilhas...", text_color=TEXTO_SECUNDARIO)
        Thread(
            target=self._executar_processamento,
            args=(list(self.caminhos_entrada), formato, pasta_blip),
            daemon=True,
        ).start()

    def _bloquear_interface(self, texto_processar: str) -> None:
        """Desabilita controles para impedir processamentos simultâneos."""
        self.processando = True
        self.botao_processar.configure(state="disabled", text=texto_processar)
        self.botao_selecionar.configure(state="disabled")
        self.botao_revisar.configure(state="disabled")
        for botao in self.botoes_formato:
            botao.configure(state="disabled")
        self.botao_abrir_pasta.configure(state="disabled")

    def _executar_processamento(
        self, caminhos: list[Path], formato: str, pasta_blip: Path | None
    ) -> None:
        """Executa leitura, validação e geração fora da thread da interface.

        Roteamento (V3.1): a Opção B (Blip por DDD) usa o pipeline V2 +
        preparação Blip separada por grupo de DDD. A Opção A (planilha completa
        + números convertidos) usa o pipeline V2/canônico com
        ``FORMATO_AMBAS_SAIDAS``, produzindo ``_tratado.xlsx`` + ``_telefones.xlsx``.
        """
        try:
            if formato == OPCAO_BLIP_DDD:
                assert pasta_blip is not None
                resultado = gerar_blip_por_ddd_de_planilhas(caminhos, pasta_blip)
            else:
                resultado = gerar_saidas_v2(caminhos, FORMATO_AMBAS_SAIDAS)
        except Exception as erro:
            self.after(0, self._concluir_processamento, None, erro)
        else:
            self.after(0, self._concluir_processamento, resultado, None)

    def _reabilitar_interface(self) -> None:
        """Reabilita os controles após o término do worker."""
        self.processando = False
        self.botao_selecionar.configure(state="normal")
        self.botao_processar.configure(
            state="normal" if self.caminhos_entrada else "disabled",
            text="PROCESSAR PLANILHA",
        )
        self.botao_revisar.configure(state="normal")
        for botao in self.botoes_formato:
            botao.configure(state="normal")

    def _concluir_processamento(
        self, resultado: dict[str, object] | None, erro: Exception | None
    ) -> None:
        """Atualiza widgets somente na thread principal após o worker terminar."""
        self._reabilitar_interface()
        if erro is not None:
            mensagem = self._mensagem_amigavel(erro)
            self.status.configure(text=mensagem, text_color=TEXTO_SECUNDARIO)
            messagebox.showerror("Não foi possível processar", mensagem, parent=self)
            return
        assert resultado is not None
        if self._operacao_atual == "blip_preparacao":
            self._mostrar_resultado_blip(resultado, revisao=False)
        elif self._operacao_atual == "blip_revisao":
            self._mostrar_resultado_blip(resultado, revisao=True)
        else:
            self._mostrar_resultado_padrao(resultado)

    def _mostrar_resultado_padrao(self, resultado: dict[str, object]) -> None:
        """Exibe métricas e arquivos das saídas clássicas/estendida."""
        for chave, valor in self.valores_metricas.items():
            valor.configure(text=str(resultado[chave]))
        arquivos = [Path(arquivo) for arquivo in resultado["arquivos_gerados"]]  # type: ignore[index]
        self.arquivos_gerados.configure(
            text="Arquivos gerados:\n" + "\n".join(f"• {arquivo.name}" for arquivo in arquivos),
        )
        self.pasta_saida = arquivos[0].parent
        self.botao_abrir_pasta.configure(state="normal")
        self.status.configure(
            text="Processamento concluído com sucesso!", text_color=VERDE_SUCESSO
        )
        texto = self._texto_sucesso_padrao(resultado, self.pasta_saida)
        messagebox.showinfo("Processamento concluído", texto, parent=self)

    def _mostrar_resultado_blip(
        self, resumo: dict[str, object], revisao: bool = False
    ) -> None:
        """Exibe o resumo do modo Blip por grupo de DDD (contatos, CSVs, SEM_DDD)."""
        metricas = resumo.get("metricas_processamento", {})
        if isinstance(metricas, dict):
            for chave, valor in self.valores_metricas.items():
                self.valores_metricas[chave].configure(text=str(metricas.get(chave, "—")))

        totais_por_grupo = resumo.get("totais_por_grupo", {})
        por_grupo = resumo.get("por_grupo", {})
        linhas = [
            f"Total de contatos processados: {resumo.get('total_contatos', '—')}",
            f"Contatos exportados: {resumo.get('contatos_exportados', '—')}",
        ]
        if revisao and "mudaram_de_grupo" in resumo:
            linhas.append(f"Mudaram de grupo após revisão: {resumo['mudaram_de_grupo']}")
        if isinstance(totais_por_grupo, dict):
            for grupo in GRUPOS_DDD:
                total_grupo = totais_por_grupo.get(grupo, 0)
                csvs_grupo = 0
                if isinstance(por_grupo, dict) and grupo in por_grupo:
                    detalhe = por_grupo[grupo]
                    if isinstance(detalhe, dict):
                        csvs_grupo = detalhe.get("lotes", 0)
                linhas.append(f"{grupo}: {total_grupo} contato(s), {csvs_grupo} CSV(s)")
        linhas.append(f"SEM_DDD: {resumo.get('total_sem_ddd', 0)} contato(s)")
        linhas.append(f"CSVs/lotes no total: {resumo.get('lotes', '—')}")
        if resumo.get("total_invalidos"):
            linhas.append(f"Registros inválidos (não exportados): {resumo['total_invalidos']}")

        pasta_raiz = resumo.get("pasta_raiz")
        if pasta_raiz is not None:
            linhas.append(f"Pasta Blip: {Path(pasta_raiz)}")  # type: ignore[arg-type]

        self.arquivos_gerados.configure(text="\n".join(linhas))
        if pasta_raiz is not None:
            self.pasta_saida = Path(pasta_raiz)  # type: ignore[arg-type]
        else:
            arquivos = [Path(a) for a in resumo.get("arquivos_gerados", [])]  # type: ignore[arg-type]
            if arquivos:
                self.pasta_saida = arquivos[0].parent
        if self.pasta_saida is not None:
            self.botao_abrir_pasta.configure(state="normal")

        if revisao:
            titulo = "CSVs finais gerados"
            status_texto = "CSVs finais gerados com sucesso!"
            texto = self._texto_sucesso_blip_revisao_ddd(resumo)
        else:
            titulo = "Preparação Blip concluída"
            status_texto = "Preparação para o Blip concluída com sucesso!"
            texto = self._texto_sucesso_blip_preparacao_ddd(resumo)
        self.status.configure(text=status_texto, text_color=VERDE_SUCESSO)
        messagebox.showinfo(titulo, texto, parent=self)

    def gerar_csvs_revisados(self) -> None:
        """Gera os CSVs definitivos a partir de um Excel de revisão editado."""
        if self.processando:
            return
        caminhos_excel = filedialog.askopenfilenames(
            title="Selecionar os Excel revisados do Blip (um por grupo de DDD)",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if not caminhos_excel:
            return
        escolha_pasta = filedialog.askdirectory(
            title="Escolha a pasta de saída para os CSVs definitivos"
        )
        if not escolha_pasta:
            return

        self._operacao_atual = "blip_revisao"
        self._bloquear_interface("GERANDO CSVs...")
        self.status.configure(
            text="Gerando CSVs a partir das planilhas revisadas...",
            text_color=TEXTO_SECUNDARIO,
        )
        Thread(
            target=self._executar_revisao,
            args=([Path(caminho) for caminho in caminhos_excel], Path(escolha_pasta)),
            daemon=True,
        ).start()

    def _executar_revisao(self, caminhos_excel: list[Path], pasta_saida: Path) -> None:
        """Gera os CSVs definitivos por grupo de DDD fora da thread da interface."""
        try:
            resumo = gerar_csvs_por_ddd_de_revisao(caminhos_excel, pasta_saida)
        except Exception as erro:
            self.after(0, self._concluir_processamento, None, erro)
        else:
            self.after(0, self._concluir_processamento, resumo, None)

    @staticmethod
    def _texto_sucesso_padrao(resultado: dict[str, object], pasta: Path | None) -> str:
        """Monta o texto de sucesso do fluxo tradicional/estendido (puro)."""
        arquivos = [Path(arquivo) for arquivo in resultado.get("arquivos_gerados", [])]  # type: ignore[arg-type]
        nomes = ", ".join(arquivo.name for arquivo in arquivos) if arquivos else "—"
        if arquivos:
            pasta_final = arquivos[0].parent
        else:
            pasta_final = pasta
        pasta_texto = str(pasta_final) if pasta_final is not None else "—"
        return (
            "Processamento concluído com sucesso!\n"
            f"Registros finais: {resultado.get('registros_finais', '—')}\n"
            f"Arquivos gerados: {nomes}\n"
            f"Salvos em: {pasta_texto}"
        )

    @staticmethod
    def _texto_sucesso_blip_preparacao(resumo: dict[str, object]) -> str:
        """Monta o texto de sucesso da preparação para o Blip (puro)."""
        arquivos = [Path(a) for a in resumo.get("arquivos_gerados", [])]  # type: ignore[arg-type]
        excel = resumo.get("excel_revisao")
        if arquivos:
            pasta_csvs = str(arquivos[0].parent)
        elif excel is not None:
            pasta_csvs = str(Path(excel).parent)  # type: ignore[arg-type]
        else:
            pasta_csvs = "—"
        excel_texto = str(Path(excel)) if excel is not None else "—"  # type: ignore[arg-type]
        return (
            "Preparação para o Blip concluída com sucesso!\n"
            f"Contatos processados: {resumo.get('contatos_exportados', '—')}\n"
            f"CSVs gerados: {resumo.get('lotes', '—')}\n"
            f"Registros inválidos: {resumo.get('total_invalidos', 0)}\n"
            f"CSVs salvos em: {pasta_csvs}\n"
            f"Excel de revisão: {excel_texto}"
        )

    @staticmethod
    def _texto_sucesso_blip_revisao(resumo: dict[str, object]) -> str:
        """Monta o texto de sucesso da geração final de CSVs (puro)."""
        arquivos = [Path(a) for a in resumo.get("arquivos_gerados", [])]  # type: ignore[arg-type]
        pasta_csvs = str(arquivos[0].parent) if arquivos else "—"
        return (
            "CSVs finais gerados com sucesso!\n"
            f"Contatos exportados: {resumo.get('contatos_exportados', '—')}\n"
            f"CSVs gerados: {resumo.get('lotes', '—')}\n"
            f"Registros inválidos: {resumo.get('total_invalidos', 0)}\n"
            f"Salvos em: {pasta_csvs}"
        )

    @staticmethod
    def _linhas_grupos_ddd(resumo: dict[str, object]) -> str:
        """Monta as linhas 'GRUPO: N contato(s), M CSV(s)' do resumo por DDD (puro)."""
        totais_por_grupo = resumo.get("totais_por_grupo", {})
        por_grupo = resumo.get("por_grupo", {})
        linhas: list[str] = []
        if isinstance(totais_por_grupo, dict):
            for grupo in GRUPOS_DDD:
                total_grupo = totais_por_grupo.get(grupo, 0)
                csvs_grupo = 0
                if isinstance(por_grupo, dict) and isinstance(por_grupo.get(grupo), dict):
                    csvs_grupo = por_grupo[grupo].get("lotes", 0)  # type: ignore[union-attr]
                linhas.append(f"{grupo}: {total_grupo} contato(s), {csvs_grupo} CSV(s)")
        return "\n".join(linhas)

    @staticmethod
    def _texto_sucesso_blip_preparacao_ddd(resumo: dict[str, object]) -> str:
        """Monta o texto de sucesso da preparação Blip por grupo de DDD (puro)."""
        pasta_raiz = resumo.get("pasta_raiz")
        pasta_texto = str(Path(pasta_raiz)) if pasta_raiz is not None else "—"  # type: ignore[arg-type]
        linhas_grupos = ConversorPlanilhasApp._linhas_grupos_ddd(resumo)
        return (
            "Preparação para o Blip concluída com sucesso!\n"
            f"Total de contatos processados: {resumo.get('total_contatos', '—')}\n"
            f"{linhas_grupos}\n"
            f"SEM_DDD: {resumo.get('total_sem_ddd', 0)}\n"
            f"CSVs gerados (total): {resumo.get('lotes', '—')}\n"
            f"Pasta Blip: {pasta_texto}"
        )

    @staticmethod
    def _texto_sucesso_blip_revisao_ddd(resumo: dict[str, object]) -> str:
        """Monta o texto de sucesso da geração de CSVs revisados por DDD (puro)."""
        pasta_raiz = resumo.get("pasta_raiz")
        pasta_texto = str(Path(pasta_raiz)) if pasta_raiz is not None else "—"  # type: ignore[arg-type]
        linhas_grupos = ConversorPlanilhasApp._linhas_grupos_ddd(resumo)
        return (
            "CSVs finais gerados com sucesso!\n"
            f"Total de contatos exportados: {resumo.get('contatos_exportados', '—')}\n"
            f"Contatos que mudaram de grupo: {resumo.get('mudaram_de_grupo', 0)}\n"
            f"{linhas_grupos}\n"
            f"SEM_DDD: {resumo.get('total_sem_ddd', 0)}\n"
            f"CSVs gerados (total): {resumo.get('lotes', '—')}\n"
            f"Pasta Blip: {pasta_texto}"
        )

    @staticmethod
    def _mensagem_amigavel(erro: Exception) -> str:
        if isinstance(erro, LimiteArquivosExcedidoError):
            return MENSAGEM_LIMITE_ARQUIVOS
        if isinstance(erro, FormatoDesconhecidoError):
            return MENSAGEM_SEM_TELEFONE
        if isinstance(erro, BlipError):
            return str(erro)
        if isinstance(erro, PermissionError):
            return "Não foi possível salvar o arquivo. Feche-o no Excel e tente novamente."
        if isinstance(erro, FileNotFoundError):
            return "Arquivo não encontrado. Selecione-o novamente e tente outra vez."
        return MENSAGEM_ARQUIVO_INCOMPATIVEL

    def abrir_pasta(self) -> None:
        """Abre no Explorador do Windows a pasta da última saída."""
        if self.pasta_saida is not None:
            os.startfile(self.pasta_saida)  # type: ignore[attr-defined]


def main() -> None:
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = ConversorPlanilhasApp()
    app.mainloop()


if __name__ == "__main__":
    main()
