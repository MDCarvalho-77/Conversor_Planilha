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
    MAX_BYTES_POR_ARQUIVO,
    MAX_CONTATOS_POR_ARQUIVO,
    BlipError,
    gerar_blip_de_planilhas,
    gerar_csvs_do_excel_revisado,
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

# Formatos clássicos da V1 (para um único arquivo CRM 1 usam o pipeline V1).
FORMATOS_CLASSICOS = (
    FORMATO_PLANILHA_COMPLETA,
    FORMATO_SOMENTE_NUMEROS,
    FORMATO_AMBAS_SAIDAS,
)


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
        self.geometry("720x860")
        self.minsize(720, 820)
        self.caminhos_entrada: list[Path] = []
        self.pasta_saida: Path | None = None
        self.processando = False
        self.botoes_formato: list[ctk.CTkRadioButton] = []
        self.valores_metricas: dict[str, ctk.CTkLabel] = {}
        self.grid_columnconfigure(0, weight=1)

        self._criar_cabecalho()
        self._criar_area_planilha()
        self._criar_area_formato()
        self._criar_area_revisao_blip()
        self._criar_area_resultado()

    def _criar_cabecalho(self) -> None:
        """Monta a identificação institucional do aplicativo."""
        self.logo_original = Image.open(CAMINHO_LOGOMARCA)
        self.logo_exibicao = ctk.CTkImage(
            light_image=self.logo_original,
            dark_image=self.logo_original,
            size=(345, 50),
        )
        ctk.CTkLabel(self, text="", image=self.logo_exibicao).grid(
            row=0, column=0, pady=(22, 8)
        )
        ctk.CTkLabel(
            self, text="Conversor CRM", font=ctk.CTkFont(size=27, weight="bold")
        ).grid(row=1, column=0, pady=(0, 1))
        ctk.CTkLabel(
            self,
            text="Tratamento de dados do CRM",
            text_color=TEXTO_SECUNDARIO,
            font=ctk.CTkFont(size=14),
        ).grid(row=2, column=0, pady=(0, 16))

    def _criar_area_planilha(self) -> None:
        area = self._nova_secao(f"PLANILHAS (até {LIMITE_ARQUIVOS})", 3)
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
        area = self._nova_secao("FORMATO DE SAÍDA", 4)
        self.formato_saida = StringVar(value=FORMATO_PLANILHA_COMPLETA)
        opcoes = (
            (FORMATO_PLANILHA_COMPLETA, "Planilha com todos os dados tratados."),
            (FORMATO_SOMENTE_NUMEROS, "Lista pronta com apenas os telefones."),
            (FORMATO_AMBAS_SAIDAS, "Gera a planilha completa e a lista de telefones."),
            (
                FORMATO_PLANILHA_ESTENDIDA,
                "Planilha completa + Regional/Unidade e Nome do Produto.",
            ),
            (
                FORMATO_BLIP,
                "Gera CSVs (até 100 contatos / 100 KB) + Excel de revisão.",
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
            self,
            text="PROCESSAR PLANILHA",
            command=self.processar,
            state="disabled",
            width=290,
            height=42,
            fg_color=AZUL_PRINCIPAL,
            hover_color=AZUL_DESTAQUE,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.botao_processar.grid(row=5, column=0, pady=(0, 13))

    def _criar_area_revisao_blip(self) -> None:
        area = self._nova_secao("BLIP — GERAR CSVs DA PLANILHA REVISADA", 6)
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
        area = self._nova_secao("RESULTADO", 7)
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
        self.status = ctk.CTkLabel(self, text="", text_color=TEXTO_SECUNDARIO, wraplength=600)
        self.status.grid(row=8, column=0, pady=(0, 15))

    def _nova_secao(self, titulo: str, linha: int) -> ctk.CTkFrame:
        """Cria um painel visual discreto para agrupar uma etapa do fluxo."""
        area = ctk.CTkFrame(self, fg_color=FUNDO_SECAO, corner_radius=10)
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

        # No modo Blip, a pasta de saída é escolhida pelo usuário (na thread
        # principal) antes de iniciar o processamento pesado.
        pasta_blip: Path | None = None
        if formato == FORMATO_BLIP:
            escolha = filedialog.askdirectory(
                title="Escolha a pasta de saída para os arquivos do Blip"
            )
            if not escolha:
                return
            pasta_blip = Path(escolha)

        self._bloquear_interface("PROCESSANDO...")
        self.status.configure(text="Processando suas planilhas...")
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

        Roteamento: um único arquivo CRM 1 com formato clássico usa o pipeline
        V1 intacto. O modo Blip usa o pipeline V2 + preparação Blip. Os demais
        casos usam o pipeline V2/canônico.
        """
        try:
            if formato == FORMATO_BLIP:
                assert pasta_blip is not None
                resultado = gerar_blip_de_planilhas(caminhos, pasta_blip)
            else:
                usa_v1 = (
                    len(caminhos) == 1
                    and formato in FORMATOS_CLASSICOS
                    and _e_crm1(caminhos[0])
                )
                if usa_v1:
                    _validar_estrutura_crm(caminhos[0])
                    resultado = gerar_saidas(caminhos[0], formato)
                else:
                    resultado = gerar_saidas_v2(caminhos, formato)
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
            self.status.configure(text=mensagem)
            messagebox.showerror("Não foi possível processar", mensagem, parent=self)
            return
        assert resultado is not None
        if "lotes" in resultado:  # resultado do modo Blip
            self._mostrar_resultado_blip(resultado)
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
        self.status.configure(text="Processamento concluído.")

    def _mostrar_resultado_blip(self, resumo: dict[str, object]) -> None:
        """Exibe o resumo do modo Blip (lotes, contatos, inválidos)."""
        metricas = resumo.get("metricas_processamento", {})
        if isinstance(metricas, dict):
            for chave, valor in self.valores_metricas.items():
                self.valores_metricas[chave].configure(text=str(metricas.get(chave, "—")))

        arquivos = [Path(a) for a in resumo["arquivos_gerados"]]  # type: ignore[index]
        excel = resumo.get("excel_revisao")
        linhas = [
            f"Contatos exportados: {resumo['contatos_exportados']}",
            f"Lotes/CSVs gerados: {resumo['lotes']}",
        ]
        if resumo.get("contatos_sem_nome"):
            linhas.append(f"Contatos sem nome (mantidos): {resumo['contatos_sem_nome']}")
        if resumo.get("total_invalidos"):
            linhas.append(f"Registros inválidos (não exportados): {resumo['total_invalidos']}")
        if excel is not None:
            linhas.append(f"Excel de revisão: {Path(excel).name}")
        if arquivos:
            exemplos = ", ".join(a.name for a in arquivos[:3])
            reticencias = "..." if len(arquivos) > 3 else ""
            linhas.append(f"CSVs: {exemplos}{reticencias}")

        self.arquivos_gerados.configure(text="\n".join(linhas))
        if arquivos:
            self.pasta_saida = arquivos[0].parent
        elif excel is not None:
            self.pasta_saida = Path(excel).parent
        if self.pasta_saida is not None:
            self.botao_abrir_pasta.configure(state="normal")
        self.status.configure(text="Preparação Blip concluída. Revise o Excel antes do upload.")

    def gerar_csvs_revisados(self) -> None:
        """Gera os CSVs definitivos a partir de um Excel de revisão editado."""
        if self.processando:
            return
        caminho_excel = filedialog.askopenfilename(
            title="Selecionar o Excel revisado do Blip",
            filetypes=[("Planilhas Excel", "*.xlsx")],
        )
        if not caminho_excel:
            return
        escolha_pasta = filedialog.askdirectory(
            title="Escolha a pasta de saída para os CSVs definitivos"
        )
        if not escolha_pasta:
            return

        self._bloquear_interface("GERANDO CSVs...")
        self.status.configure(text="Gerando CSVs a partir da planilha revisada...")
        Thread(
            target=self._executar_revisao,
            args=(Path(caminho_excel), Path(escolha_pasta)),
            daemon=True,
        ).start()

    def _executar_revisao(self, caminho_excel: Path, pasta_saida: Path) -> None:
        """Gera os CSVs definitivos fora da thread da interface."""
        try:
            resumo = gerar_csvs_do_excel_revisado(caminho_excel, pasta_saida)
        except Exception as erro:
            self.after(0, self._concluir_processamento, None, erro)
        else:
            self.after(0, self._concluir_processamento, resumo, None)

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
