"""Módulo do tutorial da V3.3 (REQUISITO B).

Responsabilidades:
- Persistência à prova de falha da preferência "não mostrar novamente" em
  ``%APPDATA%/ConversorCRM/config.json`` (com fallback para ``Path.home()``
  quando ``APPDATA`` não existe no ambiente).
- Lógica de navegação das 4 etapas, pura e testável sem GUI.
- Janela de tutorial em ``CTkToplevel`` modal leve (não instanciada nos testes).

Decisões de projeto:
- A leitura/gravação da preferência NUNCA propaga exceção de IO/parse: em caso
  de falha, ``carregar_preferencias`` retorna os defaults e ``deve_mostrar_tutorial``
  devolve ``True`` (mostrar por padrão). ``salvar_preferencia_nao_mostrar`` silencia
  qualquer erro de IO.
- As cores e o caminho da logomarca são definidos localmente (ou recebidos por
  parâmetro em ``abrir_tutorial``) para NÃO importar ``app.py`` e manter os
  testes livres de Tk.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import NamedTuple, Optional

try:  # customtkinter só é necessário para abrir a janela, não para a lógica/testes
    import customtkinter as ctk
except Exception:  # pragma: no cover - ambiente de teste pode não ter Tk utilizável
    ctk = None


# ---------------------------------------------------------------------------
# Identidade visual (definida localmente para não importar app.py)
# ---------------------------------------------------------------------------
PASTA_RECURSOS = Path(getattr(__import__("sys"), "_MEIPASS", Path(__file__).resolve().parent))
CAMINHO_LOGOMARCA = PASTA_RECURSOS / "assets" / "marca-SESI-SENAI_reduzida_cor.png"
AZUL_PRINCIPAL = "#1479B8"
AZUL_DESTAQUE = "#0D639A"
TEXTO_SECUNDARIO = ("#52616B", "#BAC7D0")

# Variável de ambiente que permite aos testes redirecionar o config para um tmp dir.
ENV_OVERRIDE_CONFIG = "CONVERSOR_CRM_CONFIG_PATH"

# Defaults usados sempre que a leitura falha ou não há config.
PREFERENCIAS_PADRAO = {"nao_mostrar": False}


# ---------------------------------------------------------------------------
# Persistência de preferência (à prova de falha)
# ---------------------------------------------------------------------------
def _caminho_config(override: Optional[os.PathLike | str] = None) -> Path:
    """Resolve o caminho do ``config.json``.

    Ordem de resolução:
    1. ``override`` explícito (usado por testes).
    2. Variável de ambiente ``CONVERSOR_CRM_CONFIG_PATH`` (usada por testes).
    3. ``%APPDATA%/ConversorCRM/config.json``.
    4. Fallback: ``~/ConversorCRM/config.json`` quando ``APPDATA`` está ausente.
    """
    if override is not None:
        return Path(override)

    override_env = os.getenv(ENV_OVERRIDE_CONFIG)
    if override_env:
        return Path(override_env)

    appdata = os.getenv("APPDATA")
    if appdata:
        base = Path(appdata)
    else:
        base = Path.home()
    return base / "ConversorCRM" / "config.json"


def carregar_preferencias(override: Optional[os.PathLike | str] = None) -> dict:
    """Lê e faz parse do config. Qualquer falha retorna os defaults.

    Nunca propaga exceção de IO/parse.
    """
    try:
        caminho = _caminho_config(override)
        with open(caminho, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        if not isinstance(dados, dict):
            return dict(PREFERENCIAS_PADRAO)
        preferencias = dict(PREFERENCIAS_PADRAO)
        preferencias.update(dados)
        return preferencias
    except Exception:
        return dict(PREFERENCIAS_PADRAO)


def salvar_preferencia_nao_mostrar(
    valor: bool, override: Optional[os.PathLike | str] = None
) -> None:
    """Grava a preferência "não mostrar novamente". Qualquer falha é silenciada."""
    try:
        caminho = _caminho_config(override)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        preferencias = carregar_preferencias(override)
        preferencias["nao_mostrar"] = bool(valor)
        with open(caminho, "w", encoding="utf-8") as arquivo:
            json.dump(preferencias, arquivo, ensure_ascii=False, indent=2)
    except Exception:
        # À prova de falha: nunca propaga erro de IO.
        return


def deve_mostrar_tutorial(override: Optional[os.PathLike | str] = None) -> bool:
    """Retorna True quando não há preferência de "não mostrar".

    À prova de falha: se a leitura falhar, usa os defaults e retorna True.
    """
    return not carregar_preferencias(override).get("nao_mostrar", False)


# ---------------------------------------------------------------------------
# Conteúdo das 4 etapas (dados puros, sem GUI)
# ---------------------------------------------------------------------------
class EtapaTutorial(NamedTuple):
    titulo: str
    texto: str


ETAPAS: tuple[EtapaTutorial, ...] = (
    EtapaTutorial(
        titulo="Bem-vindo ao Conversor CRM",
        texto=(
            "Este assistente converte planilhas de contatos para os formatos "
            "aceitos pelo CRM. Use os botões Anterior e Próximo para navegar "
            "pelas etapas e conhecer cada recurso. Você pode reabrir este "
            "tutorial a qualquer momento pelo botão Ajuda / Tutorial no "
            "cabeçalho da janela principal."
        ),
    ),
    EtapaTutorial(
        titulo="Importação de arquivos",
        texto=(
            "Selecione as planilhas de origem no formato Excel. As modalidades "
            "de conversão que organizam contatos por arquivo aceitam até 7 "
            "arquivos por vez. A modalidade Unificar lista — Somente telefones "
            "NÃO tem limite de arquivos: você pode selecionar quantas planilhas "
            "precisar. As colunas são reconhecidas pelo nome do cabeçalho, "
            "então mantenha os títulos das colunas de telefone legíveis."
        ),
    ),
    EtapaTutorial(
        titulo="As quatro opções de saída",
        texto=(
            "O Conversor CRM oferece quatro modalidades:\n"
            "1. Conversão padrão de planilhas (V1);\n"
            "2. Conversão com pipeline de adaptação (V2);\n"
            "3. Geração de CSVs e Excel de revisão para o Blip, com "
            "classificação por DDD (V3);\n"
            "4. Unificar lista — Somente telefones: reúne os telefones de "
            "todas as planilhas selecionadas em uma única lista, sem limite de "
            "arquivos. Escolha a modalidade adequada ao seu objetivo antes de "
            "gerar o resultado."
        ),
    ),
    EtapaTutorial(
        titulo="Resultados e revisão",
        texto=(
            "Ao concluir, o aplicativo informa onde os arquivos foram salvos e "
            "apresenta um resumo do processamento. Telefones inválidos ou não "
            "classificáveis não são descartados em silêncio: você recebe um "
            "relatório com a quantidade e, quando possível, a identificação dos "
            "registros para conferência. Arquivos existentes nunca são "
            "sobrescritos sem a sua autorização: o aplicativo avisa e permite "
            "escolher outro destino ou confirmar a substituição."
        ),
    ),
)


# ---------------------------------------------------------------------------
# Controlador de navegação (puro, testável sem GUI)
# ---------------------------------------------------------------------------
class ControladorTutorial:
    """Controla o índice da etapa atual respeitando os limites 0..len-1."""

    def __init__(self, total_etapas: int = len(ETAPAS)) -> None:
        self._total = max(1, int(total_etapas))
        self._indice = 0

    @property
    def indice(self) -> int:
        return self._indice

    @property
    def total(self) -> int:
        return self._total

    def proximo(self) -> int:
        if self._indice < self._total - 1:
            self._indice += 1
        return self._indice

    def anterior(self) -> int:
        if self._indice > 0:
            self._indice -= 1
        return self._indice

    @property
    def primeiro(self) -> bool:
        return self._indice == 0

    @property
    def ultimo(self) -> bool:
        return self._indice == self._total - 1


# ---------------------------------------------------------------------------
# Janela do tutorial (CTkToplevel modal leve) — NÃO instanciada nos testes
# ---------------------------------------------------------------------------
def abrir_tutorial(
    master,
    caminho_logomarca: os.PathLike | str = CAMINHO_LOGOMARCA,
    azul_principal: str = AZUL_PRINCIPAL,
    azul_destaque: str = AZUL_DESTAQUE,
    texto_secundario=TEXTO_SECUNDARIO,
):  # pragma: no cover - requer Tk, não exercitado nos testes
    """Abre a janela do tutorial como ``CTkToplevel`` modal leve.

    A modalidade (``transient`` + ``grab_set``) é aplicada apenas enquanto a
    janela está aberta, sem travar o worker. Ao marcar "Não mostrar novamente"
    e fechar, a preferência é persistida via ``salvar_preferencia_nao_mostrar``.
    """
    if ctk is None:
        raise RuntimeError("customtkinter não está disponível neste ambiente.")

    controlador = ControladorTutorial(len(ETAPAS))

    janela = ctk.CTkToplevel(master)
    janela.title("Ajuda / Tutorial")
    janela.geometry("620x460")
    janela.resizable(False, False)
    janela.transient(master)
    janela.grab_set()  # modal apenas enquanto aberta

    # Logomarca (identidade visual SESI/SENAI).
    try:
        from PIL import Image

        logo = Image.open(caminho_logomarca)
        imagem = ctk.CTkImage(light_image=logo, dark_image=logo, size=(276, 40))
        ctk.CTkLabel(janela, text="", image=imagem).pack(pady=(18, 8))
        janela._logo_ref = imagem  # evita coleta de lixo
    except Exception:
        pass

    label_titulo = ctk.CTkLabel(
        janela,
        text="",
        text_color=(azul_destaque, "#58BCEB"),
        font=ctk.CTkFont(size=18, weight="bold"),
    )
    label_titulo.pack(pady=(4, 8), padx=24)

    label_texto = ctk.CTkLabel(
        janela,
        text="",
        justify="left",
        wraplength=540,
        text_color=texto_secundario,
        font=ctk.CTkFont(size=14),
    )
    label_texto.pack(pady=(0, 12), padx=24, fill="both", expand=True)

    var_nao_mostrar = ctk.BooleanVar(value=False)
    check = ctk.CTkCheckBox(
        janela,
        text="Não mostrar novamente",
        variable=var_nao_mostrar,
        fg_color=azul_principal,
        hover_color=azul_destaque,
    )
    check.pack(pady=(0, 12))

    rodape = ctk.CTkFrame(janela, fg_color="transparent")
    rodape.pack(pady=(0, 18))

    def _fechar() -> None:
        salvar_preferencia_nao_mostrar(bool(var_nao_mostrar.get()))
        try:
            janela.grab_release()
        except Exception:
            pass
        janela.destroy()

    botao_anterior = ctk.CTkButton(
        rodape,
        text="Anterior",
        width=120,
        fg_color="transparent",
        border_width=1,
        border_color=azul_principal,
        text_color=(azul_destaque, "#54B7ED"),
    )
    botao_pular = ctk.CTkButton(
        rodape,
        text="Pular",
        width=120,
        fg_color="transparent",
        border_width=1,
        border_color=azul_principal,
        text_color=(azul_destaque, "#54B7ED"),
        command=_fechar,
    )
    botao_proximo = ctk.CTkButton(
        rodape,
        text="Próximo",
        width=120,
        fg_color=azul_principal,
        hover_color=azul_destaque,
    )
    botao_anterior.grid(row=0, column=0, padx=6)
    botao_pular.grid(row=0, column=1, padx=6)
    botao_proximo.grid(row=0, column=2, padx=6)

    def _atualizar() -> None:
        etapa = ETAPAS[controlador.indice]
        label_titulo.configure(text=etapa.titulo)
        label_texto.configure(text=etapa.texto)
        botao_anterior.configure(state="disabled" if controlador.primeiro else "normal")
        botao_proximo.configure(text="Concluir" if controlador.ultimo else "Próximo")

    def _ir_proximo() -> None:
        if controlador.ultimo:
            _fechar()
            return
        controlador.proximo()
        _atualizar()

    def _ir_anterior() -> None:
        controlador.anterior()
        _atualizar()

    botao_anterior.configure(command=_ir_anterior)
    botao_proximo.configure(command=_ir_proximo)
    janela.protocol("WM_DELETE_WINDOW", _fechar)

    _atualizar()
    return janela
