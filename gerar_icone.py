"""Gera o ícone do aplicativo (.ico) a partir da logomarca SESI/SENAI.

A logomarca é retangular; para o ícone quadrado do Windows a imagem é
redimensionada mantendo a proporção e centralizada em um canvas quadrado
transparente. Gera um .ico multi-resolução.

Uso:
    python gerar_icone.py
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image


PASTA_ASSETS = Path(__file__).resolve().parent / "assets"
ORIGEM = PASTA_ASSETS / "marca-SESI-SENAI_reduzida_cor.png"
DESTINO = PASTA_ASSETS / "conversor_crm.ico"
TAMANHOS = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def gerar_icone(origem: Path = ORIGEM, destino: Path = DESTINO) -> Path:
    """Cria um .ico quadrado multi-resolução a partir da logomarca."""
    logo = Image.open(origem).convert("RGBA")

    lado = max(logo.size)
    canvas = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    posicao = ((lado - logo.width) // 2, (lado - logo.height) // 2)
    canvas.paste(logo, posicao, logo)

    canvas.save(destino, format="ICO", sizes=TAMANHOS)
    return destino


if __name__ == "__main__":
    caminho = gerar_icone()
    print(f"Ícone gerado em: {caminho}")
