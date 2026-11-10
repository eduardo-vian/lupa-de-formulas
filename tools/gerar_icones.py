"""Gera os ícones PNG da Lupa (lente + cabo) sem bibliotecas externas.

Uso: python tools/gerar_icones.py
"""
import math
import struct
import zlib
from pathlib import Path

PASTA = Path(__file__).resolve().parent.parent / "docs" / "assets"
TAMANHOS = (16, 32, 64, 80, 128)
AMOSTRAS = 4  # supersampling por eixo, para suavizar as bordas

AZUL = (31, 92, 168)
VIDRO = (222, 236, 251)
CENTRO, RAIO_EXT, RAIO_INT = (0.42, 0.42), 0.32, 0.22
CABO_INI, CABO_FIM, CABO_ESP = (0.62, 0.62), (0.88, 0.88), 0.075


def distancia_segmento(p, a, b):
    (px, py), (ax, ay), (bx, by) = p, a, b
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def cor_amostra(u, v):
    d = math.hypot(u - CENTRO[0], v - CENTRO[1])
    if RAIO_INT <= d <= RAIO_EXT or distancia_segmento((u, v), CABO_INI, CABO_FIM) <= CABO_ESP:
        return AZUL
    if d < RAIO_INT:
        return VIDRO
    return None


def pixel(x, y, n):
    r = g = b = cobertas = 0
    for i in range(AMOSTRAS):
        for j in range(AMOSTRAS):
            c = cor_amostra((x + (i + 0.5) / AMOSTRAS) / n, (y + (j + 0.5) / AMOSTRAS) / n)
            if c:
                r, g, b, cobertas = r + c[0], g + c[1], b + c[2], cobertas + 1
    if not cobertas:
        return (0, 0, 0, 0)
    return (r // cobertas, g // cobertas, b // cobertas, round(255 * cobertas / AMOSTRAS ** 2))


def salvar_png(caminho, n):
    bruto = bytearray()
    for y in range(n):
        bruto.append(0)  # filtro "nenhum" no início de cada linha
        for x in range(n):
            bruto.extend(pixel(x, y, n))

    def bloco(tipo, dados):
        return struct.pack(">I", len(dados)) + tipo + dados + struct.pack(">I", zlib.crc32(tipo + dados))

    cabecalho = struct.pack(">IIBBBBB", n, n, 8, 6, 0, 0, 0)  # 8 bits por canal, RGBA
    caminho.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + bloco(b"IHDR", cabecalho)
        + bloco(b"IDAT", zlib.compress(bytes(bruto), 9))
        + bloco(b"IEND", b"")
    )


def main():
    PASTA.mkdir(parents=True, exist_ok=True)
    for n in TAMANHOS:
        salvar_png(PASTA / f"icon-{n}.png", n)
    print(f"{len(TAMANHOS)} icones gerados em {PASTA}")


if __name__ == "__main__":
    main()
