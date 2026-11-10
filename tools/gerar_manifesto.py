"""Preenche o endereço de hospedagem no manifesto.

Uso: python tools/gerar_manifesto.py https://endereco/da/lupa [saida.xml]
Sem saída indicada, grava docs/manifest.xml (o arquivo que o Excel instala).
"""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit(__doc__)
    base = sys.argv[1].rstrip("/")
    if not base.startswith("https://"):
        sys.exit("O endereço precisa começar com https://")
    saida = Path(sys.argv[2]) if len(sys.argv) == 3 else RAIZ / "docs" / "manifest.xml"
    modelo = (RAIZ / "manifest.template.xml").read_text(encoding="utf-8")
    saida.write_text(modelo.replace("{{BASE_URL}}", base), encoding="utf-8")
    print(f"{saida} gerado para {base}")


if __name__ == "__main__":
    main()
