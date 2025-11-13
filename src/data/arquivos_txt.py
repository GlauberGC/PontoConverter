from datetime import datetime, date
from pathlib import Path

def carregar_lista_txt(data_dir: Path, nome_arquivo: str) -> set[date]:
    caminho = data_dir / nome_arquivo
    if not caminho.exists():
        return set()

    dias = set()
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        try:
            dias.add(datetime.strptime(linha, "%Y-%m-%d").date())
        except Exception:
            print(f"Linha inválida em {nome_arquivo}: {linha}")
    return dias
