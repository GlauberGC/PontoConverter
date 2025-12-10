from datetime import datetime, date, timedelta
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

        # Suporta dois formatos:
        # 1) YYYY-MM-DD
        # 2) YYYY-MM-DD;YYYY-MM-DD  (intervalo inclusivo)
        if ";" in linha:
            parts = [p.strip() for p in linha.split(";")]
            if len(parts) >= 2:
                try:
                    d1 = datetime.strptime(parts[0], "%Y-%m-%d").date()
                    d2 = datetime.strptime(parts[1], "%Y-%m-%d").date()
                    if d2 < d1:
                        d1, d2 = d2, d1
                    cur = d1
                    while cur <= d2:
                        dias.add(cur)
                        cur = cur + timedelta(days=1)
                except Exception:
                    print(f"Linha inválida (intervalo) em {nome_arquivo}: {linha}")
                    continue
            else:
                print(f"Linha inválida em {nome_arquivo}: {linha}")
            continue

        try:
            dias.add(datetime.strptime(linha, "%Y-%m-%d").date())
        except Exception:
            print(f"Linha inválida em {nome_arquivo}: {linha}")
    return dias


def carregar_pontos_manuais(data_dir: Path, nome_arquivo: str = "pontos_manuais.txt") -> dict[date, dict]:
    """
    Lê `pontos_manuais.txt` com linhas no formato:
        YYYY-MM-DD;HH:MM;HH:MM;...

    Retorna um dicionário no mesmo formato usado por `extrair_horas_por_dia`:
        { date: {"entradas": [...], "saidas": [...], "total_label": "HH:MM"}, ... }
    """
    caminho = data_dir / nome_arquivo
    if not caminho.exists():
        return {}

    resultado: dict[date, dict] = {}

    def soma_periodos(periodos: list[tuple[str, str]]) -> str:
        total_min = 0
        for ent, sai in periodos:
            if not ent or not sai:
                continue
            try:
                hh_e, mm_e = map(int, ent.split(":"))
                hh_s, mm_s = map(int, sai.split(":"))
                start = hh_e * 60 + mm_e
                end = hh_s * 60 + mm_s
                if end >= start:
                    total_min += (end - start)
            except Exception:
                continue
        hh = total_min // 60
        mm = total_min % 60
        return f"{hh:02d}:{mm:02d}"

    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        parts = [p.strip() for p in linha.split(";") if p.strip()]
        if not parts:
            continue
        try:
            d = datetime.strptime(parts[0], "%Y-%m-%d").date()
        except Exception:
            print(f"Linha inválida em {nome_arquivo}: {linha}")
            continue

        times = parts[1:]
        entradas: list[str] = []
        saidas: list[str] = []

        # alterna entre entrada/saída
        for i, t in enumerate(times):
            if not t:
                continue
            if i % 2 == 0:
                entradas.append(t)
            else:
                saidas.append(t)

        # construir pares para soma (usa o mínimo entre entradas e saidas)
        pares: list[tuple[str, str]] = []
        for i in range(min(len(entradas), len(saidas))):
            pares.append((entradas[i], saidas[i]))

        total_label = soma_periodos(pares)

        resultado[d] = {
            "entradas": entradas,
            "saidas": saidas,
            "total_label": total_label,
        }

    return resultado
