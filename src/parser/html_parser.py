from bs4 import BeautifulSoup
from datetime import datetime
from collections import defaultdict
import re


def _somar_horas_str(h1: str, h2: str) -> str:
    """
    Soma duas horas no formato 'HH:MM' e devolve outra string 'HH:MM'.
    Trata valores vazios ou inválidos como 00:00.
    """
    def to_min(h: str) -> int:
        if not h:
            return 0
        h = h.strip()
        if not re.match(r"^\d{1,2}:\d{2}$", h):
            return 0
        try:
            hh, mm = map(int, h.split(":"))
            return hh * 60 + mm
        except Exception:
            return 0

    total_min = to_min(h1) + to_min(h2)
    hh = total_min // 60
    mm = total_min % 60
    return f"{hh:02d}:{mm:02d}"


def extrair_horas_por_dia(html_file):
    """
    Lê o HTML da folha de ponto e devolve um dicionário:

        {
          date(...): {
              "entradas": [ "08:02", "13:15", ... ],
              "saidas":   [ "12:14", "18:05", ... ],
              "total_label": "HH:MM"   # SOMA de todos os labels TRABALHANDO do dia
          },
          ...
        }

    A data é obtida do atributo title do input de SAÍDA:
        title="Ponto fechado em 13/01/2025"
    O total do dia é a SOMA de todos os <label>HH:MM</label> encontrados
    nas linhas daquele dia.
    """

    with open(html_file, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    dias = defaultdict(lambda: {
        "entradas": [],
        "saidas": [],
        "total_label": "00:00",
    })

    regex_data = re.compile(r"(\d{2}/\d{2}/\d{4})")
    regex_hora = re.compile(r"^\d{1,2}:\d{2}$")

    # Cada registro de ponto está em uma <tr>
    for tr in soup.find_all("tr"):

        # 1) SAÍDA -> de onde tiramos a DATA
        saida_input = tr.find("input", class_="saida")
        if not saida_input:
            continue

        title = saida_input.get("title") or ""
        m = regex_data.search(title)
        if not m:
            continue

        data_str = m.group(1)
        try:
            data_dia = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            continue

        registro_dia = dias[data_dia]

        # 2) ENTRADA / SAÍDA (para exibição detalhada)
        entrada_input = tr.find("input", class_="entrada")
        entrada_val = (entrada_input.get("value") or "").strip() if entrada_input else ""
        saida_val = (saida_input.get("value") or "").strip()

        if entrada_val and regex_hora.match(entrada_val):
            registro_dia["entradas"].append(entrada_val)

        if saida_val and regex_hora.match(saida_val):
            registro_dia["saidas"].append(saida_val)

        # 3) LABEL TRABALHANDO (Total do período) -> somar no total do dia
        labels = tr.find_all("label")
        for lb in labels:
            txt = lb.get_text(strip=True)
            if regex_hora.match(txt):
                # soma esse período ao total do dia
                registro_dia["total_label"] = _somar_horas_str(
                    registro_dia["total_label"], txt
                )
                # cada linha TR costuma ter só um label HH:MM relevante
                break

    return dias
