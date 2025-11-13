import re
from datetime import datetime, date, timedelta
from pathlib import Path
from bs4 import BeautifulSoup
from src.utils.time_utils import parse_time

def extrair_horas_por_dia(html_file: Path):
    with open(html_file, encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    regex_data = re.compile(r"\d{2}/\d{2}/\d{4}")
    spans = soup.find_all("span")

    horas_por_dia = {}

    for span in spans:
        texto = span.get_text(strip=True)
        achou = regex_data.search(texto)
        if not achou:
            continue

        data_str = achou.group()
        d = datetime.strptime(data_str, "%d/%m/%Y").date()
        horas_por_dia.setdefault(d, [])

        elem = span
        while True:
            elem = elem.find_next()
            if elem is None:
                break
            if elem.name == "span" and regex_data.search(elem.get_text(strip=True)):
                break
            if elem.name == "label" and elem.get_text(strip=True).upper() == "TRABALHANDO":
                lab = elem.find_next("label")
                if lab:
                    h = lab.get_text(strip=True)
                    if re.match(r"^\d{2}:\d{2}$", h):
                        horas_por_dia[d].append(parse_time(h))

    return horas_por_dia

def obter_mes_ano(html_file: Path, horas_por_dia: dict[date, list[timedelta]]):
    partes = html_file.stem.split("_")
    if len(partes) == 2:
        try:
            return int(partes[0]), int(partes[1])
        except Exception:
            pass

    if horas_por_dia:
        qualquer_dia = next(iter(horas_por_dia.keys()))
        return qualquer_dia.month, qualquer_dia.year

    hoje = datetime.now().date()
    return hoje.month, hoje.year
