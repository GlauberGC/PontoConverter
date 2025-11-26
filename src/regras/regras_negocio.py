import calendar
from datetime import date, timedelta
from pathlib import Path
import re

import pandas as pd

from src.parser.html_parser import extrair_horas_por_dia
from src.regras.calendario import feriados_ano
from src.utils.time_utils import parse_time, format_timedelta

# Carga padrão
CARGA_DIARIA_PADRAO = timedelta(hours=8)
CARGA_ANIVERSARIO = timedelta(hours=4)

dias_semana = {
    0: "Segunda",
    1: "Terça",
    2: "Quarta",
    3: "Quinta",
    4: "Sexta",
    5: "Sábado",
    6: "Domingo",
}


def gerar_nome_aba(mes: int, ano: int) -> str:
    """Gera nome amigável para a aba do Excel."""
    mapa = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março",
        4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro",
        10: "Outubro", 11: "Novembro", 12: "Dezembro",
    }
    nome = f"{mapa.get(mes, mes)} {ano}"
    invalid = '[]:*?/\\'
    for ch in invalid:
        nome = nome.replace(ch, ' ')
    return nome[:31]


def obter_mes_ano(html_file: Path, horas_por_dia: dict) -> tuple[int, int]:
    """
    Descobre (mês, ano) a partir:
    - nome do arquivo: 07_2025.html
    - ou da primeira data do ponto
    """
    nome = html_file.name
    m = re.search(r"(\d{2})_(\d{4})", nome)
    if m:
        return int(m.group(1)), int(m.group(2))

    if horas_por_dia:
        primeira_data = next(iter(horas_por_dia.keys()))
        return primeira_data.month, primeira_data.year

    raise ValueError(f"Não foi possível determinar mês/ano a partir de {nome}")


def _entradas_saidas(info):
    """Garante retorno seguro de entradas e saídas."""
    if info is None:
        return [], []

    entradas = info.get("entradas", []) or []
    saidas = info.get("saidas", []) or []
    return entradas, saidas


def processar_mes(html_file: Path,
                  ferias: set[date],
                  atestados: set[date],
                  aniversarios: set[date],
                  abonos: set[date]):

    horas_por_dia = extrair_horas_por_dia(html_file)

    mes, ano = obter_mes_ano(html_file, horas_por_dia)

    feriados_ano_todo = feriados_ano(ano)
    feriados_mes = {d for d in feriados_ano_todo if d.month == mes}

    _, ultimo_dia = calendar.monthrange(ano, mes)

    dias_processar = []
    for dia in range(1, ultimo_dia + 1):
        d = date(ano, mes, dia)
        if (
            d in horas_por_dia
            or d in ferias
            or d in atestados
            or d in abonos
            or d in aniversarios
            or d in feriados_mes
            or d.weekday() >= 5
        ):
            dias_processar.append(d)

    dias_processar.sort()

    linhas = []

    total_trab = timedelta()
    total_prev = timedelta()
    total_saldo = timedelta()

    for d in dias_processar:

        info = horas_por_dia.get(d)

        # -----------------------------
        # TRATAMENTO SEGURO DO DIA
        # -----------------------------
        if info is None:
            entradas = []
            saidas = []
            total_label = "00:00"
        else:
            entradas, saidas = _entradas_saidas(info)
            total_label = info.get("total_label", "00:00")

        # total trabalhado REAL (do HTML)
        try:
            total_horas = parse_time(total_label)
        except:
            total_horas = timedelta()

        eh_ferias = d in ferias
        eh_fds = d.weekday() >= 5
        eh_feriado = d in feriados_mes
        eh_atestado = d in atestados
        eh_abono = d in abonos
        eh_aniversario = d in aniversarios

        # ================================================================
        # 1) FÉRIAS
        # ================================================================
        if eh_ferias:
            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Férias",
                "Total Trabalhado": "00:00",
                "Carga Prevista": "00:00",
                "Saldo do dia": "00:00",
            }
            linhas.append(linha)
            continue

        # ================================================================
        # 2) FINAL DE SEMANA
        # ================================================================
        if eh_fds:
            carga_prevista = timedelta(0)
            saldo = total_horas

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Final de Semana",
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": "00:00",
                "Saldo do dia": format_timedelta(saldo),
            }
            # anexar entradas/saídas
            for i in range(max(len(entradas), len(saidas))):
                linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
                linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

            linhas.append(linha)
            continue

        # ================================================================
        # 3) FERIADO
        # ================================================================
        if eh_feriado:
            carga_prevista = timedelta(0)
            saldo = total_horas

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Feriado",
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": "00:00",
                "Saldo do dia": format_timedelta(saldo),
            }

            for i in range(max(len(entradas), len(saidas))):
                linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
                linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

            linhas.append(linha)
            continue

        # ================================================================
        # 4) ATESTADO — SEMPRE saldo = 0
        # ================================================================
        if eh_atestado:
            carga_prevista = CARGA_DIARIA_PADRAO
            saldo = timedelta(0)

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Atestado",
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": format_timedelta(carga_prevista),
                "Saldo do dia": "00:00",
            }

            for i in range(max(len(entradas), len(saidas))):
                linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
                linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

            linhas.append(linha)
            continue

        # ================================================================
        # 5) ABONO — SEMPRE saldo = 0
        # ================================================================
        if eh_abono:
            carga_prevista = CARGA_DIARIA_PADRAO
            saldo = timedelta(0)

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Abono",
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": format_timedelta(carga_prevista),
                "Saldo do dia": "00:00",
            }

            for i in range(max(len(entradas), len(saidas))):
                linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
                linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

            linhas.append(linha)
            continue

        # ================================================================
        # 6) ANIVERSÁRIO
        # ================================================================
        if eh_aniversario:
            carga_prevista = CARGA_ANIVERSARIO
            saldo = total_horas - carga_prevista

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linha = {
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": "Aniversário",
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": format_timedelta(carga_prevista),
                "Saldo do dia": format_timedelta(saldo),
            }

            for i in range(max(len(entradas), len(saidas))):
                linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
                linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

            linhas.append(linha)
            continue

        # ================================================================
        # 7) DIA NORMAL
        # ================================================================
        carga_prevista = CARGA_DIARIA_PADRAO
        saldo = total_horas - carga_prevista

        total_trab += total_horas
        total_prev += carga_prevista
        total_saldo += saldo

        linha = {
            "Data": d.strftime("%d/%m/%Y"),
            "Dia da Semana": dias_semana[d.weekday()],
            "Tipo do Dia": "Normal",
            "Total Trabalhado": format_timedelta(total_horas),
            "Carga Prevista": format_timedelta(carga_prevista),
            "Saldo do dia": format_timedelta(saldo),
        }

        for i in range(max(len(entradas), len(saidas))):
            linha[f"Entrada {i+1}"] = entradas[i] if i < len(entradas) else ""
            linha[f"Saída {i+1}"] = saidas[i] if i < len(saidas) else ""

        linhas.append(linha)

    # ================================================================
    # TOTAL DO MÊS
    # ================================================================
    linhas.append({
        "Data": "TOTAL MÊS",
        "Dia da Semana": "",
        "Tipo do Dia": "",
        "Total Trabalhado": format_timedelta(total_trab),
        "Carga Prevista": format_timedelta(total_prev),
        "Saldo do dia": format_timedelta(total_saldo),
    })

    df_mes = pd.DataFrame(linhas)

    resumo = {
        "nome_aba": gerar_nome_aba(mes, ano),
        "total_trabalhado": total_trab,
        "total_previsto": total_prev,
        "saldo_dia": total_saldo,
    }

    return df_mes, resumo
