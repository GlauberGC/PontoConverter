import calendar
from datetime import date, timedelta, datetime
from pathlib import Path

import pandas as pd

from src.parser.html_parser import extrair_horas_por_dia, obter_mes_ano
from src.regras.calendario import feriados_ano
from src.utils.time_utils import parse_time, format_timedelta

# Configurações gerais
CARGA_DIARIA_PADRAO = "08:00"
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
    mapa = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
    }
    nome = f"{mapa.get(mes, mes)} {ano}"
    invalid = '[]:*?/\\'
    for ch in invalid:
        nome = nome.replace(ch, ' ')
    return nome[:31]


def processar_mes(html_file: Path,
                  ferias: set[date],
                  atestados: set[date],
                  aniversarios: set[date]):

    horas_por_dia = extrair_horas_por_dia(html_file)
    if not horas_por_dia:
        print(f"⚠ Sem registros de horas em {html_file.name}")
        return None, None

    mes, ano = obter_mes_ano(html_file, horas_por_dia)
    feriados_ano_todo = feriados_ano(ano)
    feriados_mes = {d for d in feriados_ano_todo if d.month == mes}

    _, ultimo_dia = calendar.monthrange(ano, mes)
    dias_processar = []

    # Agora inclui TODOS os dias úteis e finais de semana
    for dia in range(1, ultimo_dia + 1):
        d = date(ano, mes, dia)

        # Se tiver registro, ou feriado, ou ferias, ou atestado, ou aniversário, ou for final de semana
        if (
            d in horas_por_dia
            or d in feriados_mes
            or d in atestados
            or d in ferias
            or d in aniversarios
            or d.weekday() >= 5  # inclui sábado e domingo
        ):
            dias_processar.append(d)

    dias_processar.sort()

    linhas = []
    carga_diaria = parse_time(CARGA_DIARIA_PADRAO)

    total_trab = timedelta()
    total_prev = timedelta()
    total_saldo = timedelta()

    for d in dias_processar:
        registros = horas_por_dia.get(d, [])
        total_horas = sum(registros, timedelta())

        eh_feriado = d in feriados_mes
        eh_atestado = d in atestados
        eh_ferias = d in ferias
        eh_aniversario = d in aniversarios
        eh_fds = d.weekday() >= 5  # sábado = 5, domingo = 6

        # ============================
        # 1) FÉRIAS
        # ============================
        if eh_ferias:
            tipo = "Férias"
            carga_prevista = timedelta(0)
            total_horas = timedelta(0)
            saldo = timedelta(0)

            linhas.append({
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": tipo,
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": "00:00",
                "Saldo do dia": "00:00",
            })
            continue

        # ============================
        # 2) FERIADO
        # ============================
        if eh_feriado:
            tipo = "Feriado"
            carga_prevista = timedelta(0)
            saldo = total_horas  # se trabalhou → saldo positivo

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linhas.append({
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": tipo,
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": "00:00",
                "Saldo do dia": format_timedelta(saldo),
            })
            continue

        # ============================
        # 3) ATESTADO (saldo sempre zero)
        # ============================
        if eh_atestado:
            tipo = "Atestado"
            carga_prevista = carga_diaria
            horas_exibidas = total_horas  # mostra horas reais
            saldo = timedelta(0)

            total_trab += horas_exibidas
            total_prev += carga_prevista
            total_saldo += saldo

            linhas.append({
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": tipo,
                "Total Trabalhado": format_timedelta(horas_exibidas),
                "Carga Prevista": format_timedelta(carga_prevista),
                "Saldo do dia": "00:00",
            })
            continue

        # ============================
        # 4) ANIVERSÁRIO
        # ============================
        if eh_aniversario:
            tipo = "Aniversário"
            carga_prevista = CARGA_ANIVERSARIO
            saldo = total_horas - carga_prevista

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linhas.append({
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": tipo,
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": format_timedelta(carga_prevista),
                "Saldo do dia": format_timedelta(saldo),
            })
            continue

        # ============================
        # 5) FINAL DE SEMANA
        # ============================
        if eh_fds:
            tipo = "Final de Semana"
            carga_prevista = timedelta(0)
            saldo = total_horas  # se trabalhou, vira positivo

            total_trab += total_horas
            total_prev += carga_prevista
            total_saldo += saldo

            linhas.append({
                "Data": d.strftime("%d/%m/%Y"),
                "Dia da Semana": dias_semana[d.weekday()],
                "Tipo do Dia": tipo,
                "Total Trabalhado": format_timedelta(total_horas),
                "Carga Prevista": "00:00",
                "Saldo do dia": format_timedelta(saldo),
            })
            continue

        # ============================
        # 6) DIA NORMAL (útil)
        # ============================
        tipo = "Normal"
        carga_prevista = carga_diaria
        saldo = total_horas - carga_prevista

        total_trab += total_horas
        total_prev += carga_prevista
        total_saldo += saldo

        linhas.append({
            "Data": d.strftime("%d/%m/%Y"),
            "Dia da Semana": dias_semana[d.weekday()],
            "Tipo do Dia": tipo,
            "Total Trabalhado": format_timedelta(total_horas),
            "Carga Prevista": format_timedelta(carga_prevista),
            "Saldo do dia": format_timedelta(saldo),
        })

    # ============================
    # LINHA TOTAL DO MÊS
    # ============================
    linhas.append({
        "Data": "TOTAL MÊS",
        "Dia da Semana": "",
        "Tipo do Dia": "",
        "Total Trabalhado": format_timedelta(total_trab),
        "Carga Prevista": format_timedelta(total_prev),
        "Saldo do dia": format_timedelta(total_saldo),
    })

    df_mes = pd.DataFrame(linhas)
    nome_aba = gerar_nome_aba(mes, ano)

    resumo = {
        "nome_aba": nome_aba,
        "total_trabalhado": total_trab,
        "total_previsto": total_prev,
        "saldo_dia": total_saldo,
    }

    return df_mes, resumo
