from datetime import timedelta
import pandas as pd
from src.utils.time_utils import format_timedelta

def montar_consolidado(resultados: list[dict]) -> pd.DataFrame:
    linhas = []
    total_trab_geral = timedelta()
    total_prev_geral = timedelta()
    total_diff_geral = timedelta()

    for r in resultados:
        nome_aba = r["sheet"]
        resumo = r["resumo"]

        total_trab_geral += resumo["Total Trabalhado"]
        total_prev_geral += resumo["Total Previsto"]
        total_diff_geral += resumo["Diferença"]

        linhas.append({
            "Mês": nome_aba,
            "Total Trabalhado": format_timedelta(resumo["Total Trabalhado"]),
            "Total Previsto": format_timedelta(resumo["Total Previsto"]),
            "Diferença": format_timedelta(resumo["Diferença"]),
        })

    linhas.append({
        "Mês": "TOTAL GERAL",
        "Total Trabalhado": format_timedelta(total_trab_geral),
        "Total Previsto": format_timedelta(total_prev_geral),
        "Diferença": format_timedelta(total_diff_geral),
    })

    return pd.DataFrame(linhas)
