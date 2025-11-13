import re
import calendar
from datetime import timedelta, datetime, date
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from xlsxwriter.utility import xl_col_to_name


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

CARGA_DIARIA_PADRAO = "08:00"
CARGA_ANIVERSARIO = timedelta(hours=4)  # folga de 4h
PASTA_HTML = "PONTO_HTML"
PASTA_EXCEL = "PONTO_EXCEL"
NOME_ARQUIVO_SAIDA = "PONTOS_CONSOLIDADOS.xlsx"
ARQ_FERIAS = "ferias.txt"
ARQ_ATESTADO = "atestado.txt"
ARQ_ANIVERSARIO = "aniversario.txt"


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================

def parse_time(horario: str) -> timedelta:
    h, m = horario.split(':')
    return timedelta(hours=int(h), minutes=int(m))


def format_timedelta(td: timedelta) -> str:
    total_seconds = int(td.total_seconds())
    sign = "-" if total_seconds < 0 else ""
    total_seconds = abs(total_seconds)
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    return f"{sign}{h:02d}:{m:02d}"


def str_to_timedelta(s: str) -> timedelta:
    if s is None:
        return timedelta(0)
    s = str(s).strip()
    if not s:
        return timedelta(0)
    sign = -1 if s.startswith("-") else 1
    if s[0] in "+-":
        s = s[1:]
    try:
        h, m = s.split(":")
        return sign * timedelta(hours=int(h), minutes=int(m))
    except Exception:
        return timedelta(0)


# ==========================================================
# FERIADOS (fixos + móveis)
# ==========================================================

def easter_date(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2*e + 2*i - h - k) % 7
    m = (a + 11*h + 22*l) // 451
    month = (h + l - 7*m + 114) // 31
    day = 1 + ((h + l - 7*m + 114) % 31)
    return date(year, month, day)


def feriados_ano(year: int) -> set[date]:
    feriados = set()

    # Fixos
    feriados.add(date(year, 1, 1))
    feriados.add(date(year, 4, 21))
    feriados.add(date(year, 5, 1))
    feriados.add(date(year, 9, 7))
    feriados.add(date(year, 10, 12))
    feriados.add(date(year, 11, 2))
    feriados.add(date(year, 11, 15))
    feriados.add(date(year, 11, 20))
    feriados.add(date(year, 11, 30))
    feriados.add(date(year, 12, 25))

    # Móveis
    pascoa = easter_date(year)
    feriados.add(pascoa - timedelta(days=48))  # Carnaval (seg)
    feriados.add(pascoa - timedelta(days=47))  # Carnaval (ter)
    feriados.add(pascoa - timedelta(days=2))   # Sexta Santa
    feriados.add(pascoa + timedelta(days=60))  # Corpus Christi

    return feriados


# ==========================================================
# NOME DA ABA
# ==========================================================

def sanitize_sheet_name(name: str) -> str:
    invalid = '[]:*?/\\'
    for ch in invalid:
        name = name.replace(ch, ' ')
    return name[:31]


def obter_nome_aba(stem: str) -> str:
    meses = {
        "1":"Janeiro","01":"Janeiro",
        "2":"Fevereiro","02":"Fevereiro",
        "3":"Março","03":"Março",
        "4":"Abril","04":"Abril",
        "5":"Maio","05":"Maio",
        "6":"Junho","06":"Junho",
        "7":"Julho","07":"Julho",
        "8":"Agosto","08":"Agosto",
        "9":"Setembro","09":"Setembro",
        "10":"Outubro",
        "11":"Novembro",
        "12":"Dezembro"
    }

    partes = stem.split("_")
    if len(partes) == 2 and partes[0] in meses:
        return sanitize_sheet_name(f"{meses[partes[0]]} {partes[1]}")
    return sanitize_sheet_name(stem)


def obter_mes_ano_do_arquivo(html_file: Path, horas_por_dia):
    partes = html_file.stem.split("_")
    if len(partes) == 2:
        try:
            return int(partes[0]), int(partes[1])
        except:
            pass
    if horas_por_dia:
        d = next(iter(horas_por_dia.keys()))
        return d.month, d.year
    return 1, datetime.now().year


# ==========================================================
# CARREGAMENTO DE ARQUIVOS (Férias, Atestado, Aniversário)
# ==========================================================

def carregar_lista(base_dir: Path, nome: str) -> set[date]:
    caminho = base_dir / nome
    if not caminho.exists():
        return set()
    dias = set()
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        try:
            dias.add(datetime.strptime(linha, "%Y-%m-%d").date())
        except:
            print(f"⚠ Linha inválida em {nome}: {linha}")
    return dias


# ==========================================================
# PROCESSAR HTML DO MÊS
# ==========================================================

def processar_folha(html_file: Path,
                    ferias: set[date],
                    atestados: set[date],
                    aniversarios: set[date]) -> pd.DataFrame:

    with open(html_file, encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    regex_data = re.compile(r"\d{2}/\d{2}/\d{4}")
    spans = soup.find_all("span")
    horas_por_dia: dict[date, list[timedelta]] = {}

    # Captura datas + horários TRABALHANDO
    for span in spans:
        texto = span.get_text(strip=True)
        achou = regex_data.search(texto)
        if not achou:
            continue
        d = datetime.strptime(achou.group(), "%d/%m/%Y").date()
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

    # Identificação do mês/ano
    mes, ano = obter_mes_ano_do_arquivo(html_file, horas_por_dia)
    feriados_ano_todo = feriados_ano(ano)
    feriados_mes = {d for d in feriados_ano_todo if d.month == mes}

    _, ultimo_dia = calendar.monthrange(ano, mes)
    dias_processar = []

    for dia in range(1, ultimo_dia + 1):
        d = date(ano, mes, dia)

        if d.weekday() >= 5:  # final de semana
            continue

        if (
            d in horas_por_dia
            or d in feriados_mes
            or d in atestados
            or d in ferias
            or d in aniversarios
        ):
            dias_processar.append(d)

    dias_processar.sort()

    linhas = []
    total_trab = timedelta()
    total_previsto = timedelta()
    total_diff = timedelta()

    carga_diaria = parse_time(CARGA_DIARIA_PADRAO)

    for d in dias_processar:
        lista = horas_por_dia.get(d, [])
        hh = sum(lista, timedelta())

        eh_feriado = d in feriados_mes
        eh_atestado = d in atestados
        eh_ferias = d in ferias
        eh_aniversario = d in aniversarios

        if eh_ferias:
            tipo = "Férias"
            carga_prev = timedelta(0)
            hh = timedelta(0)

        elif eh_feriado:
            tipo = "Feriado"
            carga_prev = timedelta(0)

        elif eh_atestado:
            tipo = "Atestado"
            carga_prev = carga_diaria
            hh = carga_diaria

        elif eh_aniversario:
            tipo = "Aniversário"
            carga_prev = CARGA_ANIVERSARIO

        else:
            tipo = "Normal"
            carga_prev = carga_diaria

        diff = hh - carga_prev

        # Férias não entram nos totais
        if tipo != "Férias":
            total_trab += hh
            total_previsto += carga_prev
            total_diff += diff

        linhas.append({
            "Data": d.strftime("%d/%m/%Y"),
            "Tipo do Dia": tipo,
            "Total Trabalhado": format_timedelta(hh),
            "Carga Prevista": format_timedelta(carga_prev),
            "Diferença": format_timedelta(diff),
        })

    # Linha TOTAL MÊS
    linhas.append({
        "Data": "TOTAL MÊS",
        "Tipo do Dia": "",
        "Total Trabalhado": format_timedelta(total_trab),
        "Carga Prevista": format_timedelta(total_previsto),
        "Diferença": format_timedelta(total_diff),
    })

    return pd.DataFrame(linhas)


# ==========================================================
# GERAR EXCEL CONSOLIDADO
# ==========================================================

def gerar_excel_consolidado(base_dir: Path):
    pasta_html = base_dir / PASTA_HTML
    pasta_excel = base_dir / PASTA_EXCEL
    pasta_excel.mkdir(exist_ok=True)

    ferias = carregar_lista(base_dir, ARQ_FERIAS)
    atestados = carregar_lista(base_dir, ARQ_ATESTADO)
    aniversarios = carregar_lista(base_dir, ARQ_ANIVERSARIO)

    arquivos = sorted(pasta_html.glob("*.html"))
    if not arquivos:
        print("⚠ Nenhum arquivo HTML encontrado.")
        return

    resultados = []

    for arq in arquivos:
        df = processar_folha(arq, ferias, atestados, aniversarios)
        if df.empty:
            continue

        aba = obter_nome_aba(arq.stem)
        df_mes = df[df["Data"] != "TOTAL MÊS"].copy()
        linha_total = df[df["Data"] == "TOTAL MÊS"].iloc[0]

        trab = str_to_timedelta(linha_total["Total Trabalhado"])
        prev = str_to_timedelta(linha_total["Carga Prevista"])
        diff = str_to_timedelta(linha_total["Diferença"])

        # resumo consolidado
        resultados.append({
            "sheet": aba,
            "df": df,
            "resumo": {
                "Total Trabalhado": trab,
                "Total Previsto": prev,
                "Diferença": diff,
            }
        })

    # Monta consolidado
    cons_linhas = []
    total_trab_geral = timedelta()
    total_prev_geral = timedelta()
    total_diff_geral = timedelta()

    for r in resultados:
        res = r["resumo"]
        total_trab_geral += res["Total Trabalhado"]
        total_prev_geral += res["Total Previsto"]
        total_diff_geral += res["Diferença"]

        cons_linhas.append({
            "Mês": r["sheet"],
            "Total Trabalhado": format_timedelta(res["Total Trabalhado"]),
            "Total Previsto": format_timedelta(res["Total Previsto"]),
            "Diferença": format_timedelta(res["Diferença"]),
        })

    cons_linhas.append({
        "Mês": "TOTAL GERAL",
        "Total Trabalhado": format_timedelta(total_trab_geral),
        "Total Previsto": format_timedelta(total_prev_geral),
        "Diferença": format_timedelta(total_diff_geral),
    })

    df_cons = pd.DataFrame(cons_linhas)

    caminho_saida = base_dir / PASTA_EXCEL / NOME_ARQUIVO_SAIDA
    print(f"📊 Gerando Excel em: {caminho_saida}")

    # ---------------------------------------------------------------------
    # Escrevendo no Excel (CONSOLIDADO primeiro)
    # ---------------------------------------------------------------------
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as wr:
        wb = wr.book

        # =============================== FORMATOS ===============================
        fmt_total = wb.add_format({
            "bold": True,
            "bg_color": "#D9D9D9",
            "top": 2,
            "bottom": 2,
        })

        fmt_ferias_row    = wb.add_format({"bg_color": "#DDEBF7"})
        fmt_atestado_row  = wb.add_format({"bg_color": "#FFF2CC"})
        fmt_aniv_row      = wb.add_format({"bg_color": "#E2EFDA"})
        fmt_feriado_row   = wb.add_format({"bg_color": "#816E85"})

        fmt_verde = wb.add_format({"font_color": "green"})
        fmt_vermelho = wb.add_format({"font_color": "red"})

        # ========================== ABA CONSOLIDADO ==========================
        df_cons.to_excel(wr, sheet_name="CONSOLIDADO", index=False)
        ws_c = wr.sheets["CONSOLIDADO"]

        for col in range(len(df_cons.columns)):
            ws_c.set_column(col, col, 18)

        # Condicional
        try:
            col_idx = df_cons.columns.get_loc("Diferença")
            letra = xl_col_to_name(col_idx)
            rng = f"{letra}2:{letra}{len(df_cons)+1}"
            ws_c.conditional_format(rng, {
                "type": "text", "criteria": "not containing", "value": "-", "format": fmt_verde
            })
            ws_c.conditional_format(rng, {
                "type": "text", "criteria": "containing", "value": "-", "format": fmt_vermelho
            })
        except:
            pass

        # Destacar TOTAL GERAL
        for idx, row in df_cons.iterrows():
            if row["Mês"] == "TOTAL GERAL":
                for col in range(len(df_cons.columns)):  # A:H
                    ws_c.write(idx+1, col, df_cons.iloc[idx, col], fmt_total)

        # ======================= ABAS MENSAIS ============================
        for r in resultados:
            aba = r["sheet"]
            df = r["df"]

            df.to_excel(wr, sheet_name=aba, index=False)
            ws = wr.sheets[aba]

            for col in range(len(df.columns)):
                ws.set_column(col, col, 18)

            # Condicional Diferença
            try:
                col_idx = df.columns.get_loc("Diferença")
                letra = xl_col_to_name(col_idx)
                rng = f"{letra}2:{letra}{len(df)+1}"
                ws.conditional_format(rng, {
                    "type": "text", "criteria": "not containing", "value": "-", "format": fmt_verde
                })
                ws.conditional_format(rng, {
                    "type": "text", "criteria": "containing", "value": "-", "format": fmt_vermelho
                })
            except:
                pass

            # Colunas A:E → range de 5
            for idx, row in df.iterrows():
                excel_row = idx + 1
                tipo = row["Tipo do Dia"]

                if row["Data"] == "TOTAL MÊS":
                    for col in range(5):
                        ws.write(excel_row, col, df.iloc[idx, col], fmt_total)

                elif tipo == "Férias":
                    for col in range(5):
                        ws.write(excel_row, col, df.iloc[idx, col], fmt_ferias_row)

                elif tipo == "Atestado":
                    for col in range(5):
                        ws.write(excel_row, col, df.iloc[idx, col], fmt_atestado_row)

                elif tipo == "Aniversário":
                    for col in range(5):
                        ws.write(excel_row, col, df.iloc[idx, col], fmt_aniv_row)

                elif tipo == "Feriado":
                    for col in range(5):
                        ws.write(excel_row, col, df.iloc[idx, col], fmt_feriado_row)

    print("✅ Excel gerado com sucesso!")


# ==========================================================
# MAIN
# ==========================================================

if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    gerar_excel_consolidado(base)
