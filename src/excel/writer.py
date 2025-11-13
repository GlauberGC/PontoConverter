from pathlib import Path
import pandas as pd
from xlsxwriter.utility import xl_col_to_name
from src.excel.estilos import criar_estilos

def _aplicar_condicional_diferenca(ws, df, estilos, coluna_nome="Diferença"):
    try:
        col_idx = df.columns.get_loc(coluna_nome)
        letra = xl_col_to_name(col_idx)
        faixa = f"{letra}2:{letra}{len(df) + 1}"
        ws.conditional_format(faixa, {
            "type": "text",
            "criteria": "not containing",
            "value": "-",
            "format": estilos["verde"],
        })
        ws.conditional_format(faixa, {
            "type": "text",
            "criteria": "containing",
            "value": "-",
            "format": estilos["vermelho"],
        })
    except Exception as e:
        print(f"Erro ao aplicar formatação condicional: {e}")

def gerar_arquivo_excel(caminho_saida: Path,
                        resultados: list[dict],
                        df_consolidado: pd.DataFrame):

    print(f"Gerando Excel: {caminho_saida}")

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        workbook = writer.book
        estilos = criar_estilos(workbook)

        # ================== ABA CONSOLIDADO (primeira) ==================
        df_consolidado.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
        ws_cons = writer.sheets["CONSOLIDADO"]

        for col in range(len(df_consolidado.columns)):
            ws_cons.set_column(col, col, 18)

        _aplicar_condicional_diferenca(ws_cons, df_consolidado, estilos)

        # Destaca TOTAL GERAL (linha inteira)
        for idx, row in df_consolidado.iterrows():
            if row["Mês"] == "TOTAL GERAL":
                for col in range(len(df_consolidado.columns)):  # A:H
                    ws_cons.write(idx + 1, col, df_consolidado.iloc[idx, col], estilos["total"])

        # ================== ABAS MENSAIS ==================
        for r in resultados:
            sheet_name = r["sheet"]
            df_mes = r["df"]

            df_mes.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            for col in range(len(df_mes.columns)):
                ws.set_column(col, col, 18)

            _aplicar_condicional_diferenca(ws, df_mes, estilos)

            # Pintar linhas até a última coluna de dados (A:E)
            for idx, row in df_mes.iterrows():
                excel_row = idx + 1  # +1 por causa do cabeçalho
                tipo = row["Tipo do Dia"]
                data_str = row["Data"]

                if data_str == "TOTAL MÊS":
                    for col in range(5):  # A:E
                        ws.write(excel_row, col, df_mes.iloc[idx, col], estilos["total"])
                    continue

                if tipo == "Férias":
                    fmt = estilos["ferias"]
                elif tipo == "Atestado":
                    fmt = estilos["atestado"]
                elif tipo == "Aniversário":
                    fmt = estilos["aniversario"]
                elif tipo == "Feriado":
                    fmt = estilos["feriado"]
                else:
                    continue  # dia normal, não colore

                for col in range(5):  # A:E
                    ws.write(excel_row, col, df_mes.iloc[idx, col], fmt)

    print("Excel gerado com sucesso!")
