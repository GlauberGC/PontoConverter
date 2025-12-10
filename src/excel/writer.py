from pathlib import Path
import pandas as pd
from xlsxwriter.utility import xl_col_to_name
from src.excel.estilos import criar_estilos

def _aplicar_condicional_diferenca(ws, df, estilos, coluna_nome="Saldo do dia"):
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
                        df_consolidado: pd.DataFrame,
                        df_resumo_txt: pd.DataFrame | None = None):

    print(f"Gerando Excel: {caminho_saida}")

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        workbook = writer.book
        estilos = criar_estilos(workbook)

        # ================== ABA CONSOLIDADO (primeira) ==================
        df_consolidado.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
        ws_cons = writer.sheets["CONSOLIDADO"]

        for col in range(len(df_consolidado.columns)):
            ws_cons.set_column(col, col, 18)

        _aplicar_condicional_diferenca(ws_cons, df_consolidado, estilos, coluna_nome="Saldo do dia")

        # Destaca TOTAL GERAL (linha inteira)
        for idx, row in df_consolidado.iterrows():
            if row["Mês"] == "TOTAL GERAL":
                for col in range(len(df_consolidado.columns)):  # A:H
                    ws_cons.write(idx + 1, col, df_consolidado.iloc[idx, col], estilos["total"])

        # ================== RESUMO DOS ARQUIVOS TXT ==================
        if df_resumo_txt is not None and not df_resumo_txt.empty:
            start_row = len(df_consolidado) + 3  # uma linha em branco
            # cabeçalho
            ws_cons.write(start_row, 0, "Resumo", estilos["total"])
            # tabela: colunas Tipo | Dias | Data1 | Data2 | Data3 | ...
            
            # escreve cabeçalho
            headers = ["Tipo", "Dias"] + [f"Data {i+1}" for i in range(len(df_resumo_txt.columns) - 2)]
            for c, h in enumerate(headers):
                ws_cons.write(start_row + 1, c, h, estilos["total"])

            # escreve dados
            for i, r in df_resumo_txt.iterrows():
                ws_cons.write(start_row + 2 + i, 0, str(r.get("Tipo", "")))
                ws_cons.write(start_row + 2 + i, 1, int(r.get("Dias", 0)))
                # escreve datas (colunas a partir de Data1, Data2, etc)
                col_idx = 2
                for col_name in df_resumo_txt.columns:
                    if col_name.startswith("Data"):
                        value = r.get(col_name, "")
                        # converte NaN para string vazia
                        if pd.isna(value):
                            value = ""
                        ws_cons.write(start_row + 2 + i, col_idx, str(value))
                        col_idx += 1

            # ajusta colunas para o bloco de resumo
            for col in range(len(df_resumo_txt.columns) + 1):
                ws_cons.set_column(col, col, 20)

        # ================== ABAS MENSAIS ==================
        for r in resultados:
            sheet_name = r["sheet"]
            df_mes = r["df"]

            df_mes.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            for col in range(len(df_mes.columns)):
                ws.set_column(col, col, 18)

            _aplicar_condicional_diferenca(ws, df_mes, estilos, coluna_nome="Saldo do dia")

            # Pintar linhas até a última coluna de dados (A:E)
            for idx, row in df_mes.iterrows():
                excel_row = idx + 1  # +1 por causa do cabeçalho
                tipo = row["Tipo do Dia"]
                data_str = row["Data"]

                if data_str == "TOTAL MÊS":
                    for col in range(6):  # A:F
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
                elif tipo == "Abono":
                    fmt = estilos["abono"]
                elif tipo == "Ajuste manual":
                    fmt = estilos["ajuste_manual"]
                elif tipo == "Final de Semana":
                    fmt = estilos["fds"]
                elif tipo == "Quarta-feira de Cinzas":
                    fmt = estilos["quartacinza"]                    
                else:
                    continue  # dia normal, não colore

                for col in range(6):  # A:F
                    ws.write(excel_row, col, df_mes.iloc[idx, col], fmt)

                for col in range(len(df_mes.columns)):
                    ws.set_column(col, col, 18)


    print("Excel gerado com sucesso!")
