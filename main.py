from pathlib import Path
from src.data.arquivos_txt import carregar_lista_txt, carregar_pontos_manuais
from src.regras.regras_negocio import processar_mes
from src.excel.consolidado import montar_consolidado
from src.excel.writer import gerar_arquivo_excel

PASTA_HTML = "PONTO_HTML"
PASTA_EXCEL = "PONTO_EXCEL"

ARQ_FERIAS = "ferias.txt"
ARQ_ATESTADO = "atestados.txt"
ARQ_ANIVERSARIO = "aniversario.txt"
ARQ_ABONO = "abonos.txt"

def main():
    base_dir = Path(__file__).resolve().parent
    pasta_html = base_dir / PASTA_HTML
    pasta_excel = base_dir / PASTA_EXCEL
    pasta_excel.mkdir(exist_ok=True)

    data_dir = base_dir / "src" / "data"

    # Carrega listas de dias especiais
    ferias = carregar_lista_txt(data_dir, ARQ_FERIAS)
    atestados = carregar_lista_txt(data_dir, ARQ_ATESTADO)
    aniversario = carregar_lista_txt(data_dir, ARQ_ANIVERSARIO)
    abonos = carregar_lista_txt(data_dir, ARQ_ABONO)
    pontos_manuais = carregar_pontos_manuais(data_dir)

    print("📄 Carregando arquivos TXT...\n")
    
    print(f"• Férias: {len(ferias)} dias")
    print(f"• Atestados: {len(atestados)} dias")
    print(f"• Aniversário: {len(aniversario)} dia(s)")
    print(f"• Abonos: {len(abonos)} dias")
    print(f"• Pontos manuais: {len(pontos_manuais)} registros\n")

    print("📄 Lendo arquivos HTML...\n")

    # Lista arquivos HTML/HTM na pasta
    html_files = sorted(pasta_html.glob("*.html")) + sorted(pasta_html.glob("*.htm"))
    html_files = sorted(dict.fromkeys(html_files), key=lambda p: p.name)

    if not html_files:
        print(f"Nenhum HTML/HTM encontrado em: {pasta_html}")
        return

    resultados = []

    for html_file in html_files:
        print(f"📄 Processando: {html_file.name}")
        df_mes, resumo = processar_mes(html_file, ferias, atestados, aniversario, abonos, pontos_manuais)
        if df_mes is None:
            continue

        resultados.append({
            "sheet": resumo["nome_aba"],
            "df": df_mes,
            "resumo": {
                "Total Trabalhado": resumo["total_trabalhado"],
                "Total Previsto": resumo["total_previsto"],
                "Saldo do dia": resumo["saldo_dia"],
            },
        })

    if not resultados:
        print("Nenhum dado processado.")
        return

    # Monta DataFrame da aba CONSOLIDADO
    df_consolidado = montar_consolidado(resultados)

    # Monta resumo dos arquivos TXT para análise na aba CONSOLIDADO
    # Estrutura: Tipo | Dias | Data1 | Data2 | Data3 | ...
    import pandas as pd
    
    resumo_txt_rows = []

    # Agrupa férias contíguas em intervalos (INICIO;FIM)
    def agrupar_intervalos(dates_set):
        if not dates_set:
            return []
        dates = sorted(dates_set)
        intervals = []
        start = dates[0]
        end = dates[0]
        for d in dates[1:]:
            if (d - end).days == 1:
                end = d
            else:
                intervals.append((start, end))
                start = d
                end = d
        intervals.append((start, end))
        return intervals

    # Férias com intervalos (separar início e fim em colunas diferentes)
    ferias_intervals = agrupar_intervalos(ferias)
    ferias_datas = []
    for (s, e) in ferias_intervals:
        ferias_datas.append(s.isoformat())  # data de início
        if s != e:
            ferias_datas.append(e.isoformat())  # data de fim, se houver intervalo
    if ferias_intervals:
        row = {"Tipo": "Férias", "Dias": len(ferias)}
        for i, d in enumerate(ferias_datas):
            row[f"Data{i+1}"] = d
        resumo_txt_rows.append(row)

    # Atestados
    atestados_sorted = sorted(atestados)
    if atestados_sorted:
        row = {"Tipo": "Atestado", "Dias": len(atestados_sorted)}
        for i, d in enumerate(atestados_sorted):
            row[f"Data{i+1}"] = d.isoformat()
        resumo_txt_rows.append(row)

    # Aniversários
    aniversario_sorted = sorted(aniversario)
    if aniversario_sorted:
        row = {"Tipo": "Aniversário", "Dias": len(aniversario_sorted)}
        for i, d in enumerate(aniversario_sorted):
            row[f"Data{i+1}"] = d.isoformat()
        resumo_txt_rows.append(row)

    # Abonos
    abonos_sorted = sorted(abonos)
    if abonos_sorted:
        row = {"Tipo": "Abono", "Dias": len(abonos_sorted)}
        for i, d in enumerate(abonos_sorted):
            row[f"Data{i+1}"] = d.isoformat()
        resumo_txt_rows.append(row)

    # Pontos manuais
    pontos_manuais_sorted = sorted(pontos_manuais.keys())
    if pontos_manuais_sorted:
        row = {"Tipo": "Ponto Manual", "Dias": len(pontos_manuais_sorted)}
        for i, d in enumerate(pontos_manuais_sorted):
            row[f"Data{i+1}"] = d.isoformat()
        resumo_txt_rows.append(row)

    # Transforma em DataFrame
    df_resumo_txt = pd.DataFrame(resumo_txt_rows) if resumo_txt_rows else pd.DataFrame()

    # Gera arquivo Excel final
    caminho_saida = pasta_excel / "PONTOS_CONSOLIDADOS.xlsx"
    gerar_arquivo_excel(caminho_saida, resultados, df_consolidado, df_resumo_txt)

    print("Processamento concluído com sucesso!")

if __name__ == "__main__":
    main()
