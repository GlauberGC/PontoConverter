from pathlib import Path
from src.data.arquivos_txt import carregar_lista_txt
from src.regras.regras_negocio import processar_mes
from src.excel.consolidado import montar_consolidado
from src.excel.writer import gerar_arquivo_excel

PASTA_HTML = "PONTO_HTML"
PASTA_EXCEL = "PONTO_EXCEL"

ARQ_FERIAS = "ferias.txt"
ARQ_ATESTADO = "atestado.txt"
ARQ_ANIVERSARIO = "aniversario.txt"
ARQ_ABONO = "abono.txt"

def main():
    base_dir = Path(__file__).resolve().parent
    pasta_html = base_dir / PASTA_HTML
    pasta_excel = base_dir / PASTA_EXCEL
    pasta_excel.mkdir(exist_ok=True)

    data_dir = base_dir / "src" / "data"

    # Carrega listas de dias especiais
    ferias = carregar_lista_txt(data_dir, ARQ_FERIAS)
    atestados = carregar_lista_txt(data_dir, ARQ_ATESTADO)
    aniversarios = carregar_lista_txt(data_dir, ARQ_ANIVERSARIO)
    abonos = carregar_lista_txt(data_dir, ARQ_ABONO)

    # Lista arquivos HTML/HTM na pasta
    html_files = sorted(pasta_html.glob("*.html")) + sorted(pasta_html.glob("*.htm"))
    html_files = sorted(dict.fromkeys(html_files), key=lambda p: p.name)

    if not html_files:
        print(f"Nenhum HTML/HTM encontrado em: {pasta_html}")
        return

    resultados = []

    for html_file in html_files:
        print(f"📄 Processando: {html_file.name}")
        df_mes, resumo = processar_mes(html_file, ferias, atestados, aniversarios, abonos)
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

    # Gera arquivo Excel final
    caminho_saida = pasta_excel / "PONTOS_CONSOLIDADOS.xlsx"
    gerar_arquivo_excel(caminho_saida, resultados, df_consolidado)

    print("Processamento concluído com sucesso!")

if __name__ == "__main__":
    main()
