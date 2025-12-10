def criar_estilos(workbook):
    """
    Cria e retorna um dicionário com todos os formatos usados no Excel.
    """
    estilos = {}

    estilos["total"] = workbook.add_format({
        "bold": True,
        "bg_color": "#D9D9D9",
        "top": 2,
        "bottom": 2,
    })

    estilos["ferias"] = workbook.add_format({"bg_color": "#FFECB3"})       # Amarelo Pastel (Lazer/Sol)
    estilos["atestado"] = workbook.add_format({"bg_color": "#B3E5FC"})     # Azul Claro Pastel (Saúde/Calma)
    estilos["aniversario"] = workbook.add_format({"bg_color": "#FFCCBC"})  # Laranja/Pêssego Pastel (Celebração)
    estilos["feriado"] = workbook.add_format({"bg_color": "#D1C4E9"})      # Lilás/Roxo Pastel (Especial/Diferenciado)
    estilos["fds"] = workbook.add_format({"bg_color": "#CFD8DC"})          # Cinza Claro Azulado (Neutro/Descanso)
    estilos["abono"] = workbook.add_format({"bg_color": "#C8E6C9"})        # Verde Claro Pastel (Benefício/Positivo)
    estilos["quartacinza"] = workbook.add_format({"bg_color": "#BEBEBE"})    # Cinza Super Claro (Fundo Padrão/Neutro)
    estilos["ajuste_manual"] = workbook.add_format({"bg_color": "#F8D7DA"})  # Vermelho claro (ajuste manual)


    estilos["verde"] = workbook.add_format({"font_color": "green"})
    estilos["vermelho"] = workbook.add_format({"font_color": "red"})

    return estilos
