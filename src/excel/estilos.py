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

    estilos["ferias"] = workbook.add_format({"bg_color": "#DDEBF7"})
    estilos["atestado"] = workbook.add_format({"bg_color": "#FFE699"})
    estilos["aniversario"] = workbook.add_format({"bg_color": "#F8BE8E"})
    estilos["feriado"] = workbook.add_format({"bg_color": "#F6E2F7"})
    estilos["fds"] = workbook.add_format({"bg_color": "#FFF2CC"}) 


    estilos["verde"] = workbook.add_format({"font_color": "green"})
    estilos["vermelho"] = workbook.add_format({"font_color": "red"})

    return estilos
