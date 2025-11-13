from datetime import date, timedelta

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
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = 1 + ((h + l - 7 * m + 114) % 31)
    return date(year, month, day)


def feriados_ano(year: int) -> set[date]:
    feriados = set()

    # Fixos
    feriados.add(date(year, 1, 1))   # Confraternização Universal
    feriados.add(date(year, 4, 21))  # Tiradentes
    feriados.add(date(year, 5, 1))   # Dia do Trabalho
    feriados.add(date(year, 9, 7))   # Independência
    feriados.add(date(year, 10, 12)) # Nossa Senhora Aparecida
    feriados.add(date(year, 11, 2))  # Finados
    feriados.add(date(year, 11, 15)) # Proclamação da República
    feriados.add(date(year, 11, 20)) # Dia da Consciência Negra
    feriados.add(date(year, 11, 30)) # Dia do Evangélico (DF)
    feriados.add(date(year, 12, 25)) # Natal


    # Móveis
    pascoa = easter_date(year)
    feriados.add(pascoa - timedelta(days=48))  # Carnaval (seg)
    feriados.add(pascoa - timedelta(days=47))  # Carnaval (ter)
    feriados.add(pascoa - timedelta(days=2))   # Sexta Santa
    feriados.add(pascoa + timedelta(days=60))  # Corpus Christi

    return feriados
