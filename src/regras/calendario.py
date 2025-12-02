from datetime import date, timedelta


def pascoa(ano: int) -> date:
    """Algoritmo de Meeus/Jones/Butcher para calcular a data da Páscoa (domingo)."""
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)


def feriados_moveis(ano: int) -> set[date]:
    """
    Feriados móveis principais no Brasil:

    - Segunda de Carnaval
    - Terça de Carnaval
    - Quarta-feira de Cinzas
    - Sexta-feira Santa
    - Páscoa (domingo)
    - Corpus Christi
    """
    p = pascoa(ano)

    # Regra clássica:
    # Terça-feira de Carnaval = Páscoa - 47 dias
    terca_carnaval = p - timedelta(days=47)
    segunda_carnaval = terca_carnaval - timedelta(days=1)
    quarta_cinzas = terca_carnaval + timedelta(days=1)

    sexta_santa = p - timedelta(days=2)
    corpus_christi = p + timedelta(days=60)

    return {
        segunda_carnaval,
        terca_carnaval,
        quarta_cinzas,
        sexta_santa,
        p,               # Páscoa (domingo)
        corpus_christi,
    }


def feriados_ano(ano: int) -> set[date]:
    """Retorna todos os feriados (fixos + móveis) para o ano informado."""
    feriados: set[date] = set()

    # --- FERIADOS FIXOS NACIONAIS + DF ---
    feriados.add(date(ano, 1, 1))    # Confraternização Universal
    feriados.add(date(ano, 4, 21))   # Tiradentes
    feriados.add(date(ano, 5, 1))    # Dia do Trabalho
    feriados.add(date(ano, 9, 7))    # Independência
    feriados.add(date(ano, 10, 12))  # Nossa Senhora Aparecida
    feriados.add(date(ano, 11, 2))   # Finados
    feriados.add(date(ano, 11, 15))  # Proclamação da República
    feriados.add(date(ano, 11, 20))  # Dia da Consciência Negra
    feriados.add(date(ano, 11, 30))  # Dia do Evangélico (DF)
    feriados.add(date(ano, 12, 25))  # Natal

    # --- FERIADOS MÓVEIS ---
    feriados |= feriados_moveis(ano)

    return feriados
