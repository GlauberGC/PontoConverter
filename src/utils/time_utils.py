from datetime import timedelta

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
