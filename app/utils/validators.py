from openpyxl.utils.cell import column_index_from_string


def normalize_column(value: str) -> str:
    column = (value or "").strip().upper()
    if not column:
        raise ValueError("Informe uma coluna (ex: A, B, C).")

    try:
        column_index_from_string(column)
    except ValueError as exc:
        raise ValueError("Coluna inválida. Use apenas letras, ex: A, AA, XFD.") from exc

    return column


def parse_percentage(value: str) -> float:
    raw = (value or "").strip().replace(",", ".")
    if not raw:
        raise ValueError("Informe uma porcentagem.")

    try:
        percentage = float(raw)
    except ValueError as exc:
        raise ValueError("Porcentagem inválida. Exemplo: 15 ou 2.5") from exc

    if percentage < 0:
        raise ValueError("A porcentagem não pode ser negativa.")

    return percentage
