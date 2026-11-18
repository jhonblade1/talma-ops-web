"""Motor completo para itinerarios de StarPerú (2I) desde PDF en Base Lima (LIM)."""

import io
import re
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader

CITY_LOOKUP = {
    "LIM": "PELIM", "CUZ": "PECUZ", "AQP": "PEAQP", "PIU": "PEPIU", "TRU": "PETRU",
    "IQT": "PEIQT", "TPP": "PETPP", "CIX": "PECIX", "CJA": "PECJA", "TCQ": "PETCQ",
    "JUL": "PEJUL", "PEM": "PEPEM", "PCL": "PEPCL", "AYP": "PEAYP", "TBP": "PETBP",
    "TYL": "PETYL", "JAU": "PEJAU", "CHX": "PECHX", "ILQ": "PEILQ", "RIJ": "PERIJ",
    "HUU": "PEHUU", "ANS": "PEANS", "ATG": "PEATG", "VVN": "PEVVN",
}

SPANISH_MONTHS = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "setiembre": "09", "octubre": "10",
    "noviembre": "11", "diciembre": "12",
}

MAX_TURNAROUND_STAR = 120


def _get_aircraft_star(matricula: str, vuelo: int) -> str:
    """Asigna B-733 para flota Boeing 737 (OB...) y Q-400 para Dash 8 Q400 (N... / serie 4xxx / 1xx)."""
    mat_clean = (matricula or "").strip().upper()
    if mat_clean.startswith("N") or vuelo >= 4000 or vuelo < 1000:
        return "Q-400"
    return "B-733"


def _get_city_code(city: str) -> str:
    if not city:
        return ""
    city_str = str(city).strip().upper()
    return CITY_LOOKUP.get(city_str, f"PE{city_str}")


def _parse_starperu_date(full_text: str, filename: str = "") -> str:
    """Extrae la fecha del encabezado 'JUEVES 5 DE NOVIEMBRE DEL 2026' -> '05/11/2026'."""
    m = re.search(
        r"(\d{1,2})\s+DE\s+([A-ZÁÉÍÓÚÑ]+)\s+(?:DEL?|DE)\s+(\d{4})",
        full_text.upper(),
    )
    if m:
        day = m.group(1).zfill(2)
        month = SPANISH_MONTHS.get(m.group(2).lower(), "01")
        year = m.group(3)
        return f"{day}/{month}/{year}"

    m_num = re.search(r"(\d{2})[/-](\d{2})[/-](\d{4})", full_text)
    if m_num:
        return f"{m_num.group(1)}/{m_num.group(2)}/{m_num.group(3)}"

    if filename:
        m_fn = re.search(r"(\d{2})[._-](\d{2})[._-](\d{2,4})", filename)
        if m_fn:
            yr = m_fn.group(3)
            year = f"20{yr}" if len(yr) == 2 else yr
            return f"{m_fn.group(1)}/{m_fn.group(2)}/{year}"

    return datetime.today().strftime("%d/%m/%Y")


def _to_excel_date(ddmmyyyy_str: str) -> datetime:
    try:
        return datetime.strptime(ddmmyyyy_str, "%d/%m/%Y")
    except ValueError:
        return datetime.today()


def _to_excel_time(hhmm_str: str):
    if not hhmm_str:
        return None
    try:
        h, m = map(int, hhmm_str.split(":"))
        return time(h, m)
    except (ValueError, AttributeError):
        return None


def _add_days(ddmmyyyy_str: str, days: int = 1) -> str:
    return (_to_excel_date(ddmmyyyy_str) + timedelta(days=days)).strftime("%d/%m/%Y")


def _to_mins(hhmm_str: str) -> int:
    try:
        h, m = map(int, hhmm_str.split(":"))
        return h * 60 + m
    except (ValueError, AttributeError):
        return -1


def _is_next_day_arr(etd_str: str, eta_str: str) -> bool:
    m_etd, m_eta = _to_mins(etd_str), _to_mins(eta_str)
    if m_eta < 0:
        return False
    return (m_etd >= 18 * 60 and m_eta <= 4 * 60) or m_eta == 0


def _diff_mins(eta_str: str, eta_next: bool, etd_str: str, etd_next: bool) -> int:
    m_eta, m_etd = _to_mins(eta_str), _to_mins(etd_str)
    if m_eta < 0 or m_etd < 0:
        return -1
    if eta_next:
        m_eta += 1440
    if etd_next:
        m_etd += 1440
    return m_etd - m_eta


def _calc_tvuelo(eta_str: str, eta_next: bool, etd_str: str, etd_next: bool) -> str:
    mins = _diff_mins(eta_str, eta_next, etd_str, etd_next)
    if mins < 0:
        return ""
    h, m = divmod(mins, 60)
    return f"{h:02d}:{m:02d}"


def _sort_key(hhmm_str: str, next_day: bool = False) -> int:
    mins = _to_mins(hhmm_str)
    if mins < 0:
        return 9999
    return mins + 1440 if next_day else mins


def _extraer_starperu_pdf(file_bytes: bytes, filename: str = "", base_iata: str = "LIM"):
    """Extrae llegadas y salidas de Lima desde el PDF SRU-FOR-DSP-012 de StarPerú."""
    reader = PdfReader(io.BytesIO(file_bytes))
    plain_text = "\n".join(p.extract_text() or "" for p in reader.pages)
    op_date = _parse_starperu_date(plain_text, filename=filename)

    flight_line_re = re.compile(
        r"^\s*(\d{1,2})\s+(\d{2,4})\s+([A-Z]{3})\s+([A-Z]{3})\s+(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})\b"
    )
    reg_line_re = re.compile(r"^\s*\d{1,2}\s+((?:OB|N)-?[A-Z0-9]{4,6})\s*$")

    parsed_flights = []
    reg_markers = []
    section = "REGULARES"

    for raw_line in plain_text.split("\n"):
        upper_line = raw_line.strip().upper()
        if not upper_line:
            continue
        if "VUELOS NO REGULARES" in upper_line:
            section = "NO_REGULARES"

        m_f = flight_line_re.match(upper_line)
        if m_f:
            parsed_flights.append({
                "section": section,
                "seq": int(m_f.group(1)),
                "vuelo": int(m_f.group(2)),
                "orig": m_f.group(3),
                "dest": m_f.group(4),
                "etd": m_f.group(5).zfill(5),
                "eta": m_f.group(6).zfill(5),
                "matricula": "",
            })
            continue

        m_r = reg_line_re.match(upper_line)
        if m_r:
            reg_markers.append(m_r.group(1).strip())

    # Agrupar vuelos en cadenas continuas por aeronave (LIM -> ... -> LIM)
    chains = []
    current_chain = []
    for fl in parsed_flights:
        if not current_chain:
            current_chain.append(fl)
        else:
            prev = current_chain[-1]
            if (
                prev["section"] == fl["section"]
                and prev["dest"] == fl["orig"]
                and prev["dest"] != base_iata
            ):
                current_chain.append(fl)
            else:
                chains.append(current_chain)
                current_chain = [fl]
    if current_chain:
        chains.append(current_chain)

    for idx_c, chain in enumerate(chains):
        mat = reg_markers[idx_c] if idx_c < len(reg_markers) else ""
        for fl in chain:
            fl["matricula"] = mat

    vuelos_llegada = []
    vuelos_salida = []
    vio_turno_tarde = False

    for fl in parsed_flights:
        orig = fl["orig"]
        dest = fl["dest"]
        etd_str = fl["etd"]
        eta_str = fl["eta"]
        mat = fl["matricula"]
        vlo = fl["vuelo"]

        m_etd = _to_mins(etd_str)
        if m_etd >= 18 * 60:
            vio_turno_tarde = True

        if dest == base_iata and eta_str:
            vuelos_llegada.append({
                "matricula": mat,
                "vuelo": vlo,
                "orig": orig,
                "std": etd_str,
                "sta": eta_str,
                "aeronave": _get_aircraft_star(mat, vlo),
                "next_day": _is_next_day_arr(etd_str, eta_str),
            })
        elif orig == base_iata and etd_str:
            vuelos_salida.append({
                "matricula": mat,
                "vuelo": vlo,
                "dest": dest,
                "std": etd_str,
                "sta": eta_str,
                "aeronave": _get_aircraft_star(mat, vlo),
                "next_day": vio_turno_tarde and (0 <= m_etd < 4 * 60),
            })

    return op_date, vuelos_llegada, vuelos_salida


def procesar_starperu_talma(file_bytes: bytes, filename: str, airline_code="2I", base_iata="LIM"):
    """Procesa el PDF de StarPerú y genera el archivo Excel oficial TALMA."""
    op_date, vuelos_llegada, vuelos_salida = _extraer_starperu_pdf(
        file_bytes, filename=filename, base_iata=base_iata
    )

    eventos_por_mat = {}
    for arr in vuelos_llegada:
        eventos_por_mat.setdefault(arr["matricula"], []).append(
            {"tipo": "L", "hora": arr["sta"],
                "next_day": arr["next_day"], "data": arr}
        )
    for dep in vuelos_salida:
        eventos_por_mat.setdefault(dep["matricula"], []).append(
            {"tipo": "S", "hora": dep["std"],
                "next_day": dep["next_day"], "data": dep}
        )

    operaciones = []
    for _mat, eventos in eventos_por_mat.items():
        eventos.sort(
            key=lambda e: (
                _sort_key(e["hora"], e["next_day"]), 0 if e["tipo"] == "L" else 1)
        )
        idx, total = 0, len(eventos)

        while idx < total:
            actual = eventos[idx]
            if actual["tipo"] == "L":
                arr = actual["data"]
                arr_next = arr["next_day"]
                can_pair = False
                if idx + 1 < total and eventos[idx + 1]["tipo"] == "S":
                    dep_cand = eventos[idx + 1]["data"]
                    mins = _diff_mins(arr["sta"], arr_next,
                                      dep_cand["std"], dep_cand["next_day"])
                    can_pair = 0 <= mins <= MAX_TURNAROUND_STAR

                if can_pair:
                    dep = eventos[idx + 1]["data"]
                    dep_next = dep["next_day"]
                    row_next = arr_next or dep_next
                    row_date = _add_days(op_date, 1) if row_next else op_date
                    operaciones.append({
                        "FECHA": row_date, "NEXT_DAY": row_next,
                        "AEROLINEA_L": airline_code, "MATRICULA_L": arr["matricula"],
                        "AERONAVE_L": arr["aeronave"], "VUELO_L": arr["vuelo"],
                        "CIUDAD_L": _get_city_code(arr["orig"]), "PEA_L": "", "ETA": arr["sta"],
                        "AEROLINEA_S": airline_code, "MATRICULA_S": dep["matricula"],
                        "AERONAVE_S": dep["aeronave"], "VUELO_S": dep["vuelo"],
                        "CIUDAD_S": _get_city_code(dep["dest"]), "PEA_S": "", "ETD": dep["std"],
                        "T_VUELO": _calc_tvuelo(arr["sta"], arr_next, dep["std"], dep_next),
                    })
                    idx += 2
                else:
                    row_date = _add_days(op_date, 1) if arr_next else op_date
                    operaciones.append({
                        "FECHA": row_date, "NEXT_DAY": arr_next,
                        "AEROLINEA_L": airline_code, "MATRICULA_L": arr["matricula"],
                        "AERONAVE_L": arr["aeronave"], "VUELO_L": arr["vuelo"],
                        "CIUDAD_L": _get_city_code(arr["orig"]), "PEA_L": "", "ETA": arr["sta"],
                        "AEROLINEA_S": "", "MATRICULA_S": "", "AERONAVE_S": "",
                        "VUELO_S": "", "CIUDAD_S": "", "PEA_S": "", "ETD": "", "T_VUELO": "",
                    })
                    idx += 1
            else:
                dep = actual["data"]
                dep_next = dep["next_day"]
                row_date = _add_days(op_date, 1) if dep_next else op_date
                operaciones.append({
                    "FECHA": row_date, "NEXT_DAY": dep_next,
                    "AEROLINEA_L": "", "MATRICULA_L": "", "AERONAVE_L": "",
                    "VUELO_L": "", "CIUDAD_L": "", "PEA_L": "", "ETA": "",
                    "AEROLINEA_S": airline_code, "MATRICULA_S": dep["matricula"],
                    "AERONAVE_S": dep["aeronave"], "VUELO_S": dep["vuelo"],
                    "CIUDAD_S": _get_city_code(dep["dest"]), "PEA_S": "", "ETD": dep["std"],
                    "T_VUELO": "",
                })
                idx += 1

    operaciones_validas = [
        op for op in operaciones
        if op.get("VUELO_L") not in ("", None) or op.get("VUELO_S") not in ("", None)
    ]
    operaciones_validas.sort(
        key=lambda x: _sort_key(x["ETA"] if x["ETA"]
                                else x["ETD"], x.get("NEXT_DAY", False))
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ITINERARIOS"

    headers = [
        "FECHA", "AEROLINEA_L", "MATRICULA_L", "AERONAVE_L", "VUELO_L", "CIUDAD_L", "PEA_L", "ETA",
        "AEROLINEA_S", "MATRICULA_S", "AERONAVE_S", "VUELO_S", "CIUDAD_S", "PEA_S", "ETD", "T_VUELO",
    ]
    ws.append(headers)

    yellow_fill = PatternFill(start_color="FFFF00",
                              end_color="FFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="92D050",
                             end_color="92D050", fill_type="solid")
    base_font = Font(name="Calibri", size=11, bold=False)
    thin_side = Side(style="thin", color="000000")
    cell_border = Border(left=thin_side, right=thin_side,
                         top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal="center", vertical="center")

    col_widths = {
        "A": 11, "B": 12, "C": 13, "D": 13, "E": 10, "F": 12, "G": 10, "H": 10,
        "I": 12, "J": 13, "K": 14, "L": 10, "M": 12, "N": 10, "O": 10, "P": 14,
    }
    for letter, width in col_widths.items():
        ws.column_dimensions[letter].width = width

    for col_idx in range(1, 17):
        c = ws.cell(1, col_idx)
        c.font = base_font
        c.alignment = center_align
        c.border = cell_border
        if 2 <= col_idx <= 8:
            c.fill = yellow_fill
        elif 9 <= col_idx <= 16:
            c.fill = green_fill

    for idx_row, op in enumerate(operaciones_validas, start=2):
        excel_date_obj = _to_excel_date(op["FECHA"])
        eta_obj = _to_excel_time(op["ETA"])
        etd_obj = _to_excel_time(op["ETD"])

        row_vals = [
            excel_date_obj,
            op["AEROLINEA_L"] or None, op["MATRICULA_L"] or None, op["AERONAVE_L"] or None,
            int(op["VUELO_L"]) if op["VUELO_L"] not in ("", None) else None,
            op["CIUDAD_L"] or None, None, eta_obj,
            op["AEROLINEA_S"] or None, op["MATRICULA_S"] or None, op["AERONAVE_S"] or None,
            int(op["VUELO_S"]) if op["VUELO_S"] not in ("", None) else None,
            op["CIUDAD_S"] or None, None, etd_obj, None,
        ]

        for col_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=idx_row, column=col_idx, value=val)
            c.font = base_font
            c.alignment = center_align
            c.border = cell_border
            if 2 <= col_idx <= 8:
                c.fill = yellow_fill
            elif 9 <= col_idx <= 16:
                c.fill = green_fill

        ws.cell(idx_row, 1).number_format = "mm-dd-yy"
        if eta_obj:
            ws.cell(idx_row, 8).number_format = "h:mm"
        if etd_obj:
            ws.cell(idx_row, 15).number_format = "h:mm"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return op_date, operaciones_validas, output.getvalue()
