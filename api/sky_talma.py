"""Motor completo e independiente para itinerarios de SKY Airline (H8 / H2) en Base Lima."""

import io
import re
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader

CITY_LOOKUP = {
    # Perú
    "LIM": "PELIM", "CUZ": "PECUZ", "AQP": "PEAQP", "PIU": "PEPIU", "TRU": "PETRU",
    "IQT": "PEIQT", "TPP": "PETPP", "CIX": "PECIX", "CJA": "PECJA", "TCQ": "PETCQ",
    "JUL": "PEJUL", "PEM": "PEPEM", "PCL": "PEPCL", "AYP": "PEAYP", "TBP": "PETBP",
    "TYL": "PETYL", "JAU": "PEJAU", "CHX": "PECHX", "ILQ": "PEILQ", "RIJ": "PERIJ",
    "HUU": "PEHUU", "ANS": "PEANS", "ATG": "PEATG",
    # Colombia
    "BOG": "COBOG", "MDE": "COMDE", "CTG": "COCTG", "CLO": "COCLO", "ADZ": "COADZ",
    "BAQ": "COBAQ", "SMR": "COSMR",
    # Chile
    "SCL": "CLSCL", "ANF": "CLANF", "IQQ": "CLIQQ", "CJC": "CLCJC", "CCP": "CLCCP",
    "PMC": "CLPMC", "PUQ": "CLPUQ",
    # Argentina
    "AEP": "ARAEP", "EZE": "AREZE", "COR": "ARCOR", "MDZ": "ARMDZ", "ROS": "ARROS",
    "SLA": "ARSLA", "IGR": "ARIGR",
    # Ecuador, Brasil, Uruguay, Paraguay, Bolivia
    "UIO": "ECUIO", "GYE": "ECGYE", "GRU": "BRGRU", "GIG": "BRGIG", "FLN": "BRFLN",
    "BSB": "BRBSB", "MVD": "UYMVD", "ASU": "PYASU", "VVI": "BOVVI", "LPB": "BOLPB",
    "CBB": "BOCBB",
    # Norteamérica, Caribe y Europa
    "MIA": "USMIA", "FLL": "USFLL", "JFK": "USJFK", "LAX": "USLAX", "MCO": "USMCO",
    "PUJ": "DOPUJ", "SDQ": "DOSDQ", "CUN": "MXCUN", "MEX": "MXMEX", "SJO": "CRSJO",
    "PTY": "PAPTY", "MAD": "ESMAD", "BCN": "ESBCN"
}

FLEET_LOOKUP_SKY = {
    "CC-DCA": "A-321", "CC-DCB": "A-321", "CC-DCC": "A-321", "CC-DCD": "A-321",
    "CC-DCE": "A-321", "CC-DCF": "A-321", "CC-DCG": "A-321", "CC-DCH": "A-321",
    "CC-DCI": "A-321", "CC-DCJ": "A-321"
}

SPANISH_MONTHS = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "setiembre": "09", "octubre": "10",
    "noviembre": "11", "diciembre": "12"
}

MONTH_ABBR = {
    "ENE": "01", "JAN": "01", "FEB": "02", "MAR": "03", "ABR": "04", "APR": "04",
    "MAY": "05", "JUN": "06", "JUL": "07", "AGO": "08", "AUG": "08",
    "SEP": "09", "SET": "09", "OCT": "10", "NOV": "11", "DIC": "12", "DEC": "12"
}

MAX_TURNAROUND_SKY = 90


def _resolve_sky_iata(operador_val, default_code="H8"):
    """Mapea SKU (SKY Chile) -> H2 y SKX (SKY Perú) -> H8."""
    if not operador_val:
        return default_code
    op_clean = str(operador_val).strip().upper()
    if op_clean in ("SKU", "H2"):
        return "H2"
    if op_clean in ("SKX", "H8"):
        return "H8"
    return default_code


def _get_aircraft(tail, flota_val=None):
    """Determina A-321 o A-320 usando la columna FLOTA de SKY o la matrícula."""
    if flota_val is not None:
        flota_str = str(flota_val).strip()
        if "321" in flota_str:
            return "A-321"
        if "320" in flota_str:
            return "A-320"
    if not tail:
        return "A-320"
    clean = str(tail).strip().upper()
    return "A-321" if (clean in FLEET_LOOKUP_SKY or clean.startswith("CC-DC")) else "A-320"


def _get_city_code(city):
    """Retorna el código de estación oficial con prefijo de país TALMA."""
    if not city:
        return ""
    city_str = str(city).strip().upper()
    return CITY_LOOKUP.get(city_str, f"PE{city_str}")


def _extract_date_from_filename(filename):
    """Extrae fecha completa desde el nombre del archivo (ej. '17NOV26' -> '17/11/2026')."""
    if not filename:
        return None
    m = re.search(r"(\d{1,2})([A-Z]{3})(\d{2,4})", filename.upper())
    if m:
        day = m.group(1).zfill(2)
        month = MONTH_ABBR.get(m.group(2))
        yr = m.group(3)
        year = f"20{yr}" if len(yr) == 2 else yr
        if month:
            return f"{day}/{month}/{year}"
    return None


def _parse_date(cell_val, filename=""):
    """Extrae la fecha operativa desde la celda del archivo o desde su nombre."""
    file_date = _extract_date_from_filename(filename)
    default_year = file_date.split(
        "/")[2] if file_date else str(datetime.today().year)

    if isinstance(cell_val, datetime):
        return cell_val.strftime("%d/%m/%Y")

    if cell_val:
        text = str(cell_val).strip()
        m = re.search(
            r"(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)(?:\s+(?:del?|de)\s+(\d{4}))?",
            text,
            re.I,
        )
        if m:
            day = m.group(1).zfill(2)
            month = SPANISH_MONTHS.get(m.group(2).lower(), "01")
            year = m.group(3) if m.group(3) else default_year
            return f"{day}/{month}/{year}"

        m_num = re.search(r"(\d{2})[/-](\d{2})[/-](\d{4})", text)
        if m_num:
            return f"{m_num.group(1)}/{m_num.group(2)}/{m_num.group(3)}"

    return file_date if file_date else datetime.today().strftime("%d/%m/%Y")


def _fmt_time(t):
    """Normaliza objetos o cadenas de hora al formato HH:MM."""
    if isinstance(t, (time, datetime)):
        return t.strftime("%H:%M")
    match = re.search(r"\b(\d{1,2}:\d{2})\b", str(t).strip())
    return match.group(1).zfill(5) if match else ""


def _to_excel_date(ddmmyyyy_str):
    try:
        return datetime.strptime(ddmmyyyy_str, "%d/%m/%Y")
    except ValueError:
        return datetime.today()


def _to_excel_time(hhmm_str):
    if not hhmm_str:
        return None
    try:
        h, m = map(int, hhmm_str.split(":"))
        return time(h, m)
    except (ValueError, AttributeError):
        return None


def _add_days(ddmmyyyy_str, days=1):
    return (_to_excel_date(ddmmyyyy_str) + timedelta(days=days)).strftime("%d/%m/%Y")


def _to_mins(hhmm_str):
    try:
        h, m = map(int, hhmm_str.split(":"))
        return h * 60 + m
    except (ValueError, AttributeError):
        return -1


def _is_next_day_arr(std_str, sta_str):
    """Detecta si una llegada a LIM despega de noche (>= 18:00) y aterriza >= 00:00."""
    m_std, m_sta = _to_mins(std_str), _to_mins(sta_str)
    if m_sta < 0:
        return False
    return (m_std >= 18 * 60 and m_sta <= 4 * 60) or m_sta == 0


def _diff_mins(eta_str, eta_next, etd_str, etd_next):
    m_eta, m_etd = _to_mins(eta_str), _to_mins(etd_str)
    if m_eta < 0 or m_etd < 0:
        return -1
    if eta_next:
        m_eta += 1440
    if etd_next:
        m_etd += 1440
    return m_etd - m_eta


def _calc_tvuelo(eta_str, eta_next, etd_str, etd_next):
    mins = _diff_mins(eta_str, eta_next, etd_str, etd_next)
    if mins < 0:
        return ""
    h, m = divmod(mins, 60)
    return f"{h:02d}:{m:02d}"


def _sort_key(hhmm_str, next_day=False):
    mins = _to_mins(hhmm_str)
    if mins < 0:
        return 9999
    return mins + 1440 if next_day else mins


def _extraer_sky_excel(file_bytes, filename="", default_code="H8", base_iata="LIM"):
    """Extrae llegadas y salidas de Lima desde el Excel oficial de SKY (.xlsx)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    op_date = _extract_date_from_filename(
        filename) or datetime.today().strftime("%d/%m/%Y")
    encontro_fecha = False
    for r_top in range(1, 5):
        for c_top in range(1, min(ws.max_column + 1, 10)):
            val_top = ws.cell(r_top, c_top).value
            if val_top and (
                isinstance(val_top, datetime)
                or re.search(r"\d{1,2}\s+de\s+[a-zA-Z]+", str(val_top), re.I)
                or re.search(r"\d{2}[/-]\d{2}[/-]\d{4}", str(val_top))
            ):
                op_date = _parse_date(val_top, filename=filename)
                encontro_fecha = True
                break
        if encontro_fecha:
            break

    header_row = None
    col_map = {}

    for r in range(1, 15):
        row_vals = [
            str(ws.cell(r, c).value or "").strip().upper()
            for c in range(1, ws.max_column + 1)
        ]
        if any("MAT" in v for v in row_vals) and (
            any("ORI" in v for v in row_vals) or any(
                "DEP" in v for v in row_vals)
        ):
            header_row = r
            for c in range(1, ws.max_column + 1):
                v = str(ws.cell(r, c).value or "").strip().upper()
                if "FEC" in v or "DATE" in v:
                    col_map["FECHA"] = c
                elif "MAT" in v or "REG" in v:
                    col_map["MATRICULA"] = c
                elif "LÍNEA" in v or "LINEA" in v or "OPER" in v or "CIA" in v or "AL" == v:
                    col_map["LINEA"] = c
                elif "VLO" in v or "VUELO" in v or "FLT" in v:
                    col_map["VUELO"] = c
                elif "FLO" in v or "EQUIP" in v or "AERO" in v:
                    col_map["FLOTA"] = c
                elif "ORI" in v or "DEP" in v:
                    col_map["ORIGEN"] = c
                elif "DEST" in v or "ARR" in v:
                    col_map["DESTINO"] = c
                elif "STD" in v or "ETD" in v or "SAL" in v:
                    col_map["STD"] = c
                elif "STA" in v or "ETA" in v or "LLE" in v:
                    col_map["STA"] = c
            break

    if not header_row:
        header_row = 4
        col_map = {
            "MATRICULA": 1, "LINEA": 2, "VUELO": 3, "FLOTA": 4,
            "ORIGEN": 5, "DESTINO": 6, "STD": 7, "STA": 8
        }

    vuelos_llegada, vuelos_salida = [], []
    vio_turno_tarde = False

    for r in range(header_row + 1, ws.max_row + 1):
        mat = str(ws.cell(r, col_map.get("MATRICULA", 1)
                          ).value or "").strip().upper()
        linea_val = ws.cell(r, col_map.get("LINEA", 2)
                            ).value if "LINEA" in col_map else None
        vlo_raw = str(ws.cell(r, col_map.get("VUELO", 3)).value or "").strip()
        flota_val = ws.cell(r, col_map.get("FLOTA", 4)
                            ).value if "FLOTA" in col_map else None
        orig = str(ws.cell(r, col_map.get("ORIGEN", 5)
                           ).value or "").strip().upper()
        dest = str(ws.cell(r, col_map.get("DESTINO", 6)
                           ).value or "").strip().upper()
        std_val = ws.cell(r, col_map.get("STD", 7)).value
        sta_val = ws.cell(r, col_map.get("STA", 8)).value

        vlo_digits = "".join(filter(str.isdigit, vlo_raw))
        if not vlo_digits or not mat or mat in ["NONE", "MATRICULA", "REG"]:
            continue

        # SKU -> H2 (SKY Chile), SKX -> H8 (SKY Perú)
        al_iata = _resolve_sky_iata(linea_val, default_code=default_code)
        if not linea_val and int(vlo_digits) < 1000:
            al_iata = "H2"

        std_str, sta_str = _fmt_time(std_val), _fmt_time(sta_val)
        m_std = _to_mins(std_str)
        if m_std >= 18 * 60:
            vio_turno_tarde = True

        if dest == base_iata and sta_str:
            next_day = _is_next_day_arr(std_str, sta_str)
            vuelos_llegada.append({
                "aerolinea": al_iata,
                "matricula": mat,
                "vuelo": int(vlo_digits),
                "orig": orig,
                "std": std_str,
                "sta": sta_str,
                "aeronave": _get_aircraft(mat, flota_val),
                "next_day": next_day,
            })
        elif orig == base_iata and std_str:
            next_day = vio_turno_tarde and (0 <= m_std < 4 * 60)
            vuelos_salida.append({
                "aerolinea": al_iata,
                "matricula": mat,
                "vuelo": int(vlo_digits),
                "dest": dest,
                "std": std_str,
                "sta": sta_str,
                "aeronave": _get_aircraft(mat, flota_val),
                "next_day": next_day,
            })

    return op_date, vuelos_llegada, vuelos_salida


def _extraer_sky_pdf(file_bytes, filename="", default_code="H8", base_iata="LIM"):
    """Extrae llegadas y salidas de Lima desde un archivo PDF oficial de SKY."""
    reader = PdfReader(io.BytesIO(file_bytes))
    full_text = "\n".join(p.extract_text() or "" for p in reader.pages)

    op_date = _parse_date(full_text[:500], filename=filename)
    vuelos_llegada, vuelos_salida = [], []
    current_tail = ""
    vio_turno_tarde = False

    flight_line_re = re.compile(
        r"(H8|H2|SKX|SKU)?\s*(\d{3,4})\s+"
        r"(?:(CC-[A-Z0-9]{3}|LV-[A-Z0-9]{3})\s+)?"
        r"([A-Z]{3})\s*[-–/]?\s*([A-Z]{3})\s+"
        r"(\d{1,2}:\d{2})\s+"
        r"(\d{1,2}:\d{2})"
    )

    for line in full_text.split("\n"):
        line_clean = line.strip().upper()
        if not line_clean:
            continue

        tail_match = re.search(
            r"\b(CC-[A-Z0-9]{3}|LV-[A-Z0-9]{3})\b", line_clean)
        if tail_match:
            current_tail = tail_match.group(1)

        m = flight_line_re.search(line_clean)
        if m:
            op_prefix, vlo_str, inline_tail, orig, dest, std_raw, sta_raw = m.groups()
            mat = inline_tail if inline_tail else current_tail
            if not mat:
                continue

            al_iata = _resolve_sky_iata(op_prefix, default_code=default_code)
            if not op_prefix and int(vlo_str) < 1000:
                al_iata = "H2"

            std_str, sta_str = std_raw.zfill(5), sta_raw.zfill(5)
            m_std = _to_mins(std_str)
            if m_std >= 18 * 60:
                vio_turno_tarde = True

            if dest == base_iata:
                vuelos_llegada.append({
                    "aerolinea": al_iata,
                    "matricula": mat,
                    "vuelo": int(vlo_str),
                    "orig": orig,
                    "std": std_str,
                    "sta": sta_str,
                    "aeronave": _get_aircraft(mat),
                    "next_day": _is_next_day_arr(std_str, sta_str),
                })
            elif orig == base_iata:
                vuelos_salida.append({
                    "aerolinea": al_iata,
                    "matricula": mat,
                    "vuelo": int(vlo_str),
                    "dest": dest,
                    "std": std_str,
                    "sta": sta_str,
                    "aeronave": _get_aircraft(mat),
                    "next_day": vio_turno_tarde and (0 <= m_std < 4 * 60),
                })

    return op_date, vuelos_llegada, vuelos_salida


def procesar_sky_talma(file_bytes: bytes, filename: str, airline_code="H8", base_iata="LIM"):
    """Procesa archivos Excel o PDF de SKY Airline y genera el Excel oficial TALMA."""
    if filename.lower().endswith(".pdf"):
        op_date, vuelos_llegada, vuelos_salida = _extraer_sky_pdf(
            file_bytes, filename=filename, default_code=airline_code, base_iata=base_iata
        )
    else:
        op_date, vuelos_llegada, vuelos_salida = _extraer_sky_excel(
            file_bytes, filename=filename, default_code=airline_code, base_iata=base_iata
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
        eventos.sort(key=lambda e: (
            _sort_key(e["hora"], e["next_day"]), 0 if e["tipo"] == "L" else 1))
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
                    can_pair = 0 <= mins <= MAX_TURNAROUND_SKY

                if can_pair:
                    dep = eventos[idx + 1]["data"]
                    dep_next = dep["next_day"]
                    row_next = arr_next or dep_next
                    row_date = _add_days(op_date, 1) if row_next else op_date
                    operaciones.append({
                        "FECHA": row_date, "NEXT_DAY": row_next,
                        "AEROLINEA_L": arr.get("aerolinea", airline_code),
                        "MATRICULA_L": arr["matricula"],
                        "AERONAVE_L": arr["aeronave"], "VUELO_L": arr["vuelo"],
                        "CIUDAD_L": _get_city_code(arr["orig"]), "PEA_L": "", "ETA": arr["sta"],
                        "AEROLINEA_S": dep.get("aerolinea", airline_code),
                        "MATRICULA_S": dep["matricula"],
                        "AERONAVE_S": dep["aeronave"], "VUELO_S": dep["vuelo"],
                        "CIUDAD_S": _get_city_code(dep["dest"]), "PEA_S": "", "ETD": dep["std"],
                        "T_VUELO": _calc_tvuelo(arr["sta"], arr_next, dep["std"], dep_next),
                    })
                    idx += 2
                else:
                    row_date = _add_days(op_date, 1) if arr_next else op_date
                    operaciones.append({
                        "FECHA": row_date, "NEXT_DAY": arr_next,
                        "AEROLINEA_L": arr.get("aerolinea", airline_code),
                        "MATRICULA_L": arr["matricula"],
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
                    "AEROLINEA_S": dep.get("aerolinea", airline_code),
                    "MATRICULA_S": dep["matricula"],
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
        "AEROLINEA_S", "MATRICULA_S", "AERONAVE_S", "VUELO_S", "CIUDAD_S", "PEA_S", "ETD", "T_VUELO"
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
        "I": 12, "J": 13, "K": 14, "L": 10, "M": 12, "N": 10, "O": 10, "P": 14
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
