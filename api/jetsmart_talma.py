"""Motor completo e independiente para itinerarios de JetSMART (JZ) en Base Lima."""

import io
import re
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

CITY_LOOKUP = {
    "LIM": "PELIM", "CUZ": "PECUZ", "AQP": "PEAQP", "PIU": "PEPIU", "TRU": "PETRU",
    "IQT": "PEIQT", "TPP": "PETPP", "CIX": "PECIX", "CJA": "PECJA", "TCQ": "PETCQ",
    "JUL": "PEJUL", "PEM": "PEPEM", "PCL": "PEPCL", "AYP": "PEAYP", "TBP": "PETBP",
    "TYL": "PETYL", "JAU": "PEJAU", "CHX": "PECHX", "ILQ": "PEILQ", "RIJ": "PERIJ",
    "HUU": "PEHUU", "ANS": "PEANS", "ATG": "PEATG",
    "BOG": "COBOG", "MDE": "COMDE", "CTG": "COCTG", "CLO": "COCLO", "ADZ": "COADZ",
    "BAQ": "COBAQ", "SMR": "COSMR",
    "SCL": "CLSCL", "ANF": "CLANF", "IQQ": "CLIQQ", "CJC": "CLCJC", "CCP": "CLCCP",
    "PMC": "CLPMC", "PUQ": "CLPUQ",
    "AEP": "ARAEP", "EZE": "AREZE", "COR": "ARCOR", "MDZ": "ARMDZ", "ROS": "ARROS",
    "SLA": "ARSLA", "IGR": "ARIGR",
    "UIO": "ECUIO", "GYE": "ECGYE", "GRU": "BRGRU", "GIG": "BRGIG", "FLN": "BRFLN",
    "BSB": "BRBSB", "MVD": "UYMVD", "ASU": "PYASU", "VVI": "BOVVI", "LPB": "BOLPB",
    "CBB": "BOCBB",
    "MIA": "USMIA", "FLL": "USFLL", "JFK": "USJFK", "LAX": "USLAX", "MCO": "USMCO",
    "PUJ": "DOPUJ", "SDQ": "DOSDQ", "CUN": "MXCUN", "MEX": "MXMEX", "SJO": "CRSJO",
    "PTY": "PAPTY", "MAD": "ESMAD", "BCN": "ESBCN"
}

FLEET_LOOKUP_JZ = {
    "CC-DCA": "A-321", "CC-DCB": "A-321", "CC-DCC": "A-321", "CC-DCD": "A-321",
    "CC-DCE": "A-321", "CC-DCF": "A-321"
}

MAX_TURNAROUND_JZ = 90


def _get_aircraft(tail):
    if not tail:
        return "A-320"
    clean = str(tail).strip().upper()
    return "A-321" if (clean in FLEET_LOOKUP_JZ or clean.startswith("CC-DC")) else "A-320"


def _get_city_code(city):
    if not city:
        return ""
    city_str = str(city).strip().upper()
    return CITY_LOOKUP.get(city_str, f"PE{city_str}")


def _parse_date(cell_val):
    if isinstance(cell_val, datetime):
        return cell_val.strftime("%d/%m/%Y")
    if not cell_val:
        return datetime.today().strftime("%d/%m/%Y")
    m_num = re.search(r"(\d{2})[/-](\d{2})[/-](\d{4})", str(cell_val))
    if m_num:
        return f"{m_num.group(1)}/{m_num.group(2)}/{m_num.group(3)}"
    return datetime.today().strftime("%d/%m/%Y")


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
    m_std, m_sta = _to_mins(std_str), _to_mins(sta_str)
    if m_sta < 0:
        return False
    return (m_std >= 18 * 60 and m_sta <= 4 * 60) or m_sta == 0


def _is_next_day_dep(std_str):
    m_std = _to_mins(std_str)
    return 0 <= m_std < 4 * 60


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
    return mins + 1440 if (next_day or mins == 0) else mins


def procesar_jetsmart_talma(raw_text: str, airline_code="JZ", base_iata="LIM"):
    """Procesa itinerarios pegados de JetSMART y genera el Excel oficial TALMA."""
    vuelos_llegada = []
    vuelos_salida = []
    op_date = datetime.today().strftime("%d/%m/%Y")

    flat_text = re.sub(r"\s+", " ", raw_text.strip()).upper()
    row_pattern = re.compile(
        r"(\d{2}[/-]\d{2}[/-]\d{4})\s+"
        r"(\d+)\s+"
        r"(\d+)\s+"
        r"((?:CC|LV)-[A-Z0-9]{3})\s+"
        r"([A-Z]{3})\s+"
        r"([A-Z]{3})\s+"
        r"(\d{1,2}:\d{2})\s+"
        r"(\d{1,2}:\d{2})"
    )

    for m in row_pattern.finditer(flat_text):
        d_str, _route, vlo_str, reg_str, dep_str, arr_str, std_raw, sta_raw = m.groups()
        op_date = _parse_date(d_str)
        std_str, sta_str = std_raw.zfill(5), sta_raw.zfill(5)

        if arr_str == base_iata:
            vuelos_llegada.append({
                "matricula": reg_str,
                "vuelo": int(vlo_str),
                "orig": dep_str,
                "std": std_str,
                "sta": sta_str,
                "aeronave": _get_aircraft(reg_str),
                "next_day": _is_next_day_arr(std_str, sta_str),
            })
        elif dep_str == base_iata:
            vuelos_salida.append({
                "matricula": reg_str,
                "vuelo": int(vlo_str),
                "dest": arr_str,
                "std": std_str,
                "sta": sta_str,
                "aeronave": _get_aircraft(reg_str),
                "next_day": _is_next_day_dep(std_str),
            })

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
                    can_pair = 0 <= mins <= MAX_TURNAROUND_JZ

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
