"""Enrutador principal de itinerarios TALMA por aerolínea con empaquetado Excel nativo."""

import io
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import openpyxl

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

try:
    from api.jetsmart_talma import procesar_jetsmart_talma
    from api.sky_talma import procesar_sky_talma
except ImportError:
    from jetsmart_talma import procesar_jetsmart_talma
    from sky_talma import procesar_sky_talma


def _convertir_a_excel_talma_nativo(openpyxl_bytes: bytes) -> bytes:
    """Convierte inlineStr de openpyxl a sharedStrings.xml (t='s') compatible con TALMA."""
    wb = openpyxl.load_workbook(io.BytesIO(openpyxl_bytes))
    if "Hoja1" not in wb.sheetnames:
        wb.create_sheet("Hoja1")
    wb.active = wb["ITINERARIOS"]

    buf_temp = io.BytesIO()
    wb.save(buf_temp)
    buf_temp.seek(0)

    shared_strings = []
    ss_index = {}

    def get_ss_idx(text: str) -> int:
        if text not in ss_index:
            ss_index[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_index[text]

    with zipfile.ZipFile(buf_temp, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", ns)
    sheet_root = ET.fromstring(files["xl/worksheets/sheet1.xml"])

    for cell in sheet_root.findall(f".//{{{ns}}}c"):
        t_attr = cell.get("t")
        if t_attr == "inlineStr":
            is_el = cell.find(f"{{{ns}}}is")
            t_el = is_el.find(f"{{{ns}}}t") if is_el is not None else None
            txt = t_el.text if (
                t_el is not None and t_el.text is not None) else ""
            idx = get_ss_idx(txt)
            cell.set("t", "s")
            for child in list(cell):
                cell.remove(child)
            v_el = ET.SubElement(cell, f"{{{ns}}}v")
            v_el.text = str(idx)
        elif t_attr == "n" and len(cell) == 0:
            del cell.attrib["t"]

    files["xl/worksheets/sheet1.xml"] = ET.tostring(
        sheet_root, encoding="utf-8", xml_declaration=True
    )

    sst = ET.Element(
        f"{{{ns}}}sst",
        count=str(len(shared_strings)),
        uniqueCount=str(len(shared_strings)),
    )
    for text_val in shared_strings:
        si = ET.SubElement(sst, f"{{{ns}}}si")
        t_node = ET.SubElement(si, f"{{{ns}}}t")
        if text_val.startswith(" ") or text_val.endswith(" "):
            t_node.set(
                "{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_node.text = text_val
    files["xl/sharedStrings.xml"] = ET.tostring(
        sst, encoding="utf-8", xml_declaration=True
    )

    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", ct_ns)
    ct_root = ET.fromstring(files["[Content_Types].xml"])
    if not any(
        el.get("PartName") == "/xl/sharedStrings.xml"
        for el in ct_root.findall(f"{{{ct_ns}}}Override")
    ):
        ET.SubElement(
            ct_root,
            f"{{{ct_ns}}}Override",
            PartName="/xl/sharedStrings.xml",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )
    files["[Content_Types].xml"] = ET.tostring(
        ct_root, encoding="utf-8", xml_declaration=True
    )

    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", rel_ns)
    rels_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    ss_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    if not any(
        el.get("Type") == ss_rel_type
        for el in rels_root.findall(f"{{{rel_ns}}}Relationship")
    ):
        existing_ids = {
            el.get("Id") for el in rels_root.findall(f"{{{rel_ns}}}Relationship")
        }
        idx_rel = len(existing_ids) + 1
        while f"rId{idx_rel}" in existing_ids:
            idx_rel += 1
        ET.SubElement(
            rels_root,
            f"{{{rel_ns}}}Relationship",
            Id=f"rId{idx_rel}",
            Type=ss_rel_type,
            Target="sharedStrings.xml",
        )
    files["xl/_rels/workbook.xml.rels"] = ET.tostring(
        rels_root, encoding="utf-8", xml_declaration=True
    )

    out_mem = io.BytesIO()
    with zipfile.ZipFile(out_mem, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    out_mem.seek(0)
    return out_mem.getvalue()


def procesar_vuelos_stream(file_bytes: bytes, filename: str, airline_code="H8", base_iata="LIM"):
    """Enruta el procesamiento de archivos (Excel/PDF) hacia sky_talma."""
    op_date, operaciones, raw_excel = procesar_sky_talma(
        file_bytes, filename, airline_code=airline_code, base_iata=base_iata
    )
    return op_date, operaciones, _convertir_a_excel_talma_nativo(raw_excel)


def procesar_texto_directo(raw_text: str, airline_code="JZ", base_iata="LIM"):
    """Enruta el procesamiento de texto/HTML pegado hacia jetsmart_talma."""
    op_date, operaciones, raw_excel = procesar_jetsmart_talma(
        raw_text, airline_code=airline_code, base_iata=base_iata
    )
    return op_date, operaciones, _convertir_a_excel_talma_nativo(raw_excel)
