"""Endpoints de FastAPI para TALMA: Previsualización de rotaciones y descarga de Excel."""

from fastapi import FastAPI, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

try:
    from api.talma_processor import procesar_vuelos_stream, procesar_texto_directo
except ImportError:
    from talma_processor import procesar_vuelos_stream, procesar_texto_directo

app = FastAPI(title="TALMA Operations Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health_check():
    """Verifica que la API esté activa."""
    return {"status": "ok"}


@app.post("/api/procesar-preview")
async def procesar_preview(request: Request):
    """Recibe el archivo (SKY) o texto (JetSMART) y devuelve fecha_detectada y rotaciones."""
    try:
        content_type = request.headers.get("content-type", "")

        if "multipart/form-data" in content_type:
            form = await request.form()
            file = form.get("file")
            airline_code = str(form.get("airline") or form.get(
                "airline_code") or "H8").strip().upper()
            base_iata = str(form.get("base") or form.get(
                "base_iata") or "LIM").strip().upper()

            if not file:
                return JSONResponse(
                    status_code=400,
                    content={"detail": "No se recibió ningún archivo."},
                )

            file_bytes = await file.read()
            filename = getattr(file, "filename", "") or ""
            op_date, operaciones, _ = procesar_vuelos_stream(
                file_bytes, filename, airline_code=airline_code, base_iata=base_iata
            )
        else:
            body = await request.json()
            raw_text = str(body.get("texto") or body.get(
                "raw_text") or body.get("text") or "")
            airline_code = str(body.get("airline") or body.get(
                "airline_code") or "JZ").strip().upper()
            base_iata = str(body.get("base") or body.get(
                "base_iata") or "LIM").strip().upper()

            if not raw_text.strip():
                return JSONResponse(
                    status_code=400,
                    content={"detail": "El texto del itinerario está vacío."},
                )

            op_date, operaciones, _ = procesar_texto_directo(
                raw_text, airline_code=airline_code, base_iata=base_iata
            )

        return {
            "fecha_detectada": op_date,
            "rotaciones": operaciones,
            "total": len(operaciones),
        }
    except Exception as exc:  # pylint: disable=broad-exception-caught
        return JSONResponse(status_code=500, content={"detail": str(exc)})


@app.post("/api/convertir-talma")
@app.post("/api/procesar")
@app.post("/api/descargar-excel")
async def convertir_talma(request: Request):
    """Genera y descarga el archivo Excel oficial TALMA con la fecha extraída del archivo."""
    try:
        content_type = request.headers.get("content-type", "")

        if "multipart/form-data" in content_type:
            form = await request.form()
            file = form.get("file")
            airline_code = str(form.get("airline") or form.get(
                "airline_code") or "H8").strip().upper()
            base_iata = str(form.get("base") or form.get(
                "base_iata") or "LIM").strip().upper()

            file_bytes = await file.read()
            filename = getattr(file, "filename", "") or ""
            op_date, _, excel_bytes = procesar_vuelos_stream(
                file_bytes, filename, airline_code=airline_code, base_iata=base_iata
            )
        else:
            body = await request.json()
            raw_text = str(body.get("texto") or body.get(
                "raw_text") or body.get("text") or "")
            airline_code = str(body.get("airline") or body.get(
                "airline_code") or "JZ").strip().upper()
            base_iata = str(body.get("base") or body.get(
                "base_iata") or "LIM").strip().upper()

            op_date, _, excel_bytes = procesar_texto_directo(
                raw_text, airline_code=airline_code, base_iata=base_iata
            )

        safe_date = op_date.replace("/", "-")
        out_name = f"ITINERARIOS_{airline_code}_{base_iata}_{safe_date}.xlsx"

        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
    except Exception as exc:  # pylint: disable=broad-exception-caught
        return JSONResponse(status_code=500, content={"detail": str(exc)})
