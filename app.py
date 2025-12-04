from __future__ import annotations

import os
import time
import tempfile
import logging
import unicodedata
import threading
import json
from datetime import datetime
from typing import Any, Dict, List, Tuple, Optional

from flask import (
    Flask, render_template, request, jsonify, send_file,
    after_this_request, Response
)

# ---- Google Sheets / Data ----
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

import pandas as pd

# ---- ReportLab (PDF) ----
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm, inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas as rl_canvas


# ============================================================================
# Configuración
# ============================================================================

class Config:
    # IDs / rutas
    SHEET_ID: str = os.getenv("SHEET_ID", "1GrPYixg14z76tea7PPvb58BCTRsN96wjikitCDal2OA")
    GOOGLE_CREDENTIALS_FILE: str = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")

    # Caché (segundos)
    CACHE_TTL_SECONDS: int = int(os.getenv("CACHE_TTL", "60"))

    # Límites y JSON
    JSON_SORT_KEYS: bool = False
    MAX_CONTENT_LENGTH: int = 10 * 1024 * 1024  # 10 MB

    # Columnas esperadas en la hoja
    COLUMNAS: List[str] = [
        'Proyecto/Articulo',
        'Programa',
        'Estudiante 1',
        'Estudiante 2',
        'Asesor',
        'Evaluador 1',
        'Evaluador 2',
        'Evaluador 3',
        'Hora',
        'Propuesta',
        'Anteproyecto',
        'Trabajo final',
        'Fecha sustentación',
        'Convocatoria',
        'ARTICULO/MONOGRAFIA',
        'Año'
    ]

    PAYLOAD_TO_SHEET: Dict[str, str] = {
        'proyecto_articulo': 'Proyecto/Articulo',
        'programa': 'Programa',
        'estudiante1': 'Estudiante 1',
        'estudiante2': 'Estudiante 2',
        'asesor': 'Asesor',
        'evaluador1': 'Evaluador 1',
        'evaluador2': 'Evaluador 2',
        'evaluador3': 'Evaluador 3',
        'hora': 'Hora',
        'propuesta': 'Propuesta',
        'anteproyecto': 'Anteproyecto',
        'trabajo_final': 'Trabajo final',
        'fecha_sustentacion': 'Fecha sustentación',
        'convocatoria': 'Convocatoria',
        'articulo_monografia': 'ARTICULO/MONOGRAFIA',
        'ano': 'Año',
    }

    LOCAL_DATA_JSON: str = os.getenv("LOCAL_DATA_JSON", os.path.join("static", "data.json"))


# Logger
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("proyectos")

# Flask 
app = Flask(__name__, static_folder="static", template_folder="templates")
app.config.from_object(Config)


# ============================================================================
# Seguridad / errores
# ============================================================================

@app.after_request
def add_security_headers(resp: Response) -> Response:
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    resp.headers.setdefault("X-XSS-Protection", "1; mode=block")
    if resp.mimetype == "application/json":
        resp.headers.setdefault("Cache-Control", "no-store")
    return resp

@app.errorhandler(400)
def bad_request(e): return jsonify({"error": "Solicitud inválida", "detalle": str(e)}), 400

@app.errorhandler(404)
def not_found(e): return jsonify({"error": "Recurso no encontrado"}), 404

@app.errorhandler(413)
def too_large(e): return jsonify({"error": "Payload demasiado grande"}), 413

@app.errorhandler(Exception)
def internal_error(e):
    logger.exception("Error no controlado")
    return jsonify({"error": "Error interno del servidor"}), 500



# ============================================================================
# Utilidades
# ============================================================================

def normalizar_texto(texto: Any) -> str:
    if texto is None: return ""
    s = str(texto).lower().strip()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    return s

def payload_to_row(encabezados: List[str], payload: Dict[str, Any]) -> List[str]:
    fila: List[str] = []
    for h in encabezados:
        valor = ""
        for k_payload, h_sheet in app.config["PAYLOAD_TO_SHEET"].items():
            if h_sheet == h.strip():
                valor = str(payload.get(k_payload, "")).strip()
                break
        fila.append(valor)
    return fila


# ============================================================================
# Google Sheets 
# ============================================================================

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_gs_client = None
_gs_client_lock = threading.Lock()

_cache_lock = threading.Lock()
_cache_data = {"ts": 0.0, "headers": [], "rows": [], "worksheet_title": ""}

def get_gspread_client():
    global _gs_client
    with _gs_client_lock:
        if _gs_client is None:
            # 1) Primero intentamos leer credenciales desde variable de entorno (Render)
            creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
            if creds_json:
                try:
                    info = json.loads(creds_json)
                    creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
                    logger.info("gspread autorizado con GOOGLE_CREDENTIALS_JSON")
                except Exception as e:
                    logger.exception("Error leyendo GOOGLE_CREDENTIALS_JSON: %s", e)
                    raise
            else:
                # 2) Fallback: archivo local (desarrollo en tu PC)
                cred_path = app.config["GOOGLE_CREDENTIALS_FILE"]
                if not os.path.exists(cred_path):
                    raise FileNotFoundError(
                        f"No se encontró credencial en {cred_path}. "
                        "Configura GOOGLE_APPLICATION_CREDENTIALS o GOOGLE_CREDENTIALS_JSON."
                    )
                creds = Credentials.from_service_account_file(cred_path, scopes=_SCOPES)
                logger.info("gspread autorizado con archivo de credenciales local")

            _gs_client = gspread.authorize(creds)
        return _gs_client

def get_worksheet():
    client = get_gspread_client()
    sheet = client.open_by_key(app.config["SHEET_ID"])
    return sheet.get_worksheet(0)

def _load_sheet_values() -> Tuple[List[str], List[List[str]], str]:
    ws = get_worksheet()
    values = ws.get_all_values()
    headers = values[0] if values else []
    rows = values[1:] if values and len(values) > 1 else []
    return headers, rows, ws.title

def _load_local_values() -> Tuple[List[str], List[List[str]], str]:
    """Fallback para trabajar con static/data.json si falla Sheets."""
    path = app.config["LOCAL_DATA_JSON"]
    if not os.path.exists(path):
        return [], [], "LOCAL"
    try:
        df = pd.read_json(path)
       
        if "proyectos" in df.columns:
            proyectos = pd.read_json(path)["proyectos"].tolist()
            if isinstance(proyectos, list):
                df = pd.DataFrame(proyectos)
        # Reordenar columnas si existen las esperadas
        cols = [c for c in app.config["COLUMNAS"] if c in df.columns]
        if cols:
            df = df[cols]
        headers = list(df.columns)
        rows = df.astype(str).fillna("").values.tolist()
        return headers, rows, "LOCAL"
    except Exception as e:
        logger.warning("No se pudo cargar fallback local: %s", e)
        return [], [], "LOCAL"

def _get_cached_values(force: bool = False) -> Tuple[List[str], List[List[str]], str]:
    ttl = app.config["CACHE_TTL_SECONDS"]
    now = time.time()
    with _cache_lock:
        if not force and _cache_data["rows"] and (now - _cache_data["ts"] < ttl):
            return _cache_data["headers"], _cache_data["rows"], _cache_data["worksheet_title"]

    try:
        headers, rows, title = _load_sheet_values()
    except Exception as e:
        logger.warning("Falla Sheets, usando fallback local: %s", e)
        headers, rows, title = _load_local_values()

    with _cache_lock:
        _cache_data.update({"headers": headers, "rows": rows, "worksheet_title": title, "ts": time.time()})
    return headers, rows, title

def get_registros(force: bool = False) -> List[Dict[str, Any]]:
    headers, rows, title = _get_cached_values(force=force)
    registros: List[Dict[str, Any]] = []
    if not headers or not rows:
        return registros
    for idx, row in enumerate(rows, start=2):
        item = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
        item["hoja_origen"] = title
        item["numero_fila"] = idx
        registros.append(item)
    return registros


# ============================================================================
# Rutas
# ============================================================================

@app.route("/")
def index():
    return render_template("index.html")  

@app.route("/verificar-conexion", methods=["GET"])
def verificar_conexion():
    try:
        headers, rows, title = _get_cached_values(force=True)
        total = len(rows)
        return jsonify({
            "estado": "conectado",
            "mensaje": f'Conexión exitosa. {total} registros en hoja "{title}"',
            "total_registros": total
        })
    except Exception as e:
        logger.warning("Conexión parcial o error: %s", e)
        return jsonify({"estado": "error", "mensaje": "No se pudo conectar con Google Sheets"})

@app.route("/mostrar_todos", methods=["GET"])
def mostrar_todos():
    try:
        registros = get_registros(force=False)
        resp = jsonify({"resultados": registros})
        resp.headers["Cache-Control"] = "public, max-age=30"
        return resp
    except Exception as e:
        logger.exception("Error mostrando todos")
        return jsonify({"error": f"Error al obtener registros: {e}"}), 500

@app.route("/buscar", methods=["POST"])
def buscar():
    try:
        data = request.get_json(silent=True) or {}
        termino: str = (data.get("termino") or "").strip()
        registros = get_registros(force=False)
        if not termino:
            return jsonify({"resultados": registros})

        needle = normalizar_texto(termino)
        columnas = app.config["COLUMNAS"]
        resultados: List[Dict[str, Any]] = []
        for r in registros:
            campos = [normalizar_texto(r.get(c, "")) for c in columnas]
            if any(needle in c for c in campos if c):
                resultados.append(r)
        return jsonify({"resultados": resultados})
    except Exception as e:
        logger.exception("Error en busqueda")
        return jsonify({"error": f"Error en la busqueda: {e}"}), 500

@app.route("/agregar", methods=["POST"])
def agregar():
    try:
        payload = request.get_json(silent=True) or {}
        if not payload.get("proyecto_articulo"):
            return jsonify({"error": "El campo Proyecto/Articulo es obligatorio"}), 400
        if not payload.get("estudiante1"):
            return jsonify({"error": "El campo Estudiante 1 es obligatorio"}), 400

        # Solo intentamos escribir si hay Google Sheets
        ws = get_worksheet()
        headers = ws.row_values(1) or app.config["COLUMNAS"]
        nueva_fila = payload_to_row(headers, payload)
        ws.append_row(nueva_fila, value_input_option="USER_ENTERED")

        _get_cached_values(force=True)
        return jsonify({"mensaje": "Registro agregado exitosamente a Google Sheets"})
    except Exception as e:
        logger.exception("Error al agregar")
        return jsonify({"error": f"Error al agregar registro: {e}"}), 500

@app.route("/actualizar", methods=["POST"])
def actualizar():
    try:
        payload = request.get_json(silent=True) or {}
        if not payload.get("proyecto_articulo"):
            return jsonify({"error": "El campo Proyecto/Articulo es obligatorio"}), 400
        if not payload.get("estudiante1"):
            return jsonify({"error": "El campo Estudiante 1 es obligatorio"}), 400
        if not payload.get("numero_fila"):
            return jsonify({"error": "Número de fila no especificado"}), 400

        numero_fila = int(payload["numero_fila"])
        ws = get_worksheet()
        headers = ws.row_values(1) or app.config["COLUMNAS"]
        fila_actualizada = payload_to_row(headers, payload)
        start_a1 = rowcol_to_a1(numero_fila, 1)
        ws.update(start_a1, [fila_actualizada], value_input_option="USER_ENTERED")

        _get_cached_values(force=True)
        return jsonify({"mensaje": "Registro actualizado exitosamente en Google Sheets"})
    except Exception as e:
        logger.exception("Error al actualizar")
        return jsonify({"error": f"Error al actualizar registro: {e}"}), 500

# ----------------------------------------------------------------------------
# PDF - Solo columnas especificas
# ----------------------------------------------------------------------------

class NumberedCanvas(rl_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(num_pages)
            super().showPage()
        super().save()

    def _draw_page_number(self, page_count):
        self.setFont("Helvetica", 8)
        txt = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(self._pagesize[0] - 1.5*cm, 1.1*cm, txt)

def _styles():
    base = getSampleStyleSheet()
    
    # Estilo para título principal
    title = ParagraphStyle(
        "TitleXL", 
        parent=base["Heading1"], 
        fontName="Helvetica-Bold",
        fontSize=16, 
        textColor=colors.HexColor("#B71C1C"),
        alignment=1, 
        spaceAfter=6,
        spaceBefore=15
    )
    
    # Estilo para subtítulo
    subtitle = ParagraphStyle(
        "Subtitle", 
        parent=base["Normal"], 
        fontSize=9,
        textColor=colors.HexColor("#666666"), 
        alignment=1, 
        spaceAfter=12, 
        leading=10,
    )
    
    # Estilo para encabezados de tabla
    header = ParagraphStyle(
        "Header", 
        parent=base["Normal"], 
        fontName="Helvetica-Bold",
        fontSize=9, 
        textColor=colors.white,
        alignment=1,
        leading=10
    )
    
    # Estilo para celdas normales
    cell = ParagraphStyle(
        "Cell", 
        parent=base["Normal"], 
        fontName="Helvetica",
        fontSize=8, 
        leading=9.5, 
        wordWrap="CJK",
    )
    
    # Estilo para celdas con texto importante
    cell_bold = ParagraphStyle(
        "CellBold", 
        parent=cell, 
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A4B8C")
    )
    
    return title, subtitle, header, cell, cell_bold

def _p(text, style): 
    if text is None:
        text = ""
    text = str(text).strip()
    # Limitar texto muy largo para evitar desbordamiento
    if len(text) > 200:
        text = text[:197] + "..."
    return Paragraph(text, style)

def _row_from_record(r, cell, cell_bold):
    # Procesar evaluadores
    evaluadores = ", ".join([x for x in [
        r.get('Evaluador 1',''), 
        r.get('Evaluador 2',''), 
        r.get('Evaluador 3','')
    ] if x and str(x).strip()])
    
    # Obtener los campos específicos solicitados
    proyecto = r.get('Proyecto/Articulo', '') or ''
    programa = r.get('Programa', '') or ''
    estudiante1 = r.get('Estudiante 1', '') or ''
    estudiante2 = r.get('Estudiante 2', '') or ''
    articulo_monografia = r.get('ARTICULO/MONOGRAFIA', '') or ''
    
    return [
        _p(proyecto, cell_bold),
        _p(programa, cell),
        _p(estudiante1, cell),
        _p(estudiante2, cell),
        _p(evaluadores, cell),
        _p(articulo_monografia, cell),
    ]

def _header_logo(canvas, doc):
    width, height = doc.pagesize
    canvas.saveState()
    
    # Fondo de encabezado
    canvas.setFillColor(colors.HexColor("#B71C1C"))
    canvas.rect(0, height-2.0*cm, width, 2.0*cm, stroke=0, fill=1)
    
    # Logo 
    try:
        logo_path = "Logo_UNILIBRE.png"
        if os.path.exists(logo_path):
            from reportlab.lib.utils import ImageReader
            img = ImageReader(logo_path)
            iw, ih = img.getSize()
            max_h = 1.5*cm
            scale = min(2.5*cm/iw, max_h/ih)
            lw, lh = iw*scale, ih*scale
            canvas.drawImage(
                img, 1.0*cm, height - 1.8*cm, 
                width=lw, height=lh, mask='auto'
            )
    except Exception:
        pass
    
    # Texto institucional
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 12)
    canvas.drawString(4*cm, height - 1.4*cm, "UNIVERSIDAD LIBRE")
    canvas.setFont("Helvetica", 9)
    canvas.drawString(4*cm, height - 1.8*cm, "Sistema de Gestion de Proyectos Academicos")
    
    canvas.restoreState()

def _footer_info(canvas, doc):
    width, height = doc.pagesize
    canvas.saveState()
    
    # Información del footer
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#666666"))
    
    fecha_export = datetime.now().strftime("%d/%m/%Y %H:%M")
    canvas.drawString(1.5*cm, 1.0*cm, f"Generado: {fecha_export}")
    canvas.drawCentredString(width/2, 1.0*cm, "Confidencial - Uso interno")
    
    canvas.restoreState()

@app.route("/exportar_pdf", methods=["POST"])
def exportar_pdf():
    try:
        payload = request.get_json(silent=True) or {}
        datos = payload.get("datos", [])
        if not datos:
            return jsonify({"error": "No hay datos para exportar"}), 400

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_path = tmp.name
        tmp.close()

       
        page_size = landscape(A4)
        doc = SimpleDocTemplate(
            pdf_path, pagesize=page_size,
            leftMargin=1.0*cm, rightMargin=1.0*cm, 
            topMargin=2.5*cm, bottomMargin=1.5*cm,
            title="Listado de Proyectos Académicos - Universidad Libre",
            author="Sistema de Gestion de Proyectos",
        )

        title, subtitle, header, cell, cell_bold = _styles()
        elements: List[Any] = []

        # Títulos
        elements.append(Spacer(1, 5))
        elements.append(Paragraph("LISTADO DE PROYECTOS ACADEMICOS", title))
        
        fecha_export = datetime.now().strftime("%d/%m/%Y %H:%M")
        elements.append(Paragraph(
            f"Exportado el {fecha_export} • {len(datos)} registros encontrados", 
            subtitle
        ))
        elements.append(Spacer(1, 8))

        # Cabeceras de tabla - SOLO LAS COLUMNAS SOLICITADAS
        headers = [
            Paragraph("Proyecto/Artículo", header),
            Paragraph("Programa", header),
            Paragraph("Estudiante 1", header),
            Paragraph("Estudiante 2", header),
            Paragraph("Evaluadores", header),
            Paragraph("Artículo/Monografía", header),
        ]

        # Preparar datos de la tabla
        table_data = [headers]
        for r in datos:
            try:
                table_data.append(_row_from_record(r, cell, cell_bold))
            except Exception as e:
                logger.warning("Error procesando registro para PDF: %s", e)
                continue

        # Anchos de columnas optimizados para las 6 columnas solicitadas
        col_widths = [
            4.5*cm,   # Proyecto/Artículo
            2.5*cm,   # Programa
            3.0*cm,   # Estudiante 1
            3.0*cm,   # Estudiante 2  
            4.0*cm,   # Evaluadores
            3.0*cm,   # Artículo/Monografía
        ]
        
        # Calcular ancho total
        total_width = sum(col_widths)
        available_width = page_size[0] - 2.0*cm
        
        # Ajustar anchos si es necesario
        if total_width > available_width:
            scale_factor = available_width / total_width
            col_widths = [w * scale_factor for w in col_widths]
        
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Estilo de tabla optimizado
        table.setStyle(TableStyle([
            # Encabezados
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#B71C1C")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',  (0,0), (-1,0), 9),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('TOPPADDING', (0,0), (-1,0), 8),

            # Filas alternas
            ('ROWBACKGROUNDS', (0,1), (-1,-1), 
             [colors.HexColor("#F8F9FA"), colors.white]),
            
            # Bordes y alineación
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,1), (-1,-1), 9.5),
            ('VALIGN', (0,1), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('RIGHTPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,1), (-1,-1), 4),
            ('BOTTOMPADDING', (0,1), (-1,-1), 4),
            
            # Alineación específica
            ('ALIGN', (1,1), (1,-1), 'CENTER'),  # Programa al centro
            ('ALIGN', (5,1), (5,-1), 'CENTER'),  # Artículo/Monografía al centro
        ]))

        elements.append(table)
        
        # Resumen al final
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(
            f"<b>Resumen:</b> Se exportaron {len(datos)} proyectos académicos con información básica.", 
            ParagraphStyle(
                'Summary', 
                parent=cell, 
                fontSize=8, 
                textColor=colors.HexColor("#666666"),
                alignment=1
            )
        ))

        def _on_each_page(canvas, doc_):
            _header_logo(canvas, doc_)
            _footer_info(canvas, doc_)

        doc.build(
            elements, 
            onFirstPage=_on_each_page, 
            onLaterPages=_on_each_page, 
            canvasmaker=NumberedCanvas
        )

        @after_this_request
        def cleanup(response):
            try: 
                os.remove(pdf_path)
            except Exception: 
                pass
            return response

        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=f'proyectos_academicos_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        logger.exception("Error exportando PDF")
        return jsonify({"error": f"Error al exportar PDF: {e}"}), 500

# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
@app.route("/exportar_excel", methods=["GET"])
def exportar_excel():
    try:
        headers, rows, title = _get_cached_values(force=False)
        if not rows:
            return jsonify({"error": "No hay datos para exportar"}), 400

        columnas = headers + ["Hoja Origen"]
        datos_tabla: List[List[str]] = []
        for row in rows:
            full_row = row + [""] * max(0, len(headers) - len(row))
            datos_tabla.append(full_row + [title])

        df = pd.DataFrame(datos_tabla, columns=columnas)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        excel_path = tmp.name
        tmp.close()

        @after_this_request
        def cleanup(response):
            try: os.remove(excel_path)
            except Exception: pass
            return response

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Proyectos_Academicos", index=False)
            ws = writer.sheets["Proyectos_Academicos"]
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    try: max_len = max(max_len, len(str(cell.value)))
                    except Exception: pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

        return send_file(
            excel_path,
            as_attachment=True,
            download_name=f'base_datos_proyectos_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.exception("Error exportando Excel")
        return jsonify({"error": f"Error exportando Excel: {e}"}), 500

# ----------------------------------------------------------------------------
# Estadísticas Mejoradas
# ----------------------------------------------------------------------------
@app.route("/estadisticas-detalladas", methods=["GET"])
def obtener_estadisticas_detalladas():
    try:
        registros = get_registros(force=False)
        total = len(registros)

        def _limpio(val: Any) -> str:
            v = str(val or "").replace("\xa0", " ").strip()
            if not v:
                return ""
            return unicodedata.normalize("NFKD", v).encode("ASCII", "ignore").decode("ASCII").lower()

        # Conteo directo por celda no vacía en las columnas observadas
        total_propuestas = total_anteproyectos = total_trabajos_finales = 0
        propuestas_aprobadas = trabajos_finales_aprobados = 0

        for r in registros:
            propuesta_val = _limpio(
                r.get("Propuesta", "")
                or r.get("Propuesta ", "")
                or r.get("propuesta", "")
            )
            anteproyecto_val = _limpio(
                r.get("Anteproyecto ", "")
                or r.get("Anteproyecto", "")
                or r.get("anteproyecto", "")
            )
            trabajo_final_val = _limpio(
                r.get("Trabajo final ", "")
                or r.get("Trabajo final", "")
                or r.get("Trabajo Final", "")
                or r.get("trabajo_final", "")
            )

            if propuesta_val:
                total_propuestas += 1
                if any(term in propuesta_val for term in ["aprobado", "aprobada", "approved", "si", "si ", "yes"]):
                    propuestas_aprobadas += 1

            if anteproyecto_val:
                total_anteproyectos += 1

            if trabajo_final_val:
                total_trabajos_finales += 1
                if any(term in trabajo_final_val for term in ["aprobado", "aprobada", "approved", "si", "si ", "yes"]):
                    trabajos_finales_aprobados += 1

        # Estados usando las mismas columnas con espacios finales
        def _contar_estado(candidatos) -> Dict[str, int]:
            aprobados = revision = no_aprobados = no_especificado = 0
            for r in registros:
                bruto = ""
                for c in candidatos:
                    val = r.get(c, "")
                    if val and str(val).strip():
                        bruto = val
                        break
                valor = _limpio(bruto)
                if not valor:
                    no_especificado += 1
                    continue
                if any(term in valor for term in ["aprobado", "aprobada", "approved", "si", "si ", "yes"]):
                    aprobados += 1
                elif any(term in valor for term in ["revision", "revisando", "pendiente", "en proceso"]):
                    revision += 1
                elif any(term in valor for term in ["no aprobado", "rechazado", "rejected", "no"]):
                    no_aprobados += 1
                else:
                    aprobados += 1
            return {
                "aprobados": aprobados,
                "revision": revision,
                "no_aprobados": no_aprobados,
                "no_especificado": no_especificado,
            }

        propuestas_stats = _contar_estado(["Propuesta", "Propuesta ", "propuesta"])
        anteproyecto_stats = _contar_estado(["Anteproyecto ", "Anteproyecto", "anteproyecto"])
        trabajo_final_stats = _contar_estado(["Trabajo final ", "Trabajo final", "Trabajo Final", "trabajo_final"])

        programas = {}
        for r in registros:
            programa = (r.get("Programa") or "No especificado").strip() or "No especificado"
            programas[programa] = programas.get(programa, 0) + 1

        asesores = {}
        for r in registros:
            asesor = (r.get("Asesor") or "").strip()
            if asesor and asesor.lower() not in ["no especificado", "sin especificar", "none", ""]:
                asesores[asesor] = asesores.get(asesor, 0) + 1
        top_asesores = dict(sorted(asesores.items(), key=lambda x: x[1], reverse=True)[:10])

        fechas_sustentacion = {}
        for r in registros:
            fecha = str(r.get("Fecha sustentación", "") or r.get("Fecha sustentaci\u00f3n", "")).strip()
            if fecha and fecha.lower() not in ["no especificado", "none", ""]:
                fechas_sustentacion[fecha] = fechas_sustentacion.get(fecha, 0) + 1
        fechas_ordenadas = dict(sorted(fechas_sustentacion.items(), key=lambda x: x[0])[-15:])

        anos = {}
        for r in registros:
            ano = str(r.get("A\u00f1o", "") or r.get("Año", "") or r.get("Ano", "")).strip()
            if ano and ano.isdigit():
                anos[ano] = anos.get(ano, 0) + 1

        data = {
            "totales": {
                "total_proyectos": total,
                "total_propuestas": total_propuestas,
                "total_anteproyectos": total_anteproyectos,
                "total_trabajos_finales": total_trabajos_finales,
            },
            "estados_contadores": {
                "propuestas_aprobadas": propuestas_aprobadas,
                "trabajos_finales_aprobados": trabajos_finales_aprobados,
            },
            "por_programa": programas,
            "por_asesor": top_asesores,
            "estados": {
                "propuestas": propuestas_stats,
                "anteproyectos": anteproyecto_stats,
                "trabajos_finales": trabajo_final_stats,
            },
            "por_fecha": fechas_ordenadas,
            "por_ano": anos,
            "ultima_actualizacion": datetime.now().isoformat(),
        }

        resp = jsonify(data)
        resp.headers["Cache-Control"] = "no-cache"
        return resp

    except Exception as e:
        logger.exception("Error calculando estadisticas detalladas")
        return jsonify({"error": f"Error al obtener estad\u00edsticas: {e}"}), 500

# ----------------------------------------------------------------------------
# control 
# ----------------------------------------------------------------------------
@app.route("/health", methods=["GET"])
def health_check():
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "sheet_id": app.config["SHEET_ID"],
        "cache_ttl": app.config["CACHE_TTL_SECONDS"]
    })


# ============================================================================
# Bootstrap
# ============================================================================
if __name__ == "__main__":
    logger.info("Iniciando servidor Flask…")
    logger.info("http://127.0.0.1:5000")
    app.run(host="0.0.0.0", port=5000, debug=False)
