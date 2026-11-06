import json
import os
import re
from datetime import datetime

import requests
from docx import Document as DocxDocument
from flask import Flask, jsonify, request, send_file, send_from_directory
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader

app = Flask(__name__, static_folder="static", static_url_path="")

# --- Configuración ---
OLLAMA_URL = "http://localhost:11434/api/chat"
MODEL_NAME = "phi4-mini"  # cambia esto si quieres usar otro modelo instalado en Ollama

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

SYSTEM_PROMPT_TEMPLATE = """Eres un asistente evaluador académico. Vas a calificar la tarea de UN estudiante siguiendo ESTRICTAMENTE la siguiente rúbrica. Si se te da una solución de referencia, úsala para verificar si las respuestas del estudiante son correctas.

RÚBRICA DE EVALUACIÓN:
{rubric}

SOLUCIÓN DE REFERENCIA:
{solution}

La nota final debe estar en una escala de 1 a 10, donde 10 es la nota más alta y 1 la más baja.

Debes identificar el nombre del estudiante a partir del contenido del archivo (encabezados, firmas, etc.) y, si no aparece en el contenido, a partir del nombre del archivo que se te indica.

INSTRUCCIONES DE SALIDA (MUY IMPORTANTE):
Responde ÚNICAMENTE con un objeto JSON válido, sin texto antes ni después, sin bloques de código markdown, con exactamente esta forma:

{{
  "nombre_completo": "<nombre completo del estudiante tal como lo identificaste; si no lo encuentras usa 'Desconocido'>",
  "apellido": "<apellido(s) del estudiante, usado para ordenar alfabéticamente; si no lo encuentras usa 'Desconocido'>",
  "nota": <número entre 1 y 10, puede tener un decimal>,
  "justificacion": "<justificación breve y específica de la nota, citando partes del texto del estudiante y explicando qué criterios de la rúbrica cumplió o no>"
}}
"""


def extract_text_from_file(file_storage):
    """Extrae texto plano de un archivo subido (.txt, .docx o .pdf)."""
    filename = file_storage.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    file_storage.stream.seek(0)

    if ext == ".txt":
        raw = file_storage.read()
        return raw.decode("utf-8", errors="ignore")

    if ext == ".docx":
        doc = DocxDocument(file_storage.stream)
        return "\n".join(p.text for p in doc.paragraphs)

    if ext == ".pdf":
        reader = PdfReader(file_storage.stream)
        parts = []
        for page in reader.pages:
            parts.append(page.extract_text() or "")
        return "\n".join(parts)

    raise ValueError(f"Formato no soportado: '{ext}'. Usa .txt, .docx o .pdf")


def call_ollama(system_prompt, user_content):
    resp = requests.post(
        OLLAMA_URL,
        json={
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
            "stream": False,
            "format": "json",
        },
        timeout=300,
    )
    resp.raise_for_status()
    return resp.json()["message"]["content"]


def parse_model_json(raw_text, fallback_name):
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw_text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                pass
    return {
        "nombre_completo": fallback_name,
        "apellido": fallback_name,
        "nota": None,
        "justificacion": (
            "No se pudo interpretar la respuesta del modelo como JSON. "
            f"Respuesta cruda (primeros 300 caracteres): {raw_text[:300]}"
        ),
    }


def build_excel(results, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Calificaciones"

    headers = ["Apellido", "Nombre completo", "Nota (1-10)", "Justificación", "Archivo"]
    ws.append(headers)

    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    for r in results:
        ws.append(
            [
                r.get("apellido", ""),
                r.get("nombre_completo", ""),
                r.get("nota", ""),
                r.get("justificacion", ""),
                r.get("archivo", ""),
            ]
        )

    widths = [20, 28, 12, 80, 28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    wb.save(path)


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/process", methods=["POST"])
def process():
    rubric_file = request.files.get("rubric")
    solution_file = request.files.get("solution")
    student_files = request.files.getlist("students")

    if not rubric_file or not rubric_file.filename:
        return jsonify({"error": "Debes subir el archivo de la rúbrica."}), 400
    if not student_files:
        return jsonify({"error": "Debes subir al menos una tarea de estudiante."}), 400

    try:
        rubric_text = extract_text_from_file(rubric_file)
    except Exception as e:
        return jsonify({"error": f"Error leyendo la rúbrica: {e}"}), 400

    solution_text = "(No se proporcionó solución de referencia)"
    if solution_file and solution_file.filename:
        try:
            solution_text = extract_text_from_file(solution_file)
        except Exception as e:
            return jsonify({"error": f"Error leyendo la solución: {e}"}), 400

    system_prompt = SYSTEM_PROMPT_TEMPLATE.format(rubric=rubric_text, solution=solution_text)

    results = []
    for f in student_files:
        filename = f.filename or "archivo_sin_nombre"

        try:
            student_text = extract_text_from_file(f)
        except Exception as e:
            results.append(
                {
                    "nombre_completo": filename,
                    "apellido": filename,
                    "nota": None,
                    "justificacion": f"Error leyendo el archivo: {e}",
                    "archivo": filename,
                }
            )
            continue

        if not student_text.strip():
            results.append(
                {
                    "nombre_completo": filename,
                    "apellido": filename,
                    "nota": None,
                    "justificacion": (
                        "No se pudo extraer texto del archivo "
                        "(puede ser un PDF escaneado como imagen)."
                    ),
                    "archivo": filename,
                }
            )
            continue

        user_content = (
            f"Nombre del archivo: {filename}\n\n"
            f"Contenido de la tarea del estudiante:\n{student_text}"
        )

        try:
            raw = call_ollama(system_prompt, user_content)
        except requests.exceptions.ConnectionError:
            return jsonify(
                {"error": "No se pudo conectar con Ollama en localhost:11434. ¿Está corriendo?"}
            ), 503
        except requests.exceptions.ReadTimeout:
            results.append(
                {
                    "nombre_completo": filename,
                    "apellido": filename,
                    "nota": None,
                    "justificacion": "El modelo tardó demasiado en responder (timeout) para este archivo.",
                    "archivo": filename,
                }
            )
            continue
        except Exception as e:
            results.append(
                {
                    "nombre_completo": filename,
                    "apellido": filename,
                    "nota": None,
                    "justificacion": f"Error al consultar el modelo: {e}",
                    "archivo": filename,
                }
            )
            continue

        parsed = parse_model_json(raw, fallback_name=filename)
        parsed["archivo"] = filename
        results.append(parsed)

    results.sort(key=lambda r: (r.get("apellido") or "").strip().lower())

    excel_filename = f"calificaciones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_path = os.path.join(OUTPUT_DIR, excel_filename)
    build_excel(results, excel_path)

    return jsonify({"results": results, "excel_url": f"/api/download/{excel_filename}"})


@app.route("/api/download/<path:filename>")
def download(filename):
    return send_file(os.path.join(OUTPUT_DIR, filename), as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
