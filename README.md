# Validador de Tareas — Prototipo

Herramienta web local que califica tareas de estudiantes (texto, Word o PDF)
siguiendo una rúbrica y, opcionalmente, una solución de referencia. Usa un
modelo de lenguaje corriendo localmente en Ollama (nada sale de tu equipo) y
genera un único Excel con las notas ordenadas alfabéticamente por apellido.

## Requisitos

- Windows con [Ollama](https://ollama.ai) instalado y el modelo descargado:
  ```
  ollama pull phi4-mini
  ```
- Python 3.9+ instalado (verifica con `python --version`).

## Instalación

Desde la carpeta `validador-tareas`:

```
pip install -r requirements.txt
```

## Ejecución

1. Asegúrate de que Ollama esté corriendo (se inicia solo al abrir la app de
   Ollama; si no, ejecuta `ollama serve` en otra terminal).
2. Levanta el backend:
   ```
   python app.py
   ```
3. Abre en el navegador: http://localhost:5000

## Cómo se usa

1. Sube el archivo de la **rúbrica** (.txt, .docx o .pdf).
2. (Opcional) Sube un archivo con la **solución de referencia** de la tarea.
3. Sube **uno o varios archivos** con las tareas de los estudiantes. El
   sistema intenta identificar el nombre de cada estudiante a partir del
   contenido del archivo (y usa el nombre del archivo como respaldo).
4. Presiona **Calificar y generar Excel**. Procesa los archivos uno por uno
   (puede tardar varios minutos si son muchos), y al terminar:
   - Verás una tabla de resultados en pantalla.
   - Podrás descargar un Excel (`calificaciones_FECHA.xlsx`) con las columnas
     Apellido, Nombre completo, Nota (1-10), Justificación y Archivo,
     ordenado alfabéticamente por apellido.
5. Los Excel generados también quedan guardados en la carpeta
   `validador-tareas/outputs/`.

## Notas importantes del prototipo

- La identificación del estudiante depende de que el nombre aparezca en el
  documento (encabezado, firma, etc.) o en el nombre del archivo. Si el
  modelo no logra identificarlo, la fila mostrará "Desconocido" — revisa esos
  casos manualmente.
- Los PDFs escaneados como imagen (sin texto seleccionable) no se pueden leer
  con este prototipo; mostrará un aviso en la columna de justificación.
- Cada tarea se evalúa en una llamada independiente al modelo (no hay
  "memoria" entre estudiantes), así que la rúbrica se aplica de forma
  consistente a todos.

## Cómo cambiar el modelo

En `app.py`, cambia la constante `MODEL_NAME` por el nombre de cualquier otro
modelo que tengas instalado en Ollama (verifica con `ollama list`), por
ejemplo:

```python
MODEL_NAME = "deepseek-coder-v2:16b"
```

## Próximos pasos sugeridos (fuera del alcance de este prototipo)

- Procesar los estudiantes en paralelo para acelerar lotes grandes.
- Permitir editar manualmente una nota/justificación antes de exportar.
- Guardar rúbricas y soluciones predefinidas por curso/asignación.
- Mostrar una barra de progreso real (estudiante X de Y) vía streaming o
  WebSockets en vez de esperar a que termine todo el lote.
- Detectar y marcar automáticamente estudiantes "Desconocido" para revisión
  manual antes de exportar el Excel.
