# Validador de Tareas — Prototipo

Herramienta web local que califica tareas de estudiantes (texto, Word, PDF,
PDF escaneado o fotos, incluida letra a mano) siguiendo una rúbrica y, opcionalmente, una solución de referencia. Usa un
modelo de lenguaje corriendo localmente en Ollama (nada sale de tu equipo) y
genera un único Excel con las notas ordenadas alfabéticamente por apellido.

## Requisitos

- Windows con [Ollama](https://ollama.ai) 0.12.7 o superior instalado y los
  modelos descargados:
  ```
  ollama pull phi4-mini
  ollama pull qwen3-vl:8b
  ```
  `phi4-mini` califica. `qwen3-vl:8b` es un modelo de visión que lee los PDF
  escaneados y las fotos; pesa 6.2 GB y está pensado para una GPU de 12 GB
  (por ejemplo una RTX 4070 Ti). Sin él la app funciona igual, pero los
  escaneos quedan marcados sin calificar.
- [uv](https://docs.astral.sh/uv/) instalado. No necesitas instalar Python
  aparte: uv descarga la versión del proyecto (3.12, fijada en `.python-version`).
  En Windows, desde PowerShell:
  ```
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```
  Cierra y vuelve a abrir la terminal, y verifica con `uv --version`.

## Instalación

Desde la carpeta `validador-tareas`:

```
uv sync
```

Esto crea el entorno virtual `.venv` dentro del proyecto, instala las
dependencias de `pyproject.toml` y genera `uv.lock` con las versiones exactas.
Sube `uv.lock` y `.python-version` a git para que otra persona obtenga el mismo
entorno con `uv sync`; `.venv` no se sube (ya está en `.gitignore`).

Para agregar una librería más adelante: `uv add nombre-de-la-libreria`.

## Ejecución

1. Asegúrate de que Ollama esté corriendo (se inicia solo al abrir la app de
   Ollama; si no, ejecuta `ollama serve` en otra terminal).
2. Levanta el backend:
   ```
   uv run app.py
   ```
3. Abre en el navegador: http://localhost:5000

## Cómo se usa

1. Sube el archivo de la **rúbrica** (.docx, .pdf, .txt o foto .jpg/.png).
2. (Opcional) Sube un archivo con la **solución de referencia** de la tarea.
3. Sube **uno o varios archivos** con las tareas de los estudiantes. El
   sistema intenta identificar el nombre de cada estudiante a partir del
   contenido del archivo (y usa el nombre del archivo como respaldo).
4. Elige la forma de calificar (ver abajo) y presiona **Calificar**. Verás el
   avance real: primero se leen los escaneos y fotos ("Leyendo escaneo 3 de 8"),
   luego se divide la rúbrica en criterios y se califica cada tarea
   ("Calificando tarea 5 de 30"). El proceso corre en el servidor: si recargas
   la página, el avance se recupera. Al terminar:
   - Verás una tabla de resultados en pantalla.
   - Podrás descargar un Excel (`calificaciones_FECHA.xlsx`) con las columnas
     Apellido, Nombre completo, Nota (1-10), un puntaje por criterio (si
     calificaste por criterios), Justificación, Revisar y Archivo,
     ordenado alfabéticamente por apellido (sin importar tildes, la ñ va después
     de la n y los estudiantes sin nombre quedan al final).
   - Las filas que necesitan tu atención salen en **amarillo**, con el motivo en
     la columna **Revisar** (nombre no encontrado o tomado del nombre del
     archivo, nota que no se pudo determinar, tarea leída desde imagen, posible
     intento de manipular la nota, documento demasiado largo, etc.).
   - Si hubo escaneos o fotos, el Excel trae una segunda hoja,
     **Transcripciones OCR**, con el texto que se leyó de cada imagen (también
     se ve en pantalla con "Ver texto leído de la imagen").
5. Los Excel generados también quedan guardados en la carpeta
   `validador-tareas/outputs/`.

## Formas de calificar

Antes de calificar eliges una de tres formas (selector "Forma de calificar"):

| Forma | Cómo funciona |
| --- | --- |
| **Por criterios** (recomendada) | Un *orquestador* divide la rúbrica en criterios una sola vez por grupo. Luego un *evaluador* califica cada criterio por separado, con una cita textual de la tarea, y el *verificador* (código, no IA) comprueba que la cita exista, que el puntaje no pase del máximo y que la suma cuadre. Si algo no cuadra, se le pide al evaluador que lo corrija una vez. |
| Por criterios, sin verificador | Igual, sin la verificación. Sirve para medir cuánto aporta el verificador. |
| Un solo evaluador | Un modelo califica toda la rúbrica de una vez (la versión anterior). Más rápido. |

Con la forma por criterios:

- La nota es la suma de los criterios sobre el total de la rúbrica, llevada a 1-10.
- El Excel trae una columna por criterio y una hoja **Rúbrica interpretada** con
  cómo el orquestador dividió la rúbrica; revísala la primera vez que uses una
  rúbrica nueva. En pantalla se ve con "Ver cómo se dividió la rúbrica".
- Si el orquestador no logra dividir la rúbrica, el grupo se califica con un solo
  evaluador y se avisa arriba de los resultados.
- Cuando el nombre del estudiante no está etiquetado en el documento, un agente
  pequeño (*identificador*) lo busca.
- Hace más llamadas al modelo que la forma de un solo evaluador (una por criterio,
  más los reintentos), así que tarda más. Para que Ollama atienda dos criterios a
  la vez, define la variable de entorno `OLLAMA_NUM_PARALLEL=2` antes de abrir
  Ollama.

En `app.py` puedes cambiar la forma por defecto (`MODO_POR_DEFECTO`), cuántos
criterios se califican a la vez (`EVALUADORES_SIMULTANEOS`) y el modelo de cada
rol (`MODELOS`), por ejemplo un modelo más grande para el orquestador.

## Medir contra tus notas

Para saber qué forma de calificar se parece más a tu criterio:

1. Toma 20 a 30 tareas que ya calificaste a mano y anota tus notas en un CSV con
   las columnas `archivo` y `nota` (el nombre del archivo tal como lo subes).
2. Califica ese mismo grupo con cada forma. Cada corrida genera su Excel en
   `outputs/`.
3. Compara:
   ```
   uv run medir_contra_docente.py mis_notas.csv outputs/calificaciones_A.xlsx outputs/calificaciones_C.xlsx
   ```

El informe da, por corrida, el error promedio frente a tus notas, el porcentaje de
tareas a 0,5 y a 1 punto o menos, el sesgo (si el sistema califica más alto o más
bajo que tú) y las tareas con más diferencia. Si pasas dos corridas de la misma
forma, también ves cuánto cambia la nota entre corridas.

## Notas importantes del prototipo

- **Nombre del estudiante.** Se busca en este orden:
  1. Líneas como "Nombre:", "Estudiante:", "Alumno:", "Presentado por:" al
     inicio del documento, incluidas tablas de portada, encabezados de página y
     cuadros de texto de Word.
  2. Lo que identifique el modelo.
  3. El nombre del archivo (por ejemplo `Juan Pérez - Taller 2.pdf` o el
     formato de Moodle); en ese caso la fila queda marcada para verificar.

  Si nada funciona, la fila muestra "Desconocido" y queda marcada. El apellido
  lo da el modelo; si no, se deduce (2 palabras: la última; 3: las dos últimas,
  salvo nombres compuestos frecuentes como "Juan Carlos Pérez"; 4 o más: las
  dos últimas).
- **Nota.** El modelo devuelve el puntaje obtenido y el puntaje máximo de la
  rúbrica, y la conversión a la escala 1-10 la hace el programa
  (`puntaje / máximo × 10`, mínimo 1). Así una rúbrica sobre 20 ya no produce
  notas como 18. Si la nota no se puede determinar, se le pide al modelo que la
  corrija una vez; si sigue mal, la celda queda vacía y marcada para revisión.
- **Consistencia.** El modelo corre con `temperature: 0` y semilla fija, así
  que la misma tarea recibe la misma nota en cada corrida.
- **Textos largos.** `NUM_CTX` en `app.py` (16384 por defecto) es la cantidad
  de texto que el modelo puede leer a la vez. Si una tarea no cabe, la fila se
  marca; súbelo si tu equipo tiene memoria, o bájalo a 8192 si Ollama da error
  de memoria.
- Si Ollama se cae a mitad del lote, las tareas ya calificadas se conservan y
  el Excel se genera igual; las restantes quedan marcadas como no calificadas.
- **Escaneos y fotos (OCR).** En los PDF, cada página sin texto
  seleccionable se convierte en imagen y la transcribe `qwen3-vl:8b`; las
  páginas con texto se leen normal, así que un PDF mixto funciona. Las fotos
  de celular se enderezan solas. Todos los escaneos del grupo se leen primero
  y luego se libera la GPU para calificar, en vez de alternar los dos modelos.
  - El modelo de visión puede equivocarse o inventar texto cuando no logra
    leer algo. Por eso toda tarea leída desde imagen queda marcada para que
    revises la transcripción, y lo que no se pudo leer aparece como
    `[ilegible]`. Al calificar, el modelo sabe que el texto viene de una imagen.
  - Es lento: cuenta con decenas de segundos por página manuscrita. Si va muy
    lento, cambia `OCR_MODEL` a `qwen3-vl:4b` en `app.py`.
  - Se leen hasta `OCR_MAX_PAGINAS` (15) páginas escaneadas por documento.
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
- Poder cancelar un grupo que ya se está calificando.
