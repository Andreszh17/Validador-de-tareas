"""Compara las notas del Validador con las que puso el docente a mano.

Uso:
    python medir_contra_docente.py notas_docente.csv outputs/calificaciones_A.xlsx outputs/calificaciones_C.xlsx

- notas_docente.csv (o .xlsx): columnas "archivo" y "nota", una fila por tarea,
  con el nombre del archivo tal como se subió al Validador.
- Cada Excel es una corrida del Validador (por ejemplo, una por cada forma de calificar).
  El nombre del archivo identifica la corrida en el informe.

Para cada corrida informa el error promedio, el acuerdo a 0,5 y a 1 punto, el sesgo y
las tareas con más diferencia. Con dos o más corridas, también cuánto difieren entre sí
(sirve para medir la estabilidad si corres dos veces la misma configuración).
"""
import csv
import os
import sys
import unicodedata

from openpyxl import load_workbook


def _clave(nombre):
    nombre = unicodedata.normalize("NFC", os.path.basename(str(nombre or "")).strip())
    return nombre.casefold()


def _numero(valor):
    if valor is None or valor == "":
        return None
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return None


def _filas_tabla(ruta):
    """Devuelve una lista de dicts (encabezado en minúsculas → valor) de un .csv o .xlsx."""
    if ruta.lower().endswith(".csv"):
        with open(ruta, encoding="utf-8-sig", newline="") as f:
            muestra = f.read(2048)
            f.seek(0)
            dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t") if muestra else csv.excel
            return [{(k or "").strip().casefold(): v for k, v in fila.items()} for fila in csv.DictReader(f, dialect=dialecto)]
    ws = load_workbook(ruta, read_only=True, data_only=True).worksheets[0]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return []
    encabezados = [str(h or "").strip().casefold() for h in filas[0]]
    return [dict(zip(encabezados, fila)) for fila in filas[1:]]


def notas_docente(ruta):
    notas = {}
    for fila in _filas_tabla(ruta):
        archivo = fila.get("archivo")
        nota = _numero(fila.get("nota", fila.get("nota_docente")))
        if archivo and nota is not None:
            notas[_clave(archivo)] = (str(archivo), nota)
    if not notas:
        sys.exit(f"No encontré filas con columnas «archivo» y «nota» en {ruta}.")
    return notas


def notas_validador(ruta):
    notas = {}
    for fila in _filas_tabla(ruta):
        archivo = fila.get("archivo")
        nota = _numero(fila.get("nota (1-10)"))
        if archivo:
            notas[_clave(archivo)] = nota
    return notas


def informe(nombre, sistema, docente):
    comunes = [k for k in docente if k in sistema]
    calificadas = [k for k in comunes if sistema[k] is not None]
    print(f"\n=== {nombre}")
    print(f"Tareas con nota del docente: {len(docente)} · en esta corrida: {len(comunes)} · con nota del sistema: {len(calificadas)}")
    faltan = [docente[k][0] for k in docente if k not in sistema]
    if faltan:
        print(f"No aparecen en esta corrida: {', '.join(faltan[:8])}{'…' if len(faltan) > 8 else ''}")
    if not calificadas:
        print("Sin tareas para comparar.")
        return None
    difs = [sistema[k] - docente[k][1] for k in calificadas]
    abs_difs = [abs(d) for d in difs]
    n = len(difs)
    error = sum(abs_difs) / n
    sesgo = sum(difs) / n
    print(f"Error promedio: {error:.2f} puntos")
    print(f"Acuerdo: {100 * sum(d <= 0.5 for d in abs_difs) / n:.0f} % a 0,5 o menos · "
          f"{100 * sum(d <= 1 for d in abs_difs) / n:.0f} % a 1 punto o menos")
    tendencia = "más alto" if sesgo > 0 else "más bajo"
    print(f"Sesgo: el sistema califica en promedio {abs(sesgo):.2f} puntos {tendencia} que el docente")
    sin_nota = len(comunes) - len(calificadas)
    if sin_nota:
        print(f"Sin nota del sistema (para revisión manual): {sin_nota}")
    peores = sorted(calificadas, key=lambda k: -abs(sistema[k] - docente[k][1]))[:5]
    print("Mayores diferencias:")
    for k in peores:
        print(f"  {docente[k][0]}: docente {docente[k][1]:.1f} · sistema {sistema[k]:.1f}")
    return {"nombre": nombre, "error": error, "sesgo": sesgo, "n": n}


def main(argv):
    if len(argv) < 3:
        sys.exit(__doc__)
    docente = notas_docente(argv[1])
    corridas = [(os.path.basename(r), notas_validador(r)) for r in argv[2:]]
    resumen = [informe(nombre, notas, docente) for nombre, notas in corridas]

    resumen = [r for r in resumen if r]
    if len(resumen) > 1:
        print("\n=== Resumen (menor error primero)")
        for r in sorted(resumen, key=lambda r: r["error"]):
            print(f"  {r['nombre']}: error {r['error']:.2f} · sesgo {r['sesgo']:+.2f} · {r['n']} tareas")

    if len(corridas) > 1:
        print("\n=== Diferencia promedio entre corridas (estabilidad)")
        for i, (na, a) in enumerate(corridas):
            for nb, b in corridas[i + 1:]:
                comunes = [k for k in a if k in b and a[k] is not None and b[k] is not None]
                if comunes:
                    d = sum(abs(a[k] - b[k]) for k in comunes) / len(comunes)
                    print(f"  {na} vs {nb}: {d:.2f} puntos en {len(comunes)} tareas")


if __name__ == "__main__":
    main(sys.argv)
