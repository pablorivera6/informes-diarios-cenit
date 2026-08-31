"""
Exporta los catálogos del libro maestro a CSV para cargarlos como listas
desplegables en FastField.

    python3 scripts/exportar_catalogos.py [ruta_del_informe.xlsx]

Genera en catalogos/:
    items_de_pago.csv   — para el dropdown 'Item de pago ' de subform_2
    cargos.csv          — para 'Cargo' de subform_3
    equipos.csv         — para 'Tipo de equipo' de subform_4

El formato de la etiqueta de ítem es "012 — DESCRIPCIÓN  [m]", el mismo que
parsea utils/fastfield.py (heredado del proyecto Ecopetrol).
"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from utils import cenit_report as cr

DEFECTO = Path.home() / "Downloads" / "2026-08-22_8000008746_ODS03_Informe_Diario_139.xlsx"
DESTINO = Path(__file__).resolve().parent.parent / "catalogos"


def etiqueta_item(it: dict) -> str:
    unidad = f"  [{it['unidad']}]" if it["unidad"] else ""
    sitio = it["und_funcional"] or it["estacion"] or ""
    sitio = f" ({sitio})" if sitio else ""
    return f"{it['num']:03d} — {it['descripcion']}{sitio}{unidad}"


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFECTO
    if not ruta.exists():
        sys.exit(f"No se encontró el informe en {ruta}")

    wb = cr.abrir_libro(ruta.read_bytes())
    DESTINO.mkdir(exist_ok=True)

    items = [i for i in cr.leer_items(wb) if not i["es_encabezado"]]
    with open(DESTINO / "items_de_pago.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["etiqueta", "num", "fila", "codigo", "unidad", "cantidad_contractual"])
        for it in items:
            w.writerow([etiqueta_item(it), it["num"], it["row_num"],
                        it["codigo"], it["unidad"], it["cantidad_total"]])

    cargos = cr.leer_cargos(wb)
    with open(DESTINO / "cargos.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["etiqueta", "cargo", "fila"])
        vistos: dict[str, int] = {}
        for c in cargos:
            vistos[c["cargo"]] = vistos.get(c["cargo"], 0) + 1
        repetidos = {k for k, v in vistos.items() if v > 1}
        for c in cargos:
            # Desambiguar los cargos que aparecen en más de una fila de HH
            etq = f"{c['cargo']} (fila {c['row_num']})" if c["cargo"] in repetidos else c["cargo"]
            w.writerow([etq, c["cargo"], c["row_num"]])

    equipos = cr.leer_equipos(wb)
    with open(DESTINO / "equipos.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["etiqueta", "tipo", "fila"])
        for e in equipos:
            w.writerow([e["tipo"], e["tipo"], e["row_num"]])

    print(f"Libro: {ruta.name}")
    print(f"Consecutivo actual: {cr.leer_consecutivo_actual(wb)}")
    print(f"\nEscrito en {DESTINO}/")
    print(f"  items_de_pago.csv  {len(items)} ítems")
    print(f"  cargos.csv         {len(cargos)} cargos"
          + (f"  (ojo: {sorted(repetidos)} aparecen repetidos)" if repetidos else ""))
    print(f"  equipos.csv        {len(equipos)} equipos")
    print("\nMuestra de ítems:")
    for it in items[:5]:
        print(f"  {etiqueta_item(it)}")


if __name__ == "__main__":
    main()
