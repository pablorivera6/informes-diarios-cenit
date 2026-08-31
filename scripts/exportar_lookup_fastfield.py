"""
Genera el lookup list de ítems de pago para importar en FastField.

Mismo formato que el archivo de Ecopetrol (Label, Item, Especialidad, Cod_SAP, Und),
con dos diferencias obligadas por la estructura del contrato de Cenit:

  1. El SITIO va en la etiqueta, y va ADELANTE de la descripción.
     37 de las 43 descripciones se repiten entre sitios ("Movilización y
     desmovilización" aparece 5 veces). Si el sitio fuera al final, como en
     Ecopetrol, y FastField truncara la etiqueta —las descripciones de Cenit
     llegan a 93 caracteres— quedarían varias opciones visualmente idénticas.

  2. Los números NO son consecutivos: son el 'Numero Asociado' de la matriz
     'Costo Real PDT', y saltan el 1, 12, 77 y 90, que son los encabezados de
     alcance. El número de la etiqueta es la llave con la que la app resuelve
     a qué fila escribir, así que no se puede renumerar.

    python3 scripts/exportar_lookup_fastfield.py [ruta_del_informe.xlsx]
"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from utils import cenit_report as cr

DEFECTO = Path.home() / "Downloads" / "2026-08-22_8000008746_ODS03_Informe_Diario_139.xlsx"
DESTINO = Path(__file__).resolve().parent.parent / "catalogos"

# Códigos de 3 letras, al estilo del archivo de Ecopetrol (CIV, CAT, ELE, ...).
# La columna C de RESUMEN PXQ viene con mayúsculas inconsistentes.
ESPECIALIDAD = {
    "civil": "CIV",
    "corrosion": "COR",
    "electrica": "ELE",
}


def _sin_tildes(s: str) -> str:
    import unicodedata
    n = unicodedata.normalize("NFKD", s)
    return "".join(c for c in n if not unicodedata.combining(c))


def leer_metadatos(wb) -> dict[int, dict]:
    """
    Especialidad, código SAP, catálogo, tipo y alcance por 'Numero Asociado'.
    Viven en RESUMEN PXQ, no en Costo Real PDT.
    """
    ws = wb[cr.SH_PXQ]
    meta, alcance = {}, ""

    for row in range(10, 210):
        num = ws.cell(row=row, column=2).value          # B  Numero Asociado
        desc = ws.cell(row=row, column=7).value         # G  Descripción
        if not isinstance(num, (int, float)) or not desc:
            continue
        if str(desc).strip() in ("", "0"):
            continue

        unidad = ws.cell(row=row, column=9).value       # I  Unidad de medida
        if not unidad or str(unidad).strip() in ("", "0"):
            alcance = str(desc).strip()                 # fila de encabezado
            continue

        esp_bruta = str(ws.cell(row=row, column=3).value or "").strip()
        meta[int(num)] = {
            "especialidad": ESPECIALIDAD.get(_sin_tildes(esp_bruta).lower(),
                                             esp_bruta[:3].upper()),
            "esp_completa": esp_bruta,
            "cod_sap": ws.cell(row=row, column=5).value or "",     # E
            "catalogo": ws.cell(row=row, column=4).value or "",    # D
            "tipo": str(ws.cell(row=row, column=6).value or "").strip(),  # F
            "alcance": alcance,
        }
    return meta


def etiqueta(item: dict) -> str:
    sitio = str(item["und_funcional"] or item["estacion"] or "").strip()
    unidad = str(item["unidad"] or "").strip()
    partes = f"{item['num']:03d} — "
    if sitio:
        partes += f"{sitio} · "
    partes += item["descripcion"]
    if unidad:
        partes += f"  [{unidad}]"
    return partes


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFECTO
    if not ruta.exists():
        sys.exit(f"No se encontró el informe en {ruta}")

    wb = cr.abrir_libro(ruta.read_bytes())
    items = [i for i in cr.leer_items(wb) if not i["es_encabezado"]]
    meta = leer_metadatos(wb)
    DESTINO.mkdir(exist_ok=True)

    sin_meta = []
    filas = []
    for it in items:
        m = meta.get(it["num"])
        if m is None:
            sin_meta.append(it["num"])
            m = {"especialidad": "", "cod_sap": "", "catalogo": "",
                 "tipo": "", "alcance": "", "esp_completa": ""}
        filas.append((it, m))

    # ── Archivo principal: mismas 5 columnas que el de Ecopetrol ────────────
    principal = DESTINO / "items_cenit_fastfield.csv"
    with open(principal, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Label", "Item", "Especialidad", "Cod_SAP", "Und"])
        for it, m in filas:
            w.writerow([etiqueta(it), it["num"], m["especialidad"],
                        m["cod_sap"], it["unidad"]])

    # ── Versión extendida, por si quieren filtros en cascada en FastField ───
    extendido = DESTINO / "items_cenit_fastfield_extendido.csv"
    with open(extendido, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Label", "Item", "Especialidad", "Cod_SAP", "Und",
                    "Sitio", "Tipo", "Alcance", "Catalogo", "Cant_Contractual"])
        for it, m in filas:
            w.writerow([
                etiqueta(it), it["num"], m["especialidad"], m["cod_sap"],
                it["unidad"], it["und_funcional"] or it["estacion"] or "",
                m["tipo"], m["alcance"], m["catalogo"], it["cantidad_total"],
            ])

    print(f"Libro: {ruta.name}")
    print(f"Consecutivo actual: {cr.leer_consecutivo_actual(wb)}\n")
    print(f"  {principal.name:42s} {len(filas)} ítems · 5 columnas")
    print(f"  {extendido.name:42s} {len(filas)} ítems · 10 columnas")
    if sin_meta:
        print(f"\n  AVISO: sin metadatos en RESUMEN PXQ: {sin_meta}")

    from collections import Counter
    print("\nEspecialidades:", dict(Counter(m["especialidad"] for _, m in filas)))
    print("Alcances:", dict(Counter(m["alcance"] for _, m in filas)))
    print("\nMuestra:")
    for it, m in filas[:4]:
        print(f"  {etiqueta(it)}")
    print("  ...")
    print(f"  {etiqueta(filas[-1][0])}")

    # ── Verificación: la app tiene que poder resolver cada etiqueta ─────────
    from utils.fastfield import parse_etiqueta_item
    fallos = [
        (it["num"], etiqueta(it))
        for it, _ in filas
        if parse_etiqueta_item(etiqueta(it))[0] != it["num"]
    ]
    print(f"\nEtiquetas que la app resuelve al ítem correcto: "
          f"{len(filas) - len(fallos)}/{len(filas)}")
    for num, etq in fallos:
        print(f"  FALLA {num}: {etq}")
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())
