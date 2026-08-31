"""
Puente entre el submission de FastField y el escritor de Excel.

`utils/fastfield.py` devuelve lo que se capturó en campo (etiquetas de texto).
Aquí se resuelve cada etiqueta contra el libro maestro para obtener la FILA
exacta donde hay que escribir, y se arma el dict que consume
`cenit_report.generar_informe`.
"""
from __future__ import annotations

import re

from utils import cenit_report as cr
from utils.fastfield import TOPES, norm

_FILA_EN_ETIQUETA = re.compile(r"\(fila\s+(\d+)\)", re.I)


# ── Resolución de etiquetas contra el libro ───────────────────────────────────

def _resolver_items(capturados, catalogo, avisos):
    por_num = {i["num"]: i for i in catalogo}
    por_desc = {norm(i["descripcion"]): i for i in catalogo}

    resueltos = []
    for it in capturados:
        cat = por_num.get(it["num"]) or por_desc.get(norm(it["descripcion"]))
        if not cat:
            avisos.append(
                f"Ítem {it['num']} ('{it['descripcion'][:50]}') no está en el "
                f"catálogo del libro. Se omite."
            )
            continue
        if cat["num"] != it["num"]:
            avisos.append(
                f"El ítem llegó como {it['num']} pero coincide por descripción "
                f"con el {cat['num']}. Se usó el {cat['num']}."
            )
        resueltos.append({
            **it,
            "row_num":     cat["row_num"],
            "unidad":      cat["unidad"],
            "cat_desc":    cat["descripcion"],
            "sitio":       cat["und_funcional"] or cat["estacion"] or "",
            "contractual": cat["cantidad_total"],
            "ejecutado":   cat["ejecutado"],
        })
    return resueltos


def _resolver_recursos(capturados, catalogo, campo_cat, campo_cap, etiqueta, avisos):
    """Resuelve cargos u equipos contra su matriz. Soporta el sufijo '(fila N)'."""
    por_fila = {c["row_num"]: c for c in catalogo}
    por_nombre: dict[str, list] = {}
    for c in catalogo:
        por_nombre.setdefault(norm(c[campo_cat]), []).append(c)

    resueltos = []
    for reg in capturados:
        bruto = reg[campo_cap]
        m = _FILA_EN_ETIQUETA.search(bruto)
        cat = None
        if m and int(m.group(1)) in por_fila:
            cat = por_fila[int(m.group(1))]
        else:
            limpio = norm(_FILA_EN_ETIQUETA.sub("", bruto))
            candidatos = por_nombre.get(limpio, [])
            if len(candidatos) == 1:
                cat = candidatos[0]
            elif len(candidatos) > 1:
                avisos.append(
                    f"'{bruto}' aparece en {len(candidatos)} filas de la matriz "
                    f"({', '.join(str(c['row_num']) for c in candidatos)}). "
                    f"Se usó la {candidatos[0]['row_num']} — desambigua la opción "
                    f"en FastField con el sufijo '(fila N)'."
                )
                cat = candidatos[0]

        if not cat:
            avisos.append(f"{etiqueta} '{bruto}' no existe en el libro. Se omite.")
            continue

        resueltos.append({
            **reg,
            "cantidad": reg.get("cantidad") or 1,
            "row_num": cat["row_num"],
            "resuelto": cat[campo_cat],
        })
    return resueltos


# ── Armado del dict final ─────────────────────────────────────────────────────

def construir(capturado: dict, wb, consecutivo: int,
              fotos_bytes: dict[str, bytes] | None = None) -> tuple[dict, list[str]]:
    """
    capturado    salida de fastfield.parse_submission
    wb           libro maestro abierto con cenit_report.abrir_libro
    consecutivo  del informe a generar
    fotos_bytes  {filename: bytes} descargados de la API de FastField

    Devuelve (datos_para_generar_informe, avisos).
    """
    avisos = list(capturado.get("avisos") or [])
    fotos_bytes = fotos_bytes or {}

    items_cat   = [i for i in cr.leer_items(wb) if not i["es_encabezado"]]
    cargos_cat  = cr.leer_cargos(wb)
    equipos_cat = cr.leer_equipos(wb)

    items = _resolver_items(capturado.get("items") or [], items_cat, avisos)
    mano_obra = _resolver_recursos(
        capturado.get("mano_obra") or [], cargos_cat, "cargo", "cargo", "Cargo", avisos
    )
    equipos = _resolver_recursos(
        capturado.get("equipos") or [], equipos_cat, "tipo", "tipo", "Equipo", avisos
    )

    # Un mismo ítem puede venir varias veces en el día: se suman
    avances: dict[int, float] = {}
    for it in items:
        avances[it["row_num"]] = avances.get(it["row_num"], 0) + it["cantidad"]

    # Filas del bloque de recursos del informe (693-711), cruzadas por etiqueta
    filas_mo, filas_eq = cr.mapa_filas_recursos(wb)

    datos = {
        "consecutivo": consecutivo,
        "frente": capturado.get("frente") or "",
        "actividades": (capturado.get("actividades") or [])[: TOPES["actividades"][0]],
        "observaciones": (capturado.get("observaciones") or [])[: TOPES["observaciones"][0]],
        "motivos_disponibilidad": capturado.get("motivos_disponibilidad") or "",
        "jornadas": (capturado.get("jornadas") or [])[: TOPES["jornadas"][0]],
        "avances": [{"row_num": r, "cantidad": c} for r, c in avances.items()],
        # Las HH que van a la matriz son personas x horas: la matriz HH guarda
        # horas totales por fila, y el informe las suma con SUMIFS.
        "hh": [
            {"row_num": r["row_num"], "horas": r["cantidad"] * r["horas"]}
            for r in mano_obra if r["horas"]
        ],
        "equipos": [
            {"row_num": r["row_num"], "horas": r["cantidad"] * r["horas"]}
            for r in equipos if r["horas"]
        ],
        "mo_disponible": [
            {"row_num": filas_mo[r["row_num"]], "horas": r["disponible"]}
            for r in mano_obra if r["disponible"] and r["row_num"] in filas_mo
        ],
        "eq_disponible": [
            {"row_num": filas_eq[r["row_num"]], "horas": r["disponible"]}
            for r in equipos if r["disponible"] and r["row_num"] in filas_eq
        ],
        "eq_fuera_servicio": [
            {"row_num": filas_eq[r["row_num"]], "horas": r["fuera_servicio"]}
            for r in equipos if r["fuera_servicio"] and r["row_num"] in filas_eq
        ],
        "fotos": [
            {
                "image_bytes": fotos_bytes.get(f["filename"]),
                "descripcion": f["descripcion"],
            }
            for f in (capturado.get("fotos") or [])[: TOPES["fotos"][0]]
        ],
    }

    # Detalle resuelto, para que la app pueda mostrarlo antes de generar
    detalle = {"items": items, "mano_obra": mano_obra, "equipos": equipos}

    _avisar_sobre_cantidad_personal(mano_obra, avisos)
    _avisar_sobre_ejecucion(items, avisos)
    return {**datos, "_detalle": detalle}, avisos


def _avisar_sobre_ejecucion(items, avisos):
    """Avisa si una cantidad reportada excede el saldo pendiente del ítem."""
    acumulado: dict[int, float] = {}
    primero: dict[int, dict] = {}
    for it in items:
        acumulado[it["num"]] = acumulado.get(it["num"], 0) + it["cantidad"]
        primero.setdefault(it["num"], it)
    for it in primero.values():          # un aviso por ítem, no por repetición
        contractual = it.get("contractual") or 0
        ejecutado = it.get("ejecutado") or 0
        if not contractual:
            continue
        pendiente = contractual - ejecutado
        total_hoy = acumulado[it["num"]]
        if total_hoy > pendiente + 1e-6:
            avisos.append(
                f"Ítem {it['num']} ({it['cat_desc'][:40]}): se reportan "
                f"{total_hoy:g} {it['unidad']} pero solo quedan {pendiente:g} "
                f"pendientes de {contractual:g}."
            )


def _avisar_sobre_cantidad_personal(mano_obra, avisos):
    """
    La matriz HH tiene una fila por PERSONA (columnas NOMBRE / CARGO / CEDULA),
    y el informe cuenta el personal con COUNTIFS sobre esas filas.

    El formulario captura 'Cargo + Cantidad', sin identificar a cada persona, así
    que N personas con el mismo cargo no se pueden representar como N filas. Se
    escriben las horas totales (cantidad x horas), con lo cual la columna
    HORAS LABORADAS queda correcta pero CANT. mostrará 1 en vez de N.
    """
    for r in mano_obra:
        n = r.get("cantidad") or 1
        if n > 1 and r.get("horas"):
            avisos.append(
                f"{r['resuelto']}: {n:g} personas x {r['horas']:g} h = "
                f"{n * r['horas']:g} HH. Las horas quedan bien, pero la columna "
                f"CANT. del informe mostrará 1 (la matriz HH tiene una fila por "
                f"persona y el formulario no captura nombres)."
            )
