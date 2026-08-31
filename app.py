"""Automatización de Informes Diarios — Proyecto Cenit 8000008746 / ODS03."""
import base64
import hmac
import io
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from utils import armado, cenit_report as cr
from utils.fastfield import parse_submission
from utils.fastfield_api import download_submission_photos

st.set_page_config(
    page_title="Informes Diarios — Cenit ODS03",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def _img_b64(path: str) -> str:
    data = Path(path).read_bytes()
    ext = path.rsplit(".", 1)[-1].lower()
    mime = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
    return f"data:{mime};base64,{base64.b64encode(data).decode()}"


LOGO_PCC = _img_b64("assets/logo_pcc.png")


def secreto(clave: str, defecto: str = "") -> str:
    """st.secrets lanza si no existe .streamlit/secrets.toml."""
    try:
        return st.secrets.get(clave, defecto)
    except Exception:
        return defecto

st.markdown(Path("assets/estilos.css").read_text(encoding="utf-8"), unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# Control de acceso (opcional)
#
# Por defecto la app queda ABIERTA: quien tenga la URL entra directo.
#
# Si algún día quieres cerrarla, basta con definir `app_password` en los secrets
# (Streamlit Cloud: Settings -> Secrets). No hay que tocar código: en cuanto esa
# clave existe, la app empieza a pedirla.
#
# Ten presente que con la app abierta, las credenciales de FastField que viven
# en los secrets quedan utilizables por cualquiera que tenga el enlace.
# ═════════════════════════════════════════════════════════════════════════════

def _puerta() -> bool:
    clave = secreto("app_password")
    if not clave:
        return True                      # sin contraseña configurada: acceso libre
    if st.session_state.get("autenticado"):
        return True

    st.markdown('<div class="upload-label" style="margin-top:28px;">Acceso</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="upload-desc">Esta app es de uso interno de PCC.</div>',
                unsafe_allow_html=True)

    col, _ = st.columns([1, 2], gap="large")
    with col:
        intento = st.text_input("Contraseña", type="password",
                                label_visibility="collapsed",
                                placeholder="Contraseña de acceso")
        if st.button("Entrar", type="primary", width="stretch"):
            if hmac.compare_digest(intento, clave):
                st.session_state.autenticado = True
                st.rerun()
            else:
                # pill() se define más abajo; aquí todavía no existe.
                st.markdown(
                    '<div class="pill pill-err"><span class="pill-dot"></span>'
                    'Contraseña incorrecta.</div>',
                    unsafe_allow_html=True,
                )
    return False


if not _puerta():
    st.stop()



st.markdown(f"""
<div class="pcc-header">
    <div class="h-logos">
        <img src="{LOGO_PCC}" alt="PCC" style="height:52px;object-fit:contain;">
    </div>
    <div class="h-title">
        <h1>Informe Diario de Proyectos</h1>
        <p>Cenit &nbsp;&middot;&nbsp; Contrato 8000008746 / ODS03 &nbsp;&middot;&nbsp;
           Protección Catódica de Colombia</p>
    </div>
    <div class="h-logos"></div>
</div>
""", unsafe_allow_html=True)


def paso(num: str, titulo: str, desc: str):
    st.markdown(f"""
    <div class="step-hdr">
        <div><span class="step-num">{num}</span></div>
        <div>
            <div class="step-title">{titulo}</div>
            <div class="step-desc">{desc}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def pill(texto: str, tipo: str = "ok"):
    st.markdown(
        f'<div class="pill pill-{tipo}"><span class="pill-dot"></span>{texto}</div>',
        unsafe_allow_html=True,
    )


def bloqueo(titulo: str, texto: str):
    """Riesgo que puede corromper el libro maestro. No es un pill."""
    st.markdown(
        f'<div class="bloqueo"><div class="bloqueo-icono">!</div>'
        f'<div class="bloqueo-cuerpo"><div class="bloqueo-titulo">{titulo}</div>'
        f'<div class="bloqueo-texto">{texto}</div></div></div>',
        unsafe_allow_html=True,
    )


def notas(titulo: str, items: list[str]):
    """Informativo: se pliega para no competir con lo que sí exige acción."""
    if not items:
        return
    with st.expander(f"{titulo} ({len(items)})", expanded=False):
        st.markdown(
            '<ul class="nota-lista">'
            + "".join(f"<li>{i}</li>" for i in items)
            + "</ul>",
            unsafe_allow_html=True,
        )


def vacio(texto: str):
    """Un dato ausente no es una advertencia."""
    st.markdown(f'<div class="vacio-nota">{texto}</div>', unsafe_allow_html=True)


def resumen(datos_resumen: list[tuple[str, str, bool]]):
    st.markdown(
        '<div class="resumen">'
        + "".join(
            f'<div class="resumen-dato{" vacio" if vac else ""}">'
            f'<span class="k">{k}</span><span class="v">{v}</span></div>'
            for k, v, vac in datos_resumen
        )
        + "</div>",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def _miniatura(blob: bytes, caja: tuple[int, int], modo: str) -> bytes:
    """
    Vista previa con el recorte real que va a quedar en el Excel.

    Antes se mostraba la foto original, así que la previsualización mentía:
    el Excel recibe la imagen ajustada a una caja de 2.7:1. De paso, todas las
    miniaturas quedan de la misma proporción y la fila deja de verse dentada.
    """
    return cr.ajustar_imagen(blob, caja, modo)


for clave, defecto in [
    ("capturado", None), ("plantilla_bytes", None), ("ctx", None),
    ("fotos_bytes", {}), ("salida", None),
]:
    st.session_state.setdefault(clave, defecto)


# ═════════════════════════════════════════════════════════════════════════════
paso("01", "Cargar archivos",
     "Submission de FastField y el último informe enviado a la interventoría")

col_a, col_b = st.columns(2, gap="large")

with col_a:
    st.markdown('<div class="upload-label">Submission FastField</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-desc">Export del formulario diligenciado en campo (.xlsx)</div>',
                unsafe_allow_html=True)
    ff_file = st.file_uploader("Submission", type=["xlsx"], key="ff",
                              label_visibility="collapsed")
    if ff_file:
        try:
            cap = parse_submission(io.BytesIO(ff_file.read()))
            st.session_state.capturado = cap

            nombres = [f["filename"] for f in cap["fotos"]]
            email = secreto("fastfield_email")
            passwd = secreto("fastfield_password")
            if nombres and email and passwd:
                with st.spinner(f"Descargando {len(nombres)} foto(s) de FastField..."):
                    blobs, err = download_submission_photos(
                        nombres, email, passwd,
                        secreto("fastfield_org_id"),
                        secreto("fastfield_subscription_key"),
                    )
                st.session_state.fotos_bytes = {
                    n: b for n, b in zip(nombres, blobs) if b
                }
                ok = len(st.session_state.fotos_bytes)
                pill(f"Cargado &mdash; <strong>{cap['fecha']}</strong> &nbsp;·&nbsp; "
                     f"{cap['frente']} &nbsp;·&nbsp; {ok}/{len(nombres)} fotos",
                     "ok" if ok == len(nombres) else "warn")
                if err:
                    pill(f"API de fotos: <code>{err}</code>", "warn")
            else:
                st.session_state.fotos_bytes = {}
                pill(f"Cargado &mdash; <strong>{cap['fecha']}</strong> &nbsp;·&nbsp; "
                     f"{cap['frente']}")
                if nombres:
                    pill("Sin credenciales de FastField: sube las fotos a mano abajo.", "warn")
        except Exception as e:
            pill(f"Error al leer el submission: {e}", "err")

with col_b:
    st.markdown('<div class="upload-label">Último informe enviado</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-desc">Hace de libro maestro acumulativo (.xlsx)</div>',
                unsafe_allow_html=True)
    rpt_file = st.file_uploader("Plantilla", type=["xlsx"], key="rpt",
                                label_visibility="collapsed")
    if rpt_file:
        try:
            raw = rpt_file.read()
            st.session_state.plantilla_bytes = raw
            # El libro se lee una sola vez y se descarta: openpyxl lo expande a
            # ~1 GB y conservarlo tumbaba la app por memoria en la nube.
            with st.spinner("Leyendo el libro maestro (unos segundos)..."):
                st.session_state.ctx = cr.construir_contexto(raw)
            ctx = st.session_state.ctx
            pill(f"Informe N.° <strong>{ctx['consecutivo_actual']}</strong> "
                 f"&nbsp;·&nbsp; <strong>{len(ctx['items'])}</strong> ítems en el catálogo")
        except Exception as e:
            pill(f"Error al leer el informe: {e}", "err")

if st.session_state.capturado is None or st.session_state.ctx is None:
    pill("Carga los dos archivos para continuar.", "warn")
    st.stop()

cap = st.session_state.capturado
ctx = st.session_state.ctx
dia1 = ctx["dia1"]
consec_prev = ctx["consecutivo_actual"]


# ═════════════════════════════════════════════════════════════════════════════
paso("02", "Fecha y consecutivo", "El consecutivo se calcula de la fecha del informe")

c1, c2, c3 = st.columns(3, gap="medium")
with c1:
    fecha = st.date_input("Fecha del informe", value=cap["fecha"] or date.today())
with c2:
    # Una fecha anterior al día 1 del proyecto da un consecutivo negativo
    crudo = cr.fecha_a_consecutivo(fecha, dia1)
    sugerido = max(1, crudo)
    consecutivo = st.number_input("Consecutivo", min_value=1, value=sugerido, step=1)
with c3:
    st.metric("Fecha que quedará en el informe",
              cr.consecutivo_a_fecha(consecutivo, dia1).strftime("%d/%m/%Y"))

# Triaje: lo que puede corromper el libro no puede verse igual que un trámite.
bloqueantes: list[tuple[str, str]] = []
avisos_fecha: list[str] = []

if crudo < 1:
    bloqueantes.append((
        "Fecha anterior al inicio del proyecto",
        f"El {fecha:%d/%m/%Y} es anterior al {dia1:%d/%m/%Y}, que es el día 1 "
        f"(consecutivo 1). Revisa la fecha del submission.",
    ))
if consecutivo != sugerido:
    avisos_fecha.append(
        f"El consecutivo {consecutivo} corresponde al "
        f"{cr.consecutivo_a_fecha(consecutivo, dia1):%d/%m/%Y}, no al {fecha:%d/%m/%Y}."
    )
if consecutivo != consec_prev + 1:
    avisos_fecha.append(
        f"El informe cargado es el {consec_prev}; lo normal sería generar el "
        f"{consec_prev + 1}."
    )

ya = cr.dia_ya_tiene_datos_ctx(ctx, consecutivo)
if any(ya.values()):
    bloqueantes.append((
        f"El día {consecutivo} ya tiene datos cargados",
        f"Encontré {ya['avances']} avance(s), {ya['hh']} registro(s) de HH y "
        f"{ya['equipos']} de equipos en esa columna. Las cantidades se "
        f"<strong>suman</strong>: si generas otra vez sobre este mismo día, "
        f"quedarán duplicadas en el libro maestro y las heredarán todos los "
        f"informes siguientes.",
    ))

for titulo, texto in bloqueantes:
    bloqueo(titulo, texto)
for a in avisos_fecha:
    pill(a, "warn")
notas("Notas sobre la plantilla", ctx["avisos_manuales"])


# ═════════════════════════════════════════════════════════════════════════════
datos, avisos = armado.construir(cap, ctx, int(consecutivo), st.session_state.fotos_bytes)
det = datos["_detalle"]

# Un ítem sobre-ejecutado exige decisión; que falte una sección del formulario
# es solo información. No pueden compartir tratamiento.
exigen_accion = [a for a in avisos if "se reportan" in a or "no está en el catálogo" in a]
informativos = [a for a in avisos if a not in exigen_accion]

if avisos:
    paso("03", "Revisión", f"{len(avisos)} punto(s) detectados al cruzar con el libro")
    for a in exigen_accion:
        pill(a, "warn")
    notas("Otros avisos", informativos)
else:
    paso("03", "Revisión", "El submission cuadra con el libro maestro")
    pill("Todo el contenido del formulario se resolvió contra el catálogo.")


# ═════════════════════════════════════════════════════════════════════════════
paso("04", "Avance de ítems", "Cantidades que se escribirán en la matriz del día")

if det["items"]:
    st.dataframe(pd.DataFrame([{
        "N°": i["num"],
        "Descripción": i["cat_desc"][:60],
        "Sitio": i["sitio"],
        "Dimensiones": " × ".join(f"{d:g}" for d in i["dimensiones"]),
        "Cantidad": i["cantidad"],
        "Und.": i["unidad"],
        "Contractual": i["contractual"],
        "Ejecutado antes": round(i["ejecutado"] or 0, 3),
    } for i in det["items"]]), width="stretch", hide_index=True)
else:
    vacio("El submission no trae avance de cantidades. "
          "La página de ítems del formulario quedó vacía o no se diligenció.")

st.markdown('<div class="sub-label">Mano de obra y equipos</div>', unsafe_allow_html=True)
rc1, rc2 = st.columns(2, gap="large")
with rc1:
    if det["mano_obra"]:
        st.dataframe(pd.DataFrame([{
            "Cargo": m["resuelto"], "Fila HH": m["row_num"],
            "Horas": m["horas"], "Disponible": m["disponible"],
        } for m in det["mano_obra"]]), width="stretch", hide_index=True)
    else:
        vacio("Sin mano de obra reportada.")
with rc2:
    if det["equipos"]:
        st.dataframe(pd.DataFrame([{
            "Equipo": e["resuelto"], "Horas": e["horas"],
            "Disponible": e["disponible"], "Fuera servicio": e["fuera_servicio"],
        } for e in det["equipos"]]), width="stretch", hide_index=True)
    else:
        vacio("Sin equipos reportados: el formulario todavía no tiene "
              "página de equipos.")


# ═════════════════════════════════════════════════════════════════════════════
paso("05", "Actividades, jornada y observaciones", "Texto que va al cuerpo del informe")

tc1, tc2 = st.columns([3, 2], gap="large")
with tc1:
    st.markdown(f'<div class="sub-label">Actividades ejecutadas '
                f'({len(datos["actividades"])}/28)</div>', unsafe_allow_html=True)
    actividades = st.text_area(
        "Actividades ejecutadas, una por línea",
        value="\n".join(datos["actividades"]), height=220,
        label_visibility="collapsed",
    ).split("\n")

    st.markdown(f'<div class="sub-label">Observaciones '
                f'({len(datos["observaciones"])}/8)</div>', unsafe_allow_html=True)
    observaciones = st.text_area(
        "Observaciones, una por línea",
        value="\n".join(datos["observaciones"]), height=120,
        label_visibility="collapsed", key="obs",
    ).split("\n")

with tc2:
    st.markdown('<div class="sub-label">Frente y jornada</div>', unsafe_allow_html=True)
    frente = st.text_input("Frente / sitio", value=datos["frente"])
    for i, j in enumerate(datos["jornadas"]):
        st.markdown(f'<div class="item-title">Jornada {i+1} &nbsp;—&nbsp; {j["frente"]}</div>',
                    unsafe_allow_html=True)
        st.text(f"{j['hora_inicio']} a {j['hora_final']}  ·  total {j['total_horas']}"
                + (f"\nEvento {j['evento_inicio']}-{j['evento_fin']}: {j['evento_desc'][:60]}"
                   if j["hubo_evento"] == "SI" else ""))
    motivos = st.text_area("Motivos de disponibilidad",
                           value=datos["motivos_disponibilidad"], height=80)


# ═════════════════════════════════════════════════════════════════════════════
paso("06", "Registro fotográfico",
     f"{len(cap['fotos'])} de 10 slots · así es como van a quedar en el Excel")

modo = st.radio(
    "Ajuste a la caja del formato",
    ["contener", "llenar"],
    horizontal=True,
    format_func=lambda m: ("Contener — foto completa" if m == "contener"
                           else "Llenar — recorta los bordes"),
    help="Las cajas del formato son tiras anchas (relación 2.7:1). "
         "«Contener» no pierde nada pero deja franjas blancas; "
         "«Llenar» recorta arriba y abajo.",
)

cajas = ctx["cajas"]

faltantes = [f for f in cap["fotos"] if f["filename"] not in st.session_state.fotos_bytes]
if faltantes:
    pill(f"{len(faltantes)} de {len(cap['fotos'])} fotos no se descargaron de la API. "
         f"Súbelas a mano en su slot.", "warn")

fotos_final = []
POR_FILA = 3          # 3 por fila: las miniaturas son apaisadas (2.7:1)
lista = cap["fotos"][:cr.MAX_FOTOS]

for inicio in range(0, len(lista), POR_FILA):
    tramo = lista[inicio:inicio + POR_FILA]
    # Una fila de columnas por tramo. Antes se reusaba una sola fila de 5 con
    # cols[idx % 5], así que la foto 6 se apilaba debajo de la 1.
    columnas = st.columns(POR_FILA, gap="medium")
    for desplazamiento, f in enumerate(tramo):
        idx = inicio + desplazamiento
        with columnas[desplazamiento]:
            st.markdown(
                f'<div class="foto-cabecera"><span class="slot">SLOT {idx + 1}</span>'
                f'<span>{"descargada" if f["filename"] in st.session_state.fotos_bytes else "pendiente"}</span></div>',
                unsafe_allow_html=True,
            )
            blob = st.session_state.fotos_bytes.get(f["filename"])
            if blob is None:
                st.markdown('<div class="foto-falta">No se descargó.<br>Súbela aquí.</div>',
                            unsafe_allow_html=True)
                subida = st.file_uploader(f"Foto {idx + 1}", type=["jpg", "jpeg", "png"],
                                          key=f"up_{idx}", label_visibility="collapsed")
                blob = subida.read() if subida else None
            if blob:
                st.image(_miniatura(blob, cajas[idx], modo), width="stretch")
            desc = st.text_input(
                f"Descripción de la foto {idx + 1}", value=f["descripcion"],
                key=f"desc_{idx}", label_visibility="collapsed",
                placeholder=f"Descripción del slot {idx + 1}",
            )
            fotos_final.append({"image_bytes": blob, "descripcion": desc})

if not lista:
    vacio("El submission no trae fotos.")


# ═════════════════════════════════════════════════════════════════════════════
paso("07", "Generar informe", "Esto es lo que se va a escribir en el libro maestro")

_acts = [a for a in actividades if a.strip()][:28]
_obs = [o for o in observaciones if o.strip()][:8]
_fotos_ok = sum(1 for f in fotos_final if f["image_bytes"])

resumen([
    ("Consecutivo", str(int(consecutivo)), False),
    ("Fecha", f"{cr.consecutivo_a_fecha(int(consecutivo), dia1):%d/%m/%Y}", False),
    ("Ítems", str(len(datos["avances"])), not datos["avances"]),
    ("Registros HH", str(len(datos["hh"])), not datos["hh"]),
    ("Equipos", str(len(datos["equipos"])), not datos["equipos"]),
    ("Actividades", f"{len(_acts)}/28", not _acts),
    ("Observaciones", f"{len(_obs)}/8", not _obs),
    ("Fotos", f"{_fotos_ok}/10", not _fotos_ok),
])

# La acción no puede quedar disponible mientras haya un riesgo de corromper el
# libro. Antes el botón seguía activo aunque la app mostrara el aviso.
puede_generar = True
if bloqueantes:
    puede_generar = st.checkbox(
        f"Entiendo el riesgo ({len(bloqueantes)}) y quiero generar de todos modos",
        key="forzar",
    )
    if not puede_generar:
        st.markdown(
            '<div class="vacio-nota">Corrige lo señalado arriba, o marca la casilla '
            'para continuar bajo tu criterio.</div>',
            unsafe_allow_html=True,
        )

gen_col, _ = st.columns([1, 2], gap="medium")
with gen_col:
    generar = st.button("Generar informe", type="primary",
                        width="stretch", disabled=not puede_generar)

if generar:
    datos_final = {
        **{k: v for k, v in datos.items() if k != "_detalle"},
        "consecutivo": int(consecutivo),
        "frente": frente,
        "actividades": _acts,
        "observaciones": _obs,
        "motivos_disponibilidad": motivos,
        "fotos": fotos_final,
        "modo_foto": modo,
    }
    with st.spinner("Escribiendo el libro y armando el archivo..."):
        try:
            salida = cr.generar_informe(st.session_state.plantilla_bytes, datos_final)
            st.session_state.salida = salida
            fecha_out = cr.consecutivo_a_fecha(int(consecutivo), dia1)
            st.session_state.nombre = (
                f"{fecha_out:%Y-%m-%d}_8000008746_ODS03_Informe_Diario_{int(consecutivo)}.xlsx"
            )
        except Exception as e:
            pill(f"Error al generar: {e}", "err")
            st.exception(e)

if st.session_state.salida:
    pill(f"Informe <strong>{st.session_state.nombre}</strong> listo "
         f"({len(st.session_state.salida) / 1024 / 1024:.1f} MB).")
    dl_col, _ = st.columns([1, 2], gap="medium")
    with dl_col:
        st.download_button(
            "Descargar informe Excel",
            data=st.session_state.salida,
            file_name=st.session_state.nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    notas("Antes de enviarlo", [
        "Al abrirlo, Excel recalcula todo el libro (fullCalcOnLoad). "
        "Guárdalo una vez para que queden los valores fijos.",
        "El informe hereda el autofiltro de la plantilla: revisa que se vean "
        "todos los ítems que quieres entregar.",
    ])
