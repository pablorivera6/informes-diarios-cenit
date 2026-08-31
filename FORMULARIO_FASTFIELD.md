# Formulario FastField — Informe Diario Cenit

Contrato 8000008746 / ODS03 · formato GDA-FR-363

Este documento tiene las dos caras: **las preguntas** que ve la persona en campo
y **a qué celda del Excel** va cada respuesta.

> **El parser identifica cada subformulario por sus columnas, no por su nombre.**
> Puedes crear las páginas en el orden que quieras y FastField las numerará como
> le parezca (`subform_1`, `subform_2`, …). El texto de las etiquetas también se
> compara de forma tolerante: ignora mayúsculas, tildes y espacios de más.

---

## Página 1 — Identificación

| # | Pregunta | Tipo | Obligatoria | Destino |
|---|---|---|---|---|
| 1 | **Fecha del informe** | Date (default hoy) | Sí | Determina el consecutivo |
| 2 | **Frente o sitio de trabajo** | Dropdown | Sí | `1. Informe Diario`!B662 |

Opciones de la 2: `ADMINISTRATIVO` · `PK-47+128` · `PK-321+100` · `PK-152+428` · `PK-13+300`

**No preguntar:** consecutivo (se calcula como día corrido desde el 06-abr-2026),
contrato, contratista ni sitio de ejecución — están fijos en la plantilla.

---

## Página 2 — Jornada de trabajo · repetible, máx. **2**

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 3 | **Frente** | Dropdown | B718 / B719 |
| 4 | **Hora de inicio** | Time | G |
| 5 | **Hora final** | Time | I |
| 6 | **Total de horas trabajadas** | Number | K |
| 7 | **¿Hubo algún evento de suspensión?** | SI / NO | M |
| 8 | **Hora en que inició el evento** | Time · solo si 7 = SI | O |
| 9 | **Hora en que terminó el evento** | Time · solo si 7 = SI | Q |
| 10 | **¿Qué ocurrió?** | Multilínea · solo si 7 = SI | S |

Las horas se escriben como **texto** (`"07:00"`): la fila 718 tiene formato
`h:mm` y la 719 `General`, así que un número se vería como `0,29` en una de las dos.

---

## Página 3 — Actividades ejecutadas · repetible, máx. **28**

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 11 | **Describa la actividad ejecutada** | Multilínea | G662 … G689 |

---

## Página 4 — Avance de ítems · repetible

Alimenta la matriz oculta `Costo Real PDT` en la columna del día.

| # | Pregunta | Tipo | Notas |
|---|---|---|---|
| 12 | **Ítem de pago ejecutado** | Dropdown (lookup list) | 88 opciones · `catalogos/items_cenit_fastfield.csv` |
| 13 | **Cantidad** | Number, 3 decimales | |
| 14 | **Cantidad — dimensión 2** | Number, opcional | Se multiplica |
| 15 | **Cantidad — dimensión 3** | Number, opcional | Se multiplica |

Las dimensiones 2 y 3 son para ítems en m² o m³: se mide largo × ancho ×
profundidad y FastField hace el producto. Para `Und`, `m` o `km` basta la 13.

### Lookup list

Importa `catalogos/items_cenit_fastfield.csv`, con las mismas columnas que el de
Ecopetrol: `Label, Item, Especialidad, Cod_SAP, Und`. El campo a mostrar es `Label`.

```
010 — PK-152+428 · Instalación y tendido de cable de Protección Catódica  [m]
```

Dos diferencias con el formato de Ecopetrol, ambas obligadas por el contrato:

**El sitio va adelante, no al final.** 37 de las 43 descripciones se repiten entre
sitios —"Movilización y desmovilización" aparece 5 veces— y las descripciones de
Cenit llegan a 93 caracteres. Con el sitio al final, una etiqueta truncada
mostraría varias opciones idénticas.

**Los números no son consecutivos.** Van de 002 a 092 y saltan el 001, 012, 077 y
090, que son los encabezados de alcance. Ese número es la llave con la que la app
sabe a qué fila de `Costo Real PDT` escribir: **no se puede renumerar**.

También se genera `items_cenit_fastfield_extendido.csv` con 5 columnas más
(`Sitio`, `Tipo`, `Alcance`, `Catalogo`, `Cant_Contractual`) por si quieres filtros
en cascada — por ejemplo, elegir el sitio primero y que el desplegable de ítems se
reduzca a los de ese PK.

---

## Página 5 — Mano de obra · repetible

Alimenta la matriz oculta `HH`.

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 16 | **Cargo** | Dropdown | `HH` columna del día |
| 17 | **Horas laboradas** | Number | `HH` columna del día |
| 18 | **Horas disponible (presente pero sin laborar)** | Number, opcional | `1. Informe Diario` col. M |

Opciones (de `catalogos/cargos.csv`):
Director de Obra · Ing. especialista SPC · Residente de Obra ·
Profesional programación y Control de obra · Ingeniero · Profesional HSE ·
**Profesional QA QC (fila 13)** · Analista de operaciones ·
**Profesional QA QC (fila 15)** · Ayudante

> Los dos `Profesional QA QC` son posiciones distintas en `HH` (especialidades
> AD y CV). Si se unifican en una sola opción, la app no puede saber a qué fila
> escribir.

---

## Página 6 — Equipos · repetible

Alimenta la matriz oculta `EQUIPOS`.

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 19 | **Tipo de equipo** | Dropdown | `EQUIPOS` columna del día |
| 20 | **Horas laboradas** | Number | `EQUIPOS` columna del día |
| 21 | **Horas disponible (en sitio sin operar)** | Number, opcional | `1. Informe Diario` col. Z |
| 22 | **Horas fuera de servicio** | Number, opcional | `1. Informe Diario` col. AB |

Opciones: Camioneta 4x4 · Equipo Soldadura · Herramienta menor · Retroexcavadora ·
Pica y Pala · Equipo Topografico · Martillo neumatico · Pinza electrica ·
Multimetro · Telurometro · Planta Electrica

---

## Página 7 — Registro fotográfico · máx. **10**

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 23 | **Fotos del día** | Multi Photo Picker | `2. Reg.Fotográfico`, 10 slots |
| 24 | **Descripción de la foto** | Campo `Comment` del picker | Pie de cada foto |

La fecha del pie ya es fórmula en la plantilla, no hay que capturarla.

---

## Página 8 — Cierre

| # | Pregunta | Tipo | Destino |
|---|---|---|---|
| 25 | **Motivos de disponibilidad de personal o equipos** | Multilínea, opcional | B713 |
| 26 | **Observaciones** | Multilínea, repetible máx. **8** | B722 … B729 |

La 25 conviene condicionarla a que alguien haya reportado horas disponible o
fuera de servicio.

---

## Límites duros del formato

Configúralos como máximo de repeticiones en FastField: lo que exceda **no cabe
en el Excel y se perdería en silencio**. La app avisa si llegan de más.

| Sección | Tope | Celdas |
|---|---|---|
| Actividades ejecutadas | 28 | G662:G689 |
| Jornada de trabajo | 2 | filas 718-719 |
| Observaciones | 8 | B722:B729 |
| Fotos | 10 | 10 slots |
| Ítems / mano de obra / equipos | sin tope | matrices ocultas |

---

## Credenciales de la API

Para que la app descargue las fotos, en `.streamlit/secrets.toml`:

```toml
app_password       = "..."
fastfield_email    = "..."
fastfield_password = "..."
fastfield_org_id   = ""      # opcional, si la cuenta tiene varias orgs
```

## Regenerar los catálogos

Cuando cambie el alcance del contrato:

```bash
python3 scripts/exportar_catalogos.py
```
