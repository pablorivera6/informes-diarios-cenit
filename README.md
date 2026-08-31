# Informes Diarios — Cenit ODS03

Automatiza el diligenciamiento del informe diario del contrato 8000008746 / ODS03
(formato GDA-FR-363) a partir de un formulario de FastField diligenciado en campo.

Réplica del proyecto `informes-eco` (Ecopetrol), adaptada al formato de Cenit.

## Cómo funciona

La plantilla **es el informe del día anterior**: hace de libro maestro acumulativo.
La app lo lee, escribe encima lo del día nuevo y devuelve el archivo listo.

```
[Export .xlsx de FastField]  ─┐
                              ├─> parser ─> armado ─> XlsxZipWriter ─> descarga
[Último informe enviado]     ─┘                 ↑
                                     FastField REST v3 (fotos)
```

Todo el informe cuelga del **consecutivo** (`1. Informe Diario`!Z6), que es el día
calendario corrido desde el 06-abr-2026. De ahí sale la fecha por HLOOKUP, y de la
fecha salen los datos de las cuatro hojas ocultas.

## Uso

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

| Archivo | Qué hace |
|---|---|
| `app.py` | Interfaz Streamlit, 7 pasos |
| `utils/zip_writer.py` | Escritura sobre el ZIP/XML preservando formato |
| `utils/cenit_report.py` | Mapeo del formato Cenit (celdas, matrices, fotos) |
| `utils/fastfield.py` | Parser del export de FastField |
| `utils/armado.py` | Resuelve etiquetas contra el libro y arma el dict final |
| `utils/fastfield_api.py` | Descarga de fotos por API (copiado de Ecopetrol) |
| `FORMULARIO_FASTFIELD.md` | Preguntas del formulario y a qué celda va cada una |
| `catalogos/` | Listas para los desplegables de FastField |

## Pruebas

```bash
python3 scripts/smoke_test.py      # motor de Excel contra el informe real
python3 scripts/test_pipeline.py   # FastField -> parser -> armado -> Excel
```

## Ajustar el mapeo con un export real

```bash
python3 scripts/diagnostico_export.py <export_de_fastfield.xlsx>
```

Muestra hoja por hoja qué reconoció el parser, qué columna asoció a cada pregunta
y qué quedó sin mapear. Es el paso para afinar los nombres reales que genera
FastField sin adivinar.

Ninguna necesita credenciales ni red.

## Regenerar los catálogos

```bash
python3 scripts/exportar_lookup_fastfield.py   # lookup list de ítems
python3 scripts/exportar_catalogos.py          # cargos y equipos
```
