"""
Escritura de celdas directamente sobre el ZIP/XML de un XLSX.

Preserva el 100% del formato: gráficos, imágenes, drawings, VML, estilos y
fórmulas. Solo toca el XML mínimo necesario dentro de las hojas, sharedStrings,
drawings y [Content_Types].

Derivado del motor del proyecto Ecopetrol (informes-eco), con tres capacidades
añadidas para Cenit:

  1. `force_full_recalc()` — marca el libro con fullCalcOnLoad="1" para que Excel
     recalcule TODAS las fórmulas al abrir. Indispensable en Cenit: el informe se
     arma con HLOOKUP contra el consecutivo, y esta librería solo actualiza el
     valor cacheado (<v>), nunca recalcula.

  2. `add_part()` — permite agregar archivos NUEVOS al ZIP (el motor original solo
     podía sobrescribir partes existentes).

  3. `insert_picture()` — inserta una imagen nueva creando la parte en xl/media,
     la relación en el .rels del drawing y el anclaje twoCellAnchor. El motor de
     Ecopetrol solo sabía sobrescribir imágenes placeholder ya existentes; la
     plantilla de Cenit no tiene ninguna.
"""
from __future__ import annotations

import io
import re
import zipfile
from datetime import date, datetime
from openpyxl.utils import get_column_letter


# ── Date helpers ──────────────────────────────────────────────────────────────

_EXCEL_EPOCH = date(1899, 12, 30)


def to_excel_serial(d) -> int:
    """Convert a date to Excel serial number."""
    if isinstance(d, datetime):
        d = d.date()
    elif isinstance(d, str):
        d = datetime.strptime(d, "%Y-%m-%d").date()
    return (d - _EXCEL_EPOCH).days


# ── Sheet map ─────────────────────────────────────────────────────────────────

def _get_sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return {sheet_name: zip_path} from workbook.xml + workbook.xml.rels."""
    wb_data   = zf.read("xl/workbook.xml").decode("utf-8")
    rels_data = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    # Extract rId → target from rels (filter only worksheet type)
    rid_target = {}
    for m in re.finditer(
        r'<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]*worksheet[^"]*"[^>]+Target="([^"]+)"',
        rels_data,
    ):
        rid, target = m.group(1), m.group(2)
        if not target.startswith("xl/"):
            target = "xl/" + target
        rid_target[rid] = target

    # Extract sheet name → rId from workbook.xml
    sheet_map = {}
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for m in re.finditer(
        r'<sheet\s[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"',
        wb_data,
    ):
        name, rid = m.group(1), m.group(2)
        if rid in rid_target:
            sheet_map[name] = rid_target[rid]

    return sheet_map


# ── Shared strings ────────────────────────────────────────────────────────────

def _load_shared_strings(zf: zipfile.ZipFile) -> tuple[list[str], str]:
    """Return (list_of_strings, raw_xml_bytes)."""
    raw = zf.read("xl/sharedStrings.xml").decode("utf-8")
    strings = []
    for m in re.finditer(r"<si>(.*?)</si>", raw, re.DOTALL):
        # Extract all <t> texts and join (handles rich text / runs)
        texts = re.findall(r"<t(?:[^>]*)>(.*?)</t>", m.group(1), re.DOTALL)
        strings.append("".join(_unescape(t) for t in texts))
    return strings, raw


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )


def _unescape(text: str) -> str:
    return (
        text.replace("&amp;", "&")
            .replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&quot;", '"')
    )


def _add_shared_string(strings: list[str], ss_xml: str, text: str) -> tuple[int, str]:
    """Add text to sharedStrings if not present; return (index, updated_xml)."""
    if text in strings:
        return strings.index(text), ss_xml

    idx = len(strings)
    strings.append(text)

    # Build new <si> element preserving newlines with xml:space="preserve"
    escaped = _escape(text)
    if "\n" in text:
        new_si = f'<si><t xml:space="preserve">{escaped}</t></si>'
    else:
        new_si = f"<si><t>{escaped}</t></si>"

    # Insert before </sst>
    ss_xml = ss_xml.replace("</sst>", new_si + "</sst>")

    # Update count attribute in <sst ... count="N">
    def inc_count(m):
        return m.group(0).replace(
            f'count="{m.group(1)}"',
            f'count="{int(m.group(1))+1}"',
        )
    ss_xml = re.sub(r'count="(\d+)"', inc_count, ss_xml, count=1)
    ss_xml = re.sub(r'uniqueCount="(\d+)"', lambda m: m.group(0).replace(
        m.group(1), str(int(m.group(1)) + 1)
    ), ss_xml, count=1)

    return idx, ss_xml


# ── Cell update engine ────────────────────────────────────────────────────────

def _cell_ref(col: int, row: int) -> str:
    return get_column_letter(col) + str(row)


def _val_str(value: float | int) -> str:
    if isinstance(value, int) or (isinstance(value, float) and value == int(value)):
        return str(int(value))
    return repr(value)


# Matches both <c r="REF" ...>...</c> and <c r="REF" .../>
_CELL_RE_TEMPLATE = (
    r'<c r="{ref}"([^/]*?)(?:/>|>(.*?)</c>)'
)


def _find_cell(sheet_xml: str, ref: str):
    """Return (match, open_attrs, content, is_self_closing) or None."""
    pat = re.compile(
        r'<c r="' + re.escape(ref) + r'"([^/]*?)(?:/>|>(.*?)</c>)',
        re.DOTALL,
    )
    m = pat.search(sheet_xml)
    if not m:
        return None
    attrs = m.group(1)            # everything between ref and /> or >
    content = m.group(2) or ""   # None if self-closing
    is_self_closing = m.group(2) is None
    return m, attrs, content, is_self_closing


def _set_cell_number(sheet_xml: str, ref: str, value: float | int) -> str:
    """Set numeric value in a cell. Handles self-closing empty cells."""
    val_s = _val_str(value)
    result = _find_cell(sheet_xml, ref)
    if result:
        m, attrs, content, is_sc = result
        # Remove t="s" from attrs if present
        attrs = re.sub(r'\s*t="s"', "", attrs)
        if "<f>" in content:
            # Keep formula, update cached value
            new_content = re.sub(r"<v>[^<]*</v>", f"<v>{val_s}</v>", content)
            if "<v>" not in new_content:
                new_content += f"<v>{val_s}</v>"
        else:
            new_content = f"<v>{val_s}</v>"
        replacement = f'<c r="{ref}"{attrs}>{new_content}</c>'
        sheet_xml = sheet_xml[:m.start()] + replacement + sheet_xml[m.end():]
    else:
        sheet_xml = _insert_cell(sheet_xml, ref, f"<v>{val_s}</v>", type_attr="")
    return sheet_xml


def _set_cell_string(sheet_xml: str, ref: str, ss_idx: int) -> str:
    """Set a shared-string value in a cell (type t="s")."""
    result = _find_cell(sheet_xml, ref)
    if result:
        m, attrs, content, is_sc = result
        # Ensure t="s"
        if 't="s"' not in attrs:
            attrs = re.sub(r't="[^"]*"', "", attrs)
            attrs = attrs.rstrip() + ' t="s"'
        replacement = f'<c r="{ref}"{attrs}><v>{ss_idx}</v></c>'
        sheet_xml = sheet_xml[:m.start()] + replacement + sheet_xml[m.end():]
    else:
        sheet_xml = _insert_cell(
            sheet_xml, ref, f"<v>{ss_idx}</v>", type_attr=' t="s"'
        )
    return sheet_xml


def _insert_cell(sheet_xml: str, ref: str, value_xml: str, type_attr: str = "") -> str:
    """Insert a new <c> element into the correct <row>."""
    col_str = "".join(c for c in ref if c.isalpha())
    row_num = int("".join(c for c in ref if c.isdigit()))
    from openpyxl.utils import column_index_from_string
    col_num = column_index_from_string(col_str)

    row_pattern = re.compile(
        r'(<row r="' + str(row_num) + r'"[^>]*>)(.*?)(</row>)',
        re.DOTALL,
    )
    m = row_pattern.search(sheet_xml)
    if not m:
        return sheet_xml  # Row doesn't exist — skip for safety

    open_tag, row_content, close_tag = m.group(1), m.group(2), m.group(3)
    new_cell = f'<c r="{ref}"{type_attr}>{value_xml}</c>'

    # Insert in correct column order
    existing_cells = list(re.finditer(r'<c r="([A-Z]+)(\d+)"', row_content))
    insert_pos = len(row_content)
    for cell_m in existing_cells:
        from openpyxl.utils import column_index_from_string as c2i
        if c2i(cell_m.group(1)) > col_num:
            insert_pos = cell_m.start()
            break

    row_content = row_content[:insert_pos] + new_cell + row_content[insert_pos:]
    sheet_xml = sheet_xml[:m.start()] + open_tag + row_content + close_tag + sheet_xml[m.end():]
    return sheet_xml


# ── Rutas de partes OPC ───────────────────────────────────────────────────────

def _rels_path_for(part_path: str) -> str:
    """xl/drawings/drawing1.xml -> xl/drawings/_rels/drawing1.xml.rels"""
    folder, _, fname = part_path.rpartition("/")
    return f"{folder}/_rels/{fname}.rels"


def _resolve_target(source_part: str, target: str) -> str:
    """Resuelve un Target relativo de un .rels contra la parte que lo declara."""
    if target.startswith("/"):
        return target.lstrip("/")
    base = source_part.rpartition("/")[0]
    parts = base.split("/")
    for seg in target.split("/"):
        if seg == "..":
            parts.pop()
        elif seg not in ("", "."):
            parts.append(seg)
    return "/".join(parts)


def _relative_target(source_part: str, target_part: str) -> str:
    """Ruta relativa desde la carpeta de source_part hasta target_part."""
    src_dir = source_part.rpartition("/")[0].split("/")
    tgt = target_part.split("/")
    i = 0
    while i < len(src_dir) and i < len(tgt) - 1 and src_dir[i] == tgt[i]:
        i += 1
    return "../" * (len(src_dir) - i) + "/".join(tgt[i:])


# ── High-level writer ─────────────────────────────────────────────────────────

class XlsxZipWriter:
    """
    Modify specific cells in an XLSX without touching any other content.
    Usage:
        writer = XlsxZipWriter(xlsx_bytes)
        writer.set_number("Resumen", 10, 14, 136)        # Reporte No.
        writer.set_date("Resumen", 10, 19, date(2026,5,8))
        writer.set_text("Resumen", 18, 6, "Avance...")
        result_bytes = writer.save()
    """

    def __init__(self, xlsx_bytes: bytes):
        if hasattr(xlsx_bytes, "read"):
            xlsx_bytes = xlsx_bytes.read()
        self._original = xlsx_bytes
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zf:
            self._sheet_map = _get_sheet_map(zf)
            self._shared_strings, self._ss_xml = _load_shared_strings(zf)
            names = zf.namelist()

        # path → modified xml string (start as None = unchanged)
        self._modified_sheets: dict[str, str] = {}
        # path → bytes — sobrescribe partes existentes O agrega partes nuevas
        self._overrides: dict[str, bytes] = {}
        # nombres de partes que ya existen en el ZIP original
        self._original_names: set[str] = set(names)
        # marcar recálculo total al abrir
        self._full_recalc: bool = False
        # contador para ids de drawing únicos
        self._next_drawing_id: int = 90000

    def _get_sheet_xml(self, sheet_name: str) -> str:
        path = self._sheet_map[sheet_name]
        if path not in self._modified_sheets:
            with zipfile.ZipFile(io.BytesIO(self._original), "r") as zf:
                self._modified_sheets[path] = zf.read(path).decode("utf-8")
        return self._modified_sheets[path]

    def _save_sheet_xml(self, sheet_name: str, xml: str):
        path = self._sheet_map[sheet_name]
        self._modified_sheets[path] = xml

    def set_number(self, sheet: str, row: int, col: int, value: float | int):
        xml = self._get_sheet_xml(sheet)
        ref = _cell_ref(col, row)
        xml = _set_cell_number(xml, ref, value)
        self._save_sheet_xml(sheet, xml)

    def set_date(self, sheet: str, row: int, col: int, value):
        serial = to_excel_serial(value)
        self.set_number(sheet, row, col, serial)

    def set_text(self, sheet: str, row: int, col: int, text: str):
        if not text:
            return
        idx, self._ss_xml = _add_shared_string(self._shared_strings, self._ss_xml, text)
        xml = self._get_sheet_xml(sheet)
        ref = _cell_ref(col, row)
        xml = _set_cell_string(xml, ref, idx)
        self._save_sheet_xml(sheet, xml)

    def clear_cell(self, sheet: str, row: int, col: int):
        """
        Vacía una celda conservando su estilo (s="...") y sus bordes.

        Necesario porque la plantilla es SIEMPRE el informe del día anterior:
        sin esto, las actividades, observaciones y jornada de ayer se quedan
        pegadas en el informe de hoy.
        """
        xml = self._get_sheet_xml(sheet)
        ref = _cell_ref(col, row)
        result = _find_cell(xml, ref)
        if not result:
            return
        m, attrs, content, _ = result
        if "<f>" in content:
            return                      # nunca tocar celdas con fórmula
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        xml = xml[: m.start()] + f'<c r="{ref}"{attrs.rstrip()}/>' + xml[m.end() :]
        self._save_sheet_xml(sheet, xml)

    def find_date_col(self, sheet: str, target_date, date_row: int = 1) -> int | None:
        """Find the column index where date_row contains target_date."""
        if isinstance(target_date, datetime):
            target_date = target_date.date()
        serial = to_excel_serial(target_date)
        xml = self._get_sheet_xml(sheet)

        row_m = re.search(
            r'<row r="' + str(date_row) + r'"[^>]*>(.*?)</row>', xml, re.DOTALL
        )
        if not row_m:
            return None
        row_xml = row_m.group(1)

        pat = re.compile(
            r'<c r="([A-Z]+)' + str(date_row) + r'"[^>]*>.*?<v>(\d+)</v>',
            re.DOTALL,
        )
        for m in pat.finditer(row_xml):
            if int(m.group(2)) == serial:
                from openpyxl.utils import column_index_from_string
                return column_index_from_string(m.group(1))
        return None

    def get_number(self, sheet: str, row: int, col: int) -> float | None:
        """Read the cached numeric value of a cell (None if missing or non-numeric)."""
        xml = self._get_sheet_xml(sheet)
        ref = _cell_ref(col, row)
        result = _find_cell(xml, ref)
        if not result:
            return None
        v_m = re.search(r"<v>([^<]*)</v>", result[2])
        if not v_m or not v_m.group(1):
            return None
        try:
            return float(v_m.group(1))
        except ValueError:
            return None

    def add_to_number(self, sheet: str, row: int, col: int, delta: float):
        """Add delta to existing numeric cell value (for C.Control daily quantities)."""
        xml = self._get_sheet_xml(sheet)
        ref = _cell_ref(col, row)
        result = _find_cell(xml, ref)
        if result:
            _, attrs, content, is_sc = result
            v_match = re.search(r"<v>([^<]*)</v>", content)
            current = float(v_match.group(1)) if v_match and v_match.group(1) else 0.0
        else:
            current = 0.0
        xml = _set_cell_number(xml, ref, current + delta)
        self._save_sheet_xml(sheet, xml)

    # ── Partes arbitrarias del paquete ───────────────────────────────────────

    def read_part(self, path: str) -> bytes | None:
        """Lee una parte del paquete (ya modificada si aplica)."""
        if path in self._overrides:
            return self._overrides[path]
        if path not in self._original_names:
            return None
        with zipfile.ZipFile(io.BytesIO(self._original), "r") as zf:
            return zf.read(path)

    def add_part(self, path: str, data: bytes):
        """Agrega o reemplaza una parte del paquete (funciona con partes nuevas)."""
        self._overrides[path] = data

    # ── Recálculo ────────────────────────────────────────────────────────────

    def force_full_recalc(self, enabled: bool = True):
        """
        Marca el libro para que Excel recalcule TODAS las fórmulas al abrirlo.

        Crítico en Cenit: el informe entero se arma con HLOOKUP contra el
        consecutivo (Z6). Esta librería preserva las fórmulas pero solo puede
        actualizar el valor cacheado <v>. Sin fullCalcOnLoad, al cambiar el
        consecutivo el archivo se abriría mostrando los valores del día anterior.

        La plantilla de Cenit trae <calcPr calcId="191028" calcOnSave="0"/>,
        es decir, recálculo al guardar DESACTIVADO.
        """
        self._full_recalc = enabled

    def _apply_full_recalc(self, wb_xml: str) -> str:
        if not self._full_recalc:
            return wb_xml
        m = re.search(r"<calcPr[^>]*/>", wb_xml)
        if m:
            tag = m.group(0)
            tag = re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", tag)
            tag = re.sub(r'\s*calcOnSave="[^"]*"', "", tag)
            tag = tag[:-2].rstrip() + ' calcOnSave="1" fullCalcOnLoad="1"/>'
            return wb_xml[: m.start()] + tag + wb_xml[m.end() :]
        # No hay <calcPr>: insertarlo antes de </workbook>
        return wb_xml.replace(
            "</workbook>", '<calcPr calcId="191028" calcOnSave="1" fullCalcOnLoad="1"/></workbook>'
        )

    # ── Imágenes ─────────────────────────────────────────────────────────────

    def _drawing_path_for_sheet(self, sheet: str) -> str | None:
        """Devuelve la ruta del drawing asociado a la hoja, o None si no tiene."""
        sheet_path = self._sheet_map[sheet]                      # xl/worksheets/sheetN.xml
        rels_path = _rels_path_for(sheet_path)
        rels = self.read_part(rels_path)
        if rels is None:
            return None
        m = re.search(
            r'<Relationship[^>]+Type="[^"]*/drawing"[^>]+Target="([^"]+)"', rels.decode("utf-8")
        )
        if not m:
            m = re.search(
                r'<Relationship[^>]+Target="([^"]+)"[^>]+Type="[^"]*/drawing"', rels.decode("utf-8")
            )
        if not m:
            return None
        return _resolve_target(sheet_path, m.group(1))

    def _next_media_name(self, ext: str) -> str:
        """
        Siguiente nombre libre en xl/media.

        Compara ignorando la extensión: la plantilla de Cenit ya trae
        image2.jpeg e image3.jpeg, así que crear image2.png sería legal en OPC
        pero deja el paquete confuso y propenso a errores al depurar.
        """
        used = {
            n.rsplit("/", 1)[-1].rsplit(".", 1)[0]
            for n in (set(self._original_names) | set(self._overrides))
            if n.startswith("xl/media/")
        }
        n = 1
        while f"image{n}" in used:
            n += 1
        return f"xl/media/image{n}.{ext}"

    def remove_pictures_named(self, sheet: str, name_prefix: str):
        """
        Elimina del drawing los anclajes cuyo <xdr:cNvPr name="..."> empiece por
        name_prefix. Permite regenerar un informe sin duplicar fotos.
        """
        dpath = self._drawing_path_for_sheet(sheet)
        if not dpath:
            return
        raw = self.read_part(dpath)
        if raw is None:
            return
        xml = raw.decode("utf-8")

        out, pos, removed = [], 0, 0
        for m in re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b", xml):
            tag = m.group(1)
            close = f"</xdr:{tag}>"
            end = xml.find(close, m.start())
            if end == -1:
                continue
            end += len(close)
            if m.start() < pos:      # ya consumido por un anchor anterior
                continue
            block = xml[m.start() : end]
            nm = re.search(r'<xdr:cNvPr[^>]*name="([^"]*)"', block)
            if nm and nm.group(1).startswith(name_prefix):
                out.append(xml[pos : m.start()])
                removed += 1
            else:
                out.append(xml[pos : end])
            pos = end
        out.append(xml[pos:])

        if removed:
            self.add_part(dpath, "".join(out).encode("utf-8"))

    def insert_picture(
        self,
        sheet: str,
        image_bytes: bytes,
        from_col: int,
        from_row: int,
        to_col: int,
        to_row: int,
        name: str = "PCC_FOTO",
        from_off: tuple[int, int] = (0, 0),
        to_off: tuple[int, int] = (0, 0),
    ):
        """
        Inserta una imagen anclada al rectángulo [from_col,from_row] .. [to_col,to_row].

        Índices de fila/columna en base 0 (col B = 1, fila 12 = 11), igual que
        el esquema xdr. El anclaje es twoCellAnchor editAs="oneCell": Excel calcula
        la geometría a partir de las celdas, así que no hace falta pasar EMU.

        La imagen se normaliza a PNG. Para que no se deforme, el llamador debe
        entregarla ya ajustada a la relación de aspecto de la caja
        (ver cenit_report.fit_image_to_box).
        """
        dpath = self._drawing_path_for_sheet(sheet)
        if not dpath:
            raise RuntimeError(
                f"La hoja '{sheet}' no tiene drawing asociado; no se puede insertar imagen."
            )

        # 1. Parte de media
        media_path = self._next_media_name("png")
        self.add_part(media_path, image_bytes)

        # 2. Relación en el .rels del drawing
        rels_path = _rels_path_for(dpath)
        raw_rels = self.read_part(rels_path)
        if raw_rels is None:
            rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )
        else:
            rels_xml = raw_rels.decode("utf-8")

        used_ids = {int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)}
        rid_num = max(used_ids) + 1 if used_ids else 1
        rid = f"rId{rid_num}"
        target = _relative_target(dpath, media_path)
        new_rel = (
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="{target}"/>'
        )
        rels_xml = rels_xml.replace("</Relationships>", new_rel + "</Relationships>")
        if "</Relationships>" not in rels_xml:  # era self-closing
            rels_xml = rels_xml.replace("/>", ">" + new_rel + "</Relationships>", 1)
        self.add_part(rels_path, rels_xml.encode("utf-8"))

        # 3. Anclaje en el drawing
        drawing_xml = self.read_part(dpath).decode("utf-8")
        self._next_drawing_id += 1
        anchor = (
            '<xdr:twoCellAnchor editAs="oneCell">'
            f"<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_off[0]}</xdr:colOff>"
            f"<xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_off[1]}</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_off[0]}</xdr:colOff>"
            f"<xdr:row>{to_row}</xdr:row><xdr:rowOff>{to_off[1]}</xdr:rowOff></xdr:to>"
            "<xdr:pic>"
            "<xdr:nvPicPr>"
            f'<xdr:cNvPr id="{self._next_drawing_id}" name="{_escape(name)}"/>'
            '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
            "</xdr:nvPicPr>"
            "<xdr:blipFill>"
            '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'r:embed="{rid}"/>'
            "<a:stretch><a:fillRect/></a:stretch>"
            "</xdr:blipFill>"
            "<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
            "</xdr:pic>"
            "<xdr:clientData/>"
            "</xdr:twoCellAnchor>"
        )
        drawing_xml = drawing_xml.replace("</xdr:wsDr>", anchor + "</xdr:wsDr>")
        self.add_part(dpath, drawing_xml.encode("utf-8"))

    # ── Guardado ─────────────────────────────────────────────────────────────

    def save(self) -> bytes:
        output = io.BytesIO()
        written: set[str] = set()

        with zipfile.ZipFile(io.BytesIO(self._original), "r") as zin:
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
                # Partes existentes (respetando orden y metadatos originales)
                for item in zin.infolist():
                    name = item.filename
                    if name == "xl/sharedStrings.xml":
                        data = self._ss_xml.encode("utf-8")
                    elif name == "xl/workbook.xml":
                        data = self._apply_full_recalc(
                            zin.read(name).decode("utf-8")
                        ).encode("utf-8")
                    elif name in self._modified_sheets:
                        data = self._modified_sheets[name].encode("utf-8")
                    elif name in self._overrides:
                        data = self._overrides[name]
                    else:
                        data = zin.read(name)
                    zout.writestr(item, data)
                    written.add(name)

                # Partes nuevas (media, rels o drawings que no existían)
                for name, data in self._overrides.items():
                    if name not in written:
                        zout.writestr(name, data)

        output.seek(0)
        return output.read()
