"""
Copia las credenciales desde un archivo suelto hacia .streamlit/secrets.toml.

Acepta .rtf (TextEdit), .txt o .toml. Normaliza las comillas tipográficas que
mete TextEdit («" "» y «' '»), que romperían el TOML, y escribe el archivo con
formato válido.

Nunca imprime los valores: solo confirma cuáles encontró, enmascarados.

    python3 scripts/importar_credenciales.py "<archivo con las credenciales>"
"""
import re
import shutil
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SECRETS = RAIZ / ".streamlit" / "secrets.toml"

CLAVES = [
    "fastfield_email",
    "fastfield_password",
    "fastfield_org_id",
    "fastfield_subscription_key",
]

# Comillas tipográficas y otros caracteres que TextEdit introduce
COMILLAS = str.maketrans({
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u00ab": '"', "\u00bb": '"',
    "\u2018": "'", "\u2019": "'", "\u00a0": " ",
})


def a_texto_plano(ruta: Path) -> str:
    if ruta.suffix.lower() == ".rtf":
        # textutil es nativo de macOS y convierte RTF de forma fiable
        try:
            return subprocess.run(
                ["textutil", "-convert", "txt", "-stdout", str(ruta)],
                capture_output=True, text=True, check=True,
            ).stdout
        except (FileNotFoundError, subprocess.CalledProcessError):
            pass  # sin textutil, cae al limpiador propio
        raw = ruta.read_text(encoding="utf-8", errors="replace")
        t = re.sub(r"\{\\\*?\\[^{}]*\}", "", raw)
        t = re.sub(r"\\par[d]?\b", "\n", t)
        t = re.sub(
            r"\\'([0-9a-fA-F]{2})",
            lambda m: bytes.fromhex(m.group(1)).decode("cp1252", "replace"), t,
        )
        t = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", t)
        return t.replace("{", "").replace("}", "")
    return ruta.read_text(encoding="utf-8", errors="replace")


def extraer(texto: str) -> dict[str, str]:
    texto = texto.translate(COMILLAS)
    valores = {}
    for clave in CLAVES:
        m = re.search(
            rf"^\s*{re.escape(clave)}\s*[:=]\s*(.*)$", texto, re.M | re.I
        )
        if not m:
            continue
        v = m.group(1).strip()
        v = re.sub(r"\s*#.*$", "", v).strip()          # comentario al final
        # Anotaciones tipo "← aquí" / "<- poner aquí" que se copian sin querer
        v = re.split(r"\s*(?:\u2190|<-|<=)\s*", v)[0].strip()
        v = v.rstrip("\\").strip()                     # restos del RTF
        if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
            v = v[1:-1]                                # comillas envolventes
        valores[clave] = v.strip()
    return valores


def revisar(valores: dict[str, str]) -> list[str]:
    """Problemas de forma que harían fallar la autenticación en silencio."""
    problemas = []
    for clave, v in valores.items():
        if not v:
            continue
        if '"' in v or "'" in v:
            problemas.append(
                f"{clave}: el valor contiene comillas sueltas. Suele pasar cuando "
                f"TextEdit encierra solo parte del texto. Escribe la línea SIN "
                f"comillas: {clave} = elvalor"
            )
        if v != v.strip():
            problemas.append(f"{clave}: sobran espacios al inicio o al final.")
        if "\u2190" in v or "<-" in v:
            problemas.append(f"{clave}: quedó texto de anotación pegado al valor.")
    return problemas


def enmascarar(v: str) -> str:
    if not v:
        return "(vacío)"
    if len(v) <= 4:
        return "*" * len(v) + f"  ({len(v)} car.)"
    return f"{v[0]}{'*' * (len(v) - 2)}{v[-1]}  ({len(v)} car.)"


def escapar_toml(v: str) -> str:
    return v.replace("\\", "\\\\").replace('"', '\\"')


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    origen = Path(sys.argv[1])
    if not origen.exists():
        sys.exit(f"No existe: {origen}")

    valores = extraer(a_texto_plano(origen))
    if not valores:
        sys.exit(
            f"No encontré ninguna de estas claves en {origen.name}:\n"
            + "\n".join(f"  {c}" for c in CLAVES)
            + "\n\nEl archivo debe tener líneas del estilo:\n"
              '  fastfield_password = "..."'
        )

    print(f"\nOrigen: {origen}")
    print("\nEncontrado (valores enmascarados):")
    for clave in CLAVES:
        v = valores.get(clave)
        estado = enmascarar(v) if v is not None else "no aparece en el archivo"
        print(f"  {clave:28s} {estado}")

    problemas = revisar(valores)
    if problemas:
        print("\nPROBLEMAS DE FORMA — corrige el archivo y vuelve a correr esto:")
        for x in problemas:
            print(f"  · {x}")
        print("\nNo escribí nada en secrets.toml.")
        return 1

    if not valores.get("fastfield_email") or not valores.get("fastfield_password"):
        print("\nFaltan email o contraseña; sin eso la descarga de fotos no funciona.")

    if SECRETS.exists():
        shutil.copy2(SECRETS, SECRETS.with_suffix(".toml.bak"))
        print(f"\nRespaldo del anterior: {SECRETS.name}.bak")

    SECRETS.parent.mkdir(parents=True, exist_ok=True)
    lineas = [
        "# Credenciales — NO se sube al repositorio (está en .gitignore)",
        "# Generado por scripts/importar_credenciales.py",
        "",
    ]
    for clave in CLAVES:
        lineas.append(f'{clave} = "{escapar_toml(valores.get(clave, ""))}"')
    SECRETS.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print(f"Escrito: {SECRETS}")

    # Validar que el TOML quedó bien formado
    import tomllib
    try:
        with open(SECRETS, "rb") as f:
            leido = tomllib.load(f)
    except Exception as e:
        sys.exit(f"\nERROR: el TOML quedó mal formado: {e}")
    print(f"TOML válido — {len(leido)} clave(s).")

    print("\nAhora comprueba que sirven:")
    print("  python3 scripts/probar_fastfield.py "
          "~/Downloads/e8ff3eef-0884-4b5a-ba71-308c724b0f0e.xlsx")
    print("\nY borra el archivo suelto cuando confirmes que todo funciona:")
    print(f'  rm "{origen}"')


if __name__ == "__main__":
    sys.exit(main() or 0)
