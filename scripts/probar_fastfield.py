"""
Verifica las credenciales de FastField sin exponerlas.

Autentica contra la API v3 y, si le pasas un submission, intenta descargar una
de sus fotos. Nunca imprime la contraseña ni el token.

    python3 scripts/probar_fastfield.py [submission.xlsx]
"""
import io
import sys
import tomllib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

RAIZ = Path(__file__).resolve().parent.parent
SECRETS = RAIZ / ".streamlit" / "secrets.toml"


def cargar_secretos() -> dict:
    if not SECRETS.exists():
        sys.exit(f"No existe {SECRETS}\nCrea el archivo y vuelve a correr esto.")
    with open(SECRETS, "rb") as f:
        return tomllib.load(f)


def main():
    s = cargar_secretos()
    email = s.get("fastfield_email", "")
    passwd = s.get("fastfield_password", "")
    org = s.get("fastfield_org_id", "")
    subkey = s.get("fastfield_subscription_key", "")

    print(f"\nArchivo: {SECRETS}")
    print(f"  fastfield_email             {'OK  ' + email if email else 'VACÍO'}")
    print(f"  fastfield_password          {'OK (' + str(len(passwd)) + ' caracteres)' if passwd else 'VACÍO'}")
    print(f"  fastfield_org_id            {org or '(vacío — normal)'}")
    print(f"  fastfield_subscription_key  {'definida' if subkey else '(vacío — normal)'}")

    if not email or not passwd:
        sys.exit("\nFaltan email o contraseña. Llénalos en el archivo y repite.")

    # ── Autenticación ───────────────────────────────────────────────────────
    from utils.fastfield_api import authenticate, get_photo_bytes

    print("\nAutenticando contra https://api.fastfieldforms.com/services/v3 ...")
    try:
        token = authenticate(email, passwd, org, subkey)
    except Exception as e:
        msg = str(e)
        print(f"  FALLÓ: {msg}")

        if "subscription key" in msg.lower():
            print("\n  CAUSA: la API está detrás de Azure API Management y exige una")
            print("  llave de suscripción además del usuario y la contraseña.")
            print("\n  Dónde conseguirla:")
            print("   1. Streamlit Cloud -> app de informes de Ecopetrol -> Settings ->")
            print("      Secrets. Ahí ya está configurada como fastfield_subscription_key")
            print("      (el secrets.toml local de informes-eco solo tiene app_password).")
            print("   2. O en la consola web de FastField, en la sección de API de la")
            print("      cuenta u organización.")
            print("\n  Luego agrégala al archivo de credenciales, SIN comillas:")
            print("      fastfield_subscription_key = LA_LLAVE")
            print("  y repite el importador y esta prueba.")
        elif "401" in msg or "403" in msg:
            print("\n  Qué revisar:")
            print("   · que el email y la contraseña sean los de la consola web")
            print("   · si la cuenta pertenece a varias organizaciones, llena fastfield_org_id")
        else:
            print("\n  Qué revisar:")
            print("   · conexión a internet o proxy de la red")
            print("   · que el servicio de FastField esté disponible")
        return 1
    print(f"  OK — sesión iniciada (token de {len(token)} caracteres, no se imprime)")

    # ── Descarga de una foto ────────────────────────────────────────────────
    if len(sys.argv) < 2:
        print("\nPara probar además la descarga de fotos, pásame un submission:")
        print("    python3 scripts/probar_fastfield.py <submission.xlsx>")
        return 0

    ruta = Path(sys.argv[1])
    if not ruta.exists():
        print(f"\nNo existe el submission: {ruta}")
        return 1

    from utils.fastfield import parse_submission
    cap = parse_submission(io.BytesIO(ruta.read_bytes()))
    fotos = cap.get("fotos") or []
    if not fotos:
        print(f"\nEl submission {ruta.name} no trae fotos.")
        return 0

    print(f"\nEl submission trae {len(fotos)} foto(s). Probando la primera...")
    nombre = fotos[0]["filename"]
    print(f"  archivo: {nombre[:60]}...")
    blob = get_photo_bytes(nombre, token, subkey)
    if not blob:
        print("  FALLÓ la descarga.")
        print("   · el usuario puede no tener permiso sobre los medios de ese formulario")
        print("   · o el archivo ya no está disponible en FastField")
        return 1

    print(f"  OK — descargados {len(blob) / 1024:.0f} KB")
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(blob))
        print(f"  imagen válida: {img.format} {img.width}x{img.height} px")
    except Exception as e:
        print(f"  AVISO: se descargó pero no se pudo abrir como imagen: {e}")
        return 1

    print("\nTodo listo. La app ya puede descargar las fotos sola.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
