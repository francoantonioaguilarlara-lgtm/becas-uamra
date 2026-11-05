"""
Conexión con la base de datos en la nube (Hoja de Google + Apps Script, ver nube/Codigo.gs).

La configuración (dirección y contraseña) va en .streamlit/secrets.toml:

    [nube]
    url = "https://script.google.com/macros/s/XXXX/exec"
    token = "la-misma-clave-de-Codigo.gs"
    admin_password = "clave-para-ver-graficas"     # opcional
"""

from __future__ import annotations

import json
import tomllib
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd

SECRETS = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"
CAMPOS = ["Clave_Alumno", "Carrera", "Semestre", "Estatus_Residencia", "Procedencia", "Fuente_Ingreso",
          "Ingreso_Semanal", "Gasto_Transporte", "Gasto_Alimentacion", "Otros_Gastos_Acad"]


class ErrorNube(Exception):
    """Error al hablar con la nube (sin internet, contraseña incorrecta, etc.)."""


def leer_config(ruta: Path = SECRETS) -> dict | None:
    """Lee la sección [nube] de secrets.toml (para scripts fuera de Streamlit)."""
    if not ruta.exists():
        return None
    cfg = tomllib.loads(ruta.read_text(encoding="utf-8-sig")).get("nube")  # -sig: tolera la marca BOM
    return cfg if cfg and cfg.get("url") and cfg.get("token") else None


def _pedir(req: urllib.request.Request) -> dict:
    base = req.full_url.split("?")[0]
    req.add_header("User-Agent", "Mozilla/5.0 (BecasUAMRA)")
    try:
        with urllib.request.urlopen(req, timeout=40) as resp:
            texto = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        donde = urllib.parse.urlparse(e.geturl() or base).netloc
        pista = (" Revise en Secrets que la url esté completa, entre comillas y sin espacios."
                 if e.code == 404 else "")
        raise ErrorNube(f"No se pudo conectar con la nube: HTTP {e.code} desde {donde}. "
                        f"URL configurada: …{base[-12:]} ({len(base)} caracteres).{pista}") from e
    except Exception as e:
        raise ErrorNube(f"No se pudo conectar con la nube: {e}") from e
    try:
        datos = json.loads(texto)
    except json.JSONDecodeError as e:
        raise ErrorNube("La nube respondió algo inesperado. Revise que la URL sea la de la "
                        "'Aplicación web' (termina en /exec) y que el acceso sea 'Cualquier usuario'.") from e
    if not datos.get("ok"):
        raise ErrorNube(datos.get("error", "Error desconocido en la nube."))
    return datos


def enviar_registro(r: dict, url: str, token: str) -> int:
    """Sube un alumno. Devuelve el total de alumnos en la nube."""
    cuerpo = {c: r[c] for c in CAMPOS}
    cuerpo["Clave_Alumno"] = str(cuerpo["Clave_Alumno"]).zfill(6)
    cuerpo["token"] = token
    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode("utf-8"), method="POST",
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    return int(_pedir(req).get("total", 0))


def descargar(url: str, token: str) -> pd.DataFrame:
    """Todos los alumnos de la nube como tabla (mismas columnas que Base_Datos A-J + fecha)."""
    q = urllib.parse.urlencode({"token": token})
    datos = _pedir(urllib.request.Request(f"{url}?{q}"))
    df = pd.DataFrame(datos.get("alumnos", []), columns=CAMPOS + ["Fecha_Registro"])
    if df.empty:
        return df
    df["Clave_Alumno"] = df["Clave_Alumno"].astype(str).str.zfill(6)
    for c in ["Semestre", "Ingreso_Semanal", "Gasto_Transporte", "Gasto_Alimentacion", "Otros_Gastos_Acad"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Fecha_Registro"] = pd.to_datetime(df["Fecha_Registro"], errors="coerce")
    return df


def sincronizar_a_excel(ruta_excel: Path, url: str, token: str) -> dict:
    """Trae de la nube los alumnos que todavía no están en el Excel de la laptop y los agrega."""
    import excel_db as db

    nube = descargar(url, token)
    local = set(db.leer_registros(ruta_excel)["Clave_Alumno"])
    nuevos = nube[~nube["Clave_Alumno"].isin(local)] if not nube.empty else nube
    registros = []
    for r in nuevos.to_dict("records"):
        fecha = r.pop("Fecha_Registro")
        r["Fecha_Registro"] = fecha.to_pydatetime() if pd.notna(fecha) else datetime.now()
        registros.append(r)
    resultado = db.agregar_varios(registros, ruta_excel) if registros else {"agregados": 0, "errores": []}
    resultado["en_nube"] = len(nube)
    resultado["ya_estaban"] = len(nube) - len(nuevos)
    return resultado
