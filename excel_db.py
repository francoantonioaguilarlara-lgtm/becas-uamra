"""
Conexión entre el programa en Python y el libro de Excel Programa_Becas_UAMRA.xlsm.

- Lee los parámetros del modelo desde la hoja "Parametros" (openpyxl, solo lectura).
- Lee los alumnos registrados desde la hoja "Base_Datos" (filas 5 a 504).
- Agrega alumnos nuevos A TRAVÉS DE MICROSOFT EXCEL (automatización COM), escribiendo
  las columnas A-J (captura) y las columnas K-Y como FÓRMULAS. Así se conservan las
  macros, las gráficas y el formato del libro, y todas las hojas se recalculan solas.
- Repite en Python los mismos cálculos para poder graficar sin abrir Excel.
"""

from __future__ import annotations

import math
import random
import shutil
import warnings
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

# (excel_com solo se importa al escribir: en el servidor en línea no hay Microsoft Excel)

# openpyxl solo lee el libro; avisa de extensiones (validaciones, formatos) que no necesita.
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

CARPETA = Path(__file__).resolve().parent / "datos"
EXCEL_REAL = CARPETA / "Programa_Becas_UAMRA.xlsm"
EXCEL_DEMO = CARPETA / "Programa_Becas_DEMO.xlsm"
# Libro sin alumnos (solo parámetros y listas): es el que usa la página en línea.
PLANTILLA = CARPETA.parent / "herramientas" / "plantilla_original.xlsx"

HOJA_BD = "Base_Datos"
FILA_INICIO = 5
FILA_FIN = 504

COLUMNAS_CAPTURA = [
    "Clave_Alumno", "Carrera", "Semestre", "Estatus_Residencia", "Procedencia",
    "Fuente_Ingreso", "Ingreso_Semanal", "Gasto_Transporte", "Gasto_Alimentacion",
    "Otros_Gastos_Acad",
]

P = "Parametros!"


class ErrorRegistro(Exception):
    """Error de validación o de escritura que se muestra al usuario."""


# ----------------------------------------------------------------------------
# Lectura
# ----------------------------------------------------------------------------
def leer_parametros(ruta: Path = EXCEL_REAL) -> dict:
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=False)
    ws = wb["Parametros"]
    v = lambda celda: ws[celda].value  # noqa: E731

    def tabla(col_nombre, col_valor=None):
        datos = {}
        for fila in range(20, 35):
            nombre = ws[f"{col_nombre}{fila}"].value
            if nombre in (None, ""):
                continue
            datos[nombre] = ws[f"{col_valor}{fila}"].value if col_valor else None
        return datos

    params = {
        "saldo_critico": float(v("C5")),
        "icf_alta": float(v("C6")),
        "icf_media": float(v("C7")),
        "pct_transporte": float(v("C8")),
        "pct_alimentacion": float(v("C9")),
        "saldo_referencia": float(v("C10")),
        "puntaje_minimo": float(v("C11")),
        "pts_icf": float(v("C12")),
        "pts_geo": float(v("C13")),
        "pts_saldo": float(v("C14")),
        "pts_fuente": float(v("C15")),
        "geo": {k: float(x or 0) for k, x in tabla("B", "C").items()},
        "fuente": {k: float(x or 0) for k, x in tabla("E", "F").items()},
        "carreras": list(tabla("H").keys()),
        "semestres": [int(s) for s in tabla("I").keys()],
        "residencias": list(tabla("J").keys()),
    }
    wb.close()
    return params


def leer_registros(ruta: Path = EXCEL_REAL) -> pd.DataFrame:
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=False)
    ws = wb[HOJA_BD]
    filas = []
    for fila in ws.iter_rows(min_row=FILA_INICIO, max_row=FILA_FIN, max_col=26, values_only=True):
        if fila[0] in (None, ""):
            continue
        registro = dict(zip(COLUMNAS_CAPTURA, fila[:10]))
        registro["Fecha_Registro"] = fila[25]
        filas.append(registro)
    wb.close()
    df = pd.DataFrame(filas, columns=COLUMNAS_CAPTURA + ["Fecha_Registro"])
    if df.empty:
        return df
    df["Clave_Alumno"] = df["Clave_Alumno"].astype(str).str.zfill(6)
    for c in ["Semestre", "Ingreso_Semanal", "Gasto_Transporte", "Gasto_Alimentacion", "Otros_Gastos_Acad"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# ----------------------------------------------------------------------------
# Cálculos (los mismos que hacen las fórmulas del Excel)
# ----------------------------------------------------------------------------
def diagnosticar(r: dict, p: dict) -> dict:
    ingreso = float(r["Ingreso_Semanal"])
    transp = float(r["Gasto_Transporte"])
    alim = float(r["Gasto_Alimentacion"])
    otros = float(r["Otros_Gastos_Acad"])

    gasto = transp + alim + otros
    saldo = ingreso - gasto
    icf = (transp + alim) / ingreso if ingreso > 0 else 0

    if saldo <= p["saldo_critico"] or icf >= p["icf_alta"]:
        nivel = "ALTA PRIORIDAD"
    elif icf >= p["icf_media"]:
        nivel = "MEDIA PRIORIDAD"
    else:
        nivel = "BAJA PRIORIDAD"

    if saldo <= p["saldo_critico"] and icf >= p["icf_alta"]:
        beca = "Manutención Integral"
    elif ingreso > 0 and transp / ingreso >= p["pct_transporte"]:
        beca = "Beca Transporte"
    elif ingreso > 0 and alim / ingreso >= p["pct_alimentacion"]:
        beca = "Beca Alimentaria"
    else:
        beca = "Sin Asignación Urgente"

    pts_icf = min(1, icf / p["icf_alta"]) * p["pts_icf"]
    pts_geo = p["geo"].get(r["Procedencia"], 0) * p["pts_geo"]
    if saldo <= p["saldo_critico"]:
        pts_saldo = p["pts_saldo"]
    else:
        rango = p["saldo_referencia"] - p["saldo_critico"]
        pts_saldo = max(0, p["pts_saldo"] * (1 - (saldo - p["saldo_critico"]) / rango)) if rango else 0
    pts_fuente = p["fuente"].get(r["Fuente_Ingreso"], 0) * p["pts_fuente"]
    puntaje = pts_icf + pts_geo + pts_saldo + pts_fuente

    requiere = "SÍ" if (puntaje >= p["puntaje_minimo"] or nivel == "ALTA PRIORIDAD") else "NO"

    return {
        "Gasto_Total_Semanal": gasto,
        "Saldo_Ahorro_Semanal": saldo,
        "Gasto_Transp_%": transp / gasto if gasto > 0 else 0,
        "Gasto_Alim_%": alim / gasto if gasto > 0 else 0,
        "Indice_Carga_Fija_%": icf,
        "Nivel_Vulnerabilidad": nivel,
        "Beca_Sugerida": beca,
        "Pts_ICF": pts_icf,
        "Pts_Geografia": pts_geo,
        "Pts_Saldo": pts_saldo,
        "Pts_Fuente": pts_fuente,
        "Puntaje_Total": puntaje,
        "Requiere_Beca": requiere,
    }


def calcular(df: pd.DataFrame, p: dict) -> pd.DataFrame:
    if df.empty:
        return df
    calc = pd.DataFrame([diagnosticar(r, p) for r in df.to_dict("records")], index=df.index)
    return pd.concat([df, calc], axis=1)


# ----------------------------------------------------------------------------
# Validación (mismas reglas que la hoja Formulario_Celular)
# ----------------------------------------------------------------------------
def validar(r: dict, p: dict, claves_existentes: set[str]) -> list[str]:
    errores = []
    clave = str(r["Clave_Alumno"]).strip()
    if not (clave.isdigit() and len(clave) == 6):
        errores.append("La clave universitaria debe tener exactamente 6 dígitos numéricos.")
    elif clave in claves_existentes:
        errores.append(f"La clave {clave} ya está registrada.")
    if r["Carrera"] not in p["carreras"]:
        errores.append("Seleccione una carrera válida.")
    if not (1 <= int(r["Semestre"]) <= 12):
        errores.append("El semestre debe ser un número de 1 a 12.")
    if r["Estatus_Residencia"] not in p["residencias"]:
        errores.append("Seleccione Local o Foráneo.")
    if r["Procedencia"] not in p["geo"]:
        errores.append("Seleccione un municipio válido.")
    if r["Fuente_Ingreso"] not in p["fuente"]:
        errores.append("Seleccione una fuente de ingresos válida.")
    if r["Ingreso_Semanal"] <= 0:
        errores.append("El ingreso semanal debe ser mayor a cero.")
    for c in ["Gasto_Transporte", "Gasto_Alimentacion", "Otros_Gastos_Acad"]:
        if r[c] < 0:
            errores.append("Los gastos no pueden ser negativos.")
            break
    if r["Estatus_Residencia"] == "Local" and r["Procedencia"] != "Matehuala":
        errores.append("Un alumno Local debe tener procedencia Matehuala.")
    if r["Estatus_Residencia"] == "Foráneo" and r["Procedencia"] == "Matehuala":
        errores.append("Un alumno Foráneo no puede tener procedencia Matehuala "
                       "(si renta en Matehuala elija 'Matehuala (Renta)').")
    gasto = r["Gasto_Transporte"] + r["Gasto_Alimentacion"] + r["Otros_Gastos_Acad"]
    if r["Ingreso_Semanal"] > 0 and gasto > 2 * r["Ingreso_Semanal"]:
        errores.append("El gasto total supera 2 veces el ingreso: revise si hay un error de captura.")
    return errores


# ----------------------------------------------------------------------------
# Escritura
# ----------------------------------------------------------------------------
def _formulas(n: int) -> dict[str, str]:
    """Fórmulas de las columnas K-Y para la fila n (idénticas a la lógica del libro)."""
    return {
        "K": f"=H{n}+I{n}+J{n}",
        "L": f"=G{n}-K{n}",
        "M": f"=IF(K{n}>0,H{n}/K{n},0)",
        "N": f"=IF(K{n}>0,I{n}/K{n},0)",
        "O": f"=IF(G{n}>0,(H{n}+I{n})/G{n},0)",
        "P": f'=IF(OR(L{n}<={P}$C$5,O{n}>={P}$C$6),"ALTA PRIORIDAD",'
             f'IF(O{n}>={P}$C$7,"MEDIA PRIORIDAD","BAJA PRIORIDAD"))',
        "Q": f'=IF(AND(L{n}<={P}$C$5,O{n}>={P}$C$6),"Manutención Integral",'
             f'IF(H{n}/G{n}>={P}$C$8,"Beca Transporte",'
             f'IF(I{n}/G{n}>={P}$C$9,"Beca Alimentaria","Sin Asignación Urgente")))',
        "R": f"=MIN(1,O{n}/{P}$C$6)*{P}$C$12",
        "S": f"=IFERROR(VLOOKUP(E{n},{P}$B$20:$C$34,2,FALSE),0)*{P}$C$13",
        "T": f"=IF(L{n}<={P}$C$5,{P}$C$14,"
             f"IFERROR(MAX(0,{P}$C$14*(1-(L{n}-{P}$C$5)/({P}$C$10-{P}$C$5))),0))",
        "U": f"=IFERROR(VLOOKUP(F{n},{P}$E$20:$F$34,2,FALSE),0)*{P}$C$15",
        "V": f"=R{n}+S{n}+T{n}+U{n}",
        "W": f'=IF(OR(V{n}>={P}$C$11,P{n}="ALTA PRIORIDAD"),"SÍ","NO")',
        # Clave para ordenar el padrón: puntaje con desempate por fila (solo candidatos).
        "X": f'=IF(W{n}="SÍ",V{n}-ROW()/1000000,"")',
        "Y": f'=IF(X{n}="","",COUNTIF($X${FILA_INICIO}:$X${FILA_FIN},">"&X{n})+1)',
    }


def _serial_excel(fecha: datetime) -> float:
    return (fecha - datetime(1899, 12, 30)).total_seconds() / 86400


def _valores_fila(r: dict) -> tuple:
    return (
        str(r["Clave_Alumno"]).zfill(6), r["Carrera"], int(r["Semestre"]), r["Estatus_Residencia"],
        r["Procedencia"], r["Fuente_Ingreso"], float(r["Ingreso_Semanal"]), float(r["Gasto_Transporte"]),
        float(r["Gasto_Alimentacion"]), float(r["Otros_Gastos_Acad"]),
    )


def _escribir_fila(ws, n: int, r: dict, fecha: datetime) -> None:
    ws.Range(f"A{n}").NumberFormat = "@"
    ws.Range(f"A{n}:J{n}").Value = (_valores_fila(r),)
    ws.Range(f"K{n}:Y{n}").Formula = (tuple(_formulas(n).values()),)
    ws.Range(f"Z{n}").Value = _serial_excel(fecha)
    ws.Range(f"Z{n}").NumberFormat = "dd/mm/yyyy hh:mm"


def _con_libro(ruta: Path, accion) -> None:
    """Abre el libro en Microsoft Excel, ejecuta accion(hoja Base_Datos) y guarda."""
    from excel_com import abrir_excel

    try:
        with abrir_excel() as xl:
            wb, abierto_por_mi = None, False
            for i in range(1, xl.Workbooks.Count + 1):
                if Path(xl.Workbooks(i).FullName).resolve() == ruta.resolve():
                    wb = xl.Workbooks(i)
            if wb is None:
                wb = xl.Workbooks.Open(str(ruta))
                abierto_por_mi = True
            if wb.ReadOnly:
                if abierto_por_mi:
                    wb.Close(False)
                raise ErrorRegistro("El archivo de Excel está abierto en otra ventana en modo lectura. "
                                    "Ciérrelo e intente de nuevo.")
            accion(wb.Worksheets(HOJA_BD))
            xl.Calculate()
            wb.Save()
            if abierto_por_mi:
                wb.Close(False)
    except ErrorRegistro:
        raise
    except Exception as e:  # errores de Excel / COM
        raise ErrorRegistro(f"No se pudo escribir en Excel: {e}") from e


def agregar_registro(r: dict, ruta: Path = EXCEL_REAL) -> int:
    p = leer_parametros(ruta)
    df = leer_registros(ruta)
    errores = validar(r, p, set(df["Clave_Alumno"]))
    if errores:
        raise ErrorRegistro("\n".join(errores))
    fila = FILA_INICIO + len(df)
    if fila > FILA_FIN:
        raise ErrorRegistro("La base de datos está llena (500 alumnos).")
    _con_libro(ruta, lambda ws: _escribir_fila(ws, fila, r, datetime.now().replace(microsecond=0)))
    return fila


def agregar_varios(registros: list[dict], ruta: Path = EXCEL_REAL) -> dict:
    """Agrega varios alumnos en una sola apertura de Excel (para sincronizar con la nube).

    Cada registro puede traer "Fecha_Registro". Los que no pasan la validación se reportan.
    """
    p = leer_parametros(ruta)
    df = leer_registros(ruta)
    claves = set(df["Clave_Alumno"])
    validos, errores = [], []
    for r in registros:
        problemas = validar(r, p, claves)
        if problemas:
            errores.append(f"{r['Clave_Alumno']}: " + " ".join(problemas))
            continue
        claves.add(str(r["Clave_Alumno"]).zfill(6))
        validos.append(r)
    inicio = FILA_INICIO + len(df)
    if inicio + len(validos) - 1 > FILA_FIN:
        raise ErrorRegistro("No caben todos los alumnos: la base admite 500.")

    def accion(ws):
        for i, r in enumerate(validos):
            fecha = r.get("Fecha_Registro") or datetime.now()
            _escribir_fila(ws, inicio + i, r, fecha)

    if validos:
        _con_libro(ruta, accion)
    return {"agregados": len(validos), "errores": errores}


def eliminar_registro(clave: str, ruta: Path = EXCEL_REAL) -> None:
    """Borra un alumno y recorre los demás hacia arriba para no dejar huecos."""
    df = leer_registros(ruta)
    if clave not in set(df["Clave_Alumno"]):
        raise ErrorRegistro(f"No existe la clave {clave}.")
    restantes = df[df["Clave_Alumno"] != clave]

    def accion(ws):
        ws.Range(f"A{FILA_INICIO}:Z{FILA_INICIO + len(df) - 1}").ClearContents()
        for i, r in enumerate(restantes.to_dict("records")):
            fecha = r["Fecha_Registro"] if isinstance(r["Fecha_Registro"], datetime) else datetime.now()
            _escribir_fila(ws, FILA_INICIO + i, r, fecha)

    _con_libro(ruta, accion)


# ----------------------------------------------------------------------------
# Datos ficticios para practicar / presentar
# ----------------------------------------------------------------------------
def alumnos_ficticios(n: int, p: dict, semilla: int = 7) -> list[dict]:
    rnd = random.Random(semilla)
    foraneos = [m for m in p["geo"] if m != "Matehuala"]
    usadas, alumnos = set(), []
    for _ in range(n):
        while (clave := f"{rnd.randint(300000, 399999)}") in usadas:
            pass
        usadas.add(clave)
        local = rnd.random() < 0.45
        alumnos.append({
            "Clave_Alumno": clave,
            "Carrera": rnd.choices(p["carreras"], weights=[5, 3, 2][: len(p["carreras"])])[0],
            "Semestre": rnd.randint(1, 10),
            "Estatus_Residencia": "Local" if local else "Foráneo",
            "Procedencia": "Matehuala" if local else rnd.choice(foraneos),
            "Fuente_Ingreso": rnd.choices(list(p["fuente"]), weights=[6, 2, 3, 1, 1, 1][: len(p["fuente"])])[0],
            "Ingreso_Semanal": round(max(300, rnd.gauss(1550 if local else 1350, 450)), -1),
            "Gasto_Transporte": round(max(0, rnd.gauss(120 if local else 330, 90)), -1),
            "Gasto_Alimentacion": round(max(80, rnd.gauss(380 if local else 520, 140)), -1),
            "Otros_Gastos_Acad": round(max(0, rnd.gauss(150, 70)), -1),
        })
        a = alumnos[-1]
        gasto = a["Gasto_Transporte"] + a["Gasto_Alimentacion"] + a["Otros_Gastos_Acad"]
        if gasto > 2 * a["Ingreso_Semanal"]:  # mismo criterio que la validación
            a["Ingreso_Semanal"] = round(gasto / 2 + 50, -1)
    return alumnos


def crear_demo(n: int = 80, semilla: int = 7, origen: Path = EXCEL_REAL, destino: Path = EXCEL_DEMO) -> Path:
    """Crea una copia del libro con n alumnos FICTICIOS (no toca el archivo real)."""
    shutil.copy(origen, destino)
    alumnos = alumnos_ficticios(n, leer_parametros(destino), semilla)

    def accion(ws):
        ws.Range(f"A{FILA_INICIO}:Z{FILA_FIN}").ClearContents()
        for i, r in enumerate(alumnos):
            _escribir_fila(ws, FILA_INICIO + i, r, datetime(2026, 9, 1, 8) + pd.Timedelta(hours=7 * i))

    _con_libro(destino, accion)
    return destino


# ----------------------------------------------------------------------------
# Estadística
# ----------------------------------------------------------------------------
def tabla_sturges(serie: pd.Series) -> pd.DataFrame:
    """Distribución de frecuencias agrupadas igual que la hoja Regla_Sturges."""
    x = serie.dropna()
    n = len(x)
    if n == 0:
        return pd.DataFrame()
    k = max(1, min(12, round(1 + 3.322 * math.log10(n))))
    minimo, maximo = float(x.min()), float(x.max())
    amplitud_exacta = (maximo - minimo) / k
    amplitud = math.ceil(amplitud_exacta / 10) * 10 if amplitud_exacta > 0 else 10
    filas, acumulada = [], 0
    for i in range(k):
        li = minimo + i * amplitud
        ls = li + amplitud
        ultima = i == k - 1
        fi = int(((x >= li) & ((x <= ls) if ultima else (x < ls))).sum())
        acumulada += fi
        filas.append({
            "Clase": i + 1,
            "Intervalo": f"[${li:,.0f} - ${ls:,.0f}{']' if ultima else ')'}",
            "Lím. inf.": li, "Lím. sup.": ls, "Marca de clase": (li + ls) / 2,
            "fi": fi, "hi": fi / n, "Fi": acumulada, "Hi": acumulada / n,
        })
    return pd.DataFrame(filas)


def resumen_descriptivo(df: pd.DataFrame, columnas: dict[str, str]) -> pd.DataFrame:
    filas = {}
    for col, nombre in columnas.items():
        x = df[col].dropna()
        if x.empty:
            continue
        moda = x.mode()
        filas[nombre] = {
            "n": len(x),
            "Media": x.mean(),
            "Mediana": x.median(),
            "Moda": moda.iloc[0] if len(moda) < len(x) else float("nan"),
            "Mínimo": x.min(),
            "Q1": x.quantile(0.25),
            "Q3": x.quantile(0.75),
            "Máximo": x.max(),
            "Rango": x.max() - x.min(),
            "RIC": x.quantile(0.75) - x.quantile(0.25),
            "Varianza (s²)": x.var(),
            "Desv. estándar (s)": x.std(),
            "Coef. variación": x.std() / x.mean() if x.mean() else float("nan"),
            "Asimetría": x.skew(),
        }
    return pd.DataFrame(filas)
