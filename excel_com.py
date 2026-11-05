"""
Conexión con Microsoft Excel por automatización (COM).

En esta PC el identificador "Excel.Application" está registrado por WPS Office,
así que se abre EXCEL.EXE directamente y se toma su instancia de la tabla de
objetos activos. Así se usa el Excel de Microsoft (necesario para las macros).
"""

from __future__ import annotations

import glob
import os
import subprocess
import time
from contextlib import contextmanager

import pythoncom
import win32com.client

RUTAS_EXCEL = [
    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
    *glob.glob(r"C:\Program Files*\Microsoft Office\Office1*\EXCEL.EXE"),
]


def _ruta_excel() -> str:
    for ruta in RUTAS_EXCEL:
        if os.path.exists(ruta):
            return ruta
    raise RuntimeError("No se encontró Microsoft Excel instalado.")


def _es_microsoft_excel(xl) -> bool:
    try:
        return "Microsoft Office" in str(xl.Path)
    except Exception:
        return False


@contextmanager
def abrir_excel(visible: bool = False):
    """Devuelve una instancia de Microsoft Excel lista para usar y la cierra al terminar."""
    pythoncom.CoInitialize()
    proceso = None
    xl = None
    try:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            if not _es_microsoft_excel(xl):
                xl = None
        except pythoncom.com_error:
            xl = None
        propio = xl is None
        if propio:
            proceso = subprocess.Popen([_ruta_excel(), "/automation", "-Embedding"])
            for _ in range(60):
                time.sleep(0.5)
                try:
                    xl = win32com.client.GetActiveObject("Excel.Application")
                    if _es_microsoft_excel(xl):
                        break
                except pythoncom.com_error:
                    pass
            else:
                raise RuntimeError("Excel no respondió a tiempo.")
            xl.Visible = visible
        xl.DisplayAlerts = False
        yield xl
    finally:
        if xl is not None:
            try:
                if proceso is not None:
                    # Instancia propia: se cierra todo lo que quedó abierto sin guardar.
                    for i in range(xl.Workbooks.Count, 0, -1):
                        xl.Workbooks(i).Close(False)
                    xl.Quit()
                else:
                    xl.DisplayAlerts = True
            except Exception:
                pass
        xl = None
        pythoncom.CoUninitialize()
