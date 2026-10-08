#!/usr/bin/env python3
"""Convertidor universal: detecta las bases de datos de una carpeta y las consolida en un Excel.

Cada formato se REGISTRA con @formato(...): dice cómo reconocerse (por contenido, no solo por
extensión) y cómo leerse. El script escanea la carpeta, detecta qué es cada archivo y convierte
lo reconocido; todo queda catalogado en la pestaña "Índice" y en un archivo .log.

Formatos incluidos: DBF/FoxPro, SQLite, CSV/TSV, Excel (.xlsx, .xls), JSON, Access (.mdb/.accdb).

Uso:
    python convertidor_universal.py                       # carpeta actual
    python convertidor_universal.py -i datos -o salida.xlsx -r
    python convertidor_universal.py --listar              # formatos registrados y su estado
    python convertidor_universal.py --solo dbf sqlite     # limitar a ciertos formatos
    python convertidor_universal.py --texto-extra .txt    # tratar .txt como CSV
    python convertidor_universal.py --excluir              # no excluir nada (por defecto se excluye .bak)
    python convertidor_universal.py --url "postgresql+psycopg2://usuario@host/base"   # servidor SQL
    (la interfaz gráfica está en convertidor_gui.py)

Para agregar un formato nuevo (ejemplo), basta con registrar una función:

    @formato("parquet", "Apache Parquet", detectar=lambda r, cab, a: cab[:4] == b"PAR1",
             canonicas=(".parquet",), requiere=("pyarrow",))
    def leer_parquet(ruta, args):
        yield Tabla(None, pd.read_parquet(ruta))

Opcional: pip install tqdm (barras de progreso), xlrd (.xls), pyodbc (Access).
"""

import argparse
import csv
import datetime as dt
import importlib.util
import json
import numbers
import re
import sqlite3
import struct
import sys
import time
import warnings
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterator, Optional

import numpy as np
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

try:
    from tqdm import tqdm
except ImportError:  # funciona igual, solo sin barras de progreso
    tqdm = None

MAX_FILAS_EXCEL = 1_048_575  # 1,048,576 menos la fila del encabezado
MUESTRA_ANCHO = 200  # filas que se miran para calcular el ancho de cada columna
UMBRAL_BARRA = 20_000  # barra de filas solo para tablas de este tamaño o más
BYTES_CABECERA = 8192
NOMBRE_INDICE = "Índice"
SALIDA_POR_DEFECTO = "Resultados_Consolidados.xlsx"
FORMATO_FECHA = "DD/MM/YYYY"
FORMATO_FECHA_HORA = "DD/MM/YYYY HH:MM:SS"

CARACTERES_PROHIBIDOS_HOJA = re.compile(r"[\[\]:*?/\\]")
# Caracteres de control que Excel (openpyxl) no acepta dentro de una celda
CARACTERES_ILEGALES = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class ErrorUsuario(Exception):
    """Problema que se le explica al usuario (no es un fallo del programa)."""


_HOOK = None  # la interfaz gráfica lo usa para recibir los mensajes en vez de imprimirlos


def escribir(mensaje):
    """Imprime sin romper las barras de progreso (o lo envía a la interfaz)."""
    if _HOOK:
        _HOOK(mensaje)
    else:
        (tqdm.write if tqdm else print)(mensaje)


# =========================================================================== registro

@dataclass
class Tabla:
    """Una tabla leída. nombre=None si el archivo tiene una sola tabla."""
    nombre: Optional[str]
    df: Optional[pd.DataFrame]
    detalle: str = ""   # p. ej. la codificación usada
    error: str = ""     # si la tabla falló pero el resto del archivo puede seguir


@dataclass
class Formato:
    clave: str
    nombre: str
    detectar: Callable      # (ruta, primeros_bytes, args) -> bool
    leer: Callable          # (ruta, args) -> Iterator[Tabla]
    canonicas: tuple = ()   # extensiones "normales" del formato (para nombrar pestañas)
    requiere: tuple = ()    # módulos opcionales necesarios para leerlo


FORMATOS = []  # el orden de registro es el orden de detección


def formato(clave, nombre, detectar, canonicas=(), requiere=()):
    """Decorador que registra un formato nuevo."""
    def decorador(leer):
        FORMATOS.append(Formato(clave, nombre, detectar, leer, canonicas, requiere))
        return leer
    return decorador


def dependencias_faltantes(fmt):
    return [m for m in fmt.requiere if importlib.util.find_spec(m) is None]


def detectar_formato(ruta, args):
    try:
        with open(ruta, "rb") as f:
            cab = f.read(BYTES_CABECERA)
    except OSError:
        return None
    if not cab:
        return None
    for fmt in FORMATOS:
        if args.solo and fmt.clave not in args.solo:
            continue
        try:
            if fmt.detectar(ruta, cab, args):
                return fmt
        except Exception:
            continue
    return None


# =========================================================================== limpieza común

def limpiar_celda(valor, strip, encoding="utf-8"):
    """Deja el valor listo para Excel: decodifica bytes, serializa listas, quita caracteres ilegales."""
    if valor is None or valor is pd.NA:
        return None
    if isinstance(valor, bytes):
        valor = valor.decode(encoding, errors="replace")
    elif isinstance(valor, (list, dict)):
        valor = json.dumps(valor, ensure_ascii=False, default=str)
    if isinstance(valor, str):
        valor = CARACTERES_ILEGALES.sub("", valor)
        if strip:
            valor = valor.strip() or None  # texto vacío -> celda vacía
    elif isinstance(valor, dt.date):
        if isinstance(valor, dt.datetime) and valor.tzinfo is not None:
            valor = valor.replace(tzinfo=None)  # Excel no admite zonas horarias
        if valor.year < 1900:
            valor = valor.isoformat()  # Excel no admite fechas anteriores a 1900
    elif not isinstance(valor, (numbers.Number, np.generic, dt.time, dt.timedelta)):
        valor = str(valor)  # UUID y otros objetos que Excel no puede guardar
    return valor


def limpiar_df(df, strip):
    for i in range(df.shape[1]):
        col = df.iloc[:, i]
        if isinstance(col.dtype, pd.DatetimeTZDtype):
            df.isetitem(i, col.dt.tz_localize(None))
        elif col.dtype == object or pd.api.types.is_string_dtype(col):
            df.isetitem(i, col.map(lambda v: limpiar_celda(v, strip)))
    return df


_NUMERO = r"-?(0|[1-9]\d{0,14})(\.\d{1,15})?"


def inferir_numeros(df):
    """Convierte a número solo las columnas de texto 100% numéricas y sin ceros a la izquierda."""
    for i in range(df.shape[1]):
        col = df.iloc[:, i]
        datos = col.dropna()
        if len(datos) and (col.dtype == object or pd.api.types.is_string_dtype(col)):
            if datos.astype(str).str.fullmatch(_NUMERO).all():
                df.isetitem(i, pd.to_numeric(col))
    return df


# =========================================================================== formato: DBF / FoxPro

VERSIONES_DBF = {0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8B, 0x8E, 0xCB, 0xE5, 0xF5, 0xFB}
TIPOS_CAMPO = b"CNFDLMGPBYTIOQVW0+@"
CODEPAGE_POR_DRIVER = {0x01: "cp437", 0x02: "cp850", 0x03: "cp1252", 0x57: "cp1252", 0x64: "cp852", 0x65: "cp866", 0xC8: "cp1250"}
ENCODINGS_RESPALDO = ["cp1252", "cp850", "latin1"]


def validar_cabecera_dbf(ruta):
    """Devuelve (es_valida, motivo, byte_de_codificacion)."""
    try:
        tam = ruta.stat().st_size
        with open(ruta, "rb") as f:
            cab = f.read(32)
            if len(cab) < 32:
                return False, "archivo demasiado pequeño para ser una tabla DBF", None
            if cab[0] not in VERSIONES_DBF:
                return False, f"el primer byte (0x{cab[0]:02X}) no corresponde a una tabla DBF", None
            long_cab, long_reg = struct.unpack("<HH", cab[8:12])
            if long_cab < 33 or long_cab > tam or long_reg < 2:
                return False, "longitudes de cabecera/registro inválidas", None
            descriptores = f.read(long_cab - 32)
        pos = 0
        while pos < len(descriptores):
            if descriptores[pos] == 0x0D:
                break
            desc = descriptores[pos:pos + 32]
            if len(desc) < 32 or desc[11] not in TIPOS_CAMPO:
                return False, "descriptor de campo inválido", None
            pos += 32
        else:
            return False, "no se encontró el terminador de campos", None
        if pos == 0:
            return False, "la tabla no tiene campos", None
    except OSError as e:
        return False, f"no se pudo leer el archivo: {e}", None
    return True, "", cab[29]


def _es_dbf(ruta, cab, args):
    if args.sin_validar and ruta.suffix.lower() in (".dbf", ".tmp"):
        return True
    return validar_cabecera_dbf(ruta)[0]


def candidatos_encoding(solicitado, driver):
    lista = []
    if solicitado != "auto":
        lista.append(solicitado)
    elif driver in CODEPAGE_POR_DRIVER:
        lista.append(CODEPAGE_POR_DRIVER[driver])
    lista += ENCODINGS_RESPALDO
    return list(dict.fromkeys(lista))


@formato("dbf", "DBF / FoxPro", detectar=_es_dbf, canonicas=(".dbf",))
def leer_dbf(ruta, args):
    from dbfread import DBF  # import tardío: solo se exige si hay DBF

    _, _, driver = validar_cabecera_dbf(ruta)
    candidatos = candidatos_encoding(args.encoding, driver)
    strip = not args.sin_strip
    for i, enc in enumerate(candidatos):
        es_ultima = i == len(candidatos) - 1
        try:
            tabla = DBF(
                str(ruta),
                encoding=enc,
                ignore_missing_memofile=not args.sin_memo_estricto,
                char_decode_errors="replace" if es_ultima else "strict",
            )
            filas = (
                {k: limpiar_celda(v, strip, enc) for k, v in registro.items()}
                for registro in tabla
            )
            if not args.sin_barra and tqdm and len(tabla) >= UMBRAL_BARRA:
                filas = tqdm(filas, total=len(tabla), desc=ruta.name, unit="fila", leave=False)
            df = pd.DataFrame(list(filas), columns=tabla.field_names)
            yield Tabla(None, df, f"codificación {enc}")
            return
        except UnicodeDecodeError:
            escribir(f"  ⚠️  La codificación {enc} no sirve para este archivo; probando otra...")
    raise RuntimeError("no se pudo decodificar el archivo")


# =========================================================================== formato: SQLite

@formato(
    "sqlite", "SQLite",
    detectar=lambda ruta, cab, args: cab.startswith(b"SQLite format 3\x00"),
    canonicas=(".db", ".sqlite", ".sqlite3", ".db3"),
)
def leer_sqlite(ruta, args):
    con = sqlite3.connect(f"{ruta.as_uri()}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode("utf-8", errors="replace")
    try:
        nombres = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")]
        for nombre in nombres:
            try:
                sql = 'SELECT * FROM "{}"'.format(nombre.replace('"', '""'))
                df = pd.read_sql_query(sql, con)
                yield Tabla(nombre, limpiar_tabla(df, args))
            except Exception as e:
                yield Tabla(nombre, None, error=f"{type(e).__name__}: {e}")
    finally:
        con.close()


def limpiar_tabla(df, args):
    return limpiar_df(df, not args.sin_strip)


# =========================================================================== formato: Access

def _es_access(ruta, cab, args):
    return cab[4:19] in (b"Standard Jet DB", b"Standard ACE DB")


@formato("access", "Microsoft Access", detectar=_es_access, canonicas=(".mdb", ".accdb"), requiere=("pyodbc",))
def leer_access(ruta, args):
    import pyodbc  # requiere el controlador "Microsoft Access Driver" instalado en Windows

    con = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={ruta};", readonly=True)
    try:
        nombres = [t.table_name for t in con.cursor().tables(tableType="TABLE")]
        for nombre in nombres:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)  # pandas prefiere SQLAlchemy
                    df = pd.read_sql(f"SELECT * FROM [{nombre}]", con)
                yield Tabla(nombre, limpiar_tabla(df, args))
            except Exception as e:
                yield Tabla(nombre, None, error=f"{type(e).__name__}: {e}")
    finally:
        con.close()


# =========================================================================== formato: Excel

def _es_xlsx(ruta, cab, args):
    if not cab.startswith(b"PK\x03\x04") or not zipfile.is_zipfile(ruta):
        return False
    with zipfile.ZipFile(ruta) as z:
        return "xl/workbook.xml" in z.namelist()


def _leer_excel(ruta, args):
    with pd.ExcelFile(ruta) as libro:
        hojas = libro.sheet_names
        for hoja in hojas:
            try:
                df = libro.parse(hoja)
                yield Tabla(hoja if len(hojas) > 1 else None, limpiar_tabla(df, args))
            except Exception as e:
                yield Tabla(hoja, None, error=f"{type(e).__name__}: {e}")


@formato("xlsx", "Excel (.xlsx/.xlsm)", detectar=_es_xlsx, canonicas=(".xlsx", ".xlsm"), requiere=("openpyxl",))
def leer_xlsx(ruta, args):
    yield from _leer_excel(ruta, args)


@formato(
    "xls", "Excel 97-2003 (.xls)",
    detectar=lambda ruta, cab, args: cab.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") and ruta.suffix.lower() == ".xls",
    canonicas=(".xls",), requiere=("xlrd",),
)
def leer_xls(ruta, args):
    yield from _leer_excel(ruta, args)


# =========================================================================== formato: JSON

def _es_json(ruta, cab, args):
    if ruta.suffix.lower() != ".json":
        return False
    texto = cab.decode("utf-8-sig", errors="ignore").lstrip()
    return texto[:1] in ("[", "{")


@formato("json", "JSON", detectar=_es_json, canonicas=(".json",))
def leer_json(ruta, args):
    with open(ruta, encoding="utf-8-sig") as f:
        datos = json.load(f)

    def a_tabla(valor):
        if isinstance(valor, list) and valor and all(isinstance(x, dict) for x in valor):
            return pd.json_normalize(valor)
        if isinstance(valor, list):
            return pd.DataFrame({"valor": valor})
        if isinstance(valor, dict):
            return pd.json_normalize(valor)
        return pd.DataFrame({"valor": [valor]})

    if isinstance(datos, dict):
        listas = {k: v for k, v in datos.items() if isinstance(v, list) and v and all(isinstance(x, dict) for x in v)}
        if listas:
            for clave, valor in listas.items():
                yield Tabla(str(clave), limpiar_tabla(a_tabla(valor), args))
            return
    yield Tabla(None, limpiar_tabla(a_tabla(datos), args))


# =========================================================================== formato: CSV / TSV

EXTENSIONES_TEXTO = {".csv", ".tsv"}


def _es_texto_delimitado(ruta, cab, args):
    ext = ruta.suffix.lower()
    if ext not in EXTENSIONES_TEXTO and ext not in args.texto_extra:
        return False
    return b"\x00" not in cab  # un archivo con bytes nulos es binario


@formato("csv", "Texto delimitado (CSV/TSV)", detectar=_es_texto_delimitado, canonicas=(".csv", ".tsv"))
def leer_csv(ruta, args):
    with open(ruta, "rb") as f:
        muestra = f.read(65536).decode("utf-8", errors="replace")
    if ruta.suffix.lower() == ".tsv":
        sep = "\t"
    else:
        try:
            sep = csv.Sniffer().sniff(muestra, delimiters=",;\t|").delimiter
        except csv.Error:
            sep = ","

    candidatos = ["utf-8-sig", "cp1252", "latin1"]
    if args.encoding != "auto":
        candidatos.insert(0, args.encoding)
    for enc in dict.fromkeys(candidatos):
        try:
            # dtype=str evita perder ceros a la izquierda; solo el campo vacío cuenta como nulo ("NA" es texto)
            df = pd.read_csv(ruta, sep=sep, encoding=enc, dtype=str, keep_default_na=False, na_values=[""])
            df = inferir_numeros(limpiar_df(df, not args.sin_strip))
            yield Tabla(None, df, f"separador {sep!r}, codificación {enc}")
            return
        except UnicodeDecodeError:
            continue
    raise RuntimeError("no se pudo decodificar el archivo")


# =========================================================================== servidores SQL (SQLAlchemy)
# No se detectan por contenido: se conectan con una URL (--url) o con el formulario de la interfaz gráfica.

MOTORES = {
    "PostgreSQL": {"driver": "postgresql+psycopg2", "puerto": 5432},
    "MySQL / MariaDB": {"driver": "mysql+pymysql", "puerto": 3306},
    "SQL Server": {"driver": "mssql+pyodbc", "puerto": 1433},
    "Oracle": {"driver": "oracle+oracledb", "puerto": 1521},
    "Otro (URL de SQLAlchemy)": {"driver": None, "puerto": None},
}
PIP_POR_MODULO = {"psycopg2": "psycopg2-binary", "pymysql": "pymysql", "pyodbc": "pyodbc", "oracledb": "oracledb"}
DRIVERS_ODBC_SQLSERVER = [
    "ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0", "SQL Server",
]


def resumir_error(e):
    """Primera línea del error, sin el relleno que agrega SQLAlchemy."""
    linea = (str(e).strip().splitlines() or [type(e).__name__])[0]
    return linea[:300]


def _falta_modulo(e):
    nombre = getattr(e, "name", None) or str(e)
    if nombre == "sqlalchemy":
        return ErrorUsuario("Falta instalar SQLAlchemy: pip install sqlalchemy")
    return ErrorUsuario(f"Falta instalar el driver de la base de datos: pip install {PIP_POR_MODULO.get(nombre, nombre)}")


def _driver_odbc_sqlserver():
    import pyodbc
    disponibles = pyodbc.drivers()
    for d in DRIVERS_ODBC_SQLSERVER:
        if d in disponibles:
            return d
    raise ErrorUsuario("No hay ningún controlador ODBC de SQL Server instalado en este equipo.")


def construir_url(motor, host="", puerto="", usuario="", clave="", base="", windows_auth=False, url_libre=""):
    """Arma la URL de conexión (escapa bien los símbolos raros de las contraseñas)."""
    cfg = MOTORES[motor]
    if cfg["driver"] is None:
        if not url_libre.strip():
            raise ErrorUsuario("Escribe la URL de conexión.")
        return url_libre.strip()
    if not host.strip():
        raise ErrorUsuario("Falta el servidor.")
    try:
        puerto = int(str(puerto).strip()) if str(puerto).strip() else None
    except ValueError:
        raise ErrorUsuario("El puerto debe ser un número.")
    try:
        from sqlalchemy.engine import URL
        usuario, clave, base = usuario or None, clave or None, base.strip() or None
        consulta = {}
        if motor == "SQL Server":
            consulta = {"driver": _driver_odbc_sqlserver(), "TrustServerCertificate": "yes"}
            if windows_auth:
                consulta["Trusted_Connection"] = "yes"
                usuario = clave = None
        elif motor == "Oracle":
            consulta, base = {"service_name": base}, None
        return URL.create(cfg["driver"], username=usuario, password=clave, host=host.strip(),
                          port=puerto, database=base, query=consulta)
    except ModuleNotFoundError as e:
        raise _falta_modulo(e)


def descripcion_url(url):
    """Texto seguro para mostrar (sin contraseña): servidor/base."""
    try:
        from sqlalchemy.engine import make_url
        u = make_url(url) if isinstance(url, str) else url
    except ModuleNotFoundError as e:
        raise _falta_modulo(e)
    base = u.database or (u.query.get("service_name") if u.query else "") or ""
    return f"{u.host or 'local'}/{base}".rstrip("/")


def abrir_engine(url):
    try:
        from sqlalchemy import create_engine
        return create_engine(url)
    except ModuleNotFoundError as e:
        raise _falta_modulo(e)


def listar_tablas_servidor(url, esquema=None):
    """Conecta y devuelve los nombres de tablas. Lanza ErrorUsuario con un mensaje legible si falla."""
    engine = abrir_engine(url)
    try:
        from sqlalchemy import inspect
        return inspect(engine).get_table_names(schema=esquema or None)
    except ErrorUsuario:
        raise
    except Exception as e:
        raise ErrorUsuario(f"No se pudo conectar: {resumir_error(e)}")
    finally:
        engine.dispose()


def _leer_consulta(engine, sql, args):
    """Lee por bloques (sirve en cualquier motor) y corta al llegar al límite de filas."""
    from sqlalchemy import text
    limite = args.limite_filas
    partes, n = [], 0
    with engine.connect() as con:
        con = con.execution_options(stream_results=True)
        for trozo in pd.read_sql_query(text(sql), con, chunksize=50_000):
            partes.append(trozo)
            n += len(trozo)
            if limite and n >= limite:
                break
    df = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if limite:
        df = df.iloc[:limite]
    return limpiar_tabla(df, args)


def leer_servidor(ruta, args):
    engine = abrir_engine(args.url)
    try:
        from sqlalchemy import inspect
        esquema = args.esquema or None
        nombres = list(args.tablas) if args.tablas else inspect(engine).get_table_names(schema=esquema)
        p = engine.dialect.identifier_preparer
        for nombre in nombres:
            try:
                ident = p.quote(nombre)
                if esquema:
                    ident = f"{p.quote_schema(esquema)}.{ident}"
                yield Tabla(nombre, _leer_consulta(engine, f"SELECT * FROM {ident}", args))
            except Exception as e:
                yield Tabla(nombre, None, error=resumir_error(e))
    finally:
        engine.dispose()


# No se registra en FORMATOS (no se detecta en carpetas): se usa solo con una URL de conexión.
FORMATO_SERVIDOR = Formato("servidor", "Servidor SQL", detectar=lambda *a: False, leer=leer_servidor,
                           requiere=("sqlalchemy",))


def fuente_servidor(args):
    nombre = descripcion_url(args.url)
    return [(Path(re.sub(r"[^\w.-]+", "_", nombre)), nombre, FORMATO_SERVIDOR)]


# =========================================================================== escritura

def nombre_hoja_unico(base, usadas, sufijo=""):
    """Nombre de pestaña válido (<=31 caracteres, sin símbolos prohibidos, sin repetir)."""
    base = CARACTERES_PROHIBIDOS_HOJA.sub("_", str(base)).strip("'") or "Hoja"
    n = 1
    while True:
        marca = (f" ({n})" if n > 1 else "") + sufijo
        candidato = base[: 31 - len(marca)] + marca
        if candidato.lower() not in usadas:
            usadas.add(candidato.lower())
            return candidato
        n += 1


def base_hoja(ruta, fmt, tabla):
    """INCOTE.DBF -> 'INCOTE'; INCOTE.TMP -> 'INCOTE_tmp'; datos.db + tabla 'ventas' -> 'datos_ventas'."""
    if fmt.clave == "servidor":
        return str(tabla)
    ext = ruta.suffix.lower()
    base = ruta.stem if (not ext or ext in fmt.canonicas) else f"{ruta.stem}_{ext[1:]}"
    if tabla:
        corto = max(3, 31 - len(str(tabla)) - 1)
        return f"{base[:corto]}_{tabla}"
    return base


def dar_formato(hoja, df):
    """Congela el encabezado, activa filtros y ajusta el ancho de las columnas."""
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for i in range(df.shape[1]):
        largos = df.iloc[:MUESTRA_ANCHO, i].astype(str).str.len()
        ancho = max(len(str(df.columns[i])), int(largos.max()) if len(largos) else 0)
        hoja.column_dimensions[get_column_letter(i + 1)].width = min(ancho + 2, 50)


def escribir_hojas(writer, df, base, usadas, max_filas):
    """Escribe el DataFrame; si excede max_filas lo reparte en pestañas _1, _2, ..."""
    partes = max(1, -(-len(df) // max_filas))
    hojas = []
    for n in range(partes):
        trozo = df.iloc[n * max_filas:(n + 1) * max_filas]
        hoja = nombre_hoja_unico(base, usadas, f"_{n + 1}" if partes > 1 else "")
        trozo.to_excel(writer, sheet_name=hoja, index=False)
        dar_formato(writer.sheets[hoja], trozo)
        hojas.append(hoja)
    return hojas


def escribir_indice(writer, registro):
    """Hoja 'Índice' (primera del libro) con el catálogo de todo lo detectado y su resultado."""
    columnas = ["Archivo", "Formato", "Tabla", "Pestaña", "Filas", "Columnas", "Estado", "Detalle"]
    df = pd.DataFrame(registro, columns=columnas)
    df.to_excel(writer, sheet_name=NOMBRE_INDICE, index=False)
    ws = writer.sheets[NOMBRE_INDICE]
    dar_formato(ws, df)
    for fila, item in enumerate(registro, start=2):
        if item.get("_hoja"):
            celda = ws.cell(row=fila, column=4)
            celda.hyperlink = Hyperlink(ref=celda.coordinate, location=f"'{item['_hoja']}'!A1")
            celda.font = Font(color="0563C1", underline="single")
    libro = writer.book
    libro.move_sheet(NOMBRE_INDICE, offset=-(len(libro.sheetnames) - 1))
    libro.active = 0
    for hoja in libro.worksheets:  # evita que Excel abra varias pestañas agrupadas
        hoja.sheet_view.tabSelected = hoja.title == NOMBRE_INDICE


# =========================================================================== proceso

def listar_archivos(carpeta, recursivo, excluir_rutas, excluir_ext):
    candidatos = carpeta.rglob("*") if recursivo else carpeta.iterdir()
    archivos, excluidos = [], []
    for p in candidatos:
        if not p.is_file() or p.resolve() in excluir_rutas:
            continue
        (excluidos if p.suffix.lower() in excluir_ext else archivos).append(p.resolve())
    key = lambda p: str(p).lower()
    return sorted(set(archivos), key=key), sorted(set(excluidos), key=key)


def procesar(ruta, nombre, fmt, writer, usadas, registro, args, cancelar=None):
    """Convierte una fuente (archivo detectado o servidor). Devuelve su resultado para el log."""
    t0 = time.perf_counter()
    res = {"archivo": nombre, "formato": fmt.nombre, "estado": "OK", "tablas": 0, "filas": 0, "detalle": ""}
    escribir(f"Procesando: {nombre}  [{fmt.nombre}] ...")

    def fila(tabla, hojas, filas, columnas, estado, detalle):
        registro.append({
            "Archivo": nombre, "Formato": fmt.nombre, "Tabla": tabla or "", "Pestaña": ", ".join(hojas),
            "Filas": filas, "Columnas": columnas, "Estado": estado, "Detalle": detalle,
            "_hoja": hojas[0] if hojas else "",
        })

    faltan = dependencias_faltantes(fmt)
    if faltan:
        detalle = f"falta instalar: pip install {' '.join(faltan)}"
        res.update(estado="OMITIDO", detalle=detalle)
        fila(None, [], None, None, "OMITIDO", detalle)
        escribir(f"  ⏭️  Omitido: {detalle}")
        res["segundos"] = time.perf_counter() - t0
        return res

    errores = []
    try:
        for t in fmt.leer(ruta, args):
            if t.error:
                errores.append(f"{t.nombre}: {t.error}")
                fila(t.nombre, [], None, None, "ERROR", t.error)
                escribir(f"  ❌ Tabla '{t.nombre}': {t.error}")
            else:
                hojas = escribir_hojas(writer, t.df, base_hoja(ruta, fmt, t.nombre), usadas, args.max_filas)
                res["tablas"] += 1
                res["filas"] += len(t.df)
                fila(t.nombre, hojas, len(t.df), t.df.shape[1], "OK", t.detalle)
                extra = f" (dividida en {len(hojas)} pestañas)" if len(hojas) > 1 else ""
                escribir(f"  ✅ {t.nombre + ': ' if t.nombre else ''}{len(t.df):,} filas en '{hojas[0]}'{extra}"
                         + (f" [{t.detalle}]" if t.detalle else ""))
            del t
            if cancelar and cancelar.is_set():
                break
    except ErrorUsuario as e:
        errores.append(str(e))
        fila(None, [], None, None, "ERROR", str(e))
        escribir(f"  ❌ {e}")
    except Exception as e:  # un archivo dañado no debe detener el resto
        errores.append(f"{type(e).__name__}: {e}")
        fila(None, [], None, None, "ERROR", errores[-1])
        escribir(f"  ❌ Error al procesar '{nombre}': {e}")

    if errores:
        res.update(estado="ERROR", detalle=" | ".join(errores))
    elif res["tablas"] == 0:
        res.update(estado="OMITIDO", detalle="el archivo no contiene tablas")
        fila(None, [], None, None, "OMITIDO", res["detalle"])
        escribir("  ⏭️  Omitido: el archivo no contiene tablas.")
    res["segundos"] = time.perf_counter() - t0
    return res


def guardar_log(ruta_log, args, origen, salida, resultados, no_reconocidos, excluidos, segundos, cancelado=False):
    cuenta = Counter(r["estado"] for r in resultados)
    lineas = [
        "REGISTRO DE CONVERSIÓN UNIVERSAL -> EXCEL",
        f"Fecha: {dt.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Origen: {origen}",
        f"Excel de salida: {salida}",
        f"Codificación solicitada: {args.encoding}",
    ]
    if cancelado:
        lineas.append("⚠ CONVERSIÓN CANCELADA POR EL USUARIO: el Excel solo contiene lo procesado hasta ese momento.")
    lineas += ["", "DETALLE POR ARCHIVO", "-" * 70]
    for r in resultados:
        lineas.append(
            f"[{r['estado']:<7}] {r['archivo']} | {r['formato']} | tablas: {r['tablas']} | "
            f"filas: {r['filas']:,} | {r.get('segundos', 0):.1f} s"
        )
        if r["detalle"]:
            lineas.append(f"           detalle: {r['detalle']}")
    if excluidos:
        lineas += ["", f"EXCLUIDOS POR EXTENSIÓN ({', '.join(args.excluir)}): {len(excluidos)}", "-" * 70]
        lineas += [f"  {p.name}" for p in excluidos[:200]]
    if no_reconocidos:
        lineas += ["", f"NO RECONOCIDOS COMO BASE DE DATOS: {len(no_reconocidos)}", "-" * 70]
        lineas += [f"  {p.name}" for p in no_reconocidos[:500]]
        if len(no_reconocidos) > 500:
            lineas.append(f"  ... y {len(no_reconocidos) - 500} más")
    lineas += [
        "", "RESUMEN", "-" * 70,
        f"Convertidos: {cuenta['OK']} | Omitidos: {cuenta['OMITIDO']} | Con error: {cuenta['ERROR']}",
        f"Filas totales: {sum(r['filas'] for r in resultados):,}",
        f"Tiempo total: {segundos:.1f} s",
    ]
    ruta_log.write_text("\n".join(lineas) + "\n", encoding="utf-8-sig")


def listar_formatos():
    print("Formatos registrados (en orden de detección):\n")
    for f in FORMATOS:
        faltan = dependencias_faltantes(f)
        estado = "listo" if not faltan else f"falta: pip install {' '.join(faltan)}"
        print(f"  {f.clave:<8} {f.nombre:<28} {', '.join(f.canonicas) or '(por contenido)':<30} {estado}")
    print("\nServidores SQL (PostgreSQL, MySQL/MariaDB, SQL Server, Oracle y otros): se usan con --url o con la interfaz gráfica.")


def parsear(argv=None):
    p = argparse.ArgumentParser(description="Detecta bases de datos en una carpeta (o en un servidor) y las consolida en un Excel.")
    p.add_argument("-i", "--carpeta", default=".", help="Carpeta a escanear (por defecto: actual)")
    p.add_argument("-o", "--salida", default=SALIDA_POR_DEFECTO, help="Archivo Excel de salida")
    p.add_argument("-r", "--recursivo", action="store_true", help="Buscar también en subcarpetas")
    p.add_argument("--solo", nargs="+", metavar="FORMATO", help="Convertir solo estos formatos (ver --listar)")
    p.add_argument("--listar", action="store_true", help="Mostrar los formatos registrados y salir")
    p.add_argument("--excluir", nargs="*", default=[".bak"], metavar="EXT",
                   help="Extensiones a ignorar aunque se detecten (def.: .bak). Sin valores = no excluir nada")
    p.add_argument("--texto-extra", nargs="+", default=[], metavar="EXT",
                   help="Extensiones adicionales que se tratan como texto delimitado (ej.: .txt)")
    p.add_argument("-e", "--encoding", default="auto", help="Codificación (auto, cp850, cp1252, latin1...)")
    p.add_argument("--max-filas", type=int, default=MAX_FILAS_EXCEL, help="Filas por pestaña antes de dividir")
    p.add_argument("--log", help="Ruta del archivo de registro (def.: junto al Excel, extensión .log)")
    p.add_argument("--url", help="Convertir un servidor SQL en vez de una carpeta (URL de SQLAlchemy)")
    p.add_argument("--esquema", help="Servidor: esquema donde buscar las tablas (def.: el predeterminado)")
    p.add_argument("--tablas", nargs="+", help="Servidor: convertir solo estas tablas (def.: todas)")
    p.add_argument("--limite-filas", type=int, help="Servidor: máximo de filas por tabla (def.: todas)")
    p.add_argument("--servidor", action="store_true", help="(interfaz gráfica) abrir directo el formulario de conexión")
    p.add_argument("--sin-memo-estricto", action="store_true", help="DBF: no fallar si falta el .FPT/.DBT")
    p.add_argument("--sin-validar", action="store_true", help="Tratar todo .dbf/.tmp como DBF sin validar la cabecera")
    p.add_argument("--sin-strip", action="store_true", help="No quitar espacios al inicio/fin de los textos")
    p.add_argument("--sin-barra", action="store_true", help="No mostrar barras de progreso")
    args = p.parse_args(argv)
    norm = lambda lista: {("." + e.lstrip(".")).lower() for e in lista}
    args.excluir_ext = norm(args.excluir)
    args.texto_extra = norm(args.texto_extra)
    args.solo = {s.lower() for s in args.solo} if args.solo else None
    return args


# --------------------------------------------------------------------------- API (usada por la consola y la interfaz)

def validar_opciones(args):
    claves = {f.clave for f in FORMATOS}
    if args.solo and not args.solo <= claves:
        raise ErrorUsuario(f"Formato(s) desconocido(s): {', '.join(sorted(args.solo - claves))}. Usa --listar.")
    if not 1 <= args.max_filas <= MAX_FILAS_EXCEL:
        raise ErrorUsuario(f"--max-filas debe estar entre 1 y {MAX_FILAS_EXCEL}.")


def preparar_salida(args):
    salida = Path(args.salida).resolve()
    ruta_log = Path(args.log).resolve() if args.log else salida.with_suffix(".log")
    if not salida.parent.is_dir():
        raise ErrorUsuario(f"La carpeta de salida '{salida.parent}' no existe.")
    if salida.exists():  # Excel bloquea el archivo si lo tienes abierto
        try:
            salida.open("ab").close()
        except PermissionError:
            raise ErrorUsuario(f"No se puede escribir en '{salida.name}'. Si lo tienes abierto en Excel, ciérralo.")
    return salida, ruta_log


def escanear(args):
    """Detecta las bases de datos de la carpeta. Devuelve fuentes = [(ruta, nombre, formato)]."""
    validar_opciones(args)
    carpeta = Path(args.carpeta).resolve()
    if not carpeta.is_dir():
        raise ErrorUsuario(f"La carpeta '{carpeta}' no existe.")
    salida = Path(args.salida).resolve()
    ruta_log = Path(args.log).resolve() if args.log else salida.with_suffix(".log")
    archivos, excluidos = listar_archivos(carpeta, args.recursivo, {salida, ruta_log}, args.excluir_ext)

    def nombre_mostrado(ruta):
        try:
            return str(ruta.relative_to(carpeta))
        except ValueError:
            return ruta.name

    fuentes, no_reconocidos = [], []
    for ruta in archivos:
        fmt = detectar_formato(ruta, args)
        if fmt:
            fuentes.append((ruta, nombre_mostrado(ruta), fmt))
        else:
            no_reconocidos.append(ruta)
    origen = f"{carpeta}" + (" (con subcarpetas)" if args.recursivo else "")
    return {"carpeta": carpeta, "origen": origen, "total_archivos": len(archivos), "fuentes": fuentes,
            "no_reconocidos": no_reconocidos, "excluidos": excluidos}


def convertir(args, fuentes, origen, no_reconocidos=(), excluidos=(), evento=None, cancelar=None):
    """Convierte las fuentes a un Excel + log.

    evento(tipo, **datos) recibe: "inicio"(total), "archivo"(i, total, nombre), "mensaje"(texto).
    cancelar: threading.Event; si se activa, se detiene tras la tabla en curso y guarda lo ya procesado.
    """
    global _HOOK
    validar_opciones(args)
    salida, ruta_log = preparar_salida(args)
    emitir = evento or (lambda tipo, **datos: None)
    resultados, registro = [], []
    usadas = {NOMBRE_INDICE.lower()}
    cancelado = False
    inicio = time.perf_counter()
    _HOOK = (lambda m: emitir("mensaje", texto=m)) if evento else None
    emitir("inicio", total=len(fuentes))
    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            # pandas ignora date_format/datetime_format en el writer de openpyxl, por eso se fijan aquí.
            writer._date_format = FORMATO_FECHA
            writer._datetime_format = FORMATO_FECHA_HORA
            usar_barra = tqdm and not args.sin_barra and evento is None
            iterador = tqdm(fuentes, unit="archivo") if usar_barra else fuentes
            for i, (ruta, nombre, fmt) in enumerate(iterador, start=1):
                if cancelar and cancelar.is_set():
                    cancelado = True
                    break
                emitir("archivo", i=i, total=len(fuentes), nombre=nombre)
                resultados.append(procesar(ruta, nombre, fmt, writer, usadas, registro, args, cancelar))
                if cancelar and cancelar.is_set():
                    cancelado = True
                    break
            escribir_indice(writer, registro)
    finally:
        _HOOK = None
        total = time.perf_counter() - inicio
        guardar_log(ruta_log, args, origen, salida, resultados, no_reconocidos, excluidos, total, cancelado)
    return {"resultados": resultados, "segundos": total, "salida": salida, "ruta_log": ruta_log,
            "cancelado": cancelado, "conteo": Counter(r["estado"] for r in resultados)}


def main():
    args = parsear()
    if args.listar:
        listar_formatos()
        return 0
    try:
        validar_opciones(args)
        if args.url:
            fuentes = fuente_servidor(args)
            origen, no_reconocidos, excluidos = f"Servidor {fuentes[0][1]}", [], []
            print(f"🔌 Conectando a {fuentes[0][1]} ...\n")
        else:
            esc = escanear(args)
            fuentes, origen = esc["fuentes"], esc["origen"]
            no_reconocidos, excluidos = esc["no_reconocidos"], esc["excluidos"]
            if not fuentes:
                print(f"❌ No se detectó ninguna base de datos en '{esc['carpeta']}' ({esc['total_archivos']} archivos revisados).")
                return 1
            por_formato = Counter(f.nombre for _, _, f in fuentes)
            resumen = ", ".join(f"{n} {nombre}" for nombre, n in por_formato.most_common())
            print(f"🔎 Se revisaron {esc['total_archivos']} archivos: {len(fuentes)} reconocidos ({resumen}), "
                  f"{len(no_reconocidos)} no reconocidos, {len(excluidos)} excluidos.\n")
        res = convertir(args, fuentes, origen, no_reconocidos, excluidos)
    except ErrorUsuario as e:
        print(f"❌ {e}")
        return 1

    cuenta = res["conteo"]
    print("\n" + "=" * 60)
    print(f"✅ Convertidos: {cuenta['OK']}   ⏭️  Omitidos: {cuenta['OMITIDO']}   "
          f"❌ Con error: {cuenta['ERROR']}   ⏱️  {res['segundos']:.1f} s")
    for r in res["resultados"]:
        if r["estado"] != "OK":
            print(f"   - [{r['estado']}] {r['archivo']}: {r['detalle']}")
    print(f"📄 Excel: '{res['salida']}'  (la pestaña '{NOMBRE_INDICE}' cataloga todo)")
    print(f"📝 Registro: '{res['ruta_log']}'")
    return 0 if cuenta["OK"] else 1


if __name__ == "__main__":
    sys.exit(main())
