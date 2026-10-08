#!/usr/bin/env python3
"""Consolida tablas DBF (FoxPro/dBASE) de una carpeta en un solo Excel, una pestaña por tabla.

Uso:
    python dbf_a_excel.py                          # carpeta actual (.dbf y .tmp)
    python dbf_a_excel.py -i datos -o salida.xlsx -r   # con subcarpetas
    python dbf_a_excel.py -x .dbf .tmp .bak        # más extensiones
    python dbf_a_excel.py -e cp850 --sin-memo-estricto

Opcional: pip install tqdm   (barras de progreso)
"""

import argparse
import datetime as dt
import re
import struct
import sys
import time
from pathlib import Path

import pandas as pd
from dbfread import DBF
from openpyxl.utils import get_column_letter

try:
    from tqdm import tqdm
except ImportError:  # el script funciona igual, solo sin barras de progreso
    tqdm = None

MAX_FILAS_EXCEL = 1_048_575  # 1,048,576 menos la fila del encabezado
MUESTRA_ANCHO = 200  # filas que se miran para calcular el ancho de cada columna
UMBRAL_BARRA = 20_000  # barra de filas solo para tablas de este tamaño o más
FORMATO_FECHA = "DD/MM/YYYY"
FORMATO_FECHA_HORA = "DD/MM/YYYY HH:MM:SS"

CARACTERES_PROHIBIDOS_HOJA = re.compile(r"[\[\]:*?/\\]")
# Caracteres de control que Excel (openpyxl) no acepta dentro de una celda
CARACTERES_ILEGALES = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# Primer byte de una tabla dBASE/FoxPro (FoxBase, FoxPro 2.x, Visual FoxPro, dBASE III/IV)
VERSIONES_DBF = {0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8B, 0x8E, 0xCB, 0xE5, 0xF5, 0xFB}
# Tipos de campo válidos en la cabecera
TIPOS_CAMPO = b"CNFDLMGPBYTIOQVW0+@"
# Byte 29 de la cabecera (language driver) -> codificación
CODEPAGE_POR_DRIVER = {0x01: "cp437", 0x02: "cp850", 0x03: "cp1252", 0x57: "cp1252", 0x64: "cp852", 0x65: "cp866", 0xC8: "cp1250"}
ENCODINGS_RESPALDO = ["cp1252", "cp850", "latin1"]


def escribir(mensaje):
    """Imprime sin romper las barras de progreso."""
    (tqdm.write if tqdm else print)(mensaje)


# --------------------------------------------------------------------------- validación

def validar_cabecera(ruta):
    """Comprueba que el archivo tenga estructura de tabla DBF.

    Devuelve (es_valida, motivo, byte_de_codificacion).
    """
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

            # Descriptores de campo (32 bytes c/u) hasta el terminador 0x0D
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


# --------------------------------------------------------------------------- lectura

def candidatos_encoding(solicitado, driver):
    """Lista ordenada de codificaciones a probar."""
    lista = []
    if solicitado != "auto":
        lista.append(solicitado)
    elif driver in CODEPAGE_POR_DRIVER:
        lista.append(CODEPAGE_POR_DRIVER[driver])
    lista += ENCODINGS_RESPALDO
    return list(dict.fromkeys(lista))


def limpiar_valor(valor, encoding, strip):
    """Decodifica bytes, quita caracteres ilegales y espacios; protege fechas anteriores a 1900."""
    if isinstance(valor, bytes):
        valor = valor.decode(encoding, errors="replace")
    if isinstance(valor, str):
        valor = CARACTERES_ILEGALES.sub("", valor)
        if strip:
            valor = valor.strip() or None  # texto vacío -> celda vacía
    elif isinstance(valor, dt.date) and valor.year < 1900:
        valor = valor.isoformat()  # Excel no admite fechas anteriores a 1900
    return valor


def leer_dbf(ruta, candidatos, memo_estricto, strip, con_barra):
    """Lee la tabla probando codificaciones en orden. Devuelve (DataFrame, encoding_usado).

    Con las primeras codificaciones se decodifica en modo estricto: si hay bytes que no
    pertenecen a esa codificación se pasa a la siguiente. La última usa errors="replace".
    """
    for i, enc in enumerate(candidatos):
        es_ultima = i == len(candidatos) - 1
        try:
            tabla = DBF(
                str(ruta),
                encoding=enc,
                ignore_missing_memofile=not memo_estricto,
                char_decode_errors="replace" if es_ultima else "strict",
            )
            filas = (
                {k: limpiar_valor(v, enc, strip) for k, v in registro.items()}
                for registro in tabla
            )
            if con_barra and tqdm and len(tabla) >= UMBRAL_BARRA:
                filas = tqdm(filas, total=len(tabla), desc=ruta.name, unit="fila", leave=False)
            registros = list(filas)
            return pd.DataFrame(registros, columns=tabla.field_names), enc
        except UnicodeDecodeError:
            escribir(f"  ⚠️  La codificación {enc} no sirve para este archivo; probando otra...")
    raise RuntimeError("no se pudo decodificar el archivo")  # no debería ocurrir


# --------------------------------------------------------------------------- escritura

def nombre_hoja_unico(base, usadas, sufijo=""):
    """Nombre de pestaña válido (<=31 caracteres, sin símbolos prohibidos, sin repetir)."""
    base = CARACTERES_PROHIBIDOS_HOJA.sub("_", base).strip("'") or "Hoja"
    n = 1
    while True:
        marca = (f" ({n})" if n > 1 else "") + sufijo
        candidato = base[: 31 - len(marca)] + marca
        if candidato.lower() not in usadas:
            usadas.add(candidato.lower())
            return candidato
        n += 1


def dar_formato(hoja, df):
    """Congela el encabezado, activa filtros y ajusta el ancho de las columnas."""
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for i, col in enumerate(df.columns, start=1):
        largos = df[col].head(MUESTRA_ANCHO).astype(str).str.len()
        ancho = max(len(str(col)), int(largos.max()) if len(largos) else 0)
        hoja.column_dimensions[get_column_letter(i)].width = min(ancho + 2, 50)


def escribir_hojas(writer, df, base, usadas, max_filas):
    """Escribe el DataFrame; si excede max_filas lo reparte en pestañas _1, _2, ..."""
    partes = max(1, -(-len(df) // max_filas))  # división hacia arriba
    hojas = []
    for n in range(partes):
        trozo = df.iloc[n * max_filas:(n + 1) * max_filas]
        hoja = nombre_hoja_unico(base, usadas, f"_{n + 1}" if partes > 1 else "")
        trozo.to_excel(writer, sheet_name=hoja, index=False)
        dar_formato(writer.sheets[hoja], trozo)
        hojas.append(hoja)
    return hojas


def base_hoja(ruta):
    """INCOTE.DBF -> 'INCOTE'; INCOTE.TMP -> 'INCOTE_tmp'."""
    if ruta.suffix.lower() == ".dbf":
        return ruta.stem
    return f"{ruta.stem}_{ruta.suffix[1:].lower()}"


# --------------------------------------------------------------------------- proceso

def buscar_archivos(carpeta, extensiones, recursivo, excluir):
    candidatos = carpeta.rglob("*") if recursivo else carpeta.iterdir()
    return sorted(
        {
            p.resolve()
            for p in candidatos
            if p.is_file() and p.suffix.lower() in extensiones and p.resolve() not in excluir
        },
        key=lambda p: str(p).lower(),
    )


def procesar(ruta, nombre, writer, usadas, args):
    t0 = time.perf_counter()
    res = {"archivo": nombre, "estado": "OK", "filas": 0, "hojas": [], "encoding": "", "detalle": ""}
    escribir(f"Procesando: {nombre} ...")
    try:
        valida, motivo, driver = validar_cabecera(ruta)
        if not valida and not args.sin_validar:
            res.update(estado="OMITIDO", detalle=motivo)
            escribir(f"  ⏭️  Omitido: {motivo}.")
        else:
            candidatos = candidatos_encoding(args.encoding, driver)
            df, enc = leer_dbf(ruta, candidatos, not args.sin_memo_estricto, not args.sin_strip, not args.sin_barra)
            hojas = escribir_hojas(writer, df, base_hoja(ruta), usadas, args.max_filas)
            res.update(filas=len(df), hojas=hojas, encoding=enc)
            extra = f" (dividida en {len(hojas)} pestañas)" if len(hojas) > 1 else ""
            escribir(f"  ✅ {len(df):,} filas en '{hojas[0]}'{extra} [{enc}].")
            del df
    except Exception as e:  # una tabla dañada no debe detener el resto
        res.update(estado="ERROR", detalle=f"{type(e).__name__}: {e}")
        escribir(f"  ❌ Error al procesar '{nombre}': {e}")
    res["segundos"] = time.perf_counter() - t0
    return res


def guardar_log(ruta_log, args, carpeta, salida, resultados, segundos):
    ok = [r for r in resultados if r["estado"] == "OK"]
    omitidos = [r for r in resultados if r["estado"] == "OMITIDO"]
    errores = [r for r in resultados if r["estado"] == "ERROR"]
    lineas = [
        "REGISTRO DE CONVERSIÓN DBF -> EXCEL",
        f"Fecha: {dt.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Carpeta: {carpeta}" + (" (con subcarpetas)" if args.recursivo else ""),
        f"Excel de salida: {salida}",
        f"Extensiones: {', '.join(args.extensiones)} | Codificación solicitada: {args.encoding}",
        "",
        "DETALLE POR ARCHIVO",
        "-" * 70,
    ]
    for r in resultados:
        lineas.append(
            f"[{r['estado']:<7}] {r['archivo']} | filas: {r['filas']:,} | "
            f"pestañas: {', '.join(r['hojas']) or '-'} | encoding: {r['encoding'] or '-'} | "
            f"{r['segundos']:.1f} s"
        )
        if r["detalle"]:
            lineas.append(f"           motivo: {r['detalle']}")
    lineas += [
        "",
        "RESUMEN",
        "-" * 70,
        f"Convertidos: {len(ok)} | Omitidos (no son tablas DBF): {len(omitidos)} | Con error: {len(errores)}",
        f"Filas totales: {sum(r['filas'] for r in ok):,}",
        f"Tiempo total: {segundos:.1f} s",
    ]
    ruta_log.write_text("\n".join(lineas) + "\n", encoding="utf-8-sig")


def parsear():
    p = argparse.ArgumentParser(description="Convierte tablas DBF/FoxPro a un único Excel.")
    p.add_argument("-i", "--carpeta", default=".", help="Carpeta con los archivos (por defecto: actual)")
    p.add_argument("-o", "--salida", default="Resultados_Consolidados.xlsx", help="Archivo Excel de salida")
    p.add_argument("-r", "--recursivo", action="store_true", help="Buscar también en subcarpetas")
    p.add_argument("-x", "--extensiones", nargs="+", default=[".dbf", ".tmp"], help="Extensiones a convertir (def.: .dbf .tmp)")
    p.add_argument("-e", "--encoding", default="auto", help="Codificación (auto, cp850, cp1252, latin1...). 'auto' usa la declarada en la tabla")
    p.add_argument("--max-filas", type=int, default=MAX_FILAS_EXCEL, help="Filas por pestaña antes de dividir (def.: 1,048,575)")
    p.add_argument("--log", help="Ruta del archivo de registro (def.: junto al Excel, extensión .log)")
    p.add_argument("--sin-memo-estricto", action="store_true", help="No fallar si falta el .FPT/.DBT; esos campos quedan vacíos")
    p.add_argument("--sin-validar", action="store_true", help="No verificar la cabecera DBF antes de leer")
    p.add_argument("--sin-strip", action="store_true", help="No quitar espacios al inicio/fin de los textos")
    p.add_argument("--sin-barra", action="store_true", help="No mostrar barras de progreso")
    return p.parse_args()


def main():
    args = parsear()
    carpeta = Path(args.carpeta).resolve()
    salida = Path(args.salida).resolve()
    ruta_log = Path(args.log).resolve() if args.log else salida.with_suffix(".log")
    extensiones = {("." + e.lstrip(".")).lower() for e in args.extensiones}

    if not carpeta.is_dir():
        print(f"❌ La carpeta '{carpeta}' no existe.")
        return 1
    if not 1 <= args.max_filas <= MAX_FILAS_EXCEL:
        print(f"❌ --max-filas debe estar entre 1 y {MAX_FILAS_EXCEL}.")
        return 1
    if salida.exists():  # Excel bloquea el archivo si lo tienes abierto
        try:
            salida.open("ab").close()
        except PermissionError:
            print(f"❌ No se puede escribir en '{salida.name}'. Si lo tienes abierto en Excel, ciérralo.")
            return 1

    archivos = buscar_archivos(carpeta, extensiones, args.recursivo, {salida})
    if not archivos:
        print(f"❌ No se encontraron archivos {', '.join(sorted(extensiones))} en '{carpeta}'.")
        return 1

    print(f"📦 Se encontraron {len(archivos)} archivos ({', '.join(sorted(extensiones))}) para procesar.\n")

    def nombre_mostrado(ruta):
        try:
            return str(ruta.relative_to(carpeta))
        except ValueError:
            return ruta.name

    resultados, usadas = [], set()
    inicio = time.perf_counter()
    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            # pandas ignora date_format/datetime_format en el writer de openpyxl, por eso se fijan aquí.
            # Si una versión futura cambia estos atributos, solo se vuelve al formato AAAA-MM-DD.
            writer._date_format = FORMATO_FECHA
            writer._datetime_format = FORMATO_FECHA_HORA
            iterador = tqdm(archivos, unit="archivo") if (tqdm and not args.sin_barra) else archivos
            for ruta in iterador:
                resultados.append(procesar(ruta, nombre_mostrado(ruta), writer, usadas, args))

            # openpyxl falla al cerrar si no se escribió ninguna hoja
            if not any(r["estado"] == "OK" for r in resultados):
                pd.DataFrame({"Aviso": ["No se pudo convertir ningún archivo"]}).to_excel(
                    writer, sheet_name="Aviso", index=False
                )
    finally:
        total = time.perf_counter() - inicio
        guardar_log(ruta_log, args, carpeta, salida, resultados, total)

    n_ok = sum(r["estado"] == "OK" for r in resultados)
    n_omit = sum(r["estado"] == "OMITIDO" for r in resultados)
    errores = [r for r in resultados if r["estado"] == "ERROR"]
    print("\n" + "=" * 50)
    print(f"✅ Convertidos: {n_ok}   ⏭️  Omitidos: {n_omit}   ❌ Con error: {len(errores)}   ⏱️  {total:.1f} s")
    for r in errores:
        print(f"   - {r['archivo']}: {r['detalle']}")
    print(f"📄 Excel: '{salida}'")
    print(f"📝 Registro: '{ruta_log}'")
    return 0 if n_ok else 1


if __name__ == "__main__":
    sys.exit(main())
