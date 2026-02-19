# -*- coding: utf-8 -*-
"""
Script para cargar datos desde archivos Excel fuente hacia un archivo maestro.
- Evita duplicados por Rut + Fecha de emisión.
- Mapea columnas aunque los nombres no coincidan exactamente.
"""

import pandas as pd
import sys
import os
import unicodedata


# Columnas del maestro que se calculan desde otras (no vienen del archivo fuente)
COLUMNAS_CALCULADAS = {
    "ID Blotter": "N° OP",  # dígitos de N° OP antes de la primera letra
    "Dif. Tasa": "Tasa Arriendo o Compra - Tasa Venta",
}
# Columnas que se rellenan desde columnas del fuente (fechas y columna anterior a 2ª fecha)
COLUMNAS_DESDE_FECHA = ("VPN 1ra fecha", "VPN 2da fecha", "Precio -1UF")
# Hojas del Excel fuente: datos principales y hoja auxiliar
HOJA_VALO = "Valo"
HOJA_BASE = "Base"
# Llave para enlazar filas de Valo con Base (debe existir en ambas). Si en Base tiene otro nombre, ponlo en COLUMNA_LLAVE_EN_BASE.
LLAVE_PARA_BASE = "Rut"
COLUMNA_LLAVE_EN_BASE = None  # None = buscar columna con mismo nombre normalizado en Base
# 3 columnas del maestro que se rellenan desde la hoja Base: { nombre_en_maestro: nombre_columna_en_Base }
COLUMNAS_DESDE_BASE = {
    "Fecha de emisión": "Fecha de suscripción",
    "Tasa Arriendo o Compra": "Tasa anual de emisión",
    "Tasa Venta": "Tasa anual de endoso",
}

# Columnas del archivo maestro (orden exacto)
COLUMNAS_MAESTRO = [
    "Fecha de compra",
    "N°",
    "N° OP",
    "ID Blotter",
    "Apellido Paterno",
    "Apellido materno",
    "Nombres",
    "Rut",
    "DV",
    "Fecha de emisión",
    "Monto Crédito - Cap (UF)",
    "Subsidio",
    "Pie",
    "Valor Vivienda",
    "Morosidad",
    "N° Cuotas",
    "Cuota Mes",
    "Tasa Arriendo o Compra",
    "Fecha 1er Aporte",
    "Fecha Last Aporte",
    "Fecha Corte",
    "VPN 1ra fecha",   # valor de la columna del fuente cuya cabecera es la primera fecha (orden cronológico)
    "VPN 2da fecha",   # valor de la columna cuya cabecera es la segunda fecha (siempre > primera)
    "Precio -1UF",     # valor de la columna de la 2ª fecha menos 1 (ej. col 09-02-2024 - 1)
    "Saldo insoluto Teorico al 31-07-2019",
    "Tasacion",
    "Precio Venta/Tasación",
    "Tasa Venta",
    "Dif. Tasa",
    "Monto dividendo",
    "DivRenta",
    "Carga Financiera",
    "Dirección",
    "Comuna",
]


def extraer_id_blotter_desde_n_op(val) -> str:
    """
    Extrae el ID Blotter desde el valor de N° OP: todos los dígitos
    al inicio, antes de la primera letra. Ej: '12345ABC' -> '12345', '987' -> '987'.
    """
    if pd.isna(val):
        return ""
    s = str(val).strip()
    numeros = []
    for c in s:
        if c.isdigit():
            numeros.append(c)
        else:
            break
    return "".join(numeros)


def _parsear_nombre_como_fecha(nombre) -> pd.Timestamp | None:
    """Intenta parsear el nombre de columna como fecha. Devuelve Timestamp o None."""
    if pd.isna(nombre) or str(nombre).strip() == "":
        return None
    s = str(nombre).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d-%m-%y", "%d/%m/%y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(s, dayfirst=True)
    except (ValueError, TypeError):
        return None


def obtener_columnas_fecha_vpn(df_fuente: pd.DataFrame) -> tuple:
    """
    Encuentra columnas del fuente cuyo encabezado es una fecha.
    Devuelve (col_primera_fecha, col_segunda_fecha) ordenadas por fecha ascendente
    (la segunda fecha siempre será mayor que la primera). Si hay menos de 2, devuelve None.
    """
    candidatos = []
    for col in df_fuente.columns:
        d = _parsear_nombre_como_fecha(col)
        if d is not None and not pd.isna(d):
            candidatos.append((col, d))
    candidatos.sort(key=lambda x: x[1])
    if len(candidatos) >= 2:
        return candidatos[0][0], candidatos[1][0]
    if len(candidatos) == 1:
        return candidatos[0][0], None
    return None, None


def normalizar_nombre(col: str) -> str:
    """Normaliza nombre de columna para comparación: minúsculas, sin acentos, sin espacios extra."""
    if pd.isna(col) or not isinstance(col, str):
        return ""
    s = str(col).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s


def _buscar_columna_en_df(df: pd.DataFrame, nombre: str) -> str | None:
    """Devuelve el nombre real de la columna en df que coincide con nombre (normalizado), o None."""
    if nombre is None or df is None or df.empty:
        return None
    norm = normalizar_nombre(nombre)
    for c in df.columns:
        if normalizar_nombre(c) == norm:
            return c
    return None


def rellenar_desde_hoja_base(df_out: pd.DataFrame, df_base: pd.DataFrame) -> None:
    """
    Rellena las columnas definidas en COLUMNAS_DESDE_BASE con los valores de la hoja Base,
    enlazando por la columna LLAVE_PARA_BASE (ej. Rut). Si hay varias filas en Base con la misma
    llave, se usa la primera. Modifica df_out in-place.
    """
    if not COLUMNAS_DESDE_BASE or df_base is None or df_base.empty:
        return
    llave_maestro = LLAVE_PARA_BASE
    if llave_maestro not in df_out.columns:
        return
    # Nombre de la columna llave en Base
    llave_base = COLUMNA_LLAVE_EN_BASE if COLUMNA_LLAVE_EN_BASE else _buscar_columna_en_df(df_base, llave_maestro)
    if llave_base is None:
        return
    # Columnas de Base que necesitamos (maestro -> nombre en Base)
    cols_base = [llave_base]
    mapeo_base_a_maestro = {}
    for col_maestro, col_base_nombre in COLUMNAS_DESDE_BASE.items():
        col_base_real = _buscar_columna_en_df(df_base, col_base_nombre) if col_base_nombre else None
        if col_base_real and col_maestro in df_out.columns:
            cols_base.append(col_base_real)
            mapeo_base_a_maestro[col_base_real] = col_maestro
    cols_base = list(dict.fromkeys(cols_base))
    df_lookup = df_base[cols_base].drop_duplicates(subset=[llave_base], keep="first")
    # Normalizar llave para el merge (ambos a string)
    df_out["_llave_merge_"] = df_out[llave_maestro].astype(str).str.strip()
    df_lookup["_llave_merge_"] = df_lookup[llave_base].astype(str).str.strip()
    merged = df_out[["_llave_merge_"]].merge(
        df_lookup.drop(columns=[llave_base]),
        on="_llave_merge_",
        how="left",
        suffixes=("", "_base"),
    )
    for col_base_real, col_maestro in mapeo_base_a_maestro.items():
        if col_base_real in merged.columns:
            df_out[col_maestro] = merged[col_base_real].values
    df_out.drop(columns=["_llave_merge_"], inplace=True)


def mapear_columnas_fuente_a_maestro(df_fuente: pd.DataFrame) -> dict:
    """
    Devuelve un diccionario: nombre_columna_maestro -> nombre_columna_en_fuente.
    Si no hay coincidencia, la columna maestro no estará en el dict (quedará NaN).
    """
    nombres_fuente = list(df_fuente.columns)
    normalizados_fuente = {normalizar_nombre(c): c for c in nombres_fuente}

    mapeo = {}
    for col_maestro in COLUMNAS_MAESTRO:
        norm_maestro = normalizar_nombre(col_maestro)
        if norm_maestro in normalizados_fuente:
            mapeo[col_maestro] = normalizados_fuente[norm_maestro]
        else:
            # Variantes comunes (puedes ampliar esta lista)
            variantes = {
                "fecha de compra": ["fecha compra", "fecha_compra", "fechacompra"],
                "n°": ["n", "numero", "num", "nº", "no"],
                "n° op": ["n op", "n op.", "numero op", "nop", "n° operacion"],
                "id blotter": ["id blotter", "id_blotter", "blotter"],
                "apellido paterno": ["apellido paterno", "ap paterno", "paterno"],
                "apellido materno": ["apellido materno", "ap materno", "materno"],
                "nombres": ["nombres", "nombre completo"],
                "rut": ["rut", "run", "rut cliente"],
                "dv": ["dv", "digito verificador"],
                "fecha de emision": ["fecha emision", "fecha_emision", "fecha emisión"],
                "monto credito + cap (uf)": ["monto credito + cap (uf)", "monto credito uf", "credito uf", "monto cap"],
                "subsidio": ["subsidio"],
                "pie": ["pie", "pie inicial"],
                "valor vivienda": ["valor vivienda", "valor_vivienda", "valor propiedad"],
                "morosidad": ["morosidad", "dias morosidad"],
                "n° cuotas": ["n cuotas", "numero cuotas", "cuotas", "nº cuotas"],
                "cuota mes": ["cuota mes", "cuota", "cuota mensual", "primera cuota a endosar"],
                "tasa arriendo o compra": ["tasa arriendo", "tasa compra", "tasa"],
                "fecha 1er aporte": ["fecha primer aporte", "1er aporte", "fecha 1er aporte", "fecha 1er aporte a endosar"],
                "fecha last aporte": ["fecha ultimo aporte", "last aporte", "fecha last aporte"],
                "fecha corte": ["fecha corte", "corte"],
                "saldo insoluto teorico al 31-07-2019": ["saldo insoluto", "saldo teorico", "saldo 31-07-2019"],
                "tasacion": ["tasacion", "tasación"],
                "precio venta/tasacion": ["precio venta", "precio tasacion", "precio venta tasacion", "precio venta/tasacion"],
                "tasa venta": ["tasa venta", "tasa_venta"],
                "dif. tasa": ["dif tasa", "diferencia tasa", "dif tasa"],
                "monto dividendo": ["monto dividendo", "dividendo", "monto dividendo"],
                "div/renta": ["div/renta", "divrenta", "dividendo/renta"],
                "carga financiera": ["carga financiera", "carga_financiera", "carga financiera/ renta"],
                "direccion": ["direccion", "dirección", "domicilio", "direccion de la propiedad"],
                "comuna": ["comuna"],
            }
            clave = norm_maestro
            if clave in variantes:
                for variante in variantes[clave]:
                    if variante in normalizados_fuente:
                        mapeo[col_maestro] = normalizados_fuente[variante]
                        break
    return mapeo


def imprimir_reporte_mapeo(mapeo: dict, columnas_fuente: list, columnas_fuente_adicionales_usadas: tuple = ()) -> None:
    """
    Imprime qué columnas del archivo fuente se mapearon a cada columna del maestro
    y cuáles del maestro quedaron sin mapear. También lista columnas del fuente no usadas.
    columnas_fuente_adicionales_usadas: columnas usadas por lógica (ej. columnas con fecha para VPN).
    """
    columnas_fuente = [c for c in columnas_fuente if isinstance(c, str) or not pd.isna(c)]
    usadas = set(mapeo.values()) | set(c for c in columnas_fuente_adicionales_usadas if c is not None)

    print("\n" + "=" * 60)
    print("REPORTE DE MAPEO DE COLUMNAS")
    print("=" * 60)
    print("\n--- Columnas del maestro (mapeadas) ---")
    for col_maestro in COLUMNAS_MAESTRO:
        if col_maestro in mapeo:
            col_fuente = mapeo[col_maestro]
            print(f"  [OK] {col_maestro!r}")
            print(f"       <- {col_fuente!r}")
        elif col_maestro in COLUMNAS_CALCULADAS:
            print(f"  [CALCULADO] {col_maestro!r} (desde {COLUMNAS_CALCULADAS[col_maestro]!r})")
        elif col_maestro in COLUMNAS_DESDE_FECHA:
            print(f"  [DESDE COL. FECHA] {col_maestro!r} (1ª/2ª col. del fuente con cabecera fecha)")
        elif col_maestro in COLUMNAS_DESDE_BASE:
            print(f"  [DESDE HOJA BASE] {col_maestro!r} <- {COLUMNAS_DESDE_BASE[col_maestro]!r}")
        else:
            print(f"  [SIN MAPEAR] {col_maestro!r} (quedará vacío en filas nuevas)")
    print("\n--- Columnas del archivo fuente no usadas ---")
    sin_uso = [c for c in columnas_fuente if c not in usadas]
    if sin_uso:
        for c in sin_uso:
            print(f"  - {c!r}")
    else:
        print("  (ninguna; todas las columnas del fuente se usaron)")
    print("=" * 60 + "\n")


def aplicar_columnas_calculadas(df_out: pd.DataFrame) -> None:
    """
    Rellena columnas del maestro que no vienen del archivo fuente,
    calculándolas a partir de otras columnas. Modifica df_out in-place.
    """
    # ID Blotter = todos los dígitos de N° OP antes de la primera letra
    if "N° OP" in df_out.columns:
        id_blotter = df_out["N° OP"].map(extraer_id_blotter_desde_n_op)
        df_out["ID Blotter"] = id_blotter.replace("", pd.NA)

    # Dif. Tasa = Tasa Arriendo o Compra - Tasa Venta
    if "Tasa Arriendo o Compra" in df_out.columns and "Tasa Venta" in df_out.columns:
        tasa_arriendo = pd.to_numeric(df_out["Tasa Arriendo o Compra"], errors="coerce")
        tasa_venta = pd.to_numeric(df_out["Tasa Venta"], errors="coerce")
        df_out["Dif. Tasa"] = tasa_arriendo - tasa_venta


def rellenar_vpn_desde_columnas_fecha(df_out: pd.DataFrame, df_fuente: pd.DataFrame) -> None:
    """
    Rellena 'VPN 1ra fecha', 'VPN 2da fecha' y 'Precio -1UF' desde el fuente:
    - VPN 1ra/2da = columnas cuyo encabezado es la 1ª y 2ª fecha (orden cronológico).
    - Precio -1UF = valor de la columna de la 2ª fecha menos 1.
    """
    col_1ra, col_2da = obtener_columnas_fecha_vpn(df_fuente)
    if col_1ra is not None and "VPN 1ra fecha" in df_out.columns:
        df_out["VPN 1ra fecha"] = df_fuente[col_1ra].values
    if col_2da is not None and "VPN 2da fecha" in df_out.columns:
        df_out["VPN 2da fecha"] = df_fuente[col_2da].values
    # Precio -1UF = valor de la columna de la segunda fecha - 1
    if col_2da is not None and "Precio -1UF" in df_out.columns:
        serie = df_fuente[col_2da].astype(str).str.replace(",", ".", regex=False)
        valores = pd.to_numeric(serie, errors="coerce")
        df_out["Precio -1UF"] = valores - 1


def dataframe_fuente_a_formato_maestro(df_fuente: pd.DataFrame, mapeo: dict = None) -> pd.DataFrame:
    """Construye un DataFrame con las columnas del maestro, rellenando desde el fuente según mapeo."""
    if mapeo is None:
        mapeo = mapear_columnas_fuente_a_maestro(df_fuente)
    df_out = pd.DataFrame(columns=COLUMNAS_MAESTRO)
    for col_maestro in COLUMNAS_MAESTRO:
        if col_maestro in mapeo:
            df_out[col_maestro] = df_fuente[mapeo[col_maestro]].values
        else:
            df_out[col_maestro] = pd.NA
    aplicar_columnas_calculadas(df_out)
    rellenar_vpn_desde_columnas_fecha(df_out, df_fuente)
    return df_out


def cargar_y_agregar_a_maestro(ruta_fuente: str, ruta_maestro: str, fecha_compra: str = None) -> None:
    """
    Lee el archivo fuente, lo mapea al formato maestro, lo concatena al maestro
    y guarda el maestro sin duplicados (llave: Rut + Fecha de emisión).
    Si se indica fecha_compra, se usa para todas las filas nuevas (sustituye al archivo fuente).
    """
    if not os.path.isfile(ruta_fuente):
        raise FileNotFoundError(f"No se encontró el archivo fuente: {ruta_fuente}")

    # Leer hoja Valo (datos principales); si no existe, usar la primera
    try:
        df_fuente = pd.read_excel(ruta_fuente, sheet_name=HOJA_VALO)
    except ValueError:
        df_fuente = pd.read_excel(ruta_fuente, sheet_name=0)
    if df_fuente.empty:
        print("El archivo fuente no tiene filas en la hoja de datos. No se agrega nada.")
        return

    # Leer hoja Base (para columnas adicionales)
    df_base = None
    try:
        df_base = pd.read_excel(ruta_fuente, sheet_name=HOJA_BASE)
    except ValueError:
        pass  # No hay hoja Base o tiene otro nombre

    # Mapeo y reporte
    mapeo = mapear_columnas_fuente_a_maestro(df_fuente)
    col_1ra, col_2da = obtener_columnas_fecha_vpn(df_fuente)
    imprimir_reporte_mapeo(mapeo, list(df_fuente.columns), columnas_fuente_adicionales_usadas=(col_1ra, col_2da))

    # Convertir al formato del maestro (desde Valo)
    df_nuevo = dataframe_fuente_a_formato_maestro(df_fuente, mapeo=mapeo)

    # Rellenar 3 columnas desde la hoja Base (enlace por Rut u otra llave)
    if df_base is not None and not df_base.empty and COLUMNAS_DESDE_BASE:
        rellenar_desde_hoja_base(df_nuevo, df_base)

    # Fecha de compra como input: se aplica a todas las filas nuevas
    if fecha_compra is not None and fecha_compra.strip():
        df_nuevo["Fecha de compra"] = fecha_compra.strip()

    # Leer maestro existente (los nombres de columna están en la fila 2 del Excel)
    if os.path.isfile(ruta_maestro):
        df_maestro = pd.read_excel(ruta_maestro, sheet_name=0, header=1)
        # Asegurar mismo orden de columnas
        for c in COLUMNAS_MAESTRO:
            if c not in df_maestro.columns:
                df_maestro[c] = pd.NA
        df_maestro = df_maestro[COLUMNAS_MAESTRO]
    else:
        df_maestro = pd.DataFrame(columns=COLUMNAS_MAESTRO)

    # Unir
    df_combinado = pd.concat([df_maestro, df_nuevo], ignore_index=True)

    # Normalizar llave para deduplicar (Rut + Fecha de emisión)
    def llave(r):
        rut = r["Rut"] if pd.notna(r["Rut"]) else ""
        fec = r["Fecha de emisión"] if pd.notna(r["Fecha de emisión"]) else ""
        return (str(rut).strip(), str(fec).strip())

    df_combinado["_llave_"] = df_combinado.apply(llave, axis=1)
    df_sin_duplicados = df_combinado.drop_duplicates(subset=["_llave_"], keep="first")
    df_sin_duplicados = df_sin_duplicados.drop(columns=["_llave_"])

    # Reordenar por si acaso
    df_sin_duplicados = df_sin_duplicados[COLUMNAS_MAESTRO]

    # Guardar (cabecera en fila 2 para mantener formato del maestro)
    with pd.ExcelWriter(ruta_maestro, engine="openpyxl") as writer:
        df_sin_duplicados.to_excel(writer, index=False, sheet_name="Sheet1", startrow=1)
    print(f"Maestro actualizado: {ruta_maestro}")
    print(f"  Filas en maestro antes: {len(df_maestro)}")
    print(f"  Filas nuevas leídas: {len(df_nuevo)}")
    print(f"  Duplicados eliminados: {len(df_combinado) - len(df_sin_duplicados)}")
    print(f"  Total filas en maestro ahora: {len(df_sin_duplicados)}")


def main():
    if len(sys.argv) < 2:
        print("Uso: python cargar_a_maestro.py <ruta_archivo_fuente.xlsx> [ruta_maestro.xlsx] [fecha_compra]")
        print("  Si no se indica ruta_maestro, se usa 'maestro.xlsx' en la misma carpeta del script.")
        print("  Si no se indica fecha_compra, se pedirá por consola (ej. 2024-01-15).")
        sys.exit(1)

    ruta_fuente = sys.argv[1].strip()
    if len(sys.argv) >= 3:
        ruta_maestro = sys.argv[2].strip()
    else:
        carpeta = os.path.dirname(os.path.abspath(__file__))
        ruta_maestro = os.path.join(carpeta, "maestro.xlsx")

    # Fecha de compra: por argumento o por input
    if len(sys.argv) >= 4:
        fecha_compra = sys.argv[3].strip()
    else:
        fecha_compra = input("Fecha de compra (ej. 2024-01-15 o 15/01/2024): ").strip()

    try:
        cargar_y_agregar_a_maestro(ruta_fuente, ruta_maestro, fecha_compra=fecha_compra or None)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
