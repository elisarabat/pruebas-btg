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
from openpyxl.utils import get_column_letter


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
# Fila donde están los nombres de columna en el archivo fuente (0 = primera fila)
FILA_ENCABEZADO_VALO = 2   # fila 3 en Excel (o ver USAR_FILAS_2_Y_3_VALO)
FILA_ENCABEZADO_BASE = 0   # fila 1 en Excel
# Si True, en la hoja Valo los nombres de columna se construyen desde filas 2 y 3:
# para cada columna se usa el valor de la fila 3 si no está vacío, si no el de la fila 2.
# Así las columnas que solo tienen título en la fila 2 dejan de salir como "Unnamed".
USAR_FILAS_2_Y_3_VALO = True
# Llave para enlazar filas de Valo con Base (debe existir en ambas). Si en Base tiene otro nombre, ponlo en COLUMNA_LLAVE_EN_BASE.
LLAVE_PARA_BASE = "Rut"
COLUMNA_LLAVE_EN_BASE = None  # None = buscar columna con mismo nombre normalizado en Base
# 3 columnas del maestro que se rellenan desde la hoja Base: { nombre_en_maestro: nombre_columna_en_Base }
COLUMNAS_DESDE_BASE = {
    "Fecha de emisión": "Fecha de suscripción",
    "Tasa Arriendo o Compra": "Tasa anual de emisión",
    "Tasa Venta": "Tasa anual de endoso",
}

# Mapeo por índice cuando la columna en el fuente tiene nombre vacío (Unnamed).
# Si USAR_FILAS_2_Y_3_VALO = True, muchas de estas columnas ya tendrán nombre (desde la fila 2)
# y no hará falta listarlas aquí. Este dict es fallback por si alguna sigue sin coincidir.
# Clave = columna en el maestro; valor = índice 0-based en la hoja Valo.
COLUMNAS_POR_INDICE_FUENTE = {
    "Monto Crédito + Cap (UF)": 8,
    "Fecha 1er Aporte": 15,
    "Fecha Last Aporte": 16,
    "Fecha Corte": 17,
    "Tasacion": 23,
    "Precio Venta/Tasación": 24,
    "Monto dividendo": 25,
    "Div/Renta": 26,
    "Carga Financiera": 27,
}

# Columnas que son fechas: se convierten a datetime al leer y se formatean al escribir
COLUMNAS_FECHA = (
    "Fecha de compra",
    "Fecha de emisión",
    "Fecha 1er Aporte",
    "Fecha Last Aporte",
    "Fecha Corte",
)

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
    "Monto Crédito + Cap (UF)",
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
    "Div/Renta",
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
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d-%m-%y", "%d/%m/%y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(s, dayfirst=True)
    except (ValueError, TypeError):
        return None
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
    # Unificar variantes de ordinales (1.er, 1 er, 1ª -> 1er)
    s = s.replace("1.er", "1er").replace("1 er", "1er").replace("1. er", "1er")
    s = s.replace("1.ª", "1").replace("1ª", "1")
    # Unificar signos +/− (full-width, minus sign Unicode) y colapsar espacios
    s = s.replace("＋", "+").replace("－", "-").replace("−", "-")
    s = " ".join(s.split())
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


def _maestro_a_columnas_estandar(df_maestro: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el DataFrame leído del maestro a uno con columnas exactamente COLUMNAS_MAESTRO,
    mapeando por nombre normalizado. Así no se pierden datos de filas ya existentes cuando
    el Excel tiene nombres con espacios extra, acentos distintos, etc.
    """
    out = pd.DataFrame(index=df_maestro.index)
    for col_maestro in COLUMNAS_MAESTRO:
        col_real = _buscar_columna_en_df(df_maestro, col_maestro)
        if col_real is not None:
            out[col_maestro] = df_maestro[col_real].values
        else:
            out[col_maestro] = pd.NA
    return out


def _normalizar_rut_para_merge(val) -> str:
    """
    Normaliza un RUT para comparación: quita puntos y deja formato '12345678-9'.
    Acepta RUT con guion y DV (ej. '12.345.678-9') o solo número.
    """
    if pd.isna(val):
        return ""
    s = str(val).strip().replace(".", "").upper()
    return s


def _construir_llave_rut_desde_separados(rut_serie: pd.Series, dv_serie: pd.Series) -> pd.Series:
    """
    Construye una llave RUT normalizada desde columnas Rut y DV separadas (hoja Valo).
    Formato resultado: '12345678-9' (sin puntos, DV en mayúscula).
    """
    rut = rut_serie.astype(str).str.replace(".", "", regex=False).str.strip()
    dv = dv_serie.astype(str).str.strip().str.upper()
    llave = rut + "-" + dv
    # Evitar que valores con nan coincidan con algo en Base
    llave = llave.replace(["nan-nan", "nan-NAN", "NaN-NaN"], "")
    return llave


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

    # Normalizar llave para el merge: en Valo Rut y DV suelen estar separados; en Base vienen "12345678-9"
    if llave_maestro == "Rut" and "DV" in df_out.columns:
        df_out["_llave_merge_"] = _construir_llave_rut_desde_separados(df_out["Rut"], df_out["DV"])
    else:
        df_out["_llave_merge_"] = df_out[llave_maestro].apply(_normalizar_rut_para_merge)

    df_lookup["_llave_merge_"] = df_lookup[llave_base].apply(_normalizar_rut_para_merge)
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
                "monto credito + cap (uf)": [
                    "monto credito + cap (uf)", "monto credito - cap (uf)",
                    "monto credito uf", "credito uf", "monto cap",
                    "monto credito+cap (uf)", "monto credito + cap(uf)",
                ],
                "subsidio": ["subsidio"],
                "pie": ["pie", "pie inicial"],
                "valor vivienda": ["valor vivienda", "valor_vivienda", "valor propiedad"],
                "morosidad": ["morosidad", "dias morosidad"],
                "n° cuotas": ["n cuotas", "numero cuotas", "cuotas", "nº cuotas"],
                "cuota mes": ["cuota mes", "cuota", "cuota mensual", "primera cuota a endosar"],
                "tasa arriendo o compra": ["tasa arriendo", "tasa compra", "tasa"],
                "fecha 1er aporte": ["fecha primer aporte", "1er aporte", "fecha 1er aporte", "fecha 1er aporte a endosar"],
                "fecha last aporte": ["fecha ultimo aporte", "fecha ultimo aporte a endosar", "last aporte", "fecha last aporte"],
                "fecha corte": ["fecha corte", "fecha de corte", "corte"],
                "saldo insoluto teorico al 31-07-2019": ["saldo insoluto teorico", "saldo insoluto", "saldo teorico", "saldo 31-07-2019"],
                "tasacion": ["tasacion", "tasación", "valor tasacion", "tasacion de la propiedad"],
                "precio venta/tasacion": ["precio venta", "precio tasacion", "precio venta tasacion", "precio venta/tasacion"],
                "tasa venta": ["tasa venta", "tasa_venta"],
                "dif. tasa": ["dif tasa", "diferencia tasa", "dif tasa"],
                "monto dividendo": ["monto dividendo", "dividendo", "monto dividendo"],
                "div/renta": ["div/renta", "divrenta", "dividendo/renta", "dividendo/ renta"],
                "divrenta": ["div/renta", "divrenta", "dividendo/renta", "dividendo/ renta"],
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
            # Fallback: columna del fuente cuyo nombre normalizado empieza por el del maestro
            # (ej. "Fecha 1er Aporte" -> "Fecha 1er Aporte a endosar"; "1.er" ya unificado en normalizar)
            if col_maestro not in mapeo:
                candidatos = [(k, normalizados_fuente[k]) for k in normalizados_fuente if k.startswith(norm_maestro)]
                # Prefijos alternativos (ej. fuente tiene "Último" y maestro "Last")
                prefijos_alternativos = {
                    "fecha last aporte": ["fecha ultimo aporte"],
                    "fecha corte": ["fecha de corte", "corte"],
                }
                if norm_maestro in prefijos_alternativos:
                    for prefijo in prefijos_alternativos[norm_maestro]:
                        candidatos.extend(
                            [(k, normalizados_fuente[k]) for k in normalizados_fuente if k.startswith(prefijo)]
                        )
                if candidatos:
                    # Elegir la coincidencia más específica (nombre más largo)
                    mejor = max(candidatos, key=lambda x: len(x[0]))
                    mapeo[col_maestro] = mejor[1]
            # Fallback columnas de fecha (1er/Last Aporte, Corte): buscar por contenido del nombre
            if col_maestro not in mapeo:
                if norm_maestro == "fecha 1er aporte":
                    candidatos = [(k, normalizados_fuente[k]) for k in normalizados_fuente if "fecha" in k and "aporte" in k and ("1er" in k or "1.er" in k or "primer" in k)]
                elif norm_maestro == "fecha last aporte":
                    candidatos = [(k, normalizados_fuente[k]) for k in normalizados_fuente if "fecha" in k and "aporte" in k and ("last" in k or "ultimo" in k)]
                elif norm_maestro == "fecha corte":
                    candidatos = [(k, normalizados_fuente[k]) for k in normalizados_fuente if ("fecha" in k and "corte" in k) or k == "corte" or k.startswith("corte ")]
                else:
                    candidatos = []
                if candidatos:
                    mejor = max(candidatos, key=lambda x: len(x[0]))
                    mapeo[col_maestro] = mejor[1]
            # Fallback Tasacion: columna en fuente puede llamarse "Tasación", "Valor Tasación", etc.
            if col_maestro not in mapeo and norm_maestro == "tasacion":
                candidatos = [
                    (k, normalizados_fuente[k]) for k in normalizados_fuente
                    if "tasacion" in k and "precio" not in k
                ]
                if candidatos:
                    mejor = max(candidatos, key=lambda x: len(x[0]))
                    mapeo[col_maestro] = mejor[1]
            # Fallback Monto Crédito + Cap (UF): nombre en fuente puede variar (+/- , espacios)
            if col_maestro not in mapeo and norm_maestro == "monto credito + cap (uf)":
                candidatos = [
                    (k, normalizados_fuente[k]) for k in normalizados_fuente
                    if "monto credito" in k and "cap" in k and "uf" in k
                ]
                if candidatos:
                    mejor = max(candidatos, key=lambda x: len(x[0]))
                    mapeo[col_maestro] = mejor[1]
            # Fallback por índice: columnas del fuente con encabezado vacío (Unnamed)
            if col_maestro not in mapeo and col_maestro in COLUMNAS_POR_INDICE_FUENTE:
                idx = COLUMNAS_POR_INDICE_FUENTE[col_maestro]
                if 0 <= idx < len(df_fuente.columns):
                    mapeo[col_maestro] = df_fuente.columns[idx]
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
    print("  (índice = posición 0-based para COLUMNAS_POR_INDICE_FUENTE en cargar_a_maestro.py)")
    sin_uso_con_indice = [(i, c) for i, c in enumerate(columnas_fuente) if c not in usadas]
    if sin_uso_con_indice:
        for idx, c in sin_uso_con_indice:
            print(f"  - índice {idx}: {c!r}")
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
    # 1ª col. fecha del fuente -> VPN 2da fecha; 2ª col. fecha -> VPN 1ra fecha (según convención del maestro)
    if col_1ra is not None and "VPN 2da fecha" in df_out.columns:
        df_out["VPN 2da fecha"] = df_fuente[col_1ra].values
    if col_2da is not None and "VPN 1ra fecha" in df_out.columns:
        df_out["VPN 1ra fecha"] = df_fuente[col_2da].values
    # Precio -1UF = valor de la columna de la segunda fecha - 1
    if col_2da is not None and "Precio -1UF" in df_out.columns:
        serie = df_fuente[col_2da].astype(str).str.replace(",", ".", regex=False)
        valores = pd.to_numeric(serie, errors="coerce")
        df_out["Precio -1UF"] = valores - 1


def _convertir_columna_a_fecha(serie: pd.Series) -> pd.Series:
    """Convierte una serie a datetime; acepta texto dd/mm/yyyy, números serial Excel y datetime."""
    if serie.empty:
        return serie
    # Si ya es datetime64, solo normalizar
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_datetime(serie, errors="coerce").dt.normalize()
    # Números que parecen serial Excel (días desde 1899-12-30): típicamente 1000–100000
    numeric = pd.to_numeric(serie, errors="coerce")
    if numeric.notna().any():
        sample = numeric.dropna()
        if (sample >= 1000).all() and (sample <= 100000).all():
            try:
                fechas = pd.TimedeltaIndex(numeric.astype(float), unit="d") + pd.Timestamp("1899-12-30")
                return pd.Series(fechas.normalize(), index=serie.index)
            except (ValueError, TypeError):
                pass
    # Texto o resto: convertir (dayfirst para formato chileno dd/mm/yyyy)
    return pd.to_datetime(serie, dayfirst=True, errors="coerce").dt.normalize()


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
    # Convertir columnas de fecha para que no queden en blanco (texto o número serial Excel → datetime)
    for col in COLUMNAS_FECHA:
        if col in df_out.columns:
            df_out[col] = _convertir_columna_a_fecha(df_out[col])
    aplicar_columnas_calculadas(df_out)
    rellenar_vpn_desde_columnas_fecha(df_out, df_fuente)
    return df_out


def _leer_valo_con_filas_2_y_3(ruta_fuente: str) -> pd.DataFrame:
    """
    Lee la hoja Valo sin usar una fila fija como encabezado; construye los nombres
    de columna desde las filas 2 y 3 (Excel): para cada columna usa el valor de
    la fila 3 si no está vacío, si no el de la fila 2. Los datos empiezan en la fila 4.
    """
    try:
        df_raw = pd.read_excel(ruta_fuente, sheet_name=HOJA_VALO, header=None)
    except ValueError:
        df_raw = pd.read_excel(ruta_fuente, sheet_name=0, header=None)
    if df_raw.empty or len(df_raw) < 3:
        return pd.DataFrame()
    # Excel fila 2 = índice 1, Excel fila 3 = índice 2, datos desde fila 4 = índice 3
    row2 = df_raw.iloc[1]
    row3 = df_raw.iloc[2]
    headers = []
    for i in range(len(df_raw.columns)):
        v3 = row3.iloc[i]
        v2 = row2.iloc[i]
        if pd.notna(v3) and str(v3).strip():
            headers.append(str(v3).strip())
        elif pd.notna(v2) and str(v2).strip():
            headers.append(str(v2).strip())
        else:
            headers.append(f"Unnamed: {i}")
    df_fuente = df_raw.iloc[3:].copy()
    df_fuente.columns = headers
    df_fuente.reset_index(drop=True, inplace=True)
    return df_fuente


def cargar_y_agregar_a_maestro(ruta_fuente: str, ruta_maestro: str, fecha_compra: str = None) -> None:
    """
    Lee el archivo fuente, lo mapea al formato maestro, lo concatena al maestro
    y guarda el maestro sin duplicados (llave: Rut + Fecha de emisión).
    Si se indica fecha_compra, se usa para todas las filas nuevas (sustituye al archivo fuente).
    """
    if not os.path.isfile(ruta_fuente):
        raise FileNotFoundError(f"No se encontró el archivo fuente: {ruta_fuente}")

    # Leer hoja Valo (datos principales)
    if USAR_FILAS_2_Y_3_VALO:
        df_fuente = _leer_valo_con_filas_2_y_3(ruta_fuente)
    else:
        try:
            df_fuente = pd.read_excel(ruta_fuente, sheet_name=HOJA_VALO, header=FILA_ENCABEZADO_VALO)
        except ValueError:
            df_fuente = pd.read_excel(ruta_fuente, sheet_name=0, header=FILA_ENCABEZADO_VALO)
    if df_fuente.empty:
        print("El archivo fuente no tiene filas en la hoja de datos. No se agrega nada.")
        return

    # Leer hoja Base (para columnas adicionales); nombres de columna en fila 1
    df_base = None
    try:
        df_base = pd.read_excel(ruta_fuente, sheet_name=HOJA_BASE, header=FILA_ENCABEZADO_BASE)
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
        df_maestro_raw = pd.read_excel(ruta_maestro, sheet_name=0, header=1)
        # Mapear por nombre normalizado para no perder datos de filas ya existentes
        # (si el Excel tiene "Comuna " o "Morosidad " con espacio, se reconoce igual)
        df_maestro = _maestro_a_columnas_estandar(df_maestro_raw)
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

    # Fecha de emisión: solo fecha, sin hora
    if "Fecha de emisión" in df_sin_duplicados.columns:
        df_sin_duplicados["Fecha de emisión"] = (
            pd.to_datetime(df_sin_duplicados["Fecha de emisión"], errors="coerce").dt.normalize()
        )

    # Columnas de tasa que deben mostrarse en porcentaje (Tasa endoso = Tasa Venta; Tasa Arriendo también)
    COLUMNAS_PORCENTAJE = ("Tasa Arriendo o Compra", "Tasa Venta", "Dif. Tasa")

    # Guardar (cabecera en fila 2 para mantener formato del maestro)
    with pd.ExcelWriter(ruta_maestro, engine="openpyxl") as writer:
        df_sin_duplicados.to_excel(writer, index=False, sheet_name="Sheet1", startrow=1)
        ws = writer.sheets["Sheet1"]
        for col_name in COLUMNAS_PORCENTAJE:
            if col_name not in df_sin_duplicados.columns:
                continue
            col_idx = df_sin_duplicados.columns.get_loc(col_name) + 1
            col_letter = get_column_letter(col_idx)
            for row in range(2, len(df_sin_duplicados) + 2):
                cell = ws[f"{col_letter}{row}"]
                val = cell.value
                if isinstance(val, (int, float)) and not pd.isna(val):
                    if abs(val) > 1:
                        cell.value = val / 100
                    cell.number_format = "0.00%"
        # Todas las columnas de fecha: formato dd/mm/yyyy para que se vean en Excel
        for col_name in COLUMNAS_FECHA:
            if col_name not in df_sin_duplicados.columns:
                continue
            col_idx = df_sin_duplicados.columns.get_loc(col_name) + 1
            col_letter = get_column_letter(col_idx)
            for row in range(2, len(df_sin_duplicados) + 2):
                ws[f"{col_letter}{row}"].number_format = "dd/mm/yyyy"
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
