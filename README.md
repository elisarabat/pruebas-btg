# Cargar a Maestro

Script para cargar datos desde un Excel fuente (hojas **Valo** y **Base**) hacia un archivo maestro. Evita duplicados por Rut + Fecha de emisión y mapea columnas aunque los nombres no coincidan exactamente.

---

## Requisitos

- Python 3.8 o superior
- Dependencias: `pandas`, `openpyxl`

---

## Instalación (solo la primera vez)

Abre la terminal en la carpeta del proyecto y ejecuta:

```bash
pip install -r requirements.txt
```

---

## Cómo ejecutar

### Uso básico

```bash
python cargar_a_maestro.py "ruta\de\tu\archivo_fuente.xlsx"
```

El script te pedirá la **Fecha de compra** por consola (ej. `2024-02-19` o `19/02/2024`). El maestro se guarda como **`maestro.xlsx`** en la misma carpeta del script.

### Con todos los argumentos

```bash
python cargar_a_maestro.py "archivo_fuente.xlsx" "maestro.xlsx" "2024-02-19"
```

| Argumento | Obligatorio | Descripción |
|-----------|-------------|-------------|
| 1º | Sí | Ruta del archivo Excel fuente (debe tener hojas **Valo** y **Base**). |
| 2º | No | Ruta del archivo maestro. Si no se indica, se usa `maestro.xlsx` en la carpeta del script. |
| 3º | No | Fecha de compra. Si no se indica, se pide por consola. |

**Ejemplo sin preguntas** (todo por argumentos):

```bash
python cargar_a_maestro.py "mis_datos.xlsx" "maestro.xlsx" "2024-02-19"
```

---

## Qué necesitas tener

- **Archivo Excel fuente** con al menos la hoja **Valo** (datos principales). Si tiene hoja **Base**, se usará para 3 columnas (Fecha de emisión, Tasa Arriendo o Compra, Tasa Venta), enlazando por **Rut**.
- **Archivo maestro**: puede no existir; si no existe, el script lo crea. Si existe, la cabecera debe estar en la **fila 2** del Excel.

---

## Qué hace el script

1. Lee la hoja **Valo** del archivo fuente (o la primera hoja si no existe "Valo").
2. Lee la hoja **Base** si existe.
3. Mapea las columnas del fuente a las columnas del maestro (por nombre o variantes).
4. Rellena columnas calculadas (ID Blotter, Dif. Tasa) y las que vienen de columnas con fecha (VPN 1ra/2da fecha, Precio -1UF).
5. Rellena Fecha de compra (input o argumento) y las 3 columnas desde Base (por Rut).
6. Concatena con el maestro existente (o crea uno nuevo) y elimina duplicados por **Rut + Fecha de emisión** (conserva la primera fila).
7. Guarda el maestro (cabecera en fila 2).

Al final imprime un **reporte de mapeo** en consola y un resumen de filas.

---

## Resumen rápido

| Paso | Acción |
|------|--------|
| 1 | Abrir terminal en la carpeta del proyecto. |
| 2 | `pip install -r requirements.txt` (solo la primera vez). |
| 3 | `python cargar_a_maestro.py "tu_archivo_fuente.xlsx"` |
| 4 | Cuando pida **Fecha de compra**, escribir la fecha y pulsar Enter. |
| 5 | Revisar el reporte en pantalla y abrir **maestro.xlsx** para ver el resultado. |

---

## Documentación adicional

- **`RESUMEN_COLUMNAS_MAESTRO.md`**: cuadro con cada columna del maestro y de dónde sale su valor (Valo, Base, calculado, input, etc.).
