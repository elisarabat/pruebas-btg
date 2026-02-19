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

- **Archivo Excel fuente** con al menos la hoja **Valo** (datos principales). Por defecto el script usa **filas 2 y 3** de Valo para los nombres de columna: si la fila 3 tiene texto se usa ese; si está vacía, se usa el de la fila 2. Así se leen bien las columnas cuyo título está solo en la fila 2. Los datos empiezan en la **fila 4**. En la hoja **Base** los nombres están en la **fila 1**. Si tiene hoja Base, se usará para 3 columnas (Fecha de emisión, Tasa Arriendo o Compra, Tasa Venta), enlazando por **Rut**.
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

## Por qué algunas columnas del fuente "no se usan"

El script solo puede usar una columna del fuente si sabe **a qué columna del maestro va**. Eso se hace por **nombre** (o por **índice** si el nombre falta):

1. **Columnas "Unnamed: 8", "Unnamed: 15", etc.**  
   En el Excel fuente, esas columnas tienen el **título en la fila 2** y la fila 3 vacía (o al revés). Si solo se usa la fila 3 como encabezado, pandas ve celdas vacías y les pone "Unnamed: N". Por eso el script tiene **`USAR_FILAS_2_Y_3_VALO = True`** (por defecto): construye el nombre de cada columna usando la **fila 3 si tiene texto, si no la fila 2**. Así muchas de esas columnas pasan a tener nombre (p. ej. "Monto Crédito + Cap (UF)") y el mapeo por nombre funciona. Si alguna sigue saliendo como Unnamed, puedes indicarla en **`COLUMNAS_POR_INDICE_FUENTE`**.

2. **"Leasing o Mutuo"**  
   El maestro **no tiene** una columna llamada "Leasing o Mutuo". El script no tiene destino para ese dato. Si la necesitas en el maestro, hay que añadirla a `COLUMNAS_MAESTRO` en `cargar_a_maestro.py` y, si en el fuente tiene otro nombre, añadir una variante en el mapeo (o por índice si también viene sin nombre).

3. **"Codigo Comuna"**  
   El maestro tiene **"Comuna"** (nombre de la comuna), que se rellena desde la columna "Comuna" del fuente. "Codigo Comuna" es otra columna (código numérico). El maestro no tiene "Codigo Comuna"; si la quieres, hay que añadirla a `COLUMNAS_MAESTRO` y mapearla (por nombre "Codigo Comuna" o por índice 30).

---

## Si hay columnas sin mapear (Unnamed)

Si en el reporte aparecen columnas **\[SIN MAPEAR]** y en "Columnas del archivo fuente no usadas" ves columnas **Unnamed: N**, es porque en el Excel fuente esas columnas tienen el encabezado vacío (o fusionado). Puedes mapearlas **por posición**:

1. Ejecuta el script y en el reporte revisa la sección **"Columnas del archivo fuente no usadas"**. Cada línea muestra el **índice** (posición 0-based) y el nombre de la columna.
2. Abre `cargar_a_maestro.py` y localiza el diccionario **`COLUMNAS_POR_INDICE_FUENTE`** (cerca de la línea 40).
3. Añade o descomenta líneas indicando qué columna del maestro corresponde a cada índice, por ejemplo:
   - Si "Monto Crédito + Cap (UF)" está sin mapear y en el reporte ves `índice 8: 'Unnamed: 8'`, añade: `"Monto Crédito + Cap (UF)": 8,`
   - Lo mismo para Fecha 1er Aporte, Fecha Last Aporte, Fecha Corte, Tasacion, Precio Venta/Tasación, Monto dividendo, Div/Renta, Carga Financiera, según el índice que muestre el reporte para cada una.

Guarda el script y vuelve a ejecutar; esas columnas pasarán a mapearse correctamente.

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
