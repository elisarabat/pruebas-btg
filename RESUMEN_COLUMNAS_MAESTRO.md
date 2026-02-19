# Resumen: columnas del maestro y su equivalencia

Este documento describe **cada columna del Excel maestro** y de dónde se obtiene su valor al cargar desde el archivo fuente (hojas **Valo** y **Base**).

---

## Origen de los datos

| Origen | Descripción |
|--------|-------------|
| **Valo** | Hoja principal del Excel fuente. Los nombres de columna están en la **fila 3**. Las columnas se mapean por nombre (o por variantes similares). |
| **Base** | Segunda hoja del mismo Excel. Los nombres de columna están en la **fila 1**. Se enlaza por **Rut**; solo 3 columnas del maestro vienen de aquí. |
| **Input** | Valor ingresado por consola al ejecutar el script (ej. Fecha de compra). |
| **Calculado** | Se calcula a partir de otras columnas ya rellenadas en el maestro. |
| **Col. fecha** | Viene de columnas de Valo cuyo **encabezado es una fecha** (1ª y 2ª por orden cronológico). |

---

## Cuadro: columna del maestro → equivalencia

| # | Columna en el maestro | Origen | Equivalencia / Regla |
|---|------------------------|--------|----------------------|
| 1 | Fecha de compra | **Input** | Valor que ingresas por consola al ejecutar (o por argumento). Se aplica a todas las filas nuevas. |
| 2 | N° | Valo | Mapeo desde Valo (nombre igual o variantes: N, numero, num, nº, no). |
| 3 | N° OP | Valo | Mapeo desde Valo (N op, N op., numero op, nop, n° operación). |
| 4 | ID Blotter | **Calculado** | Todos los **dígitos** de N° OP antes de la primera letra (ej. `12345ABC` → `12345`). |
| 5 | Apellido Paterno | Valo | Mapeo desde Valo (apellido paterno, ap paterno, paterno). |
| 6 | Apellido materno | Valo | Mapeo desde Valo (apellido materno, ap materno, materno). |
| 7 | Nombres | Valo | Mapeo desde Valo (nombres, nombre completo). |
| 8 | Rut | Valo | Mapeo desde Valo (rut, run, rut cliente). Se usa como llave para enlazar con Base. |
| 9 | DV | Valo | Mapeo desde Valo (dv, digito verificador). |
| 10 | Fecha de emisión | **Base** | Hoja **Base**, columna **Fecha de suscripción** (match por Rut). |
| 11 | Monto Crédito + Cap (UF) | Valo | Mapeo desde Valo (monto credito, credito uf, monto cap, etc.). |
| 12 | Subsidio | Valo | Mapeo desde Valo. |
| 13 | Pie | Valo | Mapeo desde Valo (pie, pie inicial). |
| 14 | Valor Vivienda | Valo | Mapeo desde Valo (valor vivienda, valor propiedad, etc.). |
| 15 | Morosidad | Valo | Mapeo desde Valo (morosidad, dias morosidad). |
| 16 | N° Cuotas | Valo | Mapeo desde Valo (n cuotas, numero cuotas, cuotas, etc.). |
| 17 | Cuota Mes | Valo | Mapeo desde Valo (cuota mes, cuota, cuota mensual, etc.). |
| 18 | Tasa Arriendo o Compra | **Base** | Hoja **Base**, columna **Tasa anual de emisión** (match por Rut). |
| 19 | Fecha 1er Aporte | Valo | Mapeo desde Valo (fecha primer aporte, 1er aporte, etc.). |
| 20 | Fecha Last Aporte | Valo | Mapeo desde Valo (fecha ultimo aporte, last aporte, etc.). |
| 21 | Fecha Corte | Valo | Mapeo desde Valo (fecha corte, corte). |
| 22 | VPN 1ra fecha | **Col. fecha** | Valor de la columna de Valo cuya **cabecera es la 1ª fecha** (orden cronológico). |
| 23 | VPN 2da fecha | **Col. fecha** | Valor de la columna de Valo cuya **cabecera es la 2ª fecha** (siempre posterior a la 1ª). |
| 24 | Precio -1UF | **Col. fecha** | **Valor de la 2ª columna fecha menos 1** (mismo origen que VPN 2da fecha, restado 1). |
| 25 | Saldo insoluto Teorico al 31-07-2019 | Valo | Mapeo desde Valo (saldo insoluto, saldo teorico, saldo 31-07-2019). |
| 26 | Tasacion | Valo | Mapeo desde Valo (tasacion, tasación). |
| 27 | Precio Venta/Tasación | Valo | Mapeo desde Valo (precio venta, precio tasacion, etc.). |
| 28 | Tasa Venta | **Base** | Hoja **Base**, columna **Tasa anual de endoso** (match por Rut). |
| 29 | Dif. Tasa | **Calculado** | **Tasa Arriendo o Compra − Tasa Venta** (numérico). |
| 30 | Monto dividendo | Valo | Mapeo desde Valo (monto dividendo, dividendo). |
| 31 | DivRenta | Valo | Mapeo desde Valo (div/renta, divrenta, dividendo/renta). |
| 32 | Carga Financiera | Valo | Mapeo desde Valo (carga financiera, carga_financiera, etc.). |
| 33 | Dirección | Valo | Mapeo desde Valo (direccion, domicilio, etc.). |
| 34 | Comuna | Valo | Mapeo desde Valo. |

---

## Resumen por origen

- **Valo (mapeo):** 25 columnas (N°, N° OP, Apellidos, Nombres, Rut, DV, Monto Crédito, Subsidio, Pie, Valor Vivienda, Morosidad, N° Cuotas, Cuota Mes, Fecha 1er/Last Aporte, Fecha Corte, Saldo insoluto, Tasacion, Precio Venta/Tasación, Monto dividendo, DivRenta, Carga Financiera, Dirección, Comuna).
- **Base (por Rut):** 3 columnas → Fecha de emisión, Tasa Arriendo o Compra, Tasa Venta.
- **Input:** 1 columna → Fecha de compra.
- **Calculado:** 2 columnas → ID Blotter (desde N° OP), Dif. Tasa (Tasa Arriendo − Tasa Venta).
- **Col. fecha (Valo):** 3 columnas → VPN 1ra fecha, VPN 2da fecha, Precio -1UF (2ª fecha − 1).

---

*Generado a partir de la configuración de `cargar_a_maestro.py`.*
