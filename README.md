# Automatización de Nómina Mensual (Colombia)

Programa en Python que liquida la nómina mensual de 10 empleados aplicando las reglas
laborales colombianas, genera un archivo Excel profesional con OpenPyXL e imprime un
análisis de resultados por consola.

---

## Requisitos

- Python 3.7 o superior
- OpenPyXL (única dependencia externa)

## Instalación

```bash
pip install openpyxl
```

## Uso

Desde la carpeta del proyecto:

```bash
python nomina.py
```

El programa crea el archivo **`nomina_empresa.xlsx`** en la misma carpeta y muestra el
reporte de análisis en la consola. Si el archivo ya está abierto en Excel, el programa
avisa y pide cerrarlo antes de reintentar.

---

## Reglas de negocio

Se aplican en el orden en que aparecen, ya que cada una alimenta a la siguiente:

| # | Concepto | Fórmula |
|---|----------|---------|
| 1 | Auxilio de Transporte | `$200.000` si el salario ≤ `$2.847.000`; en caso contrario `$0` |
| 2 | Valor Hora | `Salario / 240` |
| 3 | Valor Hora Extra | `Valor Hora × 1,25` (recargo del 25%) |
| 4 | Total Horas Extras | `Horas Extras × Valor Hora Extra` |
| 5 | Devengado | `Salario + Auxilio + Total Horas Extras` |
| 6 | Descuento Salud | `Devengado × 4%` |
| 7 | Descuento Pensión | `Devengado × 4%` |
| 8 | Neto a Pagar | `Devengado − Salud − Pensión` |

Todos los parámetros (tope del auxilio, horas del mes, porcentajes de salud y pensión)
están definidos como constantes al inicio de `nomina.py`, en la sección
*Parámetros y reglas de negocio*. Para actualizar la nómina a otro año basta con
cambiar esas constantes: no hay que tocar el resto del código.

---

## Estructura del archivo generado

### Hoja 1 — `Nómina <Mes> <Año>`

El nombre de la hoja toma automáticamente el mes y el año actuales (por ejemplo,
`Nómina Julio 2026`). Contiene una fila por empleado con estas 13 columnas:

| Col | Campo | Col | Campo |
|-----|-------|-----|-------|
| A | Documento | H | Valor Hora Extra |
| B | Nombre | I | Total Horas Extras |
| C | Cargo | J | Devengado |
| D | Salario | K | Salud |
| E | Horas Extras | L | Pensión |
| F | Auxilio Transporte | M | Neto a Pagar |
| G | Valor Hora | | |

Formato aplicado:

- Encabezados en **negrita**, texto blanco sobre fondo azul corporativo `#1F4E78`.
- Columnas monetarias con separador de miles (`#,##0`).
- Ancho de columna uniforme (~16) y bordes suaves en todas las celdas.
- Fila de encabezado congelada, para que siga visible al desplazarse.

### Hoja 2 — `Resumen Gerencial`

Tabla de dos columnas (**Indicador** / **Valor**) con el mismo estilo de encabezado azul.

Los valores **no** se calculan en Python: se escriben como **fórmulas de Excel** que
referencian la hoja de nómina. Así, si alguien edita un salario o unas horas extras
directamente en Excel, los indicadores se recalculan solos.

| Indicador | Fórmula |
|-----------|---------|
| Total Empleados | `=COUNTA('Nómina <Mes> <Año>'!A2:A11)` |
| Total Nómina | `=SUM('Nómina <Mes> <Año>'!J2:J11)` |
| Total Salud | `=SUM('Nómina <Mes> <Año>'!K2:K11)` |
| Total Pensión | `=SUM('Nómina <Mes> <Año>'!L2:L11)` |
| Mayor Salario | `=MAX('Nómina <Mes> <Año>'!D2:D11)` |
| Menor Salario | `=MIN('Nómina <Mes> <Año>'!D2:D11)` |
| Promedio Salarial | `=AVERAGE('Nómina <Mes> <Año>'!D2:D11)` |

> Nota: las celdas se ven vacías si el archivo se lee con OpenPyXL usando
> `data_only=True`, porque quien evalúa las fórmulas es Excel al abrir el archivo.
> Esto es normal y esperado.

---

## Análisis por consola

Después de guardar el archivo, el programa imprime:

1. Valor total pagado en nómina (total devengado).
2. Dinero descontado por salud.
3. Dinero descontado por pensión.
4. Empleado que recibió el mayor pago neto (nombre, cargo, documento y valor).
5. Ventajas de automatizar la nómina con Python.

Los valores se muestran en formato de moneda colombiana con separador de miles
(por ejemplo, `$ 35.965.885`).

---

## Estructura del código (`nomina.py`)

El archivo está organizado por secciones comentadas:

1. **Parámetros y reglas de negocio** — constantes configurables.
2. **Datos de entrada** — los 10 empleados.
3. **Cálculos de nómina** — `liquidar_empleado()` y `liquidar_nomina()`.
4. **Estilos reutilizables** — fuentes, rellenos, bordes y anchos.
5. **Hoja 1** — construcción del detalle de nómina.
6. **Hoja 2** — construcción del resumen gerencial con fórmulas.
7. **Reporte por consola** — formato de moneda y análisis.
8. **Programa principal** — orquesta el flujo y guarda el archivo.

## Archivos del proyecto

```
nomina.py            # Programa principal
nomina_empresa.xlsx  # Archivo generado al ejecutar el programa
README.md            # Este documento
```
