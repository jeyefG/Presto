# AGENTS

## Descripción general

`presto_generalized.py` es el núcleo del procesamiento. Lee un Excel exportado desde Presto, construye un grafo de dependencias a partir de las fórmulas, recorre recursivamente la jerarquía de capítulos/partidas/recursos y recalcula cantidades y totales. El resultado es una tabla plana con hasta 6 niveles de capítulos, 4 de partidas y los totales por recurso, que luego se exporta a un archivo Excel.

Puntos clave del flujo en `presto_generalized.py`:
- Carga el Excel con `openpyxl` y lo convierte a `pandas.DataFrame`.
- Detecta la celda de presupuesto total con `find_budget_cell`.
- Construye un `DependencyGraph` de celdas con fórmulas y un `FormulaEvaluator` para calcular valores.
- Recorre la jerarquía con `traverse_resources`, emitiendo filas normalizadas con `emit_row`.
- Exporta el resultado con `export_to_excel`.

## Cómo se “empaqueta” para uso comercial

El paquete de archivos `presto_commercial.py`, `presto_commercial_app.py` y `presto_commercial_app.spec` convierte el script base en una aplicación lista para entregar a terceros:

1. **`presto_commercial.py`**
   - Es una versión preparada para uso comercial del núcleo de procesamiento.
   - Expone la misma lógica de `presto_generalized.py` pero encapsulada para reutilizarse desde otros módulos.
   - Incluye `run_cli` para ejecutar desde línea de comandos y `export_to_excel` como punto de integración.

2. **`presto_commercial_app.py`**
   - Crea una interfaz gráfica (Tkinter) para usuarios no técnicos.
   - Permite elegir el Excel de Presto, opcionalmente la hoja, y la ruta de salida.
   - Ejecuta `export_to_excel` en un hilo y muestra mensajes de estado y errores.
   - Convierte el procesamiento en una app de escritorio “lista para vender”.

3. **`presto_commercial_app.spec`**
   - Es el archivo de configuración de PyInstaller que empaqueta la app.
   - Declara módulos ocultos (`pandas`, `openpyxl`) para que el ejecutable sea autosuficiente.
   - Genera un binario `PrestoCommercial` sin consola, apto para distribución.

En conjunto, este “pack” transforma el script original en una aplicación de escritorio empaquetada, con interfaz amigable y ejecutable independiente, pensada para distribución comercial.
