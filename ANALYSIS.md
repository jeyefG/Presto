# Análisis del script `PrestoV9 Price ID.py`

Este script recorre un Excel exportado desde Presto para reconstruir el desglose de costos por recurso a partir de fórmulas de subtotal que sólo entregan totales por partida. El algoritmo actual sigue las fórmulas en cascada (con `SUM`, `ROUND`, referencias cruzadas de columnas E/F/G) y va construyendo una estructura `struct` que se normaliza en filas con hasta 6 capítulos, 4 partidas y 4 tipos de recurso.

## Limitaciones observadas

* El recorrido se apoya fuertemente en el número máximo de niveles (`num_cap = 6`, `num_part = 4`) y en banderas manuales (`flag1`, `flag2`, …), lo que dificulta extenderlo a nuevas combinaciones de capítulos/partidas o recursos adicionales.
* Se mezclan dos responsabilidades: interpretar fórmulas (seguimiento de referencias, SUM/ROUND) y materializar una fila de salida. Esto hace que las ramas de control sean difíciles de verificar.
* El flujo depende de supuestos locales (por ejemplo, que el total está en la penúltima fila o que los encabezados están en fila 2), lo que complica reutilizarlo para otras variantes de reportes.

## Idea para generalizar el problema

En vez de codificar el número de niveles y sus combinaciones, se puede modelar el Excel como un **árbol de celdas referenciadas** y aplicar un recorrido genérico que vaya colapsando nodos hoja (recursos) hacia la raíz (presupuesto). El pipeline quedaría así:

1. **Construir un grafo de dependencias de celdas**
   * Parsear todas las fórmulas del rango de interés y extraer sus referencias (por ejemplo con expresiones regulares sobre `cell.value` en openpyxl).
   * Cada celda con fórmula es un nodo; las celdas referenciadas son sus hijos. Esto permite saber cuántos niveles existen sin fijar límites.

2. **Clasificar nodos por tipo semántico**
   * Detectar el tipo de cada fila leyendo las columnas de naturaleza/UM/Pres, igual que ahora, pero almacenarlo en un diccionario `row_metadata[row]`.
   * Los recursos son hojas: celdas cuyo valor final es numérico o que no contienen más referencias.
   * Las partidas y capítulos son nodos internos: su valor depende de hijos.

3. **Recorrer el grafo de manera genérica**
   * Usar DFS/BFS desde la celda de presupuesto total; cuando se visita un nodo, si es `SUM` o `ROUND`, expandir a sus hijos usando las referencias recogidas en el paso 1.
   * Mantener un stack de contexto con los metadatos de los niveles encontrados (Capítulo, Partida, etc.). El tamaño del stack se ajusta dinámicamente al profundizar o volver a subir en el árbol.

4. **Emitir filas normalizadas sin límites fijos**
   * En vez de predefinir 6 capítulos y 4 partidas, generar columnas dinámicamente (`Capítulo 1`, `Capítulo 2`, …, `Partida 1`, `Partida 2`, …) según la profundidad real. Se puede definir un ancho máximo sólo en la etapa de exportación para compatibilidad, rellenando con vacío cuando falten niveles.
   * Cada vez que se alcanza un recurso hoja, emitir una fila con el stack de contexto, la UM y cantidad del recurso y el precio unitario tomado desde la columna de Pres (o de la celda numérica si no es fórmula).

5. **Validar cobertura**
   * Comparar la suma de los totales por recurso con el total del presupuesto; si difieren, registrar qué ramas del grafo no fueron visitadas para detectar casos especiales (capítulos con monto directo, recursos sueltos, etc.).

### Ventajas del enfoque genérico

* Permite manejar cualquier profundidad de capítulos/partidas sin reconfigurar el código.
* Simplifica la lógica de control: el flujo sigue la estructura del grafo, no una secuencia de `flag` y contadores.
* Facilita agregar nuevos tipos de recurso o columnas de metadatos: basta con extender el diccionario de clasificación y el formateo de salida.
* Ayuda a auditar ramas faltantes al comparar el total agregado con el total de la hoja.

### Sugerencias de implementación incremental

* Extraer la lógica de parseo de fórmulas (`SUM`, `ROUND`, `+`, referencias E/F/G) a funciones puras que devuelvan objetos `FormulaNode` con tipo y lista de hijos.
* Encapsular el estado de recorrido en una clase `TraversalContext` que maneje el stack de capítulos/partidas y provea un método `emit_resource_row` que construya la salida.
* Incorporar tests unitarios con muestras pequeñas de Excel generadas en memoria (openpyxl) que cubran combinaciones: capítulo con monto directo sin recurso, partidas con uno y varios recursos, recursos repetidos en distintos capítulos.

Esta separación en etapas (parseo ➜ grafo ➜ recorrido ➜ salida) debería reducir el “tuning manual” y ofrecer una ruta clara para cubrir todas las ramas del árbol de dependencias.
