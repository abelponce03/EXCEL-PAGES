# Análisis de rentabilidad: negocio de paquetería

Libro de Excel que calcula de forma automática la rentabilidad del negocio de paquetería **por contenedor, por región
(Occidente, Centro, Oriente) y por mes**. Cubre todo el flujo operativo: recogida en el puerto, desagrupe en la
transitaria (ENCI / A.V), distribución regional y comercialización.

| Archivo | Contenido |
|---|---|
| [`Analisis_Rentabilidad_Paqueteria.xlsx`](Analisis_Rentabilidad_Paqueteria.xlsx) | El libro de trabajo (14 hojas, más de 32 000 fórmulas, sin macros) |
| [`GUIA_DE_USO.md`](GUIA_DE_USO.md) | Guía paso a paso para el personal (también está en la hoja GUIA_USO) |
| [`ANALISIS_Y_PROPUESTA.md`](ANALISIS_Y_PROPUESTA.md) | Diagnóstico, metodología, clasificación de costos y propuestas (también en la hoja ANALISIS) |
| `build/` | Scripts que generan y verifican el libro |

## Estructura del libro

| Hoja | Uso |
|---|---|
| INICIO | Índice, leyenda de colores y controles automáticos de cuadre |
| GUIA_USO / ANALISIS | Guía para el personal / análisis económico y propuestas |
| DASHBOARD | Indicadores, estructura de costos, rentabilidad por región y gráficos del período elegido |
| **ENTRADA** | Registro de contenedores: el único lugar donde escribe operaciones |
| PARAMETROS | Normas técnicas (litros, km, capacidades, rendimientos), % sobre ingreso, tributos |
| PRECIOS | Tarifas y precios **con fecha de vigencia** (el historial no se altera al cambiar un precio) |
| ACTIVOS | Depreciación lineal, mantenimiento y seguros del vehículo propio y de los equipos |
| COSTOS_FIJOS | Costos fijos mensuales por centro de costo; seguridad social y fuerza de trabajo sobre salarios propios |
| COMERCIAL | 21 gestores (5 + 16): pago fijo, liquidación variable mensual y conciliación |
| CALCULO | Cálculo completo por contenedor (≈95 columnas) y por región, con alertas |
| RESUMEN_MES | Resultado mensual, punto de equilibrio, margen de seguridad e impuesto anual estimado |
| SIMULADOR | Escenarios, tarifa mínima por región y sensibilidad a la tarifa, al diésel y al volumen |

## Metodología (resumen)

- **Costeo variable** para decidir (margen de contribución, punto de equilibrio) y **costeo por absorción** para conocer
  el costo completo: los fijos del mes se reparten entre los contenedores según sus kg, y los fijos regionales solo entre los kg de su región.
- **Real frente a norma:** si una celda opcional queda vacía, se usa la norma técnica. Si el consumo real de combustible de un
  tramo supera la norma más la tolerancia, aparece una alerta.
- **Controles de cuadre:** la suma de las utilidades de las tres regiones es igual a la utilidad del contenedor, y la utilidad del mes es igual a la
  suma de los contenedores menos los fijos no absorbidos. INICIO muestra «OK» o «REVISAR».

## Datos de ejemplo

El libro trae **4 contenedores de EJEMPLO** (septiembre y octubre de 2026) y valores ILUSTRATIVOS, marcados como tales,
para mostrar cómo funciona. Datos aportados por la empresa: cita del puerto de 10 USD, 60 L por viaje al puerto, vehículo
propio solo en el puerto y distribución con servicios de terceros. Con los supuestos ilustrativos, el ejemplo muestra lo que el
modelo permite detectar: con una tarifa por kg igual en todo el país, **Occidente genera casi toda la utilidad y Oriente queda en pérdida**
(el SIMULADOR estima una tarifa mínima de unos 0,80 USD/kg para Oriente y de unos 0,28 USD/kg para Occidente). Antes de usar el libro hay que
sustituir los datos según la GUIA_USO (sección 2).

## Fuentes de los valores de referencia

- Tributos de las MIPYMES en 2026 (utilidades 35 %, ventas y servicios 10 %, seguridad social 14 %, fuerza de trabajo 5 %,
  contribución territorial 1 %): [Directorio Cubano: impuestos para mipymes](https://www.directoriocubano.info/panorama/asi-seran-los-impuestos-para-las-mipymes-en-cuba/),
  [Periódico Cubano: los 17 impuestos de 2026](https://www.periodicocubano.com/gaceta-oficial-publica-los-17-impuestos-que-se-cobraran-en-cuba-durante-el-2026/),
  [IPS Cuba](https://www.ipscuba.net/economia/pago-de-impuestos-empresas-extranjeras-vs-mipymes-cubanas/).
  **Confirmar el régimen aplicable con el contador y la ONAT.**
- Precio del diésel en USD (referencia de CUPET 2026, unos 2,00 USD/L para el diésel regular; varía por operación desde mayo de 2026):
  [Directorio Cubano: precio del combustible](https://www.directoriocubano.info/?p=149626).

## Regenerar el libro

```bash
pip install openpyxl lxml
RECALC=/ruta/a/recalc.py build/construir.sh   # genera, recalcula con LibreOffice Calc, inyecta los valores y verifica
```

Sin `RECALC`, el libro se genera sin valores en caché y Excel lo calcula al abrirlo. `build/verificar.py` compara los resultados
del libro con un cálculo independiente en Python.
