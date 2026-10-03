# Análisis económico y propuesta

Objetivo: medir con rigor la rentabilidad del negocio de paquetería por contenedor, por región y por mes; identificar dónde se gana y dónde se pierde dinero; y dar a la dirección indicadores para decidir tarifas, rutas y volumen.

## 1. Diagnóstico del flujo actual

| Etapa | Actividad | Recurso | Costos identificados | Comportamiento |
|---|---|---|---|---|
| 1. Puerto | Recogida del contenedor en el muelle | Vehículo PROPIO | Cita (10 USD), 60 L de combustible, dieta del chofer, desgaste, depreciación, mantenimiento | Variable por viaje + fijo mensual |
| 2. Transitaria (ENCI / A.V) | Desagrupe y clasificación por destinatario | Mano de obra por jornada | Jornales; posibles cargos de la transitaria, almacenaje y estadía | Variable por kg / contenedor |
| 3a. Occidente | Distribución directa | Servicio contratado a terceros | Chofer y ayudante (pago fijo mensual) + combustible de distribución | Fijo mensual + variable |
| 3b. Centro y Oriente | Traslado a centro transitorio, desagrupe y distribución | Servicio contratado a terceros | Chofer y ayudante + combustible de traslado + jornales de desagrupe + combustible de distribución | Fijo mensual + variable |
| 4. Comercialización | Gestión y cobro en 16 territorios + grupo central de 5 | 21 gestores | Pago fijo igual para todos + parte variable según lo gestionado | Fijo mensual + variable |
| Transversal | Dirección, contabilidad, local, servicios | Estructura | Salarios, alquiler, electricidad, comunicaciones, seguros, tributos | Fijo mensual |

## 2. Problemas de la medición actual

- No se separan los costos fijos de los variables. Sin esa separación no se puede calcular el punto de equilibrio ni saber cuánto aporta cada kg.
- Faltan costos ocultos: depreciación y mantenimiento del vehículo propio, tributos sobre ingresos y salarios, estadía y almacenaje del contenedor, mermas y reclamaciones, y comisiones bancarias. Por eso la rentabilidad aparente es MAYOR que la real.
- Centro y Oriente tienen dos tramos de transporte y un segundo desagrupe. Con una tarifa por kg igual en todo el país, el Occidente probablemente subsidia al Oriente. Hoy no se puede ver.
- No existen normas de consumo de combustible por ruta. Los desvíos (consumo excesivo, pérdidas) no se detectan.
- Los precios cambian (el diésel en USD varía desde mayo de 2026) y no hay historial. Si se sobrescriben, se pierde el costo real de los contenedores anteriores.
- La parte variable de la comercialización no se concilia con los kg o ingresos realmente entregados y cobrados.

## 3. Metodología del modelo propuesto

- Costeo variable (direct costing) para decidir: margen de contribución por contenedor, por kg y por región, punto de equilibrio y margen de seguridad.
- Costeo por absorción para conocer el costo completo: los costos fijos mensuales se reparten entre los contenedores del mes según sus kg o en partes iguales. Los fijos regionales se reparten solo entre los kg de esa región.
- Centros de costo: Puerto, Transitaria, Occidente, Centro, Oriente, Comercialización y Generales. Cada costo tiene un inductor: viajes, kg, jornales, litros o ingreso.
- Precios con vigencia (hoja PRECIOS): cada contenedor se valora con los precios de su fecha de arribo. El historial queda intacto y es auditable.
- Real frente a norma: cada consumo tiene una norma técnica. Si no hay dato real, se usa la norma. Si el dato real la supera más allá de la tolerancia, aparece una alerta.
- Clasificación por elementos de gasto: materias primas y materiales, combustibles, energía, fuerza de trabajo, depreciación y otros gastos monetarios. Así el libro conecta con la contabilidad y con los informes estadísticos.
- Controles de cuadre: la suma de las utilidades por región es igual a la utilidad del contenedor, y la utilidad mensual es igual a la suma de los contenedores menos los fijos no absorbidos.

## 4. Clasificación de costos

| Concepto | Centro de costo | Tipo | Inductor | Elemento de gasto |
|---|---|---|---|---|
| Cita del puerto | Puerto | Variable | Viaje | Otros gastos monetarios |
| Combustible al puerto | Puerto | Variable | Litros (norma 60 L/viaje) | Combustibles y lubricantes |
| Dieta del chofer | Puerto | Variable | Viaje | Otros gastos monetarios |
| Desgaste (neumáticos, lubricantes) | Puerto | Variable | Km recorridos | Materias primas y materiales |
| Depreciación del vehículo propio | Generales | Fijo | Mes (vida útil) | Depreciación y amortización |
| Mantenimiento preventivo, seguros y licencias | Generales | Fijo | Mes | Otros gastos monetarios |
| Jornales de desagrupe (ENCI / A.V) | Transitaria | Variable | Jornal (kg / rendimiento) | Gastos de fuerza de trabajo |
| Servicio de la transitaria, estadía y almacenaje | Transitaria | Variable | Contenedor / días | Otros gastos monetarios |
| Chofer y ayudante (servicio contratado) | Occidente / Centro / Oriente | Fijo | Mes | Otros gastos monetarios |
| Combustible de traslado y distribución | Occidente / Centro / Oriente | Variable | Litros (viajes × norma) | Combustibles y lubricantes |
| Jornales de desagrupe en centro transitorio | Centro / Oriente | Variable | Jornal | Gastos de fuerza de trabajo |
| Comercialización: parte fija | Generales | Fijo | Mes × 21 gestores | Gastos de fuerza de trabajo |
| Comercialización: parte variable | Por región | Variable | % del ingreso o USD/kg | Gastos de fuerza de trabajo |
| Reclamaciones / mermas | Por región | Variable | % del ingreso | Otros gastos monetarios |
| Comisiones bancarias | Por región | Variable | % del ingreso | Otros gastos monetarios |
| Impuesto sobre ventas y servicios (10 %) | Tributos | Variable | Ingreso | Tributo |
| Contribución territorial (1 %) | Tributos | Variable | Ingreso | Tributo |
| Seguridad social (14 %) y fuerza de trabajo (5 %) | Según el salario | Fijo | Salarios propios | Gastos de fuerza de trabajo |
| Impuesto sobre utilidades (35 %) | Resultado | Sobre utilidad | Utilidad anual | Tributo |

## 5. Propuestas de mejora (ordenadas por impacto económico)

| # | Propuesta | Beneficio esperado | Prioridad |
|---|---|---|---|
| 1 | Tarifa diferenciada por región (la hoja SIMULADOR calcula la tarifa mínima de cada región) | Evita que Occidente subsidie a Centro y Oriente; protege el margen | ALTA |
| 2 | Normas de consumo de combustible por ruta, con hoja de ruta, kilometraje y vales por viaje | El combustible es el mayor costo variable: un control del 10 % se nota directamente en la utilidad | ALTA |
| 3 | Recoger el contenedor dentro de los días libres de la naviera y de la transitaria | Elimina los cargos por estadía, demora y almacenaje | ALTA |
| 4 | Liquidar la comercialización variable sobre lo COBRADO, con conciliación mensual (hoja COMERCIAL) | Evita pagar comisiones sobre carga no cobrada y detecta diferencias | ALTA |
| 5 | Pedir a ENCI y A.V sus tarifas por escrito y compararlas para cada contenedor | Elegir la transitaria más económica según el volumen | MEDIA |
| 6 | Consolidar viajes a Centro y Oriente: no salir con el camión a menos del 80 % de su capacidad | Menos viajes y menos litros por kg | MEDIA |
| 7 | Evaluar si conviene contratar servicios fijos mensuales o pagar por viaje, según el volumen de cada región | Convertir costo fijo en variable cuando hay pocos contenedores reduce el riesgo | MEDIA |
| 8 | Buscar carga de retorno para los camiones que vuelven vacíos de Centro y Oriente | Ingreso adicional casi sin costo | MEDIA |
| 9 | Crear un fondo de reposición del vehículo propio igual a la depreciación mensual | Garantiza la continuidad de la operación | MEDIA |
| 10 | Asegurar la mercancía en tránsito y crear una provisión para reclamaciones | Protege contra pérdidas extraordinarias | MEDIA |
| 11 | Registrar los kg y bultos por provincia en cada contenedor; más adelante, por cliente o agencia emisora | Permite saber la rentabilidad por cliente y por provincia | BAJA |
| 12 | Cuando el volumen supere unos 25 contenedores al mes, pasar el registro a una base de datos o a un sistema de gestión (ERP) | Escalabilidad, varios usuarios y trazabilidad | BAJA |

## 6. Indicadores que la dirección debe revisar cada mes

- Utilidad antes de impuesto y margen % del mes y del año (DASHBOARD).
- Costo total por kg e ingreso por kg. La diferencia es la utilidad por kg.
- Utilidad por región y su tendencia.
- Punto de equilibrio (kg y contenedores por mes) y margen de seguridad.
- Desvío del consumo de combustible respecto a la norma (%).
- Días de ciclo: desde el arribo hasta el fin de la distribución.
- Costos fijos no absorbidos: indican capacidad ociosa.

## 7. Supuestos del modelo y datos por confirmar

| Supuesto | Valor usado | Fuente / estado |
|---|---|---|
| Moneda de registro | USD para todo | Indicado por la empresa |
| Ingreso | Tarifa por kg por región | Indicado por la empresa; el valor 0,80 USD/kg es ILUSTRATIVO |
| Unidad de análisis | Contenedor; consolidado mensual para los fijos | Indicado por la empresa |
| Cita del puerto | 10 USD por viaje | Indicado por la empresa |
| Combustible al puerto | 60 L por viaje | Indicado por la empresa |
| Precio del diésel | 2,00 USD/L | Referencia oficial de CUPET para 2026 (diésel regular). Varía por operación: confirmar con las facturas |
| Vehículo propio | Solo en la recogida en el puerto | Indicado por la empresa |
| Distribución regional | Servicio de terceros: chofer y ayudante con pago fijo mensual | Indicado por la empresa |
| Desagrupe | Pago por jornal, igual en ENCI y A.V | Indicado por la empresa |
| Comercialización | 21 gestores (5 + 16) con fijo igual + variable según lo gestionado | Indicado por la empresa; las tasas son ILUSTRATIVAS |
| Tributos | Ventas y servicios 10 %, territorial 1 %, seguridad social 14 %, fuerza de trabajo 5 %, utilidades 35 % | Régimen general de MIPYMES (Ley 113 y Ley del Presupuesto 2026). Confirmar el régimen de la empresa con el contador y la ONAT |
| Normas técnicas (km, capacidades, rendimientos, litros regionales) | Valores de PARAMETROS | ILUSTRATIVAS: medir en 2 o 3 contenedores reales y ajustar |
| Gastos fijos y salarios | Valores de COSTOS_FIJOS | ILUSTRATIVOS: sustituir por los registros contables |

## 8. Plan de implantación

1. Semana 1: el económico carga PRECIOS, ACTIVOS, COSTOS_FIJOS y COMERCIAL con datos reales y borra los ejemplos.
2. Semanas 2 a 4: operaciones registra los contenedores del mes en curso y, si hay datos, de los 2 o 3 meses anteriores.
3. Mes 1: se miden en la práctica las normas de combustible y de desagrupe, y se ajusta PARAMETROS.
4. Mes 2: primer cierre mensual con conciliación de la comercialización. La dirección analiza la tarifa mínima por región.
5. Mes 3: decisión sobre las tarifas diferenciadas y la consolidación de viajes. Se revisa el resultado frente a la contabilidad oficial.
