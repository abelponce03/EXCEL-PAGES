# Guía de uso (resumen)

Este libro calcula de forma automática cuánto gana o pierde la empresa con cada contenedor de paquetería, en cada región y en cada mes. Lo lleva UNA sola persona, que solo escribe datos en las celdas de entrada (amarillas, letra azul). Todo lo demás se calcula con fórmulas. La guía completa, con imágenes y ejercicios paso a paso, es el PDF «Guia_de_Uso_Rentabilidad_Paqueteria.pdf» que acompaña a este archivo.

## 1. Colores y convenciones

| Aspecto de la celda | Significado | ¿Se puede escribir? |
|---|---|---|
| Fondo amarillo claro, letra azul | Dato de entrada (valor que usted escribe) | SÍ |
| Letra negra | Fórmula de cálculo | NO |
| Letra verde | Enlace a otra hoja | NO |
| Fondo gris en encabezados | Títulos de columnas y secciones | NO |
| Celda vacía en columnas «(opcional)» | El libro usa la norma técnica de PARAMETROS | Sí, si hay dato real |

Las hojas están protegidas sin contraseña para evitar borrar fórmulas por error. Si necesita modificar la estructura: Revisar > Desproteger hoja. Vuelva a protegerla al terminar.

## 2. Configuración inicial (una sola vez)

1. PARAMETROS: escriba el nombre de la empresa, el primer mes del período (Fecha_Inicio) y revise las normas técnicas: litros por viaje, kilómetros, capacidad de los camiones y rendimiento del desagrupe (kg por jornal). Revise también los porcentajes de reclamaciones, de comisiones bancarias y las tasas tributarias.
2. PRECIOS: en la primera fila escriba las tarifas por kg que cobra la empresa en cada región, el precio del diésel, la cita del puerto, la dieta, el pago por jornal y las comisiones de comercialización, con la fecha desde la que rigen.
3. ACTIVOS: registre cada vehículo propio y cada equipo con su fecha de alta, valor de compra, valor residual y vida útil. El libro calcula la depreciación mensual y deja de calcularla al terminar la vida útil o en la fecha de baja.
4. COSTOS_FIJOS: escriba en la columna «Valor base mensual» cada gasto fijo (salarios de oficina, alquiler, electricidad, servicios contratados de chofer y ayudante por región, etc.). Indique si es salario de un trabajador propio («Sí» en la columna correspondiente): así se calculan la seguridad social (14 %) y el impuesto por la utilización de la fuerza de trabajo (5 %).
5. COMERCIAL: escriba los nombres de los 21 gestores (5 del grupo central y 16 provinciales) y el pago fijo mensual de cada uno. El total pasa solo a COSTOS_FIJOS.
6. Borre los datos de EJEMPLO de las hojas ENTRADA, PRECIOS, ACTIVOS y COMERCIAL. Seleccione las celdas y pulse la tecla Supr. NUNCA elimine filas ni columnas.

## 3. Registro de cada contenedor (hoja ENTRADA)

Use una fila por contenedor. Los datos se registran en dos momentos:

1. Al recibir el contenedor: ID del contenedor, fecha de arribo, transitaria (ENCI o A.V), tipo, número de bultos y kg que van a cada región (Occidente, Centro, Oriente), según el manifiesto o desglose. Ponga el Estado en «Abierto».
2. Al terminar la distribución, escriba los datos REALES que tenga: viajes y litros al puerto, dieta pagada, jornales del desagrupe, litros de traslado y de distribución por región, jornales de los centros transitorios, cargos por estadía o almacenaje, otros gastos directos con su detalle y otros ingresos. Ponga la fecha de fin de la distribución y cambie el Estado a «Cerrado».
- Si deja vacía una celda «(opcional)», el libro usa la norma técnica: por ejemplo, 60 L por viaje al puerto. Lo correcto es escribir siempre el dato real, tomado de los vales de combustible y de las hojas de ruta.
- Las columnas «Alertas» y «Utilidad» al final de la fila muestran el resultado al instante.
- Cada importe debe tener un documento que lo respalde: factura, vale, nómina o comprobante. Anote el número del documento en Observaciones.

## 4. Cierre mensual (el primer día hábil del mes siguiente)

1. Compruebe que todos los contenedores del mes estén en «Cerrado» y sin alertas pendientes.
2. En COSTOS_FIJOS, si algún gasto fijo real del mes fue distinto del valor base, escriba el valor real en la celda de ese mes. Esa celda sustituye la fórmula solo para ese mes.
3. En COMERCIAL, seleccione el mes, escriba la base variable de cada gestor (kg o USD cobrados) y compruebe que la conciliación diga «CUADRA». El resultado es la liquidación que se paga a los 21 gestores.
4. Revise RESUMEN_MES y DASHBOARD. La columna «Control» de RESUMEN_MES debe ser cero.
5. Guarde una copia del archivo con el nombre del mes, por ejemplo Rentabilidad_2026-10.xlsx, como respaldo.

## 5. Cuando cambia un precio o una tarifa

NO sobrescriba la fila vigente de PRECIOS. Agregue una fila nueva debajo, con la fecha desde la que rige el nuevo precio, y copie los demás valores que no cambian. Cada contenedor toma automáticamente los precios vigentes en su fecha de arribo. Así los contenedores anteriores conservan su costo real y el historial no se altera. Mantenga las fechas en orden ascendente.

## 6. Cómo leer los resultados

| Indicador | Qué significa | Cómo se calcula |
|---|---|---|
| Margen de contribución (MC) | Lo que queda de los ingresos para cubrir los costos fijos | Ingresos − costos variables |
| Utilidad antes de impuesto (UAI) | Ganancia o pérdida después de todos los costos y de los tributos sobre ingresos | MC − costos fijos − impuesto sobre ventas y servicios − contribución territorial |
| Utilidad neta estimada | Ganancia después del impuesto sobre utilidades (estimado) | UAI − 35 % de la UAI positiva |
| Costo total por kg | Lo que cuesta mover 1 kg de punta a punta | Costos totales ÷ kg |
| Punto de equilibrio (kg) | kg mínimos del mes para no perder dinero | Costos fijos ÷ (MC − tributos sobre ingresos) por kg |
| Margen de seguridad | Cuánto pueden caer los kg antes de entrar en pérdida | (kg reales − kg de equilibrio) ÷ kg reales |
| Fijos no absorbidos | Costos fijos del mes que ningún contenedor cubrió (meses sin arribos o regiones sin carga) | Fijos del mes − fijos asignados a contenedores |
| Tarifa mínima (SIMULADOR) | Precio por kg por debajo del cual la región pierde dinero | Costos de la región ÷ kg ajustados por los % sobre ingreso |

- Utilidad de una región: ingresos de la región menos sus costos propios (combustible, desagrupe, comisión y fijos regionales), menos su parte de los costos compartidos (puerto, transitaria, fijos generales) según sus kg.
- El impuesto sobre utilidades es una ESTIMACIÓN. La liquidación oficial es anual: use la tabla anual de RESUMEN_MES y valídela con el contador y la ONAT.

## 7. Simulador («¿qué pasaría si...?»)

En SIMULADOR la columna «Valor actual» trae los datos reales del libro. Escriba un valor en «Valor a simular» para probar un escenario: subir la tarifa, subir el diésel, recibir más contenedores, etc. Borre el valor para volver al dato real. Las tablas de sensibilidad muestran la utilidad mensual para distintas tarifas, precios del diésel y cantidades de contenedores. El simulador no modifica ningún dato del libro.

## 8. Alertas y cómo resolverlas

| Alerta | Causa | Qué hacer |
|---|---|---|
| Falta fecha / Falta transitaria | Dato obligatorio vacío | Complete la columna en ENTRADA |
| Mes fuera del rango de COSTOS_FIJOS | La fecha cae fuera de los 24 meses del período | Corrija la fecha, o cambie Fecha_Inicio en PARAMETROS al comenzar un nuevo período (guarde antes una copia) |
| Sin kg | No se registraron kg por región | Escriba los kg de Occidente, Centro y Oriente |
| Combustible X % sobre norma | El consumo real supera la norma más la tolerancia | Revise vales y hojas de ruta. Si la norma está desactualizada, corríjala en PARAMETROS |
| Fecha anterior a la 1ª vigencia de precios | No hay precios para esa fecha | Agregue en PRECIOS una vigencia anterior |
| Contenedor con pérdida | Los costos superan los ingresos | Analice la causa en CALCULO (región, combustible, estadía) |

## 9. Reglas de oro

- Nunca elimine filas ni columnas. Para borrar datos, seleccione las celdas y pulse Supr.
- Escriba solo en las celdas amarillas. No copie y pegue formatos desde otros archivos: use «Pegar valores».
- Escriba los números sin texto: 60, no «60 L». Escriba las fechas como dd/mm/aaaa.
- Trabaje siempre sobre un único archivo maestro. No haga copias de trabajo paralelas: las copias con fecha son solo respaldos.
- Haga una copia de respaldo semanal y otra en cada cierre mensual.
- El libro admite 300 contenedores en ENTRADA y 24 meses. Al llegar al límite, guarde el archivo como histórico y empiece uno nuevo con Fecha_Inicio actualizada.

## 10. Preguntas frecuentes

- ¿Por qué cambia la utilidad de un contenedor cuando registro otro del mismo mes? Porque los costos fijos del mes se reparten entre todos los contenedores del mes según sus kg. El resultado de un mes es definitivo al cerrarlo.
- ¿Puedo cambiar el método de reparto de los fijos? Sí, en PARAMETROS > Metodo_Prorrateo: KG (recomendado) o CONTENEDOR (partes iguales).
- ¿Cómo registro un gasto que no aparece en ninguna columna? Úselo como «Otros gastos directos» en ENTRADA, si es del contenedor, o como fila «Otros fijos» en COSTOS_FIJOS, si es del mes.
- ¿Qué hago si cambian los impuestos? Actualice las tasas en PARAMETROS, sección D, después de consultarlo con el contador.
- ¿Por qué los meses futuros no muestran costos fijos? Porque todavía no han ocurrido. Los resultados llegan hasta el «mes de corte»: el último mes con contenedores, o el que se fije en PARAMETROS > Mes_Corte. Un mes YA transcurrido sin arribos sí carga sus costos fijos, porque es capacidad ociosa real.
