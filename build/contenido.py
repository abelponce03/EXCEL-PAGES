# -*- coding: utf-8 -*-
"""Textos de la guía de uso y del análisis económico.

Se usan para dos salidas: las hojas GUIA_USO y ANALISIS del Excel, y los
documentos Markdown del repositorio. Formato de cada bloque:
  ("h1", texto) título · ("h2", texto) subtítulo · ("p", texto) párrafo
  ("b", texto) viñeta · ("n", texto) paso numerado · ("t", [filas]) tabla
"""

GUIA = [
    ("h1", "GUÍA DE USO PARA EL PERSONAL"),
    ("p", "Este libro calcula de forma automática cuánto gana o pierde la empresa con cada contenedor de "
          "paquetería, en cada región y en cada mes. El personal solo escribe datos en las celdas de "
          "entrada (amarillas, letra azul). Todo lo demás se calcula con fórmulas."),

    ("h2", "1. Colores y convenciones"),
    ("t", [["Aspecto de la celda", "Significado", "¿Se puede escribir?"],
           ["Fondo amarillo claro, letra azul", "Dato de entrada (valor tecleado por el personal)", "SÍ"],
           ["Letra negra", "Fórmula de cálculo", "NO"],
           ["Letra verde", "Enlace a otra hoja", "NO"],
           ["Fondo gris en encabezados", "Títulos de columnas y secciones", "NO"],
           ["Celda vacía en columnas «(opcional)»", "El libro usa la norma técnica de PARAMETROS", "Sí, si hay dato real"]]),
    ("p", "Las hojas están protegidas sin contraseña para evitar borrar fórmulas por error. Si un responsable "
          "necesita modificar la estructura: Revisar > Desproteger hoja. Vuelva a protegerla al terminar."),

    ("h2", "2. Configuración inicial (una sola vez, la hace el económico o el contador)"),
    ("n", "PARAMETROS: escriba el nombre de la empresa, el primer mes del período (Fecha_Inicio) y revise las "
          "normas técnicas: litros por viaje, kilómetros, capacidad de los camiones y rendimiento del desagrupe "
          "(kg por jornal). Revise también los porcentajes de reclamaciones, de comisiones bancarias y las tasas tributarias."),
    ("n", "PRECIOS: en la primera fila escriba las tarifas por kg que cobra la empresa en cada región, el precio del "
          "diésel, la cita del puerto, la dieta, el pago por jornal y las comisiones de comercialización, con la "
          "fecha desde la que rigen."),
    ("n", "ACTIVOS: registre cada vehículo propio y cada equipo con su fecha de alta, valor de compra, valor "
          "residual y vida útil. El libro calcula la depreciación mensual y deja de calcularla al terminar la vida útil o en la fecha de baja."),
    ("n", "COSTOS_FIJOS: escriba en la columna «Valor base mensual» cada gasto fijo (salarios de oficina, alquiler, "
          "electricidad, servicios contratados de chofer y ayudante por región, etc.). Indique si es salario de "
          "un trabajador propio («Sí» en la columna correspondiente): así se calculan la seguridad social (14 %) y el "
          "impuesto por la utilización de la fuerza de trabajo (5 %)."),
    ("n", "COMERCIAL: escriba los nombres de los 21 gestores (5 del grupo central y 16 provinciales) y el pago "
          "fijo mensual de cada uno. El total pasa solo a COSTOS_FIJOS."),
    ("n", "Borre los datos de EJEMPLO de las hojas ENTRADA, PRECIOS, ACTIVOS y COMERCIAL. Seleccione las celdas y pulse la tecla Supr. NUNCA elimine filas ni columnas."),

    ("h2", "3. Registro de cada contenedor (hoja ENTRADA). Lo hace el especialista de operaciones"),
    ("p", "Use una fila por contenedor. Los datos se registran en dos momentos:"),
    ("n", "Al recibir el contenedor: ID del contenedor, fecha de arribo, transitaria (ENCI o A.V), tipo, número de "
          "bultos y kg que van a cada región (Occidente, Centro, Oriente), según el manifiesto o desglose. "
          "Ponga el Estado en «Abierto»."),
    ("n", "Al terminar la distribución, escriba los datos REALES que tenga: viajes y litros al puerto, dieta pagada, "
          "jornales del desagrupe, litros de traslado y de distribución por región, jornales de los centros "
          "transitorios, cargos por estadía o almacenaje, otros gastos directos con su detalle y otros ingresos. "
          "Ponga la fecha de fin de la distribución y cambie el Estado a «Cerrado»."),
    ("b", "Si deja vacía una celda «(opcional)», el libro usa la norma técnica: por ejemplo, 60 L por viaje al puerto. "
          "Lo correcto es escribir siempre el dato real, tomado de los vales de combustible y de las hojas de ruta."),
    ("b", "Las columnas «Alertas» y «Utilidad» al final de la fila muestran el resultado al instante."),
    ("b", "Cada importe debe tener un documento que lo respalde: factura, vale, nómina o comprobante. Anote el número del documento en Observaciones."),

    ("h2", "4. Cierre mensual (lo hace el económico el primer día hábil del mes siguiente)"),
    ("n", "Compruebe que todos los contenedores del mes estén en «Cerrado» y sin alertas pendientes."),
    ("n", "En COSTOS_FIJOS, si algún gasto fijo real del mes fue distinto del valor base, escriba el valor real en la "
          "celda de ese mes. Esa celda sustituye la fórmula solo para ese mes."),
    ("n", "En COMERCIAL, seleccione el mes, escriba la base variable de cada gestor (kg o USD cobrados) y compruebe que "
          "la conciliación diga «CUADRA». El resultado es la liquidación que se paga a los 21 gestores."),
    ("n", "Revise RESUMEN_MES y DASHBOARD. La columna «Control» de RESUMEN_MES debe ser cero."),
    ("n", "Guarde una copia del archivo con el nombre del mes, por ejemplo Rentabilidad_2026-10.xlsx, como respaldo."),

    ("h2", "5. Cuando cambia un precio o una tarifa"),
    ("p", "NO sobrescriba la fila vigente de PRECIOS. Agregue una fila nueva debajo, con la fecha desde la que rige el "
          "nuevo precio, y copie los demás valores que no cambian. Cada contenedor toma automáticamente los precios "
          "vigentes en su fecha de arribo. Así los contenedores anteriores conservan su costo real y el historial no se altera. "
          "Mantenga las fechas en orden ascendente."),

    ("h2", "6. Cómo leer los resultados"),
    ("t", [["Indicador", "Qué significa", "Cómo se calcula"],
           ["Margen de contribución (MC)", "Lo que queda de los ingresos para cubrir los costos fijos", "Ingresos − costos variables"],
           ["Utilidad antes de impuesto (UAI)", "Ganancia o pérdida después de todos los costos y de los tributos sobre ingresos",
            "MC − costos fijos − impuesto sobre ventas y servicios − contribución territorial"],
           ["Utilidad neta estimada", "Ganancia después del impuesto sobre utilidades (estimado)", "UAI − 35 % de la UAI positiva"],
           ["Costo total por kg", "Lo que cuesta mover 1 kg de punta a punta", "Costos totales ÷ kg"],
           ["Punto de equilibrio (kg)", "kg mínimos del mes para no perder dinero", "Costos fijos ÷ (MC − tributos sobre ingresos) por kg"],
           ["Margen de seguridad", "Cuánto pueden caer los kg antes de entrar en pérdida", "(kg reales − kg de equilibrio) ÷ kg reales"],
           ["Fijos no absorbidos", "Costos fijos del mes que ningún contenedor cubrió (meses sin arribos o regiones sin carga)",
            "Fijos del mes − fijos asignados a contenedores"],
           ["Tarifa mínima (SIMULADOR)", "Precio por kg por debajo del cual la región pierde dinero", "Costos de la región ÷ kg ajustados por los % sobre ingreso"]]),
    ("b", "Utilidad de una región: ingresos de la región menos sus costos propios (combustible, desagrupe, comisión y "
          "fijos regionales), menos su parte de los costos compartidos (puerto, transitaria, fijos generales) según sus kg."),
    ("b", "El impuesto sobre utilidades es una ESTIMACIÓN. La liquidación oficial es anual: use la tabla anual de RESUMEN_MES y valídela con el contador y la ONAT."),

    ("h2", "7. Simulador («¿qué pasaría si...?»)"),
    ("p", "En SIMULADOR la columna «Valor actual» trae los datos reales del libro. Escriba un valor en «Valor a simular» "
          "para probar un escenario: subir la tarifa, subir el diésel, recibir más contenedores, etc. Borre el valor para "
          "volver al dato real. Las tablas de sensibilidad muestran la utilidad mensual para distintas tarifas, precios del diésel y "
          "cantidades de contenedores. El simulador no modifica ningún dato del libro."),

    ("h2", "8. Alertas y cómo resolverlas"),
    ("t", [["Alerta", "Causa", "Qué hacer"],
           ["Falta fecha / Falta transitaria", "Dato obligatorio vacío", "Complete la columna en ENTRADA"],
           ["Mes fuera del rango de COSTOS_FIJOS", "La fecha cae fuera de los 24 meses del período",
            "Corrija la fecha, o cambie Fecha_Inicio en PARAMETROS al comenzar un nuevo período (guarde antes una copia)"],
           ["Sin kg", "No se registraron kg por región", "Escriba los kg de Occidente, Centro y Oriente"],
           ["Combustible X % sobre norma", "El consumo real supera la norma más la tolerancia",
            "Revise vales y hojas de ruta. Si la norma está desactualizada, corríjala en PARAMETROS"],
           ["Fecha anterior a la 1ª vigencia de precios", "No hay precios para esa fecha", "Agregue en PRECIOS una vigencia anterior"],
           ["Contenedor con pérdida", "Los costos superan los ingresos", "Analice la causa en CALCULO (región, combustible, estadía)"]]),

    ("h2", "9. Reglas de oro"),
    ("b", "Nunca elimine filas ni columnas. Para borrar datos, seleccione las celdas y pulse Supr."),
    ("b", "Escriba solo en las celdas amarillas. No copie y pegue formatos desde otros archivos: use «Pegar valores»."),
    ("b", "Escriba los números sin texto: 60, no «60 L». Escriba las fechas como dd/mm/aaaa."),
    ("b", "Un solo responsable por hoja: operaciones (ENTRADA), economía (PRECIOS, COSTOS_FIJOS, COMERCIAL), dirección (lectura)."),
    ("b", "Haga una copia de respaldo semanal y otra en cada cierre mensual."),
    ("b", "El libro admite 300 contenedores en ENTRADA y 24 meses. Al llegar al límite, guarde el archivo como "
          "histórico y empiece uno nuevo con Fecha_Inicio actualizada."),

    ("h2", "10. Preguntas frecuentes"),
    ("b", "¿Por qué cambia la utilidad de un contenedor cuando registro otro del mismo mes? Porque los costos fijos del "
          "mes se reparten entre todos los contenedores del mes según sus kg. El resultado de un mes es definitivo al cerrarlo."),
    ("b", "¿Puedo cambiar el método de reparto de los fijos? Sí, en PARAMETROS > Metodo_Prorrateo: KG (recomendado) o CONTENEDOR (partes iguales)."),
    ("b", "¿Cómo registro un gasto que no aparece en ninguna columna? Úselo como «Otros gastos directos» en ENTRADA, si es del "
          "contenedor, o como fila «Otros fijos» en COSTOS_FIJOS, si es del mes."),
    ("b", "¿Qué hago si cambian los impuestos? Actualice las tasas en PARAMETROS, sección D, después de consultarlo con el contador."),
    ("b", "¿Por qué los meses futuros no muestran costos fijos? Porque todavía no han ocurrido. Los resultados llegan hasta el «mes de corte»: "
          "el último mes con contenedores, o el que se fije en PARAMETROS > Mes_Corte. Un mes YA transcurrido sin arribos sí carga sus "
          "costos fijos, porque es capacidad ociosa real."),
]


ANALISIS = [
    ("h1", "ANÁLISIS ECONÓMICO Y PROPUESTA"),
    ("p", "Objetivo: medir con rigor la rentabilidad del negocio de paquetería por contenedor, por región y por mes; "
          "identificar dónde se gana y dónde se pierde dinero; y dar a la dirección indicadores para decidir tarifas, rutas y volumen."),

    ("h2", "1. Diagnóstico del flujo actual"),
    ("t", [["Etapa", "Actividad", "Recurso", "Costos identificados", "Comportamiento"],
           ["1. Puerto", "Recogida del contenedor en el muelle", "Vehículo PROPIO",
            "Cita (10 USD), 60 L de combustible, dieta del chofer, desgaste, depreciación, mantenimiento", "Variable por viaje + fijo mensual"],
           ["2. Transitaria (ENCI / A.V)", "Desagrupe y clasificación por destinatario", "Mano de obra por jornada",
            "Jornales; posibles cargos de la transitaria, almacenaje y estadía", "Variable por kg / contenedor"],
           ["3a. Occidente", "Distribución directa", "Servicio contratado a terceros",
            "Chofer y ayudante (pago fijo mensual) + combustible de distribución", "Fijo mensual + variable"],
           ["3b. Centro y Oriente", "Traslado a centro transitorio, desagrupe y distribución", "Servicio contratado a terceros",
            "Chofer y ayudante + combustible de traslado + jornales de desagrupe + combustible de distribución", "Fijo mensual + variable"],
           ["4. Comercialización", "Gestión y cobro en 16 territorios + grupo central de 5", "21 gestores",
            "Pago fijo igual para todos + parte variable según lo gestionado", "Fijo mensual + variable"],
           ["Transversal", "Dirección, contabilidad, local, servicios", "Estructura",
            "Salarios, alquiler, electricidad, comunicaciones, seguros, tributos", "Fijo mensual"]]),

    ("h2", "2. Problemas de la medición actual"),
    ("b", "No se separan los costos fijos de los variables. Sin esa separación no se puede calcular el punto de equilibrio ni saber cuánto aporta cada kg."),
    ("b", "Faltan costos ocultos: depreciación y mantenimiento del vehículo propio, tributos sobre ingresos y salarios, "
          "estadía y almacenaje del contenedor, mermas y reclamaciones, y comisiones bancarias. Por eso la rentabilidad aparente es MAYOR que la real."),
    ("b", "Centro y Oriente tienen dos tramos de transporte y un segundo desagrupe. Con una tarifa por kg igual en todo el país, el Occidente "
          "probablemente subsidia al Oriente. Hoy no se puede ver."),
    ("b", "No existen normas de consumo de combustible por ruta. Los desvíos (consumo excesivo, pérdidas) no se detectan."),
    ("b", "Los precios cambian (el diésel en USD varía desde mayo de 2026) y no hay historial. Si se sobrescriben, se pierde el costo real de los contenedores anteriores."),
    ("b", "La parte variable de la comercialización no se concilia con los kg o ingresos realmente entregados y cobrados."),

    ("h2", "3. Metodología del modelo propuesto"),
    ("b", "Costeo variable (direct costing) para decidir: margen de contribución por contenedor, por kg y por región, punto de equilibrio y margen de seguridad."),
    ("b", "Costeo por absorción para conocer el costo completo: los costos fijos mensuales se reparten entre los contenedores del mes "
          "según sus kg o en partes iguales. Los fijos regionales se reparten solo entre los kg de esa región."),
    ("b", "Centros de costo: Puerto, Transitaria, Occidente, Centro, Oriente, Comercialización y Generales. Cada costo tiene un "
          "inductor: viajes, kg, jornales, litros o ingreso."),
    ("b", "Precios con vigencia (hoja PRECIOS): cada contenedor se valora con los precios de su fecha de arribo. El historial queda intacto y es auditable."),
    ("b", "Real frente a norma: cada consumo tiene una norma técnica. Si no hay dato real, se usa la norma. Si el dato real la supera más allá de la tolerancia, aparece una alerta."),
    ("b", "Clasificación por elementos de gasto: materias primas y materiales, combustibles, energía, fuerza de trabajo, "
          "depreciación y otros gastos monetarios. Así el libro conecta con la contabilidad y con los informes estadísticos."),
    ("b", "Controles de cuadre: la suma de las utilidades por región es igual a la utilidad del contenedor, y la utilidad mensual es igual "
          "a la suma de los contenedores menos los fijos no absorbidos."),

    ("h2", "4. Clasificación de costos"),
    ("t", [["Concepto", "Centro de costo", "Tipo", "Inductor", "Elemento de gasto"],
           ["Cita del puerto", "Puerto", "Variable", "Viaje", "Otros gastos monetarios"],
           ["Combustible al puerto", "Puerto", "Variable", "Litros (norma 60 L/viaje)", "Combustibles y lubricantes"],
           ["Dieta del chofer", "Puerto", "Variable", "Viaje", "Otros gastos monetarios"],
           ["Desgaste (neumáticos, lubricantes)", "Puerto", "Variable", "Km recorridos", "Materias primas y materiales"],
           ["Depreciación del vehículo propio", "Generales", "Fijo", "Mes (vida útil)", "Depreciación y amortización"],
           ["Mantenimiento preventivo, seguros y licencias", "Generales", "Fijo", "Mes", "Otros gastos monetarios"],
           ["Jornales de desagrupe (ENCI / A.V)", "Transitaria", "Variable", "Jornal (kg / rendimiento)", "Gastos de fuerza de trabajo"],
           ["Servicio de la transitaria, estadía y almacenaje", "Transitaria", "Variable", "Contenedor / días", "Otros gastos monetarios"],
           ["Chofer y ayudante (servicio contratado)", "Occidente / Centro / Oriente", "Fijo", "Mes", "Otros gastos monetarios"],
           ["Combustible de traslado y distribución", "Occidente / Centro / Oriente", "Variable", "Litros (viajes × norma)", "Combustibles y lubricantes"],
           ["Jornales de desagrupe en centro transitorio", "Centro / Oriente", "Variable", "Jornal", "Gastos de fuerza de trabajo"],
           ["Comercialización: parte fija", "Generales", "Fijo", "Mes × 21 gestores", "Gastos de fuerza de trabajo"],
           ["Comercialización: parte variable", "Por región", "Variable", "% del ingreso o USD/kg", "Gastos de fuerza de trabajo"],
           ["Reclamaciones / mermas", "Por región", "Variable", "% del ingreso", "Otros gastos monetarios"],
           ["Comisiones bancarias", "Por región", "Variable", "% del ingreso", "Otros gastos monetarios"],
           ["Impuesto sobre ventas y servicios (10 %)", "Tributos", "Variable", "Ingreso", "Tributo"],
           ["Contribución territorial (1 %)", "Tributos", "Variable", "Ingreso", "Tributo"],
           ["Seguridad social (14 %) y fuerza de trabajo (5 %)", "Según el salario", "Fijo", "Salarios propios", "Gastos de fuerza de trabajo"],
           ["Impuesto sobre utilidades (35 %)", "Resultado", "Sobre utilidad", "Utilidad anual", "Tributo"]]),

    ("h2", "5. Propuestas de mejora (ordenadas por impacto económico)"),
    ("t", [["#", "Propuesta", "Beneficio esperado", "Prioridad"],
           ["1", "Tarifa diferenciada por región (la hoja SIMULADOR calcula la tarifa mínima de cada región)",
            "Evita que Occidente subsidie a Centro y Oriente; protege el margen", "ALTA"],
           ["2", "Normas de consumo de combustible por ruta, con hoja de ruta, kilometraje y vales por viaje",
            "El combustible es el mayor costo variable: un control del 10 % se nota directamente en la utilidad", "ALTA"],
           ["3", "Recoger el contenedor dentro de los días libres de la naviera y de la transitaria",
            "Elimina los cargos por estadía, demora y almacenaje", "ALTA"],
           ["4", "Liquidar la comercialización variable sobre lo COBRADO, con conciliación mensual (hoja COMERCIAL)",
            "Evita pagar comisiones sobre carga no cobrada y detecta diferencias", "ALTA"],
           ["5", "Pedir a ENCI y A.V sus tarifas por escrito y compararlas para cada contenedor",
            "Elegir la transitaria más económica según el volumen", "MEDIA"],
           ["6", "Consolidar viajes a Centro y Oriente: no salir con el camión a menos del 80 % de su capacidad",
            "Menos viajes y menos litros por kg", "MEDIA"],
           ["7", "Evaluar si conviene contratar servicios fijos mensuales o pagar por viaje, según el volumen de cada región",
            "Convertir costo fijo en variable cuando hay pocos contenedores reduce el riesgo", "MEDIA"],
           ["8", "Buscar carga de retorno para los camiones que vuelven vacíos de Centro y Oriente", "Ingreso adicional casi sin costo", "MEDIA"],
           ["9", "Crear un fondo de reposición del vehículo propio igual a la depreciación mensual", "Garantiza la continuidad de la operación", "MEDIA"],
           ["10", "Asegurar la mercancía en tránsito y crear una provisión para reclamaciones", "Protege contra pérdidas extraordinarias", "MEDIA"],
           ["11", "Registrar los kg y bultos por provincia en cada contenedor; más adelante, por cliente o agencia emisora",
            "Permite saber la rentabilidad por cliente y por provincia", "BAJA"],
           ["12", "Cuando el volumen supere unos 25 contenedores al mes, pasar el registro a una base de datos o a un sistema de gestión (ERP)",
            "Escalabilidad, varios usuarios y trazabilidad", "BAJA"]]),

    ("h2", "6. Indicadores que la dirección debe revisar cada mes"),
    ("b", "Utilidad antes de impuesto y margen % del mes y del año (DASHBOARD)."),
    ("b", "Costo total por kg e ingreso por kg. La diferencia es la utilidad por kg."),
    ("b", "Utilidad por región y su tendencia."),
    ("b", "Punto de equilibrio (kg y contenedores por mes) y margen de seguridad."),
    ("b", "Desvío del consumo de combustible respecto a la norma (%)."),
    ("b", "Días de ciclo: desde el arribo hasta el fin de la distribución."),
    ("b", "Costos fijos no absorbidos: indican capacidad ociosa."),

    ("h2", "7. Supuestos del modelo y datos por confirmar"),
    ("t", [["Supuesto", "Valor usado", "Fuente / estado"],
           ["Moneda de registro", "USD para todo", "Indicado por la empresa"],
           ["Ingreso", "Tarifa por kg por región", "Indicado por la empresa; el valor 0,80 USD/kg es ILUSTRATIVO"],
           ["Unidad de análisis", "Contenedor; consolidado mensual para los fijos", "Indicado por la empresa"],
           ["Cita del puerto", "10 USD por viaje", "Indicado por la empresa"],
           ["Combustible al puerto", "60 L por viaje", "Indicado por la empresa"],
           ["Precio del diésel", "2,00 USD/L", "Referencia oficial de CUPET para 2026 (diésel regular). Varía por operación: confirmar con las facturas"],
           ["Vehículo propio", "Solo en la recogida en el puerto", "Indicado por la empresa"],
           ["Distribución regional", "Servicio de terceros: chofer y ayudante con pago fijo mensual", "Indicado por la empresa"],
           ["Desagrupe", "Pago por jornal, igual en ENCI y A.V", "Indicado por la empresa"],
           ["Comercialización", "21 gestores (5 + 16) con fijo igual + variable según lo gestionado", "Indicado por la empresa; las tasas son ILUSTRATIVAS"],
           ["Tributos", "Ventas y servicios 10 %, territorial 1 %, seguridad social 14 %, fuerza de trabajo 5 %, utilidades 35 %",
            "Régimen general de MIPYMES (Ley 113 y Ley del Presupuesto 2026). Confirmar el régimen de la empresa con el contador y la ONAT"],
           ["Normas técnicas (km, capacidades, rendimientos, litros regionales)", "Valores de PARAMETROS", "ILUSTRATIVAS: medir en 2 o 3 contenedores reales y ajustar"],
           ["Gastos fijos y salarios", "Valores de COSTOS_FIJOS", "ILUSTRATIVOS: sustituir por los registros contables"]]),

    ("h2", "8. Plan de implantación"),
    ("n", "Semana 1: el económico carga PRECIOS, ACTIVOS, COSTOS_FIJOS y COMERCIAL con datos reales y borra los ejemplos."),
    ("n", "Semanas 2 a 4: operaciones registra los contenedores del mes en curso y, si hay datos, de los 2 o 3 meses anteriores."),
    ("n", "Mes 1: se miden en la práctica las normas de combustible y de desagrupe, y se ajusta PARAMETROS."),
    ("n", "Mes 2: primer cierre mensual con conciliación de la comercialización. La dirección analiza la tarifa mínima por región."),
    ("n", "Mes 3: decisión sobre las tarifas diferenciadas y la consolidación de viajes. Se revisa el resultado frente a la contabilidad oficial."),
]
