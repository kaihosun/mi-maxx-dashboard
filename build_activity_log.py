#!/usr/bin/env python3
"""
Build Activity Log (Admin Staff) - COMPLETE.xlsx
English translation of Registro Actividades (Personal Admin) - COMPLETO.xlsx
MI Technologies MTY MAXX — TV refurbishment maquiladora, Monterrey
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
#  COLORS
# ─────────────────────────────────────────────
WHITE      = "FFFFFF"
BLACK      = "000000"
LIGHT_GRAY = "F2F2F2"

# ─────────────────────────────────────────────
#  SHEET NAME MAP
# ─────────────────────────────────────────────
SHEET_NAMES = {
    "Almacén":      "Warehouse",
    "Auditoría":    "Audit",
    "Calidad":      "Quality",
    "Clasificación":"Sorting",
    "FFT":          "FFT Line",
    "Incoming":     "Incoming",
    "Logística":    "Logistics",
    "Open Cell":    "Open Cell",
    "Shipping B2B": "Shipping B2B",
    "Shipping B2C": "Shipping B2C",
    "Paletizado":   "Palletizing",
    "EHS":          "EHS",
    "Mantenimiento":"Maintenance",
    "IT":           "IT",
    "RH":           "HR",
}

# ─────────────────────────────────────────────
#  FULL TRANSLATION DICTIONARY
#  key = exact Spanish string; value = English
# ─────────────────────────────────────────────
TRANSLATIONS = {
    # ── Headers / meta ──
    "Semana:": "Week:",
    "NOMBRE":  "NAME",
    "LUNES":   "MON",
    "MARTES":  "TUE",
    "MIÉRCOLES": "WED",
    "JUEVES":  "THU",
    "VIERNES": "FRI",
    "Documento preparado por el área de Capacitación y Desarrollo  |  MI Technologies MTY MAXX":
        "Prepared by the Training & Development Department  |  MI Technologies MTY MAXX",

    # ── actividad / task labels ──
    "actividad 1":  "task 1",
    "actividad 2":  "task 2",
    "actividad 3":  "task 3",
    "actividad 4":  "task 4",
    "actividad 5":  "task 5",
    "actividad 6":  "task 6",
    "actividad 7":  "task 7",
    "actividad 8":  "task 8",
    "actividad 9":  "task 9",
    "actividad 10": "task 10",
    "actividad 11": "task 11",
    "actividad 12": "task 12",
    "actividad 13": "task 13",
    "actividad 14": "task 14",
    "actividad 15": "task 15",
    "actividad 16": "task 16",
    "actividad 17": "task 17",
    "actividad 18": "task 18",
    "actividad 19": "task 19",
    "actividad 20": "task 20",
    "actividad 21": "task 21",

    # ── Common short activities ──
    "Registro Ordenes": "Order Entry",
    "Atencion a correos": "Email management",
    "Revision de correos": "Check and respond to emails",
    "Revision de correos en general (cargas de tijuana, proyectos nuevos y seguimiento a temas pendientes)":
        "Check all emails (TJ loads, new projects, open items follow-up)",
    "Revision de correos pendientes, de programacion del día (unidades por llegar)":
        "Review pending emails and day's schedule (inbound units)",
    "Manejo de formatos": "Fill out and manage forms",
    "Captura de asistencia, HC y TE": "Log daily attendance, HC & OT",
    "Captura de asistencia, HC": "Log daily attendance & HC",
    "Captura de asistencia, HC del equipo de mantenimiento": "Log maintenance team attendance & HC",
    "Captura de asistencia, HC del turno": "Log shift attendance & HC",
    "Captura de asistencia, HC del turno de paletizado": "Log palletizing shift attendance & HC",
    "Captura de asistencia, HC, TE": "Log daily attendance, HC & OT",
    "Captura asistencia, HC y TE": "Log daily attendance, HC & OT",
    "Captura de asitencia, HC y TE": "Log daily attendance, HC & OT",
    "Captura de asistencia y resultados de capacitaciones aplicadas":
        "Log attendance and training results",
    "Reporte de no conformidades": "Non-conformance reporting",
    "Creacion de ayudas visuales": "Create visual aids",
    "Creacion de ayudas visuales para linea de produccion": "Create visual aids for production line",
    "Creacion de ayudas visuales y manuales de proceso": "Create visual aids and process manuals",
    "Balanceo de lineas": "Line balancing",
    "Creacion de hojas PIVOT": "Build pivot tables",
    "Reporte de consumo de cajas": "Box consumption report",
    "Generacion de shipment": "Generate shipment",
    "Generacion de shipment y BOL": "Generate shipment & BOL",
    "Seguimiento a correos": "Follow up on emails",
    "Seguimiento a ordenes (Como van con el surtido)": "Follow up on orders (check fulfillment status)",
    "Seguimiento a estatus de ordenes": "Follow up on order status",
    "Seguimiento a demoras de ordenes": "Follow up on delayed orders",
    "Seguimiento a ordenes B2B y B2C": "Follow up on B2B and B2C orders",
    "Seguimiento a ordenes B2B, B2C": "Follow up on B2B and B2C orders",
    "Seguimiento a ordenes, B2B y B2C": "Follow up on B2B and B2C orders",
    "Seguimiento a ordenes, UPT, CALLVERY, FBA, FULL": "Follow up on orders: UPT, CALLVERY, FBA, FULL",
    "Seguimiento al Dashboard de ordenes": "Monitor orders dashboard",
    "Seguimiento a unidades en transito": "Track units in transit",
    "Seguimiento a material TBD pendiente por clasificar": "Follow up on TBD units pending classification",
    "Seguimiento a planeacion de logistica": "Follow up on logistics planning",
    "Seguimiento a proyectos de mejora con otros departamentos via correo electronico":
        "Track improvement projects with other depts via email",
    "Seguimiento a plan de unidades para soriana, HEB, Rack":
        "Follow up on unit plan for Soriana, HEB, Rack",
    "Seguimiento a surtido de FBA y FULL": "Follow up on FBA & FULL fulfillment",
    "Seguimiento a solicitudes de surtido pendientes del dia":
        "Follow up on pending fulfillment requests for the day",
    "Seguimiento a ordenes pendientes de surtir con el equipo de almacen":
        "Follow up on unfilled orders with the warehouse team",
    "Seguimiento a ordenes pendientes del dia: prioridad por fecha de compromiso de entrega":
        "Follow up on open orders — priority by commit date",
    "Seguimiento a SLAs de entrega por cliente (TRG, HV, Mayoristas) en Excel":
        "Track delivery SLAs by customer (TRG, HV, Wholesalers) in Excel",
    "Seguimiento a SLAs por tipo de orden: TRG (5 dias), HV, Mayorista (segun compromiso)":
        "Track SLAs by order type: TRG (5 days), HV, Wholesale (per commitment)",
    "Seguimiento a SLAs: FBA < 7 dias, FULL < 3 dias, BULKY < 10 dias con reporte diario":
        "Track SLAs: FBA < 7 days, FULL < 3 days, BULKY < 10 days — daily report",
    "Seguimiento a check list de cargas: esquineros, emplaye, etiquetado secuencial de pallets":
        "Check load checklist: corner guards, stretch wrap, sequential pallet labeling",
    "Seguimiento a material retenido por falta de documentacion hasta 48 hrs":
        "Follow up on material held for missing docs (up to 48 hrs)",
    "Seguimiento a contratos de nuevo ingreso con fechas de vencimiento":
        "Track new hire contracts and expiration dates",
    "Seguimiento a tramites de IMSS de colaboradores nuevos":
        "Follow up on IMSS paperwork for new hires",
    "Seguimiento a tickets abiertos de soporte con usuarios":
        "Follow up on open support tickets with users",
    "Seguimiento a lista de solicitudes pendientes: prioridad por impacto en produccion":
        "Track pending request list — priority by production impact",
    "Seguimiento a cotizaciones y tiempos de respuesta de proveedores externos de mantenimiento":
        "Follow up on quotes and response times from external maintenance vendors",
    "Seguimiento a correccion de condiciones inseguras reportadas por area":
        "Follow up on reported unsafe conditions by dept",
    "Seguimiento a acciones correctivas abiertas del area":
        "Follow up on open corrective actions in the area",
    "Seguimiento a acciones correctivas de auditoria 5S pendientes de cierre":
        "Follow up on pending 5S audit corrective actions",
    "Seguimiento a acciones correctivas de auditorias 5S: estatus, responsable y fecha cierre":
        "Track 5S corrective actions: status, owner, and close date",
    "Seguimiento a indicadores clave de calidad en dashboard semanal":
        "Track key quality KPIs in weekly dashboard",
    "Seguimiento a indicadores de produccion del area (UPH y rendimiento por tipo de reparacion)":
        "Track area production KPIs (UPH and yield by repair type)",
    "Seguimiento a proyecto de capacitacion con Arturo Badillo — revision de avances":
        "Track training project with Arturo Badillo — progress review",
    "Seguimiento a alertas de Reset Factory Obligatorio (AC-MTY-PRU-ELE-001-RESET)":
        "Track mandatory factory reset alerts (AC-MTY-PRU-ELE-001-RESET)",
    "Seguimiento al proyecto de control de ordenes paletizado (coordinacion con Roman de Leon FFT)":
        "Track palletizing order control project (coordinate with Roman de Leon, FFT)",
    "Seguimiento al proyecto de estadia": "Track dwell-time project",
    "Seguimiento al log de produccion en sistema SmartControl (pallets creados y cerrados)":
        "Monitor production log in SmartControl (pallets created and closed)",
    "Seguimiento de acciones correctivas 5S en Excel con semaforo de estatus porque no hay herramienta de gestion de acciones correctivas en sistema":
        "Track 5S corrective actions in Excel with RAG status — no corrective action module in the system",
    "Seguimiento de documentacion para despacho de unidades":
        "Follow up on dispatch documentation for outbound units",
    "Seguimiento de estatus de unidades en transito en sistema Manifest":
        "Track in-transit unit status in Manifest",
    "Seguimiento de movimientos en los andenes": "Monitor dock activity",
    "Seguimiento de tags en cargas nacionales": "Track tags on domestic loads",
    "Seguimiento de unidades en mantenimiento": "Track units in maintenance",
    "Seguimiento documental y trazabilidad de recepciones del dia":
        "Document tracking and traceability for day's receipts",
    "Seguimientos a los formatos de envios del dia, ordenes atendidas y UPH":
        "Follow up on daily shipping forms, orders fulfilled, and UPH",
    "Seguimiento a promesas de horario de caja con impacto en produccion":
        "Follow up on trailer arrival commitments that impact production",
    "Seguimiento a reconciliacion de inventario fisico vs sistema por turno":
        "Follow up on physical vs system inventory reconciliation per shift",
    "Seguimiento a reconciliacion semanal de RMA recibidos vs esperados segun manifiesto":
        "Weekly reconciliation follow-up: RMAs received vs expected per manifest",
    "Seguimiento a incidencias de ticket hasta cierre con el area responsable":
        "Follow up on tickets until closed with the responsible dept",

    # ── Reconciliation / manual builds ──
    "Reconciliacion de inventario fisico vs BinManager por turno":
        "Physical vs BinManager inventory reconciliation per shift",
    "Reconciliacion semanal de RMA recibidos vs esperados segun manifiestos":
        "Weekly reconciliation: RMAs received vs expected per manifest",
    "Conciliacion diaria manual de guias generadas vs entregadas por paqueteria en Excel (DHL, FedEx, Estafeta, Redpack) porque NODO/SOKA no consolida por paqueteria":
        "Manual daily reconciliation of tracking numbers generated vs delivered by carrier (DHL, FedEx, Estafeta, Redpack) in Excel — NODO/SOKA doesn't consolidate by carrier",
    "Conciliacion manual de RMA recibidos vs manifiestos de devolucion porque los sistemas no cruzan esta informacion":
        "Manual reconciliation of RMAs received vs return manifests — systems don't cross-check this",
    "Conciliacion manual de gastos de operadores vs viajes realizados en hoja de control porque Manifest no tiene modulo de liquidaciones":
        "Manual reconciliation of driver expenses vs trips in tracking sheet — Manifest has no settlement module",
    "Conciliacion manual de manifiesto BOL vs conteo fisico en hoja Excel porque el sistema no hace este cruce automaticamente":
        "Manual reconciliation of BOL manifest vs physical count in Excel — the system doesn't do this cross-check automatically",
    "Conciliacion manual de ordenes despachadas vs plataforma del cliente (TRG portal) al cierre porque los sistemas no estan sincronizados":
        "Manual reconciliation of dispatched orders vs customer portal (TRG) at EOD — systems aren't synced",
    "Conciliacion manual de ordenes entre plataformas (ML, Amazon, Walmart) y BinManager porque los sistemas no estan integrados":
        "Manual reconciliation of orders across platforms (ML, Amazon, Walmart) vs BinManager — systems aren't integrated",
    "Conciliacion manual de pallets fisicos vs pallets registrados en SmartControl al cierre del turno porque el sistema no detecta diferencias automaticamente":
        "Manual reconciliation of physical pallets vs SmartControl at shift close — the system doesn't flag discrepancies automatically",
    "Conciliacion manual entre BinManager y conteo fisico por rack porque el sistema no genera reporte de diferencias automatico":
        "Manual reconciliation of BinManager vs physical count by rack — system doesn't generate a discrepancy report automatically",
    "Conciliacion manual entre inventario de placas en BinManager vs inventario fisico porque el sistema no refleja salidas de material de reparacion en tiempo real":
        "Manual reconciliation of boards inventory in BinManager vs physical — system doesn't reflect repair material outflows in real time",
    "Apoyo en conciliacion de manifiesto vs conteo fisico al final de descarga":
        "Assist with manifest vs physical count reconciliation at end of unload",
    "Comparacion de cantidades manifiesto BOL vs conteo fisico al descargar":
        "Compare BOL manifest quantities vs physical count during unload",

    # ── Manual reports / Excel builds ──
    "Armado de bitacora de mantenimiento por equipo en Excel con historial de fallas y reparaciones porque el sistema no tiene este registro":
        "Build maintenance log per equipment in Excel with fault/repair history — the system doesn't track this",
    "Armado de reporte de unidades en transito con estatus actualizado porque los sistemas del cliente (TRG portal) no se sincronizan con BinManager":
        "Build in-transit unit report with current status — TRG portal doesn't sync with BinManager",
    "Armado de reporte semanal de cumplimiento de SLA por cliente en Excel porque Return ProApp no tiene este reporte":
        "Build weekly SLA compliance report by customer in Excel — Return ProApp doesn't have this report",
    "Armado de tabla Excel de ordenes Not Found con numero de serie, ubicacion y estatus porque BinManager no las agrupa":
        "Build Excel table of Not Found orders with serial number, location, and status — BinManager doesn't group them",
    "Armado de tabla de compatibilidades (TV vs mainboard vs pantalla) en Excel porque no existe catalogo en sistema":
        "Build TV-mainboard-screen compatibility table in Excel — no catalog exists in the system",
    "Armado de tabla de incidencias de nomina en Excel (faltas, retardos, permisos, TE) porque el sistema no las exporta en el formato de contabilidad":
        "Build payroll incident table in Excel (absences, tardiness, leaves, OT) — the system doesn't export in accounting's format",
    "Armado de tabla de seguimiento de incidencias en Excel porque no hay modulo de tickets en sistema":
        "Build incident tracking table in Excel — no ticket module in the system",
    "Armado de tabla de unidades TBD pendientes con numero de serie y dias en espera porque el sistema no las agrupa":
        "Build TBD pending units table with serial number and days waiting — the system doesn't group them",
    "Armado manual de bitacora de incidentes de sistema en Excel (fecha, area afectada, causa, tiempo de resolucion) porque no hay herramienta ITSM":
        "Manually build system incident log in Excel (date, affected area, cause, resolution time) — no ITSM tool in place",
    "Armado manual de matriz de capacitacion de seguridad en Excel (temas, fechas, asistencia, firma) porque no hay sistema de gestion de capacitacion":
        "Manually build safety training matrix in Excel (topics, dates, attendance, signature) — no training management system",
    "Armado manual de reporte de alertas de calidad activas porque el sistema no consolida el estatus de todas las alertas":
        "Manually build active quality alerts report — the system doesn't consolidate all alert statuses",
    "Armado manual de reporte de consumo de cajas y accesorios en Excel porque BinManager no cruza consumo vs produccion":
        "Manually build box and accessory consumption report in Excel — BinManager doesn't cross-check consumption vs production",
    "Construccion de PIVOT de produccion por linea, turno y clasificacion porque SmartControl no tiene esta vista":
        "Build production PIVOT by line, shift, and classification — SmartControl doesn't have this view",
    "Construccion de estudio de balanceo de lineas en Excel (tiempos por estacion, UPH, cuellos de botella) porque SmartControl no tiene modulo de balanceo":
        "Build line balancing study in Excel (station times, UPH, bottlenecks) — SmartControl has no balancing module",
    "Construccion de grafica de tendencia semanal de defectos en Excel para presentacion a gerencia":
        "Build weekly defect trend chart in Excel for management presentation",
    "Construccion de reporte de disponibilidad de SmartControl y BinManager en Excel para presentacion semanal porque los sistemas no tienen dashboard de uptime":
        "Build SmartControl & BinManager uptime report in Excel for weekly review — systems have no uptime dashboard",
    "Construccion de reporte de indicadores EHS en Excel (frecuencia de accidentes, condiciones inseguras, horas-hombre) porque el sistema no tiene modulo EHS":
        "Build EHS KPI report in Excel (incident frequency, unsafe conditions, man-hours) — no EHS module in the system",
    "Construccion de reporte de tiempo muerto por equipo y causa en Excel para analisis de confiabilidad porque el sistema no registra paros de equipo":
        "Build downtime report by equipment and cause in Excel for reliability analysis — the system doesn't log equipment stoppages",
    "Construccion manual de log de contenedores descargados en Excel (placa, sello, manifiesto, hora, unidades, danos) porque Trackly no genera este reporte":
        "Manually build unloaded container log in Excel (plate, seal, manifest, time, units, damages) — Trackly doesn't generate this report",
    "Construccion manual de reporte UPH por clasificador en Excel porque SmartControl no desglosa productividad individual":
        "Manually build UPH report by sorter in Excel — SmartControl doesn't break down individual productivity",
    "Construccion manual de reporte de asistencia consolidado por area en Excel porque el sistema de control de personal no genera el formato que requiere corporativo":
        "Manually build consolidated attendance report by dept in Excel — the personnel system doesn't produce the format corporate needs",
    "Construccion manual de reporte de citas de carga pendientes porque el sistema no tiene modulo de programacion de andenes":
        "Manually build pending load appointment report — the system has no dock scheduling module",
    "Construccion manual de reporte de inventario por clasificacion (GRA/GRB/GRC) en Excel porque BinManager no tiene este filtro combinado":
        "Manually build inventory report by classification (GRA/GRB/GRC) in Excel — BinManager doesn't have this combined filter",
    "Construccion manual de reporte de pallets por destino y estatus en Excel porque SmartControl no tiene vista de kanban de paletizado":
        "Manually build pallet report by destination and status in Excel — SmartControl has no palletizing kanban view",
    "Construccion manual de reporte de rendimiento de reparacion por tipo (pantalla, mainboard, t-con) en Excel porque SmartControl no desglosa por tipo de reparacion":
        "Manually build repair yield report by type (screen, mainboard, t-con) in Excel — SmartControl doesn't break down by repair type",
    "Elaboracion de Pareto de no conformidades en Excel por tipo de defecto, turno y linea porque SmartControl no tiene graficas de tendencia":
        "Build Pareto chart of non-conformances in Excel by defect type, shift, and line — SmartControl has no trend charts",
    "Elaboracion de reporte de efectividad de entrega por paqueteria en Excel porque la plataforma no genera este analisis":
        "Build carrier delivery effectiveness report in Excel — the platform doesn't generate this analysis",
    "Elaboracion de reporte de headcount por area y turno en Excel semanal porque el sistema no tiene esta vista consolidada":
        "Build weekly headcount report by dept and shift in Excel — the system has no consolidated view",
    "Elaboracion del manifiesto de carga en Excel (num orden, SKU, cant, clasificacion, peso) que usan Auditoria y Calidad para auditar la carga, el sistema no genera este formato":
        "Build load manifest in Excel (order #, SKU, qty, classification, weight) used by Audit and Quality to inspect the load — the system doesn't generate this format",
    "Elaboracion manual de ordenes de trabajo en Excel porque no hay modulo de OTs en SmartControl ni en BinManager":
        "Manually build work orders in Excel — no work order module in SmartControl or BinManager",
    "Elaboracion manual de plan de carga diario en Excel con asignacion por destino (TRG/HEB/Soriana/B2B) porque el sistema no genera plan automatico":
        "Manually build daily load plan in Excel with assignment by destination (TRG/HEB/Soriana/B2B) — the system doesn't auto-generate the plan",
    "Elaboracion manual de reporte de FRM generado por turno en Excel para enviarlo a Open Cell (el sistema no genera este cruce)":
        "Manually build FRM report by shift in Excel to send to Open Cell — the system doesn't generate this cross",
    "Elaboracion manual de reporte de diferencias inventario fisico vs BinManager en Excel por ubicacion (el sistema no lo genera)":
        "Manually build physical vs BinManager inventory discrepancy report in Excel by location — the system doesn't generate it",
    "Captura manual de clasificaciones especiales y excepciones en tabla Excel porque SmartControl no tiene campo para excepciones":
        "Manually log special classifications and exceptions in Excel — SmartControl has no exceptions field",
    "Captura manual de configuracion de estiba en hoja de control porque SmartControl no tiene tabla de pulgadas por tipo de caja":
        "Manually log stack configuration in tracking sheet — SmartControl has no inches-by-box-type table",
    "Captura manual de manifiestos de scrap en Excel porque SmartControl no tiene modulo de scrap/destruccion":
        "Manually log scrap manifests in Excel — SmartControl has no scrap/destruction module",
    "Captura manual de movimientos de material en hoja de control porque SmartControl no registra traslados internos":
        "Manually log material movements in tracking sheet — SmartControl doesn't record internal transfers",
    "Captura manual de numeros de serie con dano en Excel para reporte a calidad porque SmartControl no tiene campo de dano al recibir":
        "Manually log damaged serial numbers in Excel for quality reporting — SmartControl has no damage field on receipt",
    "Captura manual de resultados de inspeccion en formato Excel porque SmartControl no tiene modulo de calidad integrado":
        "Manually log inspection results in Excel — SmartControl has no integrated quality module",
    "Captura manual de solicitudes de soporte en hoja de control porque no hay sistema de tickets formalmente implementado":
        "Manually log support requests in tracking sheet — no ticket system formally in place",
    "Control manual de inventario de equipo de computo en Excel (laptops, scanners, impresoras Zebra) porque el sistema no tiene modulo de activos IT":
        "Manually manage IT equipment inventory in Excel (laptops, scanners, Zebra printers) — no IT asset module in the system",
    "Control manual de inventario de refacciones en Excel (stock, punto de reorden, proveedor) porque el sistema no contempla inventario de mantenimiento":
        "Manually manage spare parts inventory in Excel (stock, reorder point, vendor) — the system doesn't cover maintenance inventory",
    "Control manual de vencimiento de contratos en Excel con alertas de fecha porque el sistema no tiene notificaciones automaticas":
        "Manually track contract expiration dates in Excel with date alerts — the system has no automatic notifications",
    "Transcripcion de checklists de equipos moviles de papel a Excel para reporte mensual porque no hay sistema de captura digital":
        "Transcribe mobile equipment checklists from paper to Excel for monthly report — no digital capture system",

    # ── Printing / labels ──
    "Reimpresion de TRG ID cuando etiqueta no es leida por scanner en anden de shipping":
        "Reprint TRG ID label when scanner can't read it at shipping dock",
    "Reimpresion manual de etiquetas Pallet ID cuando la impresora Zebra pierde conexion con SmartControl":
        "Manually reprint Pallet ID labels when the Zebra printer loses connection with SmartControl",
    "Verificacion de etiquetado correcto (FNSKU + BOX label) vs SKU del PDF de Ventas":
        "Verify correct labeling (FNSKU + BOX label) vs SKU on Sales PDF",
    "Apoyo a impresion de etiquetas": "Help with label printing",
    "Verificacion de etiquetas y codigos de productos en minimarket":
        "Verify product labels and barcodes in minimarket",
    "Verificacion de etiquetas y manifiestos en pallets antes de liberar a shipping":
        "Verify labels and manifests on pallets before releasing to shipping",

    # ── Coordination ──
    "Coordinacion con IT (Jose Sanchez) para baja sistemica de LPNs procesados confirmados":
        "Coordinate with IT (Jose Sanchez) for LPN closeouts on confirmed processed units",
    "Coordinacion con IT (Jose Sanchez) para baja sistemica de LPNs via WhatsApp":
        "Coordinate with IT (Jose Sanchez) for LPN closeouts via WhatsApp",
    "Coordinacion con IT para baja sistemica de LPNs procesados":
        "Coordinate with IT for LPN closeouts on processed units",
    "Coordinacion con Cain sobre areas de capacitacion prioritarias del turno":
        "Coordinate with Cain on priority training areas for the shift",
    "Coordinacion con Itzel sobre unidades en transito con destino nacional":
        "Coordinate with Itzel on in-transit units with domestic destination",
    "Coordinacion con Open Cell sobre material FRM que genera el area de clasificacion":
        "Coordinate with Open Cell on FRM material generated by Sorting",
    "Coordinacion con Open Cell sobre solicitudes de accesorios para proceso":
        "Coordinate with Open Cell on accessory requests for processing",
    "Coordinacion con Seguridad para custodia de apertura y cierre de unidades en anden":
        "Coordinate with Security for trailer opening and closing custody at dock",
    "Coordinacion con Sorting sobre material FRM que llega a Open Cell":
        "Coordinate with Sorting on FRM material arriving to Open Cell",
    "Coordinacion con calidad (Candido Rodas) sobre unidades con dano identificado en descarga":
        "Coordinate with Quality (Candido Rodas) on units with damage identified at unload",
    "Coordinacion con contabilidad para cierre de nomina semanal":
        "Coordinate with Accounting for weekly payroll close",
    "Coordinacion con ejecutivo bancario Santander para tramites del personal":
        "Coordinate with Santander bank rep for employee banking procedures",
    "Coordinacion con infraestructura de TI en Tijuana para problemas mayores":
        "Coordinate with IT infrastructure team in Tijuana for major issues",
    "Coordinacion con inspector de calidad para liberacion y sello de pallets TRG":
        "Coordinate with quality inspector for TRG pallet sign-off and seal",
    "Coordinacion con logistica sobre llegadas de contenedores TJ-MTY":
        "Coordinate with Logistics on incoming TJ-MTY container arrivals",
    "Coordinacion con mantenimiento sobre reparaciones de seguridad prioritarias":
        "Coordinate with Maintenance on priority safety repairs",
    "Coordinacion con montacarguista sobre destino de material (Sorting vs bines)":
        "Coordinate with forklift operator on material destination (Sorting vs bins)",
    "Coordinacion con operaciones sobre pool de cajas disponibles para carga":
        "Coordinate with Operations on available box pool for loading",
    "Coordinacion con otros departamentos sobre material comprometido para B2B/B2C":
        "Coordinate with other depts on material committed to B2B/B2C orders",
    "Coordinacion con pickers de almacen para surtido y re-etiquetado FBA/FULL/BULKY":
        "Coordinate with warehouse pickers for FBA/FULL/BULKY fulfillment and re-labeling",
    "Coordinacion con produccion sobre lotes con alta tasa de no conformidad":
        "Coordinate with Production on lots with high non-conformance rate",
    "Coordinacion con produccion sobre prioridades de surtido del dia":
        "Coordinate with Production on day's fulfillment priorities",
    "Coordinacion con produccion sobre unidades en cuarentena pendientes de resolucion":
        "Coordinate with Production on quarantined units pending resolution",
    "Coordinacion con proveedor externo cuando reparacion requiere especialista fuera de planta":
        "Coordinate with external vendor when repair needs an off-site specialist",
    "Coordinacion con proveedores externos de mantenimiento para reparaciones especializadas":
        "Coordinate with external maintenance vendors for specialized repairs",
    "Coordinacion con shipping sobre espacio disponible en anden para carga":
        "Coordinate with Shipping on available dock space for loading",
    "Coordinacion con shipping sobre material comprometido en ordenes del dia":
        "Coordinate with Shipping on material committed to today's orders",
    "Coordinacion con supervisores de area para correccion de condiciones inseguras en 24 hrs":
        "Coordinate with area supervisors to fix unsafe conditions within 24 hrs",
    "Coordinacion con supervisores para programacion de vacaciones del personal":
        "Coordinate with supervisors on employee vacation scheduling",
    "Coordinacion de cuadrilla para movimientos de material entre racks":
        "Coordinate crew for material moves between racks",
    "Coordinacion de espacio en piso para pallets entrantes con logistica":
        "Coordinate floor space for incoming pallets with Logistics",
    "Coordinacion de ordenes de trabajo y control de refacciones y proveedores externos":
        "Coordinate work orders and manage spare parts & external vendors",
    "Coordinacion de preparacion de ordenes y seguimiento documental de embarques":
        "Coordinate order prep and shipping documentation follow-up",
    "Coordinacion de ronda de firmas del manifiesto con las 5 areas (B2B/Seg/Cal/Alm/Log)":
        "Coordinate manifest signature round with all 5 areas (B2B/Security/Quality/Warehouse/Logistics)",
    "Coordinacion de solicitudes de unidades de transporte y seguimiento de arribo":
        "Coordinate transport unit requests and track arrivals",
    "Coordinacion de soporte de sistemas y resolucion de incidentes":
        "Coordinate system support and incident resolution",
    "Coordinacion y supervision de mantenimiento correctivo y preventivo":
        "Coordinate and supervise corrective and preventive maintenance",
    "Coordinar cajas con logistica": "Coordinate boxes with Logistics",
    "Auditoria 5S por area — coordinacion con supervisores para acciones correctivas":
        "5S audit by area — coordinate corrective actions with supervisors",

    # ── Communication ──
    "Comunicacion a areas sobre tiempo estimado de resolucion de incidentes":
        "Notify depts on estimated incident resolution time",
    "Comunicacion con almacen para confirmar baja sistematica de LPNs despachados":
        "Confirm with Warehouse that dispatched LPNs are closed out in the system",
    "Comunicacion con clientes B2B sobre estatus de pedidos pendientes":
        "Communicate with B2B customers on pending order status",
    "Comunicacion con operaciones sobre prioridades del dia via WhatsApp":
        "Communicate day's priorities to Operations via WhatsApp",
    "Comunicacion con produccion sobre disponibilidad de material en almacen":
        "Communicate material availability in warehouse to Production",
    "Comunicacion con supervisores sobre avances y problemas del turno via WhatsApp":
        "Update supervisors on shift progress and issues via WhatsApp",
    "Comunicacion de anomalias de inventario al area de operaciones":
        "Report inventory discrepancies to Operations",
    "Comunicacion de anomalias de inventario al area de operaciones responsable":
        "Report inventory discrepancies to the responsible Operations area",
    "Comunicacion de beneficios y novedades de nomina al personal":
        "Communicate payroll updates and benefits to staff",
    "Comunicacion de discrepancias de descarga al supervisor de incoming":
        "Report unload discrepancies to Incoming supervisor",
    "Comunicacion de faltantes al area de almacen via WhatsApp":
        "Report shortages to Warehouse via WhatsApp",
    "Comunicacion de novedades al supervisor Fernando Tobias via WhatsApp":
        "Update supervisor Fernando Tobias on day's news via WhatsApp",
    "Comunicacion de problemas de calidad en contenedor al area de logistica":
        "Report container quality issues to Logistics",
    "Comunicacion por el canal de whatsapp para indicaciones y actividades":
        "Communicate instructions and task updates via WhatsApp",
    "Comunicado de restablecimiento de sistemas al grupo operativo via WhatsApp":
        "Notify operations group that systems are back up via WhatsApp",
    "Comunicados internos al personal via grupos de WhatsApp":
        "Internal communications to staff via WhatsApp groups",

    # ── Inventory / warehouse ──
    "Alta de inventario": "Inventory intake / stock entry",
    "Altas y bajas en sistema de insumos": "Log inbound and outbound supplies in system",
    "Altas y salidas de material": "Log material intake and outflows",
    "Salidas de material": "Issue materials / material outflows",
    "Control de inventario": "Inventory control",
    "Descargar inventarios (Binmanager)": "Download inventory reports from BinManager",
    "Revision de inventario (Binmanager)": "Review inventory in BinManager",
    "Revision de inventarios, para contemplar la produccion de material y capacidad de material":
        "Review inventory levels to plan material production and capacity",
    "Revision a inventarios, accesorios y cajas": "Review inventory, accessories, and boxes",
    "Conteo fisico de pallets por area y conciliacion contra registros en BinManager":
        "Physical pallet count by area and reconciliation against BinManager records",
    "Validacion de stock disponible en Bin Manager al recibir orden de Ventas/Planeacion":
        "Validate available stock in BinManager when receiving order from Sales/Planning",
    "Verificacion de consistencia de registros en BinManager":
        "Verify record consistency in BinManager",
    "Se realiza inventario y auditoria por racks, se realiza la verificacion en sistema":
        "Rack-by-rack physical inventory and audit with system verification",
    "Actualizacion de ubicaciones en BinManager tras movimientos de material":
        "Update locations in BinManager after material moves",
    "Actualizacion del estatus de pallets en sistema tras ubicacion en bines o Sorting":
        "Update pallet status in system after placing in bins or Sorting",
    "Actualizacion de inventario de minimarket en hoja de control":
        "Update minimarket inventory in tracking sheet",
    "Actualizacion de formatos de control de inventarios":
        "Update inventory control forms",
    "Descargas de Excel para planeacion": "Download Excel files for planning",
    "Reconciliacion de inventario fisico vs BinManager": "Physical vs BinManager inventory reconciliation",

    # ── BinManager / SmartControl operations ──
    "Gestion de ordenes con estatus Not Found en BinManager":
        "Manage Not Found status orders in BinManager",
    "Gestion de ordenes con estatus Not Found: investigacion con almacen via manifiesto firmado":
        "Manage Not Found orders — investigate with warehouse via signed manifest",
    "Bajas sistemicas de LPNs procesados en SmartControl":
        "Systematic LPN closeouts in SmartControl",
    "Confirmacion de registro exitoso en Bin Manager y SmartControl (doble verificacion)":
        "Confirm successful entry in BinManager and SmartControl (double-check)",
    "Actualizacion de manifest": "Update Manifest",
    "Registro de datos de arribo en plataforma Manifest (numero de caja, sello, transportista)":
        "Log arrival data in Manifest platform (trailer #, seal, carrier)",
    "Registro de salida de unidades en sistema Manifest (estatus vacia y despachada)":
        "Log unit departure in Manifest (status: empty and dispatched)",
    "Creacion y cierre de Pallet ID en modulo Pallet de SmartControl con evidencia fotografica":
        "Create and close Pallet IDs in SmartControl's Pallet module with photo evidence",
    "Registro de unidades procesadas en SmartControl por turno":
        "Log units processed in SmartControl per shift",
    "Agregar modulos de Trabajo SmartControl": "Add work modules in SmartControl",
    "Soporte operativo de sistemas SmartControl y BinManager":
        "Operational support for SmartControl and BinManager",
    "Revision de configuracion de modulos nuevos en SmartControl":
        "Review configuration of new SmartControl modules",
    "Apoyo en configuracion de nuevos usuarios en SmartControl":
        "Help configure new SmartControl users",
    "Monitoreo de disponibilidad de SmartControl y BinManager":
        "Monitor SmartControl and BinManager availability",

    # ── Shipping / dispatch ──
    "Verificacion de pallets auditados antes de despacho (marcado de pallet liberado)":
        "Verify pallets are audited and stamped before dispatch",
    "Verificacion de que pallets auditados tienen marcado de auditoria antes de despacho":
        "Verify audited pallets have the audit stamp before dispatch",
    "Reporte de tiempo de caja colocada en rampa": "Track dock-to-dock dwell time per trailer",
    "Reporte de tiempo de caja colocada en rampa vs tiempo de descarga real":
        "Track dock dwell time vs actual unload time per trailer",
    "Revision de historial de LPN de salida": "Review outbound LPN history",
    "Revision de historial de LPN de salida para validacion de trazabilidad completa":
        "Review outbound LPN history to validate full traceability",
    "Apoyo en documentacion de despacho y cierre de shipments":
        "Help with dispatch documentation and shipment closeout",
    "Control de manifiestos de salida y seguimiento a ordenes con incidencias":
        "Manage outbound manifests and follow up on orders with issues",
    "Recoleccion de firmas en manifiestos de carga por las 4 areas (Logistica/Calidad/Seg/Alm)":
        "Collect signatures on load manifests from all 4 areas (Logistics/Quality/Security/Warehouse)",
    "Despacho de unidades": "Dispatch units",
    "Despacho de unidades de cargadas y colocacion de sellos de unidades cargadas":
        "Dispatch loaded units and apply seals to loaded trailers",
    "Cierre y colocacion de sellos": "Close and apply seals",
    "Generacion de correos para solicitud o gestion de material": "Send emails to request or manage material",
    "Envio de comprobantes de salida sellados a archivo de logistica":
        "Send sealed outbound proof documents to logistics archive",
    "Envio de evidencia de cargas": "Send load evidence (photos)",
    "Envio de informacion de unidad y sello para cargas del dia":
        "Send trailer and seal info for today's loads",
    "Envio de informacion, retroalimentacion de daños recibidos":
        "Send damage feedback on received units",
    "Envio de informacion: Reportes de carga": "Send load reports",
    "Envio de notificacion interna de arribo de unidad a Shipping e Incoming":
        "Send internal arrival notification to Shipping and Incoming",
    "Envio de notificaciones a Compras y Supervisor por material en cuarentena":
        "Send quarantine notifications to Purchasing and Supervisor",
    "Envio de plan semanal": "Send weekly plan",
    "Envio de reporte de costos": "Send cost report",
    "Reportar unidades que se presentan para carga": "Report units ready for loading",
    "Reporte de unidades que se presentan para carga": "Report units ready for loading",
    "Plan de envio semanal con SUP": "Weekly shipping plan with supervisor",
    "Planeacion semanal de cargas": "Weekly load planning",
    "Planeacion semanal para sig semana": "Planning next week's schedule",
    "Solicitud de unidades para carga (depende del plan del dia)":
        "Request transport units for loading (depends on the day's plan)",
    "Gestion con proveedores de transporte para asignacion de unidades segun requerimiento":
        "Manage transport providers for unit assignment per requirement",
    "Planeacion de operadores para cruce": "Schedule drivers for cross-border runs",
    "Solicitud de recurso para disel": "Request diesel budget",
    "Solicitud de recurso para disel, movimientos locales": "Request diesel budget for local moves",
    "Solicitud de ticket de gastos de los operadores que salen a cruce":
        "Request expense tickets for drivers going on cross-border runs",
    "Revision de gastos de operadores y liquidacion": "Review driver expenses and process settlements",
    "Revision de plan de envios": "Review shipping plan",
    "Actualizacion de reporte de import y export (movimientos del dia)":
        "Update import and export report (day's movements)",
    "Actualizacion de reportes de logistica": "Update logistics reports",
    "Documentacion de import o export del día": "Document today's import or export",
    "Reporte de asistencia de operadores de transporte, viajes y TE (de personal y para carga)":
        "Report transport operator attendance, trips, and OT (staff & load related)",

    # ── Receiving / incoming ──
    "Descarga de manifiesto de cada pallet que produce FRM":
        "Download manifest for each FRM pallet produced",
    "Captura de devoluciones en sistema (RMA incoming)": "Log returns in system (RMA incoming)",
    "Captura de devoluciones recibidas en BinManager": "Log received returns in BinManager",
    "Captura de devoluciones y RMA recibidos en sistema": "Log returns and RMAs received in system",
    "Apoyo en conteo fisico de unidades durante descarga de contenedores":
        "Assist with physical unit count during container unload",
    "Registro de condicion de unidades al recibir (danos visibles identificados)":
        "Log unit condition at receipt (visible damages identified)",
    "Registro de discrepancias detectadas en descarga en formato F-MTY-INC-DISC-001":
        "Log discrepancies detected at unload in format F-MTY-INC-DISC-001",
    "Registro de entradas de material": "Log material intake",
    "Registro de entradas de planas electronicas": "Log electronic board intake",
    "Reporte fotografico de danos en unidades descargadas enviado a supervisor via WhatsApp":
        "Photo report of damaged units at unload sent to supervisor via WhatsApp",
    "Reporte de contenedores (Problemas identificados)": "Container report (issues identified)",
    "Reporte de unidad descargada (Check list)": "Unloaded unit report (checklist)",
    "Reporte de unidades con dano evidente en descarga al area de calidad":
        "Report units with visible damage at unload to Quality",
    "Verificacion de manifiestos vs ordenes de compra (PO) en Bin Manager":
        "Verify manifests vs purchase orders (PO) in BinManager",
    "Verificacion de unidades descargadas": "Verify unloaded units",
    "Actualizacion de reporte de recepciones del dia (cantidades, proveedores, estatus)":
        "Update day's receipts report (quantities, vendors, status)",
    "Actualizacion de plan de programacion de descargas y asignacion de andenes":
        "Update unload schedule and dock assignment plan",
    "Recepcion de notificacion de orden por correo (PDF) y validacion de stock en Bin Manager":
        "Receive order notification by email (PDF) and validate stock in BinManager",
    "Seguimiento documental y trazabilidad de recepciones del dia":
        "Document tracking and traceability for today's receipts",

    # ── Palletizing ──
    "Reporte de unidades paletizadas por turno en formato de produccion":
        "Report palletized units per shift in production format",
    "Verificacion de altura maxima de pallet (88 pulgadas) antes de emplayado":
        "Verify max pallet height (88 inches) before wrapping",
    "Verificacion de configuracion de estiba segun tabla de pulgadas antes de armar pallet":
        "Verify stack configuration per inches table before building pallet",
    "Transferencia de pallets al area correcta en modulo Transfer (AREA04 apto / AREA05 no apto)":
        "Transfer pallets to the correct area in Transfer module (AREA04 fit / AREA05 unfit)",
    "Apoyo en registro en SmartControl y gestion de etiquetas Pallet ID":
        "Help with SmartControl entry and Pallet ID label management",
    "Actualizacion de kanban de pallets pendientes de cargar por cliente":
        "Update kanban of pallets pending loading by customer",
    "Verificacion de separacion de marca VIZIO vs otras marcas en pallets (excepto ICD/ICX)":
        "Verify VIZIO brand separation vs other brands on pallets (except ICD/ICX)",
    "Organizacion y seguimiento de pallets por destino (TRG/FBA/FULL/B2B)":
        "Organize and track pallets by destination (TRG/FBA/FULL/B2B)",
    "Proyecto a seguimiento de ordenes (Control de ordenes - palletizado)":
        "Order tracking project (order control — palletizing)",
    "Reporte de unidades cargadas vs comprometidas en formato de cierre":
        "Report loaded vs committed units in EOD close format",
    "Reporte de pallets liberados por calidad y listos para cargar":
        "Report pallets released by quality and ready to load",
    "Reporte de pallets en espera de inyeccion de TRG ID a sistemas para su liberacion":
        "Report pallets waiting for TRG ID injection into systems for release",
    "Reporte de pallets en cuarentena comunicados a Soporte MI Technologies para decision de destino":
        "Report quarantined pallets to MI Technologies Support for destination decision",
    "Control de produccion y coordinacion de flujo de pallets por destino":
        "Production control and pallet flow coordination by destination",
    "Apoyo en gestion de etiquetado TRG ID y transferencia de pallets en SmartControl":
        "Help with TRG ID labeling and pallet transfers in SmartControl",
    "Captura de info, al cierre de turno": "Log info at shift close",

    # ── Quality ──
    "Captura de novedades en formato de auditoria diario": "Log items in daily audit format",
    "Reporte de no conformidades": "Non-conformance report",
    "Modificacion a las alertas de calidad": "Update quality alerts",
    "Monitoreo de las alertas de calidad": "Monitor quality alerts",
    "Aplicacion de mejoras al sistema de calidad": "Apply improvements to the quality system",
    "Reporte de acciones correctivas": "Corrective action report",
    "Creacion de app con IA (para calidad)": "Build AI-powered app (for Quality)",
    "Apoyo en auditoria AQL y control de carga por cliente B2B":
        "Assist with AQL audit and load control by B2B customer",
    "Apoyo en auditoria de ordenes entry y revision de historial de LPN de salida":
        "Assist with order entry audit and outbound LPN history review",
    "Apoyo en inspeccion AQL de pallets: verificacion cosmetica y conteo de piezas vs manifiesto":
        "Assist with AQL pallet inspection: cosmetic check and piece count vs manifest",
    "Registro de diferencias detectadas en formato de auditoria diario con descripcion de causa":
        "Log discrepancies in daily audit format with cause description",
    "Registro de incidencias de proceso en formato de control":
        "Log process incidents in tracking format",
    "Reporte de mercancias en retencion con descripcion de causa":
        "Report merchandise on hold with cause description",
    "Reporte de ordenes B2C con riesgo de suspension de cuenta (ML/Amazon) para escalamiento":
        "Report B2C orders at risk of account suspension (ML/Amazon) for escalation",
    "Reporte de ordenes con SLA en riesgo al supervisor via WhatsApp":
        "Report SLA-at-risk orders to supervisor via WhatsApp",
    "Reporte de ordenes con riesgo de suspension de cuenta (ML/Amazon) al supervisor":
        "Report orders at risk of account suspension (ML/Amazon) to supervisor",
    "Reporte de unidades encontradas con clasificacion incorrecta en rack":
        "Report units found with incorrect classification in rack",
    "Reporte de problematicas de clasificacion identificadas en turno al supervisor":
        "Report classification issues identified during shift to supervisor",
    "Verificacion de calidad de reparaciones antes de transferir a FFT":
        "Verify repair quality before transferring to FFT",
    "Verificacion de manifiestos de pallet vs conteo fisico antes de liberar":
        "Verify pallet manifests vs physical count before releasing",
    "Verificacion de funcionamiento correcto de equipo post-reparacion":
        "Verify equipment works correctly after repair",
    "Verificacion de funcionamiento correcto de equipo post-reparacion antes de liberarlo":
        "Verify equipment works correctly after repair before releasing",
    "Reporte por pallets (producto no conforme dentro de un pallet)":
        "Pallet-level non-conformance report (non-conforming product within a pallet)",
    "Armado manual de reporte de alertas de calidad activas porque el sistema no consolida el estatus de todas las alertas":
        "Manually build active quality alerts report — the system doesn't consolidate all alert statuses",
    "Reporte de discrepancias de inventario en BinManager al area de auditoria":
        "Report BinManager inventory discrepancies to the Audit area",
    "Reporte de discrepancias fisico vs sistema a operaciones via WhatsApp":
        "Report physical vs system discrepancies to Operations via WhatsApp",

    # ── Classification / Sorting ──
    "Actualizacion de tabla de clasificaciones especiales y excepciones en Excel":
        "Update special classifications and exceptions table in Excel",
    "Actualizacion de kanban de produccion con material disponible para proceso":
        "Update production kanban with material available for processing",
    "Actualizacion de programacion del dia en dashboard tras cambios de ultimo momento":
        "Update day's schedule in dashboard after last-minute changes",
    "Busqueda de versiones de tv´s en BinManager": "Look up TV versions in BinManager",
    "Revision de material para proceso en BinManager": "Review material for processing in BinManager",
    "Revision de material DMA en binmanager": "Review DMA material in BinManager",
    "Revision de material DMA en binmanager para pruebas y material para proceso de captura de versión":
        "Review DMA material in BinManager for testing and version capture process",
    "Actualizacion de tracking en NODO/SOKA tras colecta de paqueteria":
        "Update tracking in NODO/SOKA after carrier pickup",
    "Actualizacion de tracking en NODO/SOKA tras colecta de paqueterias del dia":
        "Update tracking in NODO/SOKA after all carrier pickups for the day",
    "Atencion a paqueterias": "Handle carrier pickups",
    "Registro de piezas colectadas por paqueteria en formato de reporte diario":
        "Log pieces picked up by carrier in daily report format",
    "Reporte de cuantas piezas se lleva la paqueterias":
        "Report how many pieces each carrier picked up",
    "Creacion de ordenes, TRG, High Value, Mayoristas":
        "Create orders: TRG, High Value, Wholesale",
    "Asignar prioridad a ordenes que tienen que salir en el dia":
        "Assign priority to orders that need to ship today",
    "Asignación de datos en el Kanban": "Assign data in the Kanban",
    "Atención al dashboard (Programa)": "Monitor the dashboard (schedule)",
    "Actualizacion de programacion del dia en dashboard": "Update day's schedule in dashboard",
    "Planeaciones de produccion generales": "General production planning",
    "Revision de kanban y planeacion de produccion": "Review kanban and plan production",
    "Revision de material: que no tenga material comprometido o que no se pueda ir (validacion)":
        "Review material: confirm nothing is committed or blocked (validation)",
    "Control de docuemntacion sagrada (Manifiesto, ordenes, BOL, etc)":
        "Manage key shipping documents (Manifest, orders, BOL, etc.)",
    "Seguimiento de documentacion para despacho de unidades":
        "Follow up on dispatch documentation for outbound units",
    "Control y seguimiento a unidades con problematicas":
        "Track and follow up on units with issues",
    "Reporte de salidas de mercancia diario": "Daily outbound merchandise report",
    "Revision de order entry: estatus de ordenes pendientes y ordenes dadas de baja confirmadas":
        "Review order entry: pending order status and confirmed closed orders",
    "Revision de orden entry, estatus de ordenes pendientes":
        "Review order entry, status of pending orders",
    "Revision de ordenes que esten de baja": "Review orders that have been closed out",
    "Registro de las ordenes que ya estan dadas de baja y se registra el numero de orden":
        "Log orders that have been closed out with their order number",
    "Revision de ordenes de compra y envios TJ-MTY": "Review purchase orders and TJ-MTY shipments",
    "Revision de nuevas ordenes de compra": "Review new purchase orders",
    "Revision de check list para cargas": "Review loading checklist",

    # ── FFT / Production ──
    "Llenado de bitacora": "Fill out production log",
    "Creacion de diagramas de flujo": "Create flowcharts",
    "Creacion de diagrama de flujo de lineas": "Create line flow diagrams",
    "Creacion y llenado de reportes (Produccion, UPH, Costos)":
        "Create and fill out reports (Production, UPH, Costs)",
    "Generacion de reporte de produccion": "Generate production report",
    "Generacion de reporte diario de produccion, actividades del dia, etc":
        "Generate daily production report, day's activities, etc.",
    "Generacion de reporte UPH": "Generate UPH report",
    "Generacion de reporte de costos": "Generate cost report",
    "Reporte de hora por hora de produccion": "Hourly production report",
    "Reporte de unidades terminadas por dia al supervisor Alvaro Lugo":
        "Report finished units per day to supervisor Alvaro Lugo",
    "Creacion y seguimiento al proyecto manual para la estandarizacion y control de procesos en el area de produccion":
        "Create and track manual project for process standardization and control in production",
    "Revision al proyecto de badillo (actualizacion de formatos)":
        "Review Badillo's project (form updates)",
    "Revision al proyecto de balanceo y estandarizacion":
        "Review line balancing and standardization project",
    "Revision al proyecto de estandarizacion y documentacion de procesos":
        "Review process standardization and documentation project",
    "Creacion de manuales para linea de produccion": "Create manuals for production line",
    "Creacion de manuales de usuario para procesos": "Create user manuals for processes",
    "Organizacion de personal en lineas (Excel)": "Organize staff assignments to lines in Excel",
    "Correcciones en sistema de produccion": "Corrections in production system",
    "Actualizacion de bitacora de mantenimiento correctivo y preventivo del turno":
        "Update corrective and preventive maintenance log for the shift",
    "Actualizacion de bitacora de mantenimiento correctivo y preventivo por equipo":
        "Update corrective and preventive maintenance log by equipment",
    "actualizacion de reporte bitacora diaria de operaciones (generacion de reporte)":
        "Update daily operations log report (generate report)",
    "Capacitaciones al personal de nuevo ingreso sobre procedimientos FFT":
        "Train new hires on FFT procedures",
    "Elaboracion de procedimientos de trabajo estandar (SOP) para lineas FFT":
        "Develop standard operating procedures (SOPs) for FFT lines",
    "Reporte semanal de productividad del area a gerencia":
        "Weekly productivity report to management",
    "Creacion de manifiestos": "Create manifests",
    "Creacion de tickets para solicitud de contenedores": "Create tickets to request containers",
    "Control de matriz de actividades": "Manage activity matrix",
    "Acomodo de material capturado": "Organize captured material",
    "Validacion de captura de reportes, que los esten realizando de forma correcta":
        "Validate report entries — make sure they're being filled out correctly",
    "Planeacion/creacion de ordenes de compra": "Plan and create purchase orders",
    "Creacion de Graficas de Produccion": "Create production charts",
    "Reporte semanal de avances del proyecto de cubicaje a Cain Bautista":
        "Weekly cubicage project progress report to Cain Bautista",

    # ── Open Cell ──
    "Gestion de compatibilidades (Accesorios)": "Manage accessory compatibility catalog",
    "Actualizacion de lista de compatibilidades de accesorios en Drive":
        "Update accessory compatibility list in Drive",
    "Analisis e investigación sobre compatibilidades":
        "Research and analyze compatibility",
    "Elaboracion de tabla de compatibilidades de accesorios en Excel porque SmartControl no tiene catalogo de compatibilidades":
        "Build accessory compatibility table in Excel — SmartControl has no compatibility catalog",
    "Armado de tabla de compatibilidades (TV vs mainboard vs pantalla) en Excel porque no existe catalogo en sistema":
        "Build TV-mainboard-screen compatibility table in Excel — no catalog exists in the system",
    "Descarga de manifiesto de cada pallet que produce FRM":
        "Download manifest for each FRM pallet produced",
    "Verifica el manifiesto FRM con el area de OC": "Verify FRM manifest with Open Cell area",
    "Verificación material FRM": "Verify FRM material",
    "Solicitud de material para proceso (Reparacion/Opencell)":
        "Request materials for processing (Repair/Open Cell)",
    "solicitud de material para proceso": "Request materials for processing",
    "Revision de compatibilidades disponibles para proceso de reparacion":
        "Review available compatibility options for repair process",
    "Busqueda en inventarios internos (Flex, t-con, adaptadores, etc)":
        "Search internal inventories (Flex, t-con, adapters, etc.)",
    "Busqueda de Tarjetas T-CON y adaptadores en inventarios internos":
        "Search for T-CON boards and adapters in internal inventories",
    "Rastreo de Kits en Inventarios de mainboards": "Track kits in mainboard inventories",
    "Rastreo de LPN en registro de produccion": "Track LPNs in production log",
    "capturas de inventarios de flex": "Log flex inventory",
    "capturas de inventarios de t-con": "Log t-con inventory",
    "Registro de salidas (PEN,DMA,FRM,POC)": "Log outflows (PEN, DMA, FRM, POC)",
    "Control de salidas de placas electronicas y registro":
        "Control and log electronic board outflows",
    "Registro de datos placas electronicas": "Log electronic board data",
    "altas y bajas de consumo de open cell": "Log Open Cell material consumption (in and out)",
    "altas y bajas de entrega de accesorios": "Log accessory delivery (in and out)",
    "requerimientos de kits para proceso": "Kit requirements for processing",
    "Captura de datos placas electronicas": "Log electronic board data",
    "Control de Kamban para Opencell": "Manage Kanban for Open Cell",
    "creacion de manifiesto de entradas y salidas de scrap": "Create scrap in/out manifest",
    "Planeación listas de compra (paneles)": "Plan purchase lists (panels)",

    # ── EHS ──
    "Publicacion de capsula REHS diaria en grupos de WhatsApp de planta":
        "Post daily REHS safety capsule to plant WhatsApp groups",
    "Publicacion de capsula REHS diaria en grupos de WhatsApp por departamento":
        "Post daily REHS safety capsule to department WhatsApp groups",
    "Recorrido de seguridad en planta con registro fotografico de condiciones inseguras":
        "Safety walkthrough on the floor with photo record of unsafe conditions",
    "Auditoria 5S por area: llenado de formato y calificacion por zona":
        "5S audit by area: fill out format and score by zone",
    "Apoyo en auditoria 5S por area: llenado de formato y calificacion por zona":
        "Assist with 5S audit by area: fill out format and score by zone",
    "Apoyo en planes preventivos de mantenimiento de equipos de seguridad":
        "Assist with preventive maintenance plans for safety equipment",
    "Apoyo en recorridos de seguridad y gestion de checklist de equipos moviles":
        "Assist with safety walkthroughs and mobile equipment checklist management",
    "Generacion de comunicados de seguridad para personal via WhatsApp":
        "Send safety announcements to staff via WhatsApp",
    "Gestion de EPP por area (tapones auditivos, lentes de seguridad, chalecos)":
        "Manage PPE by area (earplugs, safety glasses, vests)",
    "Reporte de condiciones inseguras identificadas con fotografia a Angie MIT":
        "Report unsafe conditions with photos to Angie MIT",
    "Reporte de condiciones inseguras identificadas en recorrido con registro fotografico":
        "Report unsafe conditions identified during walkthrough with photos",
    "Reporte de incidentes y cuasi-accidentes en formato EHS":
        "Report incidents and near-misses in EHS format",
    "Reporte semanal de indicadores EHS: incidentes, cuasi-accidentes y condiciones inseguras":
        "Weekly EHS KPI report: incidents, near-misses, and unsafe conditions",
    "Refuerzo de cultura de seguridad al personal durante recorrido en piso":
        "Reinforce safety culture with floor staff during walkthrough",
    "Supervision de condiciones de seguridad en piso de produccion":
        "Monitor safety conditions on the production floor",
    "Verificacion de estatus de extintores y senalizacion de emergencia en toda la planta":
        "Verify fire extinguisher status and emergency signage throughout the plant",
    "Verificacion de estatus de extintores y senalizacion de emergencia por area":
        "Verify fire extinguisher status and emergency signage by area",
    "Armado manual de matriz de capacitacion de seguridad en Excel (temas, fechas, asistencia, firma) porque no hay sistema de gestion de capacitacion":
        "Manually build safety training matrix in Excel (topics, dates, attendance, signature) — no training management system",
    "Programacion y registro de capacitaciones de seguridad e induccion a personal nuevo":
        "Schedule and log safety training and new hire inductions",

    # ── Maintenance ──
    "Checklist de mantenimiento preventivo de montacargas, patines y equipos moviles":
        "Preventive maintenance checklist for forklifts, pallet jacks, and mobile equipment",
    "Ejecucion de mantenimiento correctivo en equipos de planta":
        "Execute corrective maintenance on plant equipment",
    "Ejecucion de reparaciones de equipos asignados por Carlos Rojas":
        "Execute equipment repairs assigned by Carlos Rojas",
    "Programacion de mantenimiento preventivo de montacargas, patines y equipos criticos de planta":
        "Schedule preventive maintenance for forklifts, pallet jacks, and critical plant equipment",
    "Recepcion y registro de solicitudes de mantenimiento correctivo en formato de orden de trabajo":
        "Receive and log corrective maintenance requests in work order format",
    "Registro de actividades realizadas en bitacora de mantenimiento":
        "Log completed maintenance activities in the maintenance log",
    "Revision de lista de solicitudes de mantenimiento pendientes asignadas":
        "Review list of pending assigned maintenance requests",
    "Atencion a solicitudes de mantenimiento correctivo recibidas via WhatsApp":
        "Respond to corrective maintenance requests received via WhatsApp",
    "Apoyo en seguimiento a reparaciones y documentacion de trabajos ejecutados":
        "Assist with repair follow-up and documentation of completed work",
    "Reporte de estatus de equipos en reparacion al gerente de planta":
        "Report equipment repair status to plant manager",
    "Reporte de reparaciones completadas vs pendientes al supervisor Carlos Rojas":
        "Report completed vs pending repairs to supervisor Carlos Rojas",
    "Control de inventario de refacciones y materiales: solicitud de reposicion cuando stock bajo":
        "Control spare parts and materials inventory — request replenishment when stock is low",
    "Solicitud de apoyo a proveedor externo cuando la reparacion lo requiere":
        "Request external vendor support when the repair calls for it",
    "Solicitud de refacciones y materiales para mantenimiento (correos, WhatsApp)":
        "Request spare parts and materials for maintenance (emails, WhatsApp)",
    "Coordinacion con proveedor externo cuando reparacion requiere especialista fuera de planta":
        "Coordinate with external vendor when repair needs an off-site specialist",
    "Apoyo en registro de actividades de mantenimiento realizadas en bitacora del dia":
        "Help log completed maintenance activities in the daily log",
    "Verificacion de funcionamiento correcto de equipo post-reparacion":
        "Verify equipment works correctly after repair",
    "Transcripcion de checklists de equipos moviles de papel a Excel para reporte mensual porque no hay sistema de captura digital":
        "Transcribe mobile equipment checklists from paper to Excel for monthly report — no digital capture system",
    "Construccion de reporte de tiempo muerto por equipo y causa en Excel para analisis de confiabilidad porque el sistema no registra paros de equipo":
        "Build downtime report by equipment and cause in Excel for reliability analysis — the system doesn't log equipment stoppages",

    # ── IT ──
    "Soporte tecnico a usuarios y mantenimiento de equipos de computo":
        "Tech support to users and maintenance of computer equipment",
    "Atencion a solicitudes de soporte tecnico de usuarios en planta":
        "Respond to tech support requests from plant users",
    "Configuracion de equipos nuevos (laptops, scanners, impresoras)":
        "Set up new equipment (laptops, scanners, printers)",
    "Mantenimiento de usuarios y accesos en sistemas operativos":
        "Manage users and system access on operating platforms",
    "Gestion de acceso de nuevos colaboradores a sistemas (tarjeta, escaner, usuario)":
        "Set up new employee system access (badge, scanner, user account)",
    "Monitoreo de disponibilidad de SmartControl y BinManager":
        "Monitor SmartControl and BinManager availability",
    "Atencion a escalamientos de incidentes criticos de sistema":
        "Respond to escalated critical system incidents",
    "Comunicacion a areas sobre tiempo estimado de resolucion de incidentes":
        "Notify depts on estimated incident resolution time",
    "Comunicado de restablecimiento de sistemas al grupo operativo via WhatsApp":
        "Notify operations group that systems are back up via WhatsApp",
    "Coordinacion con infraestructura de TI en Tijuana para problemas mayores":
        "Coordinate with IT infrastructure team in Tijuana for major issues",
    "Escalamiento de problemas de sistema a IT via WhatsApp (impacto en captura de ordenes)":
        "Escalate system issues to IT via WhatsApp (impact on order entry)",
    "Atencion a errores de captura en sistema reportados por produccion":
        "Respond to system entry errors reported by Production",
    "Revision de trackly": "Review Trackly",
    "Revision de tags": "Review tags",
    "Analisis/reaparcion de sistema internos": "Analyze/troubleshoot internal systems",
    "Uso de IA para busquedas de investigacion": "Use AI for research and lookups",
    "Uso del canal de whatsapp para comunicacion y seguimiento a indicaciones":
        "Use WhatsApp for communications and instructions follow-up",
    "Reporte de errores criticos de sistema al coordinador Pavel MI":
        "Report critical system errors to coordinator Pavel MI",
    "Reporte de incidencias de hardware al coordinador Pavel MI":
        "Report hardware incidents to coordinator Pavel MI",
    "Reporte de status de sistemas al gerente de operaciones":
        "Report system status to operations manager",
    "Armado manual de bitacora de incidentes de sistema en Excel (fecha, area afectada, causa, tiempo de resolucion) porque no hay herramienta ITSM":
        "Manually build system incident log in Excel (date, affected area, cause, resolution time) — no ITSM tool in place",
    "Captura manual de solicitudes de soporte en hoja de control porque no hay sistema de tickets formalmente implementado":
        "Manually log support requests in tracking sheet — no ticket system formally in place",
    "Control manual de inventario de equipo de computo en Excel (laptops, scanners, impresoras Zebra) porque el sistema no tiene modulo de activos IT":
        "Manually manage IT equipment inventory in Excel (laptops, scanners, Zebra printers) — no IT asset module in the system",
    "Actualizacion de inventario de equipo de computo de planta":
        "Update plant computer equipment inventory",

    # ── HR ──
    "Gestion de nomina semanal y procesos de recursos humanos":
        "Weekly payroll management and HR processes",
    "Captura de incidencias y ausentismo en sistema de control de personal":
        "Log incidents and absenteeism in the personnel control system",
    "Recoleccion y captura de incidencias (faltas, retardos, permisos) en sistema":
        "Collect and log incidents (absences, tardiness, leaves) in system",
    "Recoleccion y captura de tiempo extra (TE) por turno para nomina":
        "Collect and log OT by shift for payroll",
    "Recoleccion de firmas en manifiestos de carga por las 4 areas (Logistica/Calidad/Seg/Alm)":
        "Collect signatures on load manifests from all 4 areas (Logistics/Quality/Security/Warehouse)",
    "Recoleccion de listas de asistencia por area (recepcion de formatos firmados)":
        "Collect attendance sheets by area (receive signed forms)",
    "Apoyo en proceso de seleccion y contratacion de personal operativo":
        "Support operative staff selection and hiring process",
    "Actualizacion de expedientes del personal con documentos nuevos":
        "Update employee files with new documents",
    "Actualizacion de perfiles operativos (Actualizacion de los carnet de los operadores)":
        "Update operator profiles (update operator badges/IDs)",
    "Actualizacion de matriz de capacitacion (temas, fechas, asistencia) por area de planta":
        "Update training matrix (topics, dates, attendance) by plant area",
    "Actualizacion y llenado de formatos de capacitacion": "Update and fill out training forms",
    "Actualizacion de formatos de capacitacion": "Update training forms",
    "Apoyo en recoleccion de firmas de listas de asistencia por area":
        "Help collect signed attendance sheets by area",
    "Elaboracion de presentaciones para capacitacion de personal (PowerPoint, Word)":
        "Prepare training presentations for staff (PowerPoint, Word)",
    "Generacion de presentaciones para capacitacion de personal (word, power point, etc)":
        "Create training presentations for staff (Word, PowerPoint, etc.)",
    "Modulos de capacitacion interna": "Internal training modules",
    "Generacion de reportes de asistencia diaria para gerencia":
        "Generate daily attendance reports for management",
    "Gestion administrativa y apoyo en control de personal":
        "Administrative management and personnel control support",
    "Apoyo administrativo en procesos de RH y coordinacion con proveedores":
        "Administrative support for HR processes and vendor coordination",
    "Gestion de conflictos entre solicitudes de cajas de distintas areas":
        "Handle conflicts between box requests from different areas",
    "Coordinacion con contabilidad para cierre de nomina semanal":
        "Coordinate with Accounting for weekly payroll close",
    "Coordinacion con ejecutivo bancario Santander para tramites del personal":
        "Coordinate with Santander bank rep for employee banking procedures",
    "Coordinacion con supervisores para programacion de vacaciones del personal":
        "Coordinate with supervisors on employee vacation scheduling",
    "Comunicacion de beneficios y novedades de nomina al personal":
        "Communicate payroll updates and benefits to staff",
    "Seguimiento a tramites de IMSS de colaboradores nuevos":
        "Follow up on IMSS paperwork for new hires",
    "Seguimiento a contratos de nuevo ingreso con fechas de vencimiento":
        "Track new hire contracts and expiration dates",
    "Control manual de vencimiento de contratos en Excel con alertas de fecha porque el sistema no tiene notificaciones automaticas":
        "Manually track contract expiration dates in Excel with date alerts — the system has no automatic notifications",
    "Construccion manual de reporte de asistencia consolidado por area en Excel porque el sistema de control de personal no genera el formato que requiere corporativo":
        "Manually build consolidated attendance report by dept in Excel — the personnel system doesn't produce the format corporate needs",
    "Armado de tabla de incidencias de nomina en Excel (faltas, retardos, permisos, TE) porque el sistema no las exporta en el formato de contabilidad":
        "Build payroll incident table in Excel (absences, tardiness, leaves, OT) — the system doesn't export in accounting's format",
    "Registro en modulo de checador": "Log in the time-clock module",
    "Toma de asistencia al perosna y toma de headcount": "Take attendance and headcount",
    "Validacion de inventario, stock de minimarket": "Validate minimarket inventory and stock",
    "Captura de informacion de entradas y salidas al area de minimarket":
        "Log minimarket area inflows and outflows",
    "Captura de informacion de nuevo personal en base de datos de RH":
        "Enter new employee data in HR database",

    # ── Training / documentation (cross-dept) ──
    "Actualizacion de Kanban para OC": "Update Kanban for Open Cell",
    "Revision y actualizacion de Kanban": "Review and update Kanban",
    "Mejoras al dashboard de asistencia": "Improve attendance dashboard",
    "Validar información con otros departamentos": "Validate info with other departments",
    "Documentacion de procesos (planta general)": "Document processes (general plant)",
    "Control de versiones de documentos de proceso en carpeta compartida (Drive)":
        "Manage process document version control in shared folder (Drive)",
    "Verificacion de manuales actualizados en carpeta compartida":
        "Verify updated manuals in shared folder",
    "Elaboracion de presentaciones para capacitacion de personal (PowerPoint, Word)":
        "Prepare training presentations for staff (PowerPoint, Word)",
    "Registro y seguimiento de areas de mejoras identificados en produccion":
        "Log and track improvement areas identified in production",
    "Actualizacion de headcount por area en reporte semanal":
        "Update headcount by area in weekly report",
    "Verificacion de material capturado bueno": "Verify captured material is good",
    "Verificacion de pallet que se envie a PNP": "Verify pallet being sent to PNP",
    "Captura de consumo de insumos de empaque (cinchos, emplaye, esquineros, fleje)":
        "Log packaging supply consumption (zip ties, stretch wrap, corner guards, strapping)",
    "Reporte de consumo de materiales de empaque (cajas, emplaye, tarimas)":
        "Report packaging material consumption (boxes, stretch wrap, pallets)",
    "Reporte de consumo de accesorios y materiales por turno":
        "Report accessory and material consumption per shift",
    "Solicitudes de material para almacen": "Material requests for warehouse",
    "Solicitud de reposicion de materiales al supervisor de almacen":
        "Request material replenishment from warehouse supervisor",
    "Atención al proyecto de cubicaje para shipping": "Work on cubicage project for Shipping",
    "envio de plan diario": "Send daily plan",
    "Reportes generales del area (Movimientos de material, etc)":
        "General area reports (material movements, etc.)",
    "Registro de check-out de unidades en sistema y generacion de manifiesto de salida":
        "Log unit check-out in system and generate outbound manifest",
    "Registro de incidencias de carga en formato de reporte de embarque del dia":
        "Log load incidents in day's shipping report format",
    "Ubicaciones de televisiones en BinManager": "Look up TV locations in BinManager",
    "Revision de formatos de areas como: incoming, almacen y shipping (que los esten llenando)":
        "Review forms from Incoming, Warehouse, and Shipping (make sure they're filling them out)",
    "Revision de paneles para uso en reparacion": "Review panels available for repair use",
    "Revisión de material en kitsparts": "Review material in kit parts",
    "Reporte de unidades (TVS)": "TV unit report",
    "Apoyo en resolucion de reclamaciones con proveedor de transporte":
        "Assist with claims resolution with transport vendor",
    "Atencion al proyecto de cubicaje para shipping": "Work on cubicage project for Shipping",
    "Creacion y llenado de Drives en modelos de tv´s":
        "Create and fill out Drive folders for TV models",
    "Planeacion semanal para sig semana": "Plan next week's schedule",
    "Solicitud de material para proceso": "Request materials for processing",
    "Verificacion de LPN y numeros de serie en BM": "Verify LPNs and serial numbers in BinManager",
    "Reporte de estatus de estandarizacion a Cain Bautista":
        "Report standardization status to Cain Bautista",
    "Confirmacion de carga completada a Alejandro Barrientos con resumen del dia":
        "Confirm load completed to Alejandro Barrientos with day's summary",
}

# ─────────────────────────────────────────────
#  TITLE ROW TRANSLATIONS (dept-specific)
# ─────────────────────────────────────────────
def translate_title(value):
    if not value or not isinstance(value, str):
        return value
    if "Registro de Actividades" in value and "MI Technologies" in value:
        for sp, en in SHEET_NAMES.items():
            if sp in value:
                return value.replace("Registro de Actividades", "Activity Log").replace(f"— {sp}", f"— {en}")
        return value.replace("Registro de Actividades", "Activity Log")
    return value


def translate_cell(value):
    """Return the English translation of a cell value."""
    if not value or not isinstance(value, str):
        return value

    # Exact match first (includes supplemental)
    if value in TRANSLATIONS:
        return TRANSLATIONS[value]

    # Title row (Registro de Actividades → Activity Log + dept name swap)
    translated = translate_title(value)
    if translated != value:
        # Run through translations again in case the intermediate result matches
        return TRANSLATIONS.get(translated, translated)

    # actividad N (fallback)
    if value.startswith("actividad "):
        return value.replace("actividad ", "task ")

    # Leave everything else as-is (person names, codes, etc.)
    return value


# ─────────────────────────────────────────────
#  SUPPLEMENTAL TRANSLATIONS (missed entries)
# ─────────────────────────────────────────────
TRANSLATIONS.update({
    # ── Title rows with accented dept names ──
    "MI Technologies MTY MAXX  |  Activity Log  —  Almacén":
        "MI Technologies MTY MAXX  |  Activity Log  —  Warehouse",
    "MI Technologies MTY MAXX  |  Activity Log  —  Auditoría":
        "MI Technologies MTY MAXX  |  Activity Log  —  Audit",
    "MI Technologies MTY MAXX  |  Activity Log  —  Clasificación":
        "MI Technologies MTY MAXX  |  Activity Log  —  Sorting",
    "MI Technologies MTY MAXX  |  Activity Log  —  Logistica":
        "MI Technologies MTY MAXX  |  Activity Log  —  Logistics",
    "MI Technologies MTY MAXX  |  Activity Log  —  Logística":
        "MI Technologies MTY MAXX  |  Activity Log  —  Logistics",
    "MI Technologies MTY MAXX  |  Activity Log  —  Calidad":
        "MI Technologies MTY MAXX  |  Activity Log  —  Quality",
    "MI Technologies MTY MAXX  |  Activity Log  —  FFT":
        "MI Technologies MTY MAXX  |  Activity Log  —  FFT Line",
    "MI Technologies MTY MAXX  |  Activity Log  —  Paletizado":
        "MI Technologies MTY MAXX  |  Activity Log  —  Palletizing",
    "MI Technologies MTY MAXX  |  Activity Log  —  Mantenimiento":
        "MI Technologies MTY MAXX  |  Activity Log  —  Maintenance",
    "MI Technologies MTY MAXX  |  Activity Log  —  RH":
        "MI Technologies MTY MAXX  |  Activity Log  —  HR",
    "MI Technologies MTY MAXX  |  Activity Log  —  RH / Admin":
        "MI Technologies MTY MAXX  |  Activity Log  —  HR",
    "MI Technologies MTY MAXX  |  Activity Log  —  EHS / Seguridad":
        "MI Technologies MTY MAXX  |  Activity Log  —  EHS",
    "MI Technologies MTY MAXX  |  Activity Log  —  Opencel":
        "MI Technologies MTY MAXX  |  Activity Log  —  Open Cell",
    "MI Technologies MTY MAXX  |  Activity Log  —  IT / Sistemas":
        "MI Technologies MTY MAXX  |  Activity Log  —  IT",

    # ── Llenado / filling ──
    "Llenado de bitacora de produccion": "Fill out production log",
    "Llenado de bitacora de produccion del turno": "Fill out shift production log",
    "Llenado de bitacora de entradas y salidas del turno":
        "Fill out shift inbound/outbound log",
    "LLenado de tiempos de descarga (UPH)": "Fill out unload time log (UPH)",
    "Llenado de formato de UPH de descarga cuando hay actividad":
        "Fill out unload UPH form when there's activity",
    "Llenado de UPH interno": "Fill out internal UPH log",
    "Llenado de formatos de inspeccion de equipos de proteccion personal":
        "Fill out PPE inspection forms",
    "Llenado de ordenes de trabajo de mantenimiento en formato de control":
        "Fill out maintenance work orders in tracking format",
    "Llenado y recoleccion de checklists de equipos moviles (montacargas, patines)":
        "Fill out and collect mobile equipment checklists (forklifts, pallet jacks)",
    "Llenado y recoleccion de checklists diarios de equipos moviles (montacargas, patin electrico)":
        "Fill out and collect daily mobile equipment checklists (forklifts, electric pallet jack)",

    # ── Impresion / printing ──
    "Impresion de etiquetas TRG ID en modulo Print TRGID de SmartControl para pallets VIZIO":
        "Print TRG ID labels in SmartControl's Print TRGID module for VIZIO pallets",
    "Impresion manual de guias de envio cuando NODO/SOKA presenta falla o lentitud en el sistema":
        "Manually print shipping labels when NODO/SOKA is down or slow",
    "Impresion manual del check list de carga y ronda de firmas de 5 areas porque Return ProApp no tiene flujo de aprobacion":
        "Manually print the load checklist and collect signatures from 5 areas — Return ProApp has no approval workflow",
    "Impresion y archivo de manifiestos de carga firmados porque el sistema no los almacena digitalmente":
        "Print and file signed load manifests — the system doesn't store them digitally",
    "Impresion y verificacion de etiquetas Pallet ID (QR escaneable) por recepcion":
        "Print and verify Pallet ID labels (scannable QR) per receipt",
    "Generacion e impresion de manifiesto de embarque con desglose de SKUs y pallets":
        "Generate and print shipping manifest with SKU and pallet breakdown",

    # ── Levantamiento / incident logging ──
    "Levantamiento de incidencias identificadas (Ticket)":
        "Log and escalate identified incidents (ticket)",
    "Levantamiento de tickets de incidencias por diferencias sistematicas con descripcion de causa":
        "Log incident tickets for systematic discrepancies with cause description",

    # ── Inyeccion / trailer injection ──
    "Inyeccion de remolques": "Trailer check-in / injection into system",
    "Inyeccion de remolques (Solo cuando se solicita apoyo)":
        "Trailer check-in into system (only when support is requested)",

    # ── Final stragglers ──
    "Gestion de programa de capacitacion y seguimiento a acciones correctivas de seguridad":
        "Manage safety training program and follow up on corrective actions",
    "Gestion del programa de seguridad y salud en el trabajo":
        "Manage the occupational health & safety program",
    "Seguimiento a ordenes de compra de accesorios pendientes de llegada":
        "Follow up on accessory purchase orders pending arrival",
    "Gestion de tarjetas Santander del personal (altas, bajas, reposicion)":
        "Manage employee Santander cards (activations, cancellations, replacements)",
    "Seguimiento a documentos de transporte pendientes de firma con cada area":
        "Follow up on transport documents pending signature with each area",
    "Reporte de estatus de reparaciones en proceso al supervisor Carlos Rojas":
        "Report in-progress repair status to supervisor Carlos Rojas",
    "Seguimiento a ordenes de compra de pantallas y mainboards pendientes de llegada":
        "Follow up on screen and mainboard purchase orders pending arrival",
    "Captura de numero de serie de unidades con dano para rastreo en SmartControl":
        "Log serial numbers of damaged units for tracking in SmartControl",
    "Captura de las salidas de ordenes": "Log order outflows",
    "Gestion del modulo Order Received en Return ProApp para registro de material recibido":
        "Manage the Order Received module in Return ProApp to log incoming material",
    "Reporte de cumplimiento de ordenes vs programadas al final del dia":
        "Report orders fulfilled vs scheduled at EOD",
    "Gestion de ordenes FBA, FULL y BULKY: etiquetado y cumplimiento de SLA":
        "Manage FBA, FULL, and BULKY orders: labeling and SLA compliance",
    "Gestion de solicitudes de material urgente via WhatsApp":
        "Handle urgent material requests via WhatsApp",

    # ── Other missed ──
    "Implementacion de nuevos proyectos": "Implement new projects",
    "Verificacion de unidades para cargar": "Verify units ready for loading",
    "Revision de mercancia (RMA) que no sale de planta, para que no se de de baja en sistema":
        "Review RMA merchandise that hasn't left the plant — make sure it doesn't get closed out in the system",
    "Actualizacion de ayudas visuales cuando cambian clasificaciones de producto":
        "Update visual aids when product classifications change",
    "Gestion de cambios de ultima hora en ordenes del dia con almacen":
        "Handle last-minute order changes with the warehouse",
    "Gestion de ordenes con estatus Not Found con almacen via WhatsApp":
        "Manage Not Found status orders with warehouse via WhatsApp",
    "Gestion de unidades TBD pendientes de clasificacion — reporte al supervisor":
        "Manage TBD units pending classification — report to supervisor",
    "Validacion de kits de accesorios por clasificacion (GRA/GRB/GRC) antes de surtir":
        "Validate accessory kits by classification (GRA/GRB/GRC) before fulfilling",

    # ── Traffic Control ──
    "Mantener el tablero y reportes al día":
        "Keep the dashboard and reports up to date",
    "Gestión de papelería digital (recepción y archivo de documentos en PDF por correo)":
        "Digital document management (receive and archive PDF documents via email)",
    "Confirmación de citas de camiones vía correo electrónico (triangulación MI Technologies / JB Hunt)":
        "Confirm truck appointments via email (triangulation: MI Technologies / JB Hunt)",
    "Descarga de manifiestos en BinManager para lotes de high value (Global Stock)":
        "Download manifests in BinManager for high-value lots (Global Stock)",
    "Uso de Smart Control (módulo de compras recibidas)":
        "Use Smart Control (received purchases module)",
    "Uso de Excel para registro y control de información":
        "Use Excel for information logging and control",
    "Comunicación presencial con Incoming para coordinar descargas del día":
        "In-person coordination with Incoming to plan the day's unloads",
    "Coordinación con Sorting para solicitudes de pallets vírgenes":
        "Coordinate with Sorting for blank pallet requests",
    "Gestión de cambios de contenedor con Shipping":
        "Handle container changes with Shipping",
    "Recepción de copia de papelería de Calidad":
        "Receive document copies from Quality",
    "Supervisión y toma de fotografías durante auditorías previas al cierre de camión":
        "Supervise and photograph pre-closure trailer audits",
    "Manejo de incidencias de documentación (retener descargas hasta recibir papelería completa)":
        "Handle documentation incidents (hold unloads until full paperwork is received)",
    "Escalamiento de problemas con compañero o jefe directo":
        "Escalate issues to peer or direct supervisor",
})


# ─────────────────────────────────────────────
#  ROW CLASSIFICATION
# ─────────────────────────────────────────────
TASK_LABELS = {f"actividad {i}" for i in range(1, 22)} | {f"task {i}" for i in range(1, 22)}
PERSON_PLACEHOLDER = {"Persona 2","Persona 3","Persona 4","Persona 5","Persona 6",
                       "Persona 7","Persona 8","Persona 9","Persona 10","Persona 11",
                       "Persona 12","Persona 13","Persona 14","Persona 15"}
# Real person names (not headers/tasks/placeholders) — determined dynamically below
HEADER_TEXTS = {"NOMBRE", "NAME"}


def is_title_row(val):
    return isinstance(val, str) and "MI Technologies" in val and ("Registro" in val or "Activity" in val)

def is_week_row(val):
    return isinstance(val, str) and val.strip() in ("Semana:", "Week:")

def is_header_row(val):
    return isinstance(val, str) and val.strip() in HEADER_TEXTS

def is_task_row(val):
    return isinstance(val, str) and val.strip().lower() in {t.lower() for t in TASK_LABELS}

def is_footer_row(val):
    return isinstance(val, str) and ("Capacitación y Desarrollo" in val or "Training & Development" in val)

def is_person_row(val):
    if not isinstance(val, str):
        return False
    v = val.strip()
    if is_title_row(v) or is_week_row(v) or is_header_row(v) or is_task_row(v) or is_footer_row(v):
        return False
    if v in ("", "Semana:", "Week:"):
        return False
    # If it looks like a real person row: has a value that's not a task label and not a system code
    return True


# ─────────────────────────────────────────────
#  STYLES
# ─────────────────────────────────────────────
thin_side   = Side(border_style="thin", color=BLACK)
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

fill_black = PatternFill(fill_type="solid", fgColor=BLACK)
fill_gray  = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)
fill_white = PatternFill(fill_type="solid", fgColor=WHITE)

font_white_bold   = Font(name="Calibri", bold=True,   color=WHITE, size=11)
font_black_bold   = Font(name="Calibri", bold=True,   color=BLACK, size=11)
font_black_normal = Font(name="Calibri", bold=False,  color=BLACK, size=11)
font_title        = Font(name="Calibri", bold=True,   color=BLACK, size=12)
font_footer       = Font(name="Calibri", italic=True, color=BLACK, size=10)
font_week         = Font(name="Calibri", bold=True,   color=BLACK, size=11)

align_wrap_left = Alignment(horizontal="left",   vertical="top", wrap_text=True)
align_center    = Alignment(horizontal="center", vertical="center", wrap_text=False)

COL_A_WIDTH = 35
COL_BF_WIDTH = 45


# ─────────────────────────────────────────────
#  MAIN BUILD
# ─────────────────────────────────────────────
source_path = "/Users/eduardoflores/AI_Projects/Dashboards/Especiales/Registro Actividades (Personal Admin) - COMPLETO.xlsx"
out_path    = "/Users/eduardoflores/AI_Projects/Dashboards/Especiales/Activity Log (Admin Staff) - COMPLETE.xlsx"

src = openpyxl.load_workbook(source_path, data_only=True)
wb  = Workbook()
wb.remove(wb.active)  # drop default blank sheet

for sp_name in src.sheetnames:
    ws_in  = src[sp_name]
    en_name = SHEET_NAMES.get(sp_name, sp_name)
    ws_out  = wb.create_sheet(en_name)

    # Column widths
    ws_out.column_dimensions["A"].width = COL_A_WIDTH
    for col_letter in ["B","C","D","E","F"]:
        ws_out.column_dimensions[col_letter].width = COL_BF_WIDTH

    for row_idx, row in enumerate(ws_in.iter_rows(values_only=True), start=1):
        col_a = row[0] if row else None
        translated_row = [translate_cell(v) for v in row]

        # Write row
        for col_idx, val in enumerate(translated_row, start=1):
            cell = ws_out.cell(row=row_idx, column=col_idx, value=val)

        # Style by row type
        translated_a = translated_row[0] if translated_row else None

        if is_title_row(col_a) or is_title_row(translated_a):
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_title
                c.fill = fill_white
                c.alignment = align_wrap_left

        elif is_week_row(translated_a):
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_week
                c.fill = fill_white
                c.alignment = align_wrap_left

        elif is_header_row(translated_a):
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_white_bold
                c.fill = fill_black
                c.border = thin_border
                c.alignment = align_center

        elif is_footer_row(col_a) or is_footer_row(translated_a):
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_footer
                c.fill = fill_white
                c.alignment = align_wrap_left

        elif is_task_row(col_a) or is_task_row(translated_a):
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_black_normal
                c.fill = fill_white
                c.border = thin_border
                c.alignment = align_wrap_left

        elif col_a is None and all(v is None for v in row):
            # empty row — no border
            pass

        else:
            # person name row — light gray
            for col_idx in range(1, 7):
                c = ws_out.cell(row=row_idx, column=col_idx)
                c.font = font_black_bold
                c.fill = fill_gray
                c.border = thin_border
                c.alignment = align_wrap_left

    # Row heights: let openpyxl auto-size via wrap_text
    # Set a reasonable default row height for task rows
    for row_idx in range(1, ws_out.max_row + 1):
        ws_out.row_dimensions[row_idx].height = 45

# ─────────────────────────────────────────────
#  TRAFFIC CONTROL SHEET (hardcoded — not in source xlsx)
#  Abdon Alejandro Melero Alvarado training Reymundo Emmanuel Peña Galarza
# ─────────────────────────────────────────────
TC_TITLE  = "MI Technologies MTY MAXX  |  Activity Log  —  Traffic Control"
TC_FOOTER = "Prepared by the Training & Development Department  |  MI Technologies MTY MAXX"

TC_COMPUTER_TASKS = [
    "Keep the dashboard and reports up to date",
    "Digital document management (receive and archive PDF documents via email)",
    "Confirm truck appointments via email (triangulation: MI Technologies / JB Hunt)",
    "Download manifests in BinManager for high-value lots (Global Stock)",
    "Use Smart Control (received purchases module)",
    "Use Excel for information logging and control",
]

TC_PHYSICAL_TASKS = [
    "In-person coordination with Incoming to plan the day's unloads",
    "Coordinate with Sorting for blank pallet requests",
    "Handle container changes with Shipping",
    "Receive document copies from Quality",
    "Supervise and photograph pre-closure trailer audits",
    "Handle documentation incidents (hold unloads until full paperwork is received)",
    "Escalate issues to peer or direct supervisor",
]

TC_ALL_TASKS = TC_COMPUTER_TASKS + TC_PHYSICAL_TASKS

TC_PEOPLE = [
    "Abdon Alejandro Melero Alvarado",
    "Reymundo Emmanuel Peña Galarza",
]

DAYS = ["MON", "TUE", "WED", "THU", "FRI"]

ws_tc = wb.create_sheet("Traffic Control")
ws_tc.column_dimensions["A"].width = COL_A_WIDTH
for col_letter in ["B", "C", "D", "E", "F"]:
    ws_tc.column_dimensions[col_letter].width = COL_BF_WIDTH

tc_rows = []
tc_rows.append([TC_TITLE, None, None, None, None, None])
tc_rows.append(["Week:", None, None, None, None, None])
tc_rows.append([None, None, None, None, None, None])
tc_rows.append(["NAME"] + DAYS)

for person in TC_PEOPLE:
    tc_rows.append([person] + [None, None, None, None, None])
    for i, task in enumerate(TC_ALL_TASKS, start=1):
        tc_rows.append([f"  task {i}"] + [task] * 5)
    tc_rows.append([None, None, None, None, None, None])

tc_rows.append([TC_FOOTER, None, None, None, None, None])

for row_idx, row_data in enumerate(tc_rows, start=1):
    col_a = row_data[0]
    for col_idx, val in enumerate(row_data, start=1):
        ws_tc.cell(row=row_idx, column=col_idx, value=val)

    if is_title_row(col_a):
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_title
            c.fill = fill_white
            c.alignment = align_wrap_left
    elif is_week_row(col_a):
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_week
            c.fill = fill_white
            c.alignment = align_wrap_left
    elif is_header_row(col_a):
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_white_bold
            c.fill = fill_black
            c.border = thin_border
            c.alignment = align_center
    elif is_footer_row(col_a):
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_footer
            c.fill = fill_white
            c.alignment = align_wrap_left
    elif is_task_row(col_a):
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_black_normal
            c.fill = fill_white
            c.border = thin_border
            c.alignment = align_wrap_left
    elif col_a is None:
        pass
    else:
        # person name row
        for col_idx in range(1, 7):
            c = ws_tc.cell(row=row_idx, column=col_idx)
            c.font = font_black_bold
            c.fill = fill_gray
            c.border = thin_border
            c.alignment = align_wrap_left

for row_idx in range(1, ws_tc.max_row + 1):
    ws_tc.row_dimensions[row_idx].height = 45

print(f"Saving to: {out_path}")
wb.save(out_path)
print("Done.")
