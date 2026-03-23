// Configuracion.gs

// --- IDs DE PLANILLAS Y CARPETAS ---
const IDs = {
  // Carpetas de Drive (donde están los archivos a cargar)
  origenAsignacion: '1zCpeOKoY-w5f_r4hBHSspp6yG9Ob_91f',
  origenFallecidos: '14Obi9-eH6Zn1R3PKGOiYJaEL8mN9ef-w',
  origenPrevision: '1rsYfcZ4DZXtH3Ui8u0v1QnhnbUMF1MSa',
  destinoProcesados: '1LeL3kaOrZQizMwf1lcVBSRS0hKoX3MN8',

  // Planilla principal de carga y distribución
  asignacionCupos: '18UOMlj50-nvAfL91T5JUnRPra4MSkfZisn06Dbo0ZI8',

  // Planillas de especialidades
  plantillaEspecialidad: '1Nj5C7pBwRrBBDhIP_NRlxeYDSW6bsSxNO_VNgsDKWWg', // Disponible en 999.DP/Contactabilidad/Plantillas/Prod_Plantilla Agentes
  carpetaEspecialidades: '1qsz5MXOJeEMkHQ6SZJ06qkckfNgvfcuA',

  // Planillas de destino del flujo
  preEgresos: '1aflPvKntrpvfpQsqoaVGS7r_PCkqFbVk71OnPXkxfvU', // PreEgresos original
  seguimiento: '1oEky-yzr6SalNRpHko-EmVbx3TrxTH52dgpNhOeIXb8',
  cambioEstado: '1C3liKUD2VyPm8OThkhuFZvnwDuhjeqHVTyD1Eh5sMSs',
  supervisora: '157T0Z1rugwEI0eVrbvpofy1TzioTqeHWp6RDBlsjyzo',
  carpetaAgentes: '1kSrb-zocfdEVFxHZGsjcefZMTido0X4N',
  plantillaAgente: '1xFuAja2QoeQUY6USR4kShZOhDsSDgmc6KKhyiCUnESE', // Disponible en 999.DP/CUPOS/CCP/Plantillas/Prod_Plantilla Especialidades HCUCH 

  // HISTÓRICOS
  historicoSupervisora: '1gMQJkimpLeiHafcWTumJLh4lQuMlXjk9PN119KIWnUI',
  historicoAgentes: '1zngpyCyipGSmM93Yl98vUy_W_GIkDT-KS0L-qLexP84',

  origenLimpieza: '', // Carpeta carga eliminaciones (Aún no se implementa el sistema de limpieza de casos eliminados)
  bandejaRevision: '1DIYPeo6Oezo36ajF_ds5aDj2WPP82r4TamxfKP7tjkc' // Planilla revisión adherencia
};

// --- MAPEO DE ESTADOS Y OTROS ---
const MAPAS = {
  estados: {
    1: "Recibido", 2: "Asignación de cupo", 3: "En campaña Agendar", 4: "En campaña Notificar", 5: "Operador Asignado Agendar", 6: "Operador Asignado Notificar", 7: "Notificado", 8: "No se presenta 1", 9: "Pre Egresado", 10: "Gestión territorial", 11: "Egresado"
  },

  causalesEgreso: {
    0: "GES", 1: "Atención realizada", 2: "Procedimiento y exámenes de apoyo de diagnóstico y terapéuticos", 3: "Indicación médica para reevaluación", 4: "Atención otorgada en el extra sistema", 5: "No beneficiario", 6: "Renuncia o rechazo voluntario", 7: "Recuperación espontánea", 8: "Inasistencias", 9: "Fallecimiento", 10: "Solicitud de indicación duplicada", 11: "Contacto no corresponde", 12: "No corresponde realizar cirugía", 13: "Traslado coordinado", 14: "No pertinencia", 15: "Caso erróneo (error de digitación)", 16: "Atención por resolutividad", 17: "Atención por Telemedicina", 18: "Modificación de la condición clínico-diagnóstica del caso", 19: "Atención por Hospital Digital", 20: "Postergaciones", 99: "Técnico administrativo nivel central"
  },

  ssalud: {
    9: "SSMN", 10: "SSMOC"
  },

  // MAPEO 1:1 (Solo las causales que pasan directo a PreEgresos)
  causalPorEstado: { 
    "FALLECIDO": 9, // Causal "Fallecimiento"
    "POSTERGA": 20 // Causal "Postergaciones"
  }
};

// --- MAPEO DE COLUMNAS PLANILLA ASIGNACIÓN DE CUPOS Y SEGUIMIENTO ---
const COLUMNAS_CUPOS = {
  ID_CASO: 1, FECHA_RECEPCION_HCUCH: 2, SERVICIO_SALUD: 3, ORIGEN: 4, TIPO_LE: 5, RUT: 6, DV: 7, VALIDADO_REG_CIV: 8, NOMBRES: 9, PRIMER_APELLIDO: 10, SEGUNDO_APELLIDO: 11, SEXO: 12, CALLE: 13, NRO: 14, RESTO_DIRECCION: 15, CIUDAD: 16, FECHA_NAC: 17, FECHA_DEF: 18, TELEFONO_FIJO: 19, TELEFONO_MOVIL: 20, EMAIL: 21, VALIDADO_FONASA: 22, ESTADO_PREVISIONAL: 23, PRAIS: 24, ESPECIALIDAD: 25, SOSPECHA_DIAGNOSTICA: 26, CONFIRMACION_DIAGNOSTICA: 27, PLANO: 28, PRESTA_EST: 30, ESPECIALIDAD_CORREGIDA: 31, FECHA_ENTRADA: 32, ID_LOCAL: 33, FUENTE: 34, CHECKBOX_ASIGNAR_AGENDA: 35, TIPO_ACCION: 36, FECHA_AGENDA: 37, HORA_AGENDA: 38, AGENDA: 39, PROFESIONAL: 40, CHECKBOX_ENVIAR_CC: 41, EVALUACION_SIC: 44, COMENTARIO: 45, COD_SEXO: 46, FECHA_RECIBIDO: 47, FECHA_ASIGNACION_CUPOS: 48, FECHA_EN_CAMP_AGENDAR: 49, FECHA_EN_CAMP_NOTIFICAR: 50, FECHA_OP_ASIGN_AGENDAR: 51, FECHA_OP_ASIGN_NOTIFICAR: 52, FECHA_NOTIFICADO: 53, FECHA_NSP1: 54, FECHA_PRE_EGRESADO: 55, FECHA_GESTION_TERRITORIAL: 56, FECHA_EGRESADO: 57, ESTADO_COD: 58, ESTADO_TEXTO: 59, SSALUD_COD: 60, ESTADO_DISTRIBUCION: 61
};

// --- MAPEO DE COLUMNAS PLANILLA SUPERVISORA ---
const COLUMNAS_SUPERVISORA = {
  ID_CASO: 1, RUT: 2, DV: 3, NOMBRES: 4, PRIMER_APELLIDO: 5, SEGUNDO_APELLIDO: 6, TIPO_ACCION: 7, COMUNA: 8, FECHA_ENTRADA: 9, PRESTACION: 10, SOLICITADO_POR: 11, COMENTARIO_ASIGNADOR: 12, AGENDA: 14, DIA: 15, HORA: 16, ESTADO_CC: 17, FECHA_RECEPCION_CC: 18, FECHA_INICIO_GESTION: 19, CANT_GESTIONES: 20, AGENTE: 21, CHECKBOX_ASIGNAR: 22, MAIL_1: 23, MAIL_2: 24, TELEFONO_1: 25, TELEFONO_2: 26, TELEFONO_3: 27, TELEFONO_4: 28, TELEFONO_5: 29, TELEFONO_6: 30, TELEFONO_7: 31, TELEFONO_8: 32, FUENTE: 33, PROCESADO_PREEGRESO: 34
};

// --- MAPEO DE COLUMNAS HOJA "AGENTES" EN PLANILLA SUPERVISORA ---
const COLUMNAS_AGENTES = {
  NOMBRE: 6, ESTADO_CREACION: 7, URL_PLANILLA: 8, FECHA_CREACION: 9 
};

// --- MAPEO DE COLUMNAS PLANILLAS DE AGENTES ---
const COLUMNAS_AGENTE = {
  ID_CASO: 1,
  AGENDAR: { 
    ESTADO_CONTACTO: 12, 
    ESTADO_ADHERENCIA: 14, 
    FECHA_CITA: 15, 
    HORA_CITA: 16, 
    CHECK_COPIA: 13, 
    FECHA_GESTION: 24 },
  NOTIFICAR: { 
    FECHA_CITA: 11, 
    HORA_CITA: 12, 
    ESTADO_CONTACTO: 14, 
    CHECK_COPIA: 15, 
    ESTADO_ADHERENCIA: 16, 
    FECHA_GESTION: 24 },
  ESTADO_CONTACTO_AGENDAR: 12, 
  CHECK_COPIA_AGENDAR: 13, 
  ESTADO_ADHERENCIA_AGENDAR: 14, 
  FECHA_GESTION: 24,
  ESTADO_CONTACTO_NOTIFICAR: 14, 
  CHECK_COPIA_NOTIFICAR: 15, 
  ESTADO_ADHERENCIA_NOTIFICAR: 16,
  COL_A_Y: 25
}

// --- MAPEO DE COLUMNAS PLANILLA PREEGRESOS (CAUSAL 5 Y 9) ---
const COLUMNAS_PREEGRESO = {
  ID_CASO: 1, FECHA_INGRESO_CAUSAL: 2, SSALUD: 3, RUT: 4, DV: 5, NOMBRES: 6, PRIMER_APELLIDO: 7, SEGUNDO_APELLIDO: 8, FECHA_NACIMIENTO: 9, FECHA_DEFUNCION: 10, ESPECIALIDAD: 11, SOSPECHA_DIAGNOSTICA: 12, COD_CAUSAL_EGRESO:  13, GLOSA_CAUSAL_EGRESO: 14, RESPONSABLE_PRIMERA: 15, ESTADO_PRIMERA_REVISION: 16, FECHA_PRIMERA_REVISION: 17, 
  RESPONSABLE_VALIDACION: 18, ESTADO_FINAL: 19, FECHA_ENVIO_SIDRA: 20, CAUSAL_EGRESO_REAL: 21, EGRESADO_POR: 22, EST_DESTINO_SIDRA: 23, COMENTARIOS: 24, CHECK_FINALIZAR: 25, CAUSAL_COD_FORMULA: 26, ESTAB_EGRESA: 27,  ESTAB_DESTINO: 28, ESTADO_NUEVO_COD: 29
};

// --- MAPEO DE COLUMNAS PLANILLA PREEGRESOS (OTRAS CAUSALES) ---
const COLUMNAS_PREEGRESO_OTRAS = {
  ID_CASO: 1, COD_CAUSAL_EGRESO: 15, ESTADO_REVISION_CASO: 18, FECHA_REVISION: 19, FECHA_ENVIO_SIDRA: 20, COMENTARIO: 24, CHECK_FINALIZAR: 25, CAUSAL_COD_FORMULA: 26, ESTAB_EGRESA: 27, ESTAB_DESTINO: 28,ESTADO_NUEVO_COD: 29
};

const NOMBRES_HOJAS = {
  asignacionCupos: "Casos y Cupos",
  supervisora: "Casos",
  seguimiento: "Casos y Cupos",
  cambioEstado: "2025",
  especialidad: "Casos y Cupos",
  pestañasAgente: ["Agendar", "Notificar"]
};

// PROCESO DE PRE EGRESO 1:1
const NOMBRES_HOJAS_PREEGRESO = {
  causal1y10: "Causal 1 y 10",
  causal5y9: "Causal 5 y 9",
  causal4_6_7_8_11: "Causal 4, 6, 7, 8 y 11",
  causal13_14_20: "Causal 13, 14 y 20"
};

// Mapeo para enrutar desde la Bandeja de Revisión
const MAPA_DESTINO_PREEGRESOS = {
  // Causal ID : Nombre de la Hoja en PreEgresos
  1: "Causal 1 y 10",
  10: "Causal 1 y 10", 
  4: "Causal 4, 6, 7, 8 y 11", 
  6: "Causal 4, 6, 7, 8 y 11",
  7: "Causal 4, 6, 7, 8 y 11",
  8: "Causal 4, 6, 7, 8 y 11",
  11: "Causal 4, 6, 7, 8 y 11",  
  13: "Causal 13, 14 y 20",
  14: "Causal 13, 14 y 20",
  20: "Causal 13, 14 y 20"
};

// --- LISTA DE BLOQUEOS ESPECÍFICOS PARA PRE-EGRESO ---
const SET_BLOQUEOS_PREVISION = new Set([
  "BLOQUEADO POR ISAPRE",
  "BLOQUEADO POR CARGA ISAPRE",
  "BLOQUEADO CAPREDENA",
  "BLOQUEADO CARGA CAPREDENA",
  "BLOQUEADO DIPRECA",
  "BLOQUEADO CARGA DIPRECA",
  "BLOQUEADO POR AFILIADO FALLECIDO"
]);

// --- LISTA DE EDITORES GLOBALES PARA PROTECCIONES ---
const ADMIN_EDITORS = ['cledermann@hcuch.cl', 'mlara@hcuch.cl', 'mherbage@hcuch.cl', 'fzunigac@hcuch.cl', 'mariste@hcuch.cl', 'earenas@hcuch.cl', 'rvera@hcuch.cl', 'rvergara@hcuch.cl'];

// CONFIGURACIÓN DE LAS PLANILLAS DE DESTINO POR ESPECIALIDAD
const CONFIG_ESPECIALIDADES = {
  'CARDIOLOGÍA': { id: '1G4V1HSQMXvyjDMPY1xuloKhFnrfNuBf32nhT1v5MF-Y', nombreHoja: 'Casos y Cupos' },
  'CIRUGÍA CARDIOVASCULAR': { id: '1NFY6C7o2aje225gStdtc5mUGCoUTM2TMPzb3e4Xi97s', nombreHoja: 'Casos y Cupos'},
  'CIRUGÍA DE CABEZA, CUELLO Y MAXILOFACIAL': { id: '1axgvdJvohrUHTHSXWXV5lp2mUDTPmundcBCBo1B2z68', nombreHoja: 'Casos y Cupos'},
  'CIRUGÍA DE TÓRAX': { id: '12rS7GITDtvYaq2Qcd0o-bxp6tXixKjrLSvGkbOK646w', nombreHoja: 'Casos y Cupos' },
  'CIRUGÍA DIGESTIVA': { id: '1y-dRX3sssRhwiln1_Ao4QWtmgGSXseNct8k8LIOs12c', nombreHoja: 'Casos y Cupos' },
  'CIRUGÍA GENERAL': { id: '1k4pMPAwR9MB1PJthi428XY1oj1AWYJKng_AXbfkEh4A', nombreHoja: 'Casos y Cupos' },
  'CIRUGÍA PLÁSTICA Y REPARADORA': { id: '1-LX7YJXd8scjzTP__7oFA0mz6HTPNZgOaSJEMhn4Ww8', nombreHoja: 'Casos y Cupos' },
  'CIRUGÍA VASCULAR PERIFÉRICA': { id: '1KEQXFRxLh50P0Alc6SUJXjyWTDCu2NNZ2_HsoTILtgk', nombreHoja: 'Casos y Cupos' },
  'COLOPROCTOLOGÍA': { id: '1ZhCCFE60_LCHdKlg1_DKgW3aAKTExytJz0-oTS0Fl78', nombreHoja: 'Casos y Cupos' },
  'DERMATOLOGÍA': { id: '1yNjPfpFkArRSloMKo-ETTlbIn9Qcu41Ug1vDQ_o50Sw', nombreHoja: 'Casos y Cupos' },
  'DIABETOLOGÍA': { id: '1GjW0h-RgZypo7BEgApMi-bU-3fXGy9kPuSjy7ly7SWY', nombreHoja: 'Casos y Cupos' },
  'ENDOCRINOLOGÍA ADULTO': { id: '17-BrYdYDgJfWyY1iNazG9zPzfoJ-DVVU7zl9dQel0zI', nombreHoja: 'Casos y Cupos' },
  'ENFERMEDADES RESPIRATORIAS DEL ADULTO (BRONCOPULMONAR)': { id: '1dMDqf3oLVxe7fmR9iU-QD20tuCAnTRddPPnwYYNKsvU', nombreHoja: 'Casos y Cupos' },
  'GASTROENTEROLOGÍA ADULTO': { id: '1FU3Pm4TLCYWUqH3-GGsd0gQO1dB2NyFrPjesjckS-bk', nombreHoja: 'Casos y Cupos' },
  'GENÉTICA CLÍNICA': { id: '1JnePyvrVJdNAWJsRh9N3w_ar4L2ipwiLkpIweElQGQU', nombreHoja: 'Casos y Cupos' },
  'GERIATRÍA': { id: '1AY5-bGp9omSwvbg8GdsxYtfDHexL7yqvKOkPKK1KoFI', nombreHoja: 'Casos y Cupos' },
  'HEMATOLOGÍA': { id: '1V5m6D0N0Zm9ENHUYnYVzuRzuV6x8vNa-KFnPi32Hxgg', nombreHoja: 'Casos y Cupos' },
  'INFECTOLOGÍA': { id: '1F_mgOo0fe3mkc5fs01bqFCCfNzZ79xuM9aRvlB4EB88', nombreHoja: 'Casos y Cupos' },
  'INMUNOLOGÍA': { id: '1BfPh2JiGxvM9KVR5-2OXEFM4tb5sycp-vUyVDIvgaq4', nombreHoja: 'Casos y Cupos' },
  'MASTOLOGIA': { id: '1cvfpFD8vl4KrOKNKwnq_3hOb3hX5q83jd5n7SmhjZoM', nombreHoja: 'Casos y Cupos' },
  'MEDICINA FÍSICA Y REHABILITACIÓN': { id: '1OBtKqIInnszu-1BoaP6VCXjHfuq3dEaCb4Z4ggMFbx8', nombreHoja: 'Casos y Cupos' },
  'MEDICINA INTERNA': { id: '17QCAGobajh2UkAHvlzY90E3hfZt2_e3vMu5A9GqtCg8', nombreHoja: 'Casos y Cupos' },
  'NEFROLOGÍA ADULTO': { id: '1MehoCQfr-dUNCCg2SNn3UIxZew4-AAKdrSwy1NKu40M', nombreHoja: 'Casos y Cupos' },
  'NEUROCIRUGÍA': { id: '15UZ_tpIBIIErORmYEKExaQSCqRC47CWJXxvcyQ4UBOE', nombreHoja: 'Casos y Cupos' },
  'NEUROLOGÍA ADULTO': { id: '1TDchlfAeokWMAD3c-FZo4P9i5D3S-whq02RGwQDLqds', nombreHoja: 'Casos y Cupos' },
  'NUTRICION CLINICA': { id: '1GlbiKK8nmiIuBXXYgPGei5xlVMNn0q9pmlw0r-t2Qdo', nombreHoja: 'Casos y Cupos' },
  'OBSTETRICIA Y GINECOLOGÍA': { id: '1OO1d66VNUPlsm5GZy1vNr5x32UX9zmAN_WP_AY6DXCk', nombreHoja: 'Casos y Cupos' },
  'OFTALMOLOGÍA': { id: '1qdugGGICwoDFRAwyp8HcW-JQPtgFIYjaXP7Il4fG2ew', nombreHoja: 'Casos y Cupos' },
  'OTORRINOLARINGOLOGÍA': { id: '1XttoH8KxeH5MRvWIHF33bvmbF--t_Vpwo8ZEfarXUX0', nombreHoja: 'Casos y Cupos' },
  'REUMATOLOGÍA': { id: '1tynb5MvuYYIvxxBuByDf18z0nMsJyYQq52e0WNJa3CQ', nombreHoja: 'Casos y Cupos' },
  'TRAUMATOLOGÍA Y ORTOPEDIA': { id: '1BVmCVZEgLlTS0CXkO-BEEo0H97Eqyz_rjj0FWo_eBMw', nombreHoja: 'Casos y Cupos' },
  'UROLOGIA': { id: '1ZVwQF2Plkl78ysAF4K-Srhjjqrovp0H_yRnuenvpaXc', nombreHoja: 'Casos y Cupos' }
};

// MAPA PARA HOMOLOGACIÓN AUTOMÁTICA
const MAPA_HOMOLOGACION = {
  'CARDIOLOGÍA': ['CARDIOLOGIA'],
  'CIRUGÍA CARDIOVASCULAR': ['CIRUGIA CARDIOVASCULAR'],
  'CIRUGÍA DE CABEZA, CUELLO Y MAXILOFACIAL': ['CABEZA, CUELLO', 'MAXILOFACIAL'],
  'CIRUGÍA DE TÓRAX': ['CIRUGIA DE TORAX'],
  'CIRUGÍA DIGESTIVA': ['CIRUGIA DIGESTIVA'],
  'CIRUGÍA GENERAL': ['CIRUGIA GENERAL'],
  'CIRUGÍA PLÁSTICA Y REPARADORA': ['CIRUGIA PLASTICA', 'PLASTICA', 'REPARADORA'],
  'CIRUGÍA VASCULAR PERIFÉRICA': ['VASCULAR', 'CIRUGIA VASCULAR'],
  'COLOPROCTOLOGÍA': ['COLOPROCTOLOGIA'],
  'DERMATOLOGÍA': ['DERMATOLOGIA'],
  'DIABETOLOGÍA': ['DIABETOLOGIA'],
  'ENDOCRINOLOGÍA ADULTO': ['ENDOCRINOLOGIA', 'ENDOCRINO'],
  'ENFERMEDADES RESPIRATORIAS DEL ADULTO (BRONCOPULMONAR)': ['BRONCOPULMONAR', 'RESPIRATORIO', 'NEUMOLOGIA'],
  'GASTROENTEROLOGÍA ADULTO': ['GASTROENTEROLOGIA'],
  'GENÉTICA CLÍNICA': ['GENETICA CLINICA', 'GENETICA'],
  'GERIATRÍA': ['GERIATRIA'],
  'HEMATOLOGÍA': ['HEMATOLOGIA'],
  'INFECTOLOGÍA': ['INFECTOLOGIA'],
  'INMUNOLOGÍA': ['INMUNOLOGIA'],
  'MASTOLOGIA': ['MASTOLOGIA', 'PATOLOGIA MAMARIA'],
  'MEDICINA FÍSICA Y REHABILITACIÓN': ['FISIATRIA', 'REHABILITACION', 'MEDICINA FISICA'],
  'MEDICINA INTERNA': ['MEDICINA INTERNA'],
  'NEFROLOGÍA ADULTO': ['NEFROLOGIA'],
  'NEUROCIRUGÍA': ['NEUROCIRUGIA', 'NEUROCX'],
  'NEUROLOGÍA ADULTO': ['NEUROLOGIA'],
  'NUTRICION CLINICA': ['NUTRICION'],
  'OBSTETRICIA Y GINECOLOGÍA': ['GINECOLOGIA', 'OBSTETRICIA'],
  'OFTALMOLOGÍA': ['OFTALMOLOGIA'],
  'OTORRINOLARINGOLOGÍA': ['OTORRINO', 'OTORRINOLARINGOLOGIA'],
  'REUMATOLOGÍA': ['REUMATOLOGIA'],
  'TRAUMATOLOGÍA Y ORTOPEDIA': ['TRAUMATOLOGIA', 'ORTOPEDIA'],
  'UROLOGIA': ['UROLOGIA']
};