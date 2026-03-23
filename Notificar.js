// Notificar.gs

// --- FUNCIÓN PRINCIPAL ---
/**
 * Se activa al editar la hoja y actúa como un "enrutador", llamando a la función específica según la columna que fue editada.
 */
function NotificarPorCheckbox(e) {
  const celdaEditada = e.range;
  const filaEditada = celdaEditada.getRow();
  const columnaEditada = celdaEditada.getColumn();

  // Antes la lógica de validación estaba dentro de la función.
  // Ahora hacemos una validación rápida al inicio para salir inmediatamente si no es necesario procesar.
  if (filaEditada === 1) return; // Ignorar la fila de encabezados

  const hojaOrigen = celdaEditada.getSheet();
  const fechaActual = new Date();

  // --- ENRUTADOR DE ACCIONES ---
  // Identificamos la acción basada en la columna editada.
  switch (columnaEditada) {
    
    case COLUMNAS_CUPOS.FECHA_DEF: // Columna R
      if (celdaEditada.getValue() !== "") {
        NotificarEgreso(hojaOrigen, celdaEditada, fechaActual, 9); // Causal 9 = Fallecido
      }
      break;

    case COLUMNAS_CUPOS.ESTADO_PREVISIONAL: // Columna W
      const contenidoCelda = celdaEditada.getValue();
      const contenidoNormalizado = contenidoCelda ? contenidoCelda.toString().trim().toUpperCase() : "";

      // Verificamos si el valor EXACTO está en la lista global
      if (SET_BLOQUEOS_PREVISION.has(contenidoNormalizado)) {
        NotificarEgreso(hojaOrigen, celdaEditada, fechaActual, 5); // Causal 5 = No Beneficiario
      }
      break;
  }
}

/**
 * Procesa casos para Agendar/Notificar, los envía a Supervisora y los mueve a Seguimiento. (Añadido borrado en Asignación de Cupos)
 */
function NotificarCentroContacto(hojaOrigen, celdaEditada, fechaActual) {
  const filaEditada = celdaEditada.getRow();

  // --- CORRECCIÓN DE SEGURIDAD (Lectura y Expansión) ---
  // Leemos hasta la última columna con datos
  const ultimaColumna = hojaOrigen.getLastColumn();
  const rango = hojaOrigen.getRange(filaEditada, 1, 1, ultimaColumna);
  const datosLeidos = rango.getValues()[0];
  
  // Creamos un array seguro de 62 posiciones y volcamos los datos
  let valoresFila = new Array(62).fill("");
  for(let i=0; i<datosLeidos.length; i++) {
      valoresFila[i] = datosLeidos[i];
  }
  // ------------------------------------------------------

  const idCaso = valoresFila[COLUMNAS_CUPOS.ID_CASO - 1]; // COLUMNA A
  const tipoAccion = valoresFila[COLUMNAS_CUPOS.TIPO_ACCION - 1]; // Columna AJ

  // Determinar nuevo estado y columna de fecha
  const esNotificar = (tipoAccion === "Notificar");
  const nuevoEstadoCod = esNotificar ? 4 : 3; // Estado 4 (Notificar) o 3 (Agendar)
  const columnaFecha = esNotificar ? COLUMNAS_CUPOS.FECHA_EN_CAMP_NOTIFICAR : COLUMNAS_CUPOS.FECHA_EN_CAMP_AGENDAR; // AZ o AW

  // 1. Enviar a Supervisora
  const planillaSupervisora = SpreadsheetApp.openById(IDs.supervisora);
  const hojaSupervisora = planillaSupervisora.getSheetByName("Casos");
  const filaParaSupervisora = [
    valoresFila[COLUMNAS_CUPOS.ID_CASO - 1],
    valoresFila[COLUMNAS_CUPOS.RUT - 1],
    valoresFila[COLUMNAS_CUPOS.DV - 1],
    valoresFila[COLUMNAS_CUPOS.NOMBRES - 1],
    valoresFila[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1],
    valoresFila[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1],
    tipoAccion,
    valoresFila[COLUMNAS_CUPOS.ORIGEN - 1],
    valoresFila[COLUMNAS_CUPOS.FECHA_ENTRADA - 1],
    valoresFila[COLUMNAS_CUPOS.ESPECIALIDAD_CORREGIDA - 1] || valoresFila[COLUMNAS_CUPOS.PRESTA_EST - 1], // Prestación (Prioridad: Corregida > Original)
    "Oficina de Red",
    valoresFila[COLUMNAS_CUPOS.COMENTARIO - 1], 
    "",
    valoresFila[COLUMNAS_CUPOS.AGENDA - 1],
    valoresFila[COLUMNAS_CUPOS.FECHA_AGENDA - 1],
    valoresFila[COLUMNAS_CUPOS.HORA_AGENDA - 1],
    "Sin Gestión", // Estado inicial
    fechaActual,   // Fecha de recepción
    "", "", "", "",
    valoresFila[COLUMNAS_CUPOS.EMAIL - 1],
    "",
    valoresFila[COLUMNAS_CUPOS.TELEFONO_FIJO - 1],
    valoresFila[COLUMNAS_CUPOS.TELEFONO_MOVIL - 1],
    "", "", "", "", "", "",
    valoresFila[COLUMNAS_CUPOS.FUENTE - 1]
  ];

  const filaInicioSupervisora = getPrimeraFilaVacia(hojaSupervisora, 1, 2);
  hojaSupervisora.getRange(filaInicioSupervisora, 1, 1, filaParaSupervisora.length).setValues([filaParaSupervisora]);

  // 2. Crear registro en CambioEstado
  const planillaCE = SpreadsheetApp.openById(IDs.cambioEstado);
  const hojaCE = planillaCE.getSheetByName("2025");
  let proximoCorrelativo = getProximoCorrelativoCE(hojaCE);

  // --- Capturar Fecha/Hora para Estados 3 y 4 ---
  let fechaCita = "";
  let horaCita = "";

  // Si es "En campaña Notificar" (Estado 4), capturamos la fecha/hora que el usuario ingresó en las columnas AK y AL
  if (nuevoEstadoCod === 4) {
     fechaCita = valoresFila[COLUMNAS_CUPOS.FECHA_AGENDA - 1]; // Col AK
     horaCita = valoresFila[COLUMNAS_CUPOS.HORA_AGENDA - 1]; // Col AL
  }
  // Para Estado 3 (Agendar), se quedan vacíos, lo cual es correcto.


  const filaParaCE = [
    proximoCorrelativo, 
    idCaso, 
    nuevoEstadoCod, 
    fechaActual, 
    "", 
    "", 
    "", 
    "", 
    "", 
    valoresFila[COLUMNAS_CUPOS.AGENDA - 1],
    fechaCita || "",    // K (Fecha Cita)
    horaCita || ""      // L (Hora Cita)
    ];
  hojaCE.appendRow(filaParaCE);

  // 3. ACTUALIZACIÓN Y MOVIMIENTO A SEGUIMIENTO
  // Actualizamos los datos en memoria ANTES de moverlos
  valoresFila[COLUMNAS_CUPOS.ESTADO_COD - 1] = nuevoEstadoCod; // BF
  valoresFila[COLUMNAS_CUPOS.ESTADO_TEXTO - 1] = MAPAS.estados[nuevoEstadoCod]; // BG
  valoresFila[columnaFecha - 1] = fechaActual; // AZ o AW

  // Se llama a la versión SIN BORRAR para que el script pueda continuar y eliminar de "Asignación de Cupos"
  // La fila de esta hoja (Especialidad) se borra al final de esta función.
  moverASeguimientoSinBorrar(valoresFila);

  // 4. NO ELIMINAMOS NADA AQUÍ
  // La función procesarEnvioCC_lote() se encarga de las eliminaciones

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Caso ${idCaso} procesado. Pendiente de eliminación en lote.`,
    'Procesado',
    3
  );
}

/**
 * Procesa casos para Pre-Egreso, los envía a PreEgresos y los mueve a Seguimiento. (Añadido borrado en Asignación de Cupos)
 */
function NotificarEgreso(hojaOrigen, celdaEditada, fechaActual, causalCod) {
  const filaEditada = celdaEditada.getRow();

  // --- CORRECCIÓN DE SEGURIDAD (Lectura y Expansión) ---
  const ultimaColumna = hojaOrigen.getLastColumn();
  const rango = hojaOrigen.getRange(filaEditada, 1, 1, ultimaColumna);
  const datosLeidos = rango.getValues()[0];
  let valoresFila = new Array(62).fill("");
  for(let i=0; i<datosLeidos.length; i++) valoresFila[i] = datosLeidos[i];
  // ------------------------------------------------------

  const idCaso = valoresFila[COLUMNAS_CUPOS.ID_CASO - 1]; // Columna A
  const nuevoEstadoCod = 9; // Pre Egresado

  // 1. Enviar a PreEgresos
  const planillaPreEgresos = SpreadsheetApp.openById(IDs.preEgresos);
  let hojaDestinoPreEgreso = null;
  let filaParaPreEgreso = [];

  // Verificamos que sea causal 5 o 9 y obtenemos la hoja correspondiente
  if (causalCod === 5 || causalCod === 9) {
    hojaDestinoPreEgreso = planillaPreEgresos.getSheetByName(NOMBRES_HOJAS_PREEGRESO.causal5y9);
    if (hojaDestinoPreEgreso) {
      // Construimos la fila con 16 columnas, añadiendo "Pendiente" en la P
      filaParaPreEgreso = [
        idCaso,                                               // A
        fechaActual,                                          // B
        valoresFila[COLUMNAS_CUPOS.SERVICIO_SALUD - 1],       // C
        valoresFila[COLUMNAS_CUPOS.RUT - 1],                  // D
        valoresFila[COLUMNAS_CUPOS.DV - 1],                   // E
        valoresFila[COLUMNAS_CUPOS.NOMBRES - 1],              // F
        valoresFila[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1],      // G
        valoresFila[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1],     // H
        valoresFila[COLUMNAS_CUPOS.FECHA_NAC - 1],            // I
        valoresFila[COLUMNAS_CUPOS.FECHA_DEF - 1],            // J
        valoresFila[COLUMNAS_CUPOS.ESPECIALIDAD - 1],         // K
        valoresFila[COLUMNAS_CUPOS.SOSPECHA_DIAGNOSTICA - 1], // L
        causalCod,                                            // M
        null,                                                   // N - Glosa (se llena con fórmula)
        null,                                                   // O - Responsable (vacío inicialmente)
        "Pendiente"                                           // P - Estado Primera Revisión
      ];
    }
  } else {
    // Si por alguna razón llega otra causal aquí (no debería según el flujo), lo registramos.
    Logger.log(`Intento de enviar causal ${causalCod} a PreEgresos desde NotificarEgreso. Flujo no esperado.`);
  }

  // Escribimos en la hoja de destino si todo está correcto
  if (hojaDestinoPreEgreso && filaParaPreEgreso.length > 0) {
    const filaInicioPreEgresos = getPrimeraFilaVacia(hojaDestinoPreEgreso);
    const numFilasNuevas = 1; // Solo se procesa una fila a la vez

    // Ajustamos el ancho al número correcto de columnas (16)
    hojaDestinoPreEgreso.getRange(filaInicioPreEgresos, 1, 1, filaParaPreEgreso.length).setValues([filaParaPreEgreso]);

    // --- APLICAR FÓRMULAS EXPLÍCITAS ---
    // (Añadido por si acaso el de PreEgresos también falla)
    try {
      const formulaGlosaR1C1 = '=IF(RC[-1]=9;"FALLECIMIENTO";IF(RC[-1]=5;"NO BENEFICIARIO";""))';
      hojaDestinoPreEgreso.getRange(filaInicioPreEgreso, 14, numFilasNuevas).setFormulaR1C1(formulaGlosaR1C1);
      const formulaRespR1C1 = '=IF(OR(RC[-2]=9;RC[-2]=5);"Marcos Batarce";"")';
      hojaDestinoPreEgreso.getRange(filaInicioPreEgreso, 15, numFilasNuevas).setFormulaR1C1(formulaRespR1C1);
    } catch (e) {
      Logger.log(`Error al aplicar fórmulas explícitas en NotificarEgreso: ${e.message}`);
    }

  } else {
    Logger.log(`No se pudo escribir el caso ${idCaso} en PreEgresos (Causal ${causalCod}). Hoja destino no encontrada o fila no construida.`);
  }

  // 2. Crear registro en CambioEstado
  const planillaCE = SpreadsheetApp.openById(IDs.cambioEstado);
  const hojaCE = planillaCE.getSheetByName("2025");
  let proximoCorrelativo = getProximoCorrelativoCE(hojaCE);
  const filaParaCE = [proximoCorrelativo, idCaso, nuevoEstadoCod, fechaActual, causalCod, "", "", "", "", valoresFila[COLUMNAS_CUPOS.AGENDA - 1]];
  hojaCE.appendRow(filaParaCE);

  // 3. Actualización y movimiento a Seguimiento
  //const fechaFormateada = Utilities.formatDate(fechaActual, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy');

  // // Actualizamos los datos en memoria ANTES de moverlos
  valoresFila[COLUMNAS_CUPOS.ESTADO_COD - 1] = nuevoEstadoCod; // BF
  valoresFila[COLUMNAS_CUPOS.ESTADO_TEXTO - 1] = MAPAS.estados[nuevoEstadoCod]; // BG
  valoresFila[COLUMNAS_CUPOS.FECHA_PRE_EGRESADO - 1] = fechaActual; // Fecha Pre-egresado (BC)

  // Se llama a la versión SIN BORRAR para que el script pueda continuar y eliminar de "Asignación de Cupos"
  // La fila de esta hoja (Especialidad) se borra al final de esta función.
  moverASeguimientoSinBorrar(valoresFila);

  // 4. Eliminar de Asignación de Cupos (búsqueda mejorada)
  try {
    const ssAsignacion = SpreadsheetApp.openById(IDs.asignacionCupos);
    const hojaAsignacion = ssAsignacion.getSheetByName(NOMBRES_HOJAS.asignacionCupos);
    
    // Usar búsqueda optimizada
    const mapaAsignacion = construirMapaAsignacion(hojaAsignacion);
    const filaAEliminar = mapaAsignacion[idCaso];

    if (filaAEliminar && filaAEliminar > 1) {
      hojaAsignacion.deleteRow(filaAEliminar);
      Logger.log(`Caso ${idCaso} eliminado de Asignación de Cupos (fila ${filaAEliminar}).`);
    } else {
      Logger.log(`Caso ${idCaso} no encontrado en Asignación de Cupos.`);
    }
  } catch (err) {
    Logger.log(`ERROR al eliminar de Asignación: ${err.message}`);
  }

  // 5. Eliminar fila de Especialidad (MANTENER PARA EGRESOS)
  hojaOrigen.deleteRow(filaEditada);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Caso ${idCaso} enviado a Pre-Egresos y movido a seguimiento.`
  );
}