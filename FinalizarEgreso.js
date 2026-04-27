// FinalizarEgreso.gs

/**
 * Función central que maneja las ediciones en la planilla PreEgresos. Es llamada por el script conector.
 */
function handlePreEgresoEdit(eventInfo) {
  // Obtenemos los datos del paquete recibido
  const sourceId = eventInfo.sourceId;
  const rangeA1Notation = eventInfo.rangeA1Notation;
  const value = eventInfo.value;
  const sheetName = eventInfo.sheetName; // Usamos el nombre que nos envió el conector

  // Abrimos la hoja usando la información recibida
  const source = SpreadsheetApp.openById(sourceId);

  // 1. Obtenemos la HOJA por su NOMBRE (ej: "Causal 5 y 9")
  const sheet = source.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Error: No se pudo encontrar la hoja "${sheetName}"`);
    return;
  }
  // 2. Obtenemos el RANGO desde ESA hoja (ej: "Y24")
  const range = sheet.getRange(rangeA1Notation);
  const row = range.getRow();
  const col = range.getColumn();

  // Ignorar ediciones en la cabecera
  if (row === 1) return;

  let COLUMNAS_ACTUAL;
  if (sheetName === NOMBRES_HOJAS_PREEGRESO.causal5y9) {
    COLUMNAS_ACTUAL = COLUMNAS_PREEGRESO;
  } else if (sheetName === NOMBRES_HOJAS_PREEGRESO.causal4_6_7_8_11 || sheetName === NOMBRES_HOJAS_PREEGRESO.causal13_14_20) {
    COLUMNAS_ACTUAL = COLUMNAS_PREEGRESO_OTRAS;
  } else {
    return; // Hoja no reconocida
  }

  const colEstadoRevision = sheetName === NOMBRES_HOJAS_PREEGRESO.causal5y9 ? COLUMNAS_ACTUAL.ESTADO_PRIMERA_REVISION : COLUMNAS_ACTUAL.ESTADO_REVISION_CASO;
  const colFechaRevision = sheetName === NOMBRES_HOJAS_PREEGRESO.causal5y9 ? COLUMNAS_ACTUAL.FECHA_PRIMERA_REVISION : COLUMNAS_ACTUAL.FECHA_REVISION;

  if (col === colEstadoRevision) {
    // Usamos la variable 'sheet' que ya obtuvimos correctamente
    sheet.getRange(row, colFechaRevision).setValue(new Date());
    return;
  }

  // (Mejora: Aceptamos "TRUE" (texto) o true (booleano))
  if (col === COLUMNAS_ACTUAL.CHECK_FINALIZAR && (value === "TRUE" || value === true)) {
      // Pasamos la HOJA CORRECTA ('sheet') a la función auxiliar
      procesarFinalizacionEgreso(sheet, row, COLUMNAS_ACTUAL);
      return;
  }
}

/**
 * Función auxiliar que contiene la lógica para actualizar CambioEstado y Seguimiento. Es llamada por handlePreEgresoEdit cuando se marca el check final.
 */
function procesarFinalizacionEgreso(hojaOrigen, filaEditada, COLUMNAS_ACTUAL) {
  const fechaActual = new Date();
  const valoresFila = hojaOrigen.getRange(filaEditada, 1, 1, hojaOrigen.getLastColumn()).getValues()[0];
  
  const idCaso = valoresFila[COLUMNAS_ACTUAL.ID_CASO - 1];
  const nuevoEstadoCod = valoresFila[COLUMNAS_ACTUAL.ESTADO_NUEVO_COD - 1];
  const checkCell = hojaOrigen.getRange(filaEditada, COLUMNAS_ACTUAL.CHECK_FINALIZAR);

  // ACTUALIZACIÓN 23/03 - Verificación: ¿El estado es realmente 11, 12 o 13?
  if (![11, 12, 13].includes(nuevoEstadoCod)) {
    SpreadsheetApp.getUi().alert("Error: Para procesar el egreso, el 'Estado Nuevo' (columna AC) debe ser 11 (Egresado), 12 (Egresado otro establecimiento) o 13 (Menor de edad).");
    checkCell.setValue(false); // Desmarca si el estado no es correcto
    return;
  }

  // 1. CREAR REGISTRO EN CAMBIOESTADO
  const hojaCE = SpreadsheetApp.openById(IDs.cambioEstado).getSheetByName("2025");
  if (hojaCE) {
    const filaDestinoCE = getPrimeraFilaVacia(hojaCE);
    const valorCorrelativo = getProximoCorrelativoCE(hojaCE);
    
    // Leemos la causal real de la columna Z (26)
    const causalReal = hojaOrigen.getRange(filaEditada, 26).getValue(); 
    
    const nuevaFilaCE = [
      valorCorrelativo, idCaso, nuevoEstadoCod, fechaActual,
      causalReal, // Causal Real (Z)
      valoresFila[COLUMNAS_ACTUAL.FECHA_ENVIO_SIDRA - 1], // Fecha SIDRA (T)
      valoresFila[COLUMNAS_ACTUAL.ESTAB_EGRESA - 1],     // Estab Egresa (AA)
      valoresFila[COLUMNAS_ACTUAL.ESTAB_DESTINO - 1],    // Estab Destino (AB)
      valoresFila[COLUMNAS_ACTUAL.COMENTARIO - 1]       // Comentario (X)
    ];
    hojaCE.getRange(filaDestinoCE, 1, 1, nuevaFilaCE.length).setValues([nuevaFilaCE]);
  }

  // 2. ACTUALIZACIÓN DE ESTADO FINAL EN SEGUIMIENTO
  const hojaSeguimiento = SpreadsheetApp.openById(IDs.seguimiento).getSheetByName("Casos y Cupos");   
  if (hojaSeguimiento) {
    const filaEnSeguimiento = findRowPorPatron(hojaSeguimiento, COLUMNAS_CUPOS.ID_CASO, idCaso);     
    if (filaEnSeguimiento !== "0") {
      hojaSeguimiento.getRange(filaEnSeguimiento, COLUMNAS_CUPOS.FECHA_EGRESADO).setValue(fechaActual);
      hojaSeguimiento.getRange(filaEnSeguimiento, COLUMNAS_CUPOS.ESTADO_COD).setValue(nuevoEstadoCod);
      hojaSeguimiento.getRange(filaEnSeguimiento, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue(MAPAS.estados[nuevoEstadoCod]);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Caso ${idCaso} egresado exitosamente.`);
    } else {
      Logger.log(`ERROR: El caso Id ${idCaso} no fue encontrado en Seguimiento.`);
      SpreadsheetApp.getUi().alert(`El caso ${idCaso} no fue encontrado en Seguimiento. La actualización no se pudo completar.`);
      checkCell.setValue(false); // Desmarca si no se encontró en seguimiento
    }
  } else {
      checkCell.setValue(false); // Desmarca si no se encontró la hoja de seguimiento
  }
}
