// Seguimiento.gs

function moverASeguimiento(hojaOrigen, numeroFila, valoresFila) {
  try {
    copiarASeguimientoCore(valoresFila);
    hojaOrigen.deleteRow(numeroFila);
    Logger.log(`✅ Fila ${numeroFila} movida a Seguimiento.`);
  } catch (error) {
    Logger.log(`❌ Error al mover fila ${numeroFila}: ${error.message}`);
    SpreadsheetApp.getUi().alert(`No se pudo mover la fila. Error: ${error.message}`);
  }
}

function moverASeguimientoSinBorrar(valoresFila) {
  try {
    copiarASeguimientoCore(valoresFila);
    Logger.log(`✅ Fila copiada a Seguimiento.`);
  } catch (error) {
    Logger.log(`❌ Error al copiar fila: ${error.message}`);
  }
}

/**
 * FUNCIÓN BLINDADA (CORE): 
 * 1. Normaliza el tamaño del array a 62 columnas.
 * 2. Formatea fechas explícitamente para evitar celdas vacías.
 * 3. Escribe en Seguimiento.
 */
function copiarASeguimientoCore(valoresFila) {
    const planillaDestino = SpreadsheetApp.openById(IDs.seguimiento);
    const hojaDestino = planillaDestino.getSheetByName("Casos y Cupos");
    
    if (!hojaDestino) throw new Error("No se encontró la hoja 'Casos y Cupos' en Seguimiento.");

    // 1. ARRAY MAESTRO DE 62 COLUMNAS
    let filaMaestra = new Array(62).fill(""); 
    
    // 2. VOLCADO SEGURO
    for (let i = 0; i < valoresFila.length; i++) {
        if (valoresFila[i] !== undefined && valoresFila[i] !== null) {
            filaMaestra[i] = valoresFila[i];
        }
    }

    // 3. FORMATEO DE FECHAS CRÍTICAS (Col 48-51 / AW-AZ)
    const colsFecha = [48, 49, 50, 51]; 
    const zonaHoraria = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    
    colsFecha.forEach(idx => {
        if (filaMaestra[idx] && filaMaestra[idx] instanceof Date) {
            filaMaestra[idx] = Utilities.formatDate(filaMaestra[idx], zonaHoraria, "dd/MM/yyyy");
        }
    });

    // 4. ESCRITURA
    const filaInicio = getPrimeraFilaVacia(hojaDestino, 1, 2); 
    hojaDestino.getRange(filaInicio, 1, 1, filaMaestra.length).setValues([filaMaestra]);
}

// --- ACTUALIZACIÓN 23/03 ---

/**
 * Detecta cambios manuales de estado en Seguimiento y los registra en CambioEstado.
 * IMPORTANTE: Requiere crear un Trigger Instalable del tipo "Al editar".
 */
function onEditSeguimientoEgresos(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  // Validar que estemos en la hoja correcta y editando la columna "Estado Cod" (58)
  if (sheet.getName() !== "Casos y Cupos") return; 
  if (e.range.getColumn() !== COLUMNAS_CUPOS.ESTADO_COD) return; 
  
  const row = e.range.getRow();
  if (row < 2) return;

  const nuevoEstadoCod = parseInt(e.value, 10);
  
  // Si eligen 11, 12 o 13 manualmente
  if ([11, 12, 13].includes(nuevoEstadoCod)) {
    const fechaActual = new Date(); // La fecha y hora exacta en que el sistema hace el registro
    const idCaso = sheet.getRange(row, COLUMNAS_CUPOS.ID_CASO).getValue();
    
    // CAPTURAR LA FECHA DIGITADA EN LA COLUMNA BE
    const fechaEgresoDigitada = sheet.getRange(row, COLUMNAS_CUPOS.FECHA_EGRESADO).getValue();
    
    // 1. Autocompletar texto del estado (NO tocamos la fecha digitada)
    sheet.getRange(row, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue(MAPAS.estados[nuevoEstadoCod] || "");
    
    // 2. Registrar el movimiento en la planilla CambioEstado
    try {
      const hojaCE = SpreadsheetApp.openById(IDs.cambioEstado).getSheetByName("2025"); 
      if (hojaCE) {
        const valorCorrelativo = getProximoCorrelativoCE(hojaCE);
        const nuevaFilaCE = [
          valorCorrelativo,                 // 1. Correlativo
          idCaso,                           // 2. ID Caso
          nuevoEstadoCod,                   // 3. Estado Nuevo
          fechaActual,                      // 4. Fecha de Cambio (Registro del sistema)
          "",                               // 5. Causal Real (Vacío)
          fechaEgresoDigitada,              // 6. Fecha Envío SIDRA (La fecha digitada en BE)
          "",                               // 7. Estab Egresa (Vacío)
          "",                               // 8. Estab Destino (Vacío)
          "Egreso Directo en Seguimiento"   // 9. Comentario
        ];
        hojaCE.appendRow(nuevaFilaCE);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Registro de Cambio de Estado generado para el caso ${idCaso}`, "Éxito", 3);
      }
    } catch(error) { 
      Logger.log("Error al escribir en CambioEstado: " + error.message); 
      SpreadsheetApp.getUi().alert("Error", "No se pudo registrar en CambioEstado: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}
