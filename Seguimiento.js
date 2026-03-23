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