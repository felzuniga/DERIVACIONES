// LimpiezaAsignacion_Distribuido.gs

/**
 * LIMPIEZA QUIRÚRGICA DE ASIGNACIÓN DE CUPOS
 * Criterios de Borrado (Deben cumplirse AMBOS):
 * 1. El ID del caso YA EXISTE en la planilla 'Seguimiento'.
 * 2. El estado en 'Asignación de Cupos' (Columna BI) es "Distribuido".
 */
function limpiarAsignacionSoloDistribuidos() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. CARGA DE BASE MAESTRA (SEGUIMIENTO)
  SpreadsheetApp.getActiveSpreadsheet().toast('Cargando base de Seguimiento...', 'Fase 1/3', -1);
  
  const idsProcesados = new Set();
  try {
    const ssSeg = SpreadsheetApp.openById(IDs.seguimiento);
    const hojaSeg = ssSeg.getSheetByName(NOMBRES_HOJAS.seguimiento);
    const ultFilaSeg = hojaSeg.getLastRow();
    
    if (ultFilaSeg > 1) {
      // Leemos columna A (ID)
      const datosIds = hojaSeg.getRange(2, 1, ultFilaSeg - 1, 1).getValues();
      datosIds.forEach(r => {
        if (r[0]) {
          // Normalización: sin comillas, sin espacios
          const idLimpio = String(r[0]).replace(/['"]/g, '').trim();
          idsProcesados.add(idLimpio);
        }
      });
    }
  } catch (e) {
    ui.alert(`Error leyendo Seguimiento: ${e.message}`);
    return;
  }

  if (idsProcesados.size === 0) {
    ui.alert("La planilla de Seguimiento parece vacía. No se puede comparar.");
    return;
  }

  // 2. ESCANEO DE ASIGNACIÓN DE CUPOS
  SpreadsheetApp.getActiveSpreadsheet().toast('Analizando Asignación de Cupos...', 'Fase 2/3', -1);
  
  let filasParaBorrar = []; 
  let ejemplosEncontrados = [];

  try {
    const ssAsig = SpreadsheetApp.openById(IDs.asignacionCupos);
    const hojaAsig = ssAsig.getSheetByName(NOMBRES_HOJAS.asignacionCupos);
    
    if (!hojaAsig || hojaAsig.getLastRow() < 2) {
      ui.alert("La planilla de Asignación de Cupos está vacía o no se encuentra.");
      return;
    }

    // Leemos hasta la columna BI (Columna 61) para verificar el estado "Distribuido"
    // BI es la columna 61. En array base-0, el índice es 60.
    const ultFilaAsig = hojaAsig.getLastRow();
    const datosAsig = hojaAsig.getRange(2, 1, ultFilaAsig - 1, 61).getValues();
    
    // Recorremos buscando coincidencias
    datosAsig.forEach((fila, i) => {
      const idRaw = String(fila[0]); // Columna A
      const idCheck = idRaw.replace(/['"]/g, '').trim();
      const estadoDistribucion = String(fila[60] || "").trim(); // Columna BI (Índice 60)
      
      // CRITERIO DOBLE DE SEGURIDAD
      const existeEnSeguimiento = idsProcesados.has(idCheck);
      const esDistribuido = (estadoDistribucion === "Distribuido");

      if (existeEnSeguimiento && esDistribuido) {
        // Guardamos el índice visual (Fila real de Excel = índice + 2)
        filasParaBorrar.push(i + 2);
        
        if (ejemplosEncontrados.length < 5) ejemplosEncontrados.push(idRaw);
      }
    });

    // 3. REPORTE Y CONFIRMACIÓN
    SpreadsheetApp.getActiveSpreadsheet().toast(''); 
    
    if (filasParaBorrar.length === 0) {
      ui.alert("✅ Todo limpio.\n\nNo se encontraron casos en 'Asignación de Cupos' que cumplan ambas condiciones (Estar en Seguimiento Y estar 'Distribuido').");
      return;
    }

    const mensajeReporte = `🔍 REPORTE DE HALLAZGOS\n\n` +
      `Se encontraron ${filasParaBorrar.length} casos candidatos a eliminación.\n\n` +
      `Criterios cumplidos:\n` +
      `1. Ya existen en Seguimiento.\n` +
      `2. Tienen estado "Distribuido" en columna BI.\n\n` +
      `Ejemplos: ${ejemplosEncontrados.join(', ')}...\n\n` +
      `⚠️ ¿Deseas ELIMINAR estas ${filasParaBorrar.length} filas de Asignación de Cupos?`;

    const respuesta = ui.alert('Confirmar Eliminación Segura', mensajeReporte, ui.ButtonSet.YES_NO);

    if (respuesta === ui.Button.YES) {
      // 4. EJECUCIÓN DEL BORRADO
      SpreadsheetApp.getActiveSpreadsheet().toast(`Eliminando ${filasParaBorrar.length} filas...`, 'Fase 3/3', -1);
      
      // Ordenamos de mayor a menor para borrar de abajo hacia arriba (CRÍTICO)
      filasParaBorrar.sort((a, b) => b - a);
      
      filasParaBorrar.forEach(fila => {
        hojaAsig.deleteRow(fila);
      });
      
      ui.alert(`✅ Limpieza Completada.\n\nSe eliminaron ${filasParaBorrar.length} casos correctamente.`);
    } else {
      ui.alert("Operación cancelada. No se borró nada.");
    }

  } catch (e) {
    ui.alert(`Ocurrió un error inesperado: ${e.message}`);
    Logger.log(e.stack);
  }
}