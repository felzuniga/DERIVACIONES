// CrearAgente.gs

/**
 * FUNCIÓN PARA EL MENÚ (MANUAL):
 * Muestra una confirmación al usuario y luego presenta un resumen.
 */
function crearAgentes_manual() {
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert('Confirmar Creación', 'Este proceso creará una planilla para cada agente nuevo en la lista "Agentes". ¿Deseas continuar?', ui.ButtonSet.OK_CANCEL);
    if (respuesta !== ui.Button.OK) return;

    SpreadsheetApp.getActiveSpreadsheet().toast('Creando planillas de agentes...', 'Procesando', -1);
    
    // Pasa el ID de la planilla activa a la función de trabajo.
    const idPlanillaSupervisora = SpreadsheetApp.getActiveSpreadsheet().getId();
    const resultado = crearAgentes_core(idPlanillaSupervisora);

    SpreadsheetApp.getActiveSpreadsheet().toast('');
    ui.alert('Proceso Completado', `Se crearon ${resultado.creados} nuevas planillas de agente.\nSe encontraron ${resultado.existentes} agentes que ya tenían planilla.`, ui.ButtonSet.OK);
}

/**
 * FUNCIÓN DE TRABAJO (CORE):
 * Realiza la creación de las planillas. No tiene interfaz de usuario.
 */
function crearAgentes_core(idPlanillaSupervisora) {
    const hojaAgentes = SpreadsheetApp.openById(idPlanillaSupervisora).getSheetByName('Agentes');
    const ultimaFila = hojaAgentes.getLastRow();

    if (ultimaFila < 2) {
        Logger.log("No hay agentes en la lista para procesar.");
        return { creados: 0, existentes: 0 };
    }

    // Leemos hasta la última columna que necesitamos para tener todos los datos.
    const rangoCompleto = hojaAgentes.getRange(2, 1, ultimaFila - 1, COLUMNAS_AGENTES.FECHA_CREACION);
    const datos = rangoCompleto.getValues();
  
    const agentesNuevos = [];
    let agentesExistentes = 0;
  
    datos.forEach((fila, index) => {
        const nombre = fila[COLUMNAS_AGENTES.NOMBRE - 1];
        const estado = fila[COLUMNAS_AGENTES.ESTADO_CREACION - 1];
    
        if (nombre && nombre.trim() !== '') {
            if (estado !== 'Creado') {
                agentesNuevos.push({ nombre: nombre.trim(), fila: index + 2 }); // index + 2 es el número de fila real
            } else {
                agentesExistentes++;
            }
        }
    });
  
    if (agentesNuevos.length === 0) {
        Logger.log('No hay agentes nuevos por crear.');
        return { creados: 0, existentes: agentesExistentes };
    }
  
    Logger.log(`Se crearán archivos para ${agentesNuevos.length} agentes nuevos.`);
  
    const carpetaDestino = DriveApp.getFolderById(IDs.carpetaAgentes);
  
    agentesNuevos.forEach(agente => {
        try {
            const archivoPlantilla = DriveApp.getFileById(IDs.plantillaAgente);
            const copia = archivoPlantilla.makeCopy(agente.nombre, carpetaDestino);
    
            hojaAgentes.getRange(agente.fila, COLUMNAS_AGENTES.ESTADO_CREACION).setValue('Creado');
            hojaAgentes.getRange(agente.fila, COLUMNAS_AGENTES.URL_PLANILLA).setValue(copia.getUrl());
            hojaAgentes.getRange(agente.fila, COLUMNAS_AGENTES.FECHA_CREACION).setValue(new Date());
      
            Logger.log(`Creado archivo para: ${agente.nombre}`);
      
        } catch (error) {
            Logger.log(`Error creando planilla para ${agente.nombre}: ${error.toString()}`);
            hojaAgentes.getRange(agente.fila, COLUMNAS_AGENTES.ESTADO_CREACION).setValue('Error');
        }
    });
  
    SpreadsheetApp.flush();
    Logger.log(`Proceso completado. ${agentesNuevos.length} archivos creados.`);
    return { creados: agentesNuevos.length, existentes: agentesExistentes };
}