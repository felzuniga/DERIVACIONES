// Main.gs

/**
 * Añade el menú a la planilla central "Asignación de Cupos".
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Acciones de Carga')
    .addItem('📤 Distribuir por Especialidad', 'distribuirCasosPorEspecialidad')
    .addItem('⬆️ Subir Casos (⚠️ Exclusivo Depto. TI)', 'CargarTodosLosCasos')
    .addToUi();  
}

/**
 * Añade el menú en las planillas de "Especialidad".
 */
function onOpenEspecialidad() {
  SpreadsheetApp.getUi()
      .createMenu('🗓️ Agendamiento')
      //.addItem('Asignar Agenda Abierta', 'mostrarDialogoAsignarValor')
      .addItem('📤 Enviar a C.C.', 'procesarEnvioCC_lote')
      .addToUi();
}

/**
 * Añade el menú en la planilla "Supervisora".
 */
function onOpenSupervisora() {
  SpreadsheetApp.getUi()
      .createMenu('👨‍💼 Acciones de Supervisión')
      .addItem('📤 Asignar Casos Seleccionados', 'asignarAgentesEnLote')
      .addSeparator()
      .addItem('🔄 Actualizar Estado', 'Menu_SoloActualizarEstados')
      .addToUi();
}


/**
 * Añade el menú en la planilla "Bandeja de Revisión Adherencia".
 */
function onOpenBandeja() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Acciones Bandeja')
      .addItem('⏩ Procesar Casos Seleccionados', 'procesarCasosBandeja_manual')
      .addToUi();
}
