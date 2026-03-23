// Ingresar evento.gs

// --- FUNCIÓN PRINCIPAL ---
// Mantiene lógica original optimizada y tareas pendientes de mover casos a Seguimiento.

/**
 * Asigna una agenda abierta a las filas marcadas, las envía a Supervisora y las mueve a Seguimiento.
 */
function asignarValorPorCheckboxConInput(seleccion) {
  const LIMITE_CASOS = 10;
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Casos y Cupos");
  const ui = SpreadsheetApp.getUi();
  const fechaActual = new Date();
  const FILA_INICIO_DATOS = 2;
  const ANCHO_MAXIMO = COLUMNAS_CUPOS.ESTADO_DISTRIBUCION;

  // ========================================
  // FASE 1: LECTURA Y VALIDACIÓN INICIAL
  // ========================================

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_INICIO_DATOS) {
    ui.alert("Sin Datos", "No se encontraron casos para procesar en esta planilla.", ui.ButtonSet.OK);
    return;
  }

  const rangoDatos = hoja.getRange(FILA_INICIO_DATOS, 1, ultimaFila - FILA_INICIO_DATOS + 1, ANCHO_MAXIMO);
  const todosLosValores = rangoDatos.getValues();

  const filasParaProcesar = [];
  let contadorFilasVacias = 0;

  todosLosValores.forEach((fila, index) => {
    const idCaso = fila[COLUMNAS_CUPOS.ID_CASO - 1];
    if (fila[COLUMNAS_CUPOS.CHECKBOX_ASIGNAR_AGENDA - 1] === true) {
      if (idCaso && String(idCaso).trim() !== "") {
        filasParaProcesar.push({
          numero: FILA_INICIO_DATOS + index,
          datos: fila
        });
      } else {
        contadorFilasVacias++;
        hoja.getRange(FILA_INICIO_DATOS + index, COLUMNAS_CUPOS.CHECKBOX_ASIGNAR_AGENDA).uncheck();
      }
    }
  });

  if (filasParaProcesar.length === 0) {
    let msg = "No se encontraron casillas de verificación marcadas.";
    if (contadorFilasVacias > 0) {
      msg = `Se desmarcaron ${contadorFilasVacias} filas vacías. No se encontraron casos válidos para procesar.`;
    }
    ui.alert("Sin Cambios", msg, ui.ButtonSet.OK);
    return;
  }

  // VALIDAR LÍMITE
  if (filasParaProcesar.length > LIMITE_CASOS) {
    ui.alert(
      "Límite Excedido",
      `Solo puedes procesar ${LIMITE_CASOS} casos a la vez.\n\n` +
      `Actualmente tienes ${filasParaProcesar.length} casos marcados.\n\n` +
      `Por favor desmarca algunos casos e intenta nuevamente.`,
      ui.ButtonSet.OK
    );
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Procesando ${filasParaProcesar.length} caso(s)...`,
    'Asignando Agenda',
    -1
  );

  // ========================================
  // FASE 2: PREPARACIÓN DE RECURSOS
  // ========================================

  let planillaSupervisora, hojaSupervisora;
  let planillaCE, hojaCE;
  let planillaAsignacion, hojaAsignacion;

  try {
    planillaSupervisora = SpreadsheetApp.openById(IDs.supervisora);
    hojaSupervisora = planillaSupervisora.getSheetByName("Casos");

    planillaCE = SpreadsheetApp.openById(IDs.cambioEstado);
    hojaCE = planillaCE.getSheetByName("2025");

    planillaAsignacion = SpreadsheetApp.openById(IDs.asignacionCupos);
    hojaAsignacion = planillaAsignacion.getSheetByName(NOMBRES_HOJAS.asignacionCupos);
  } catch (e) {
    Logger.log(`ERROR FATAL al abrir planillas: ${e.message}`);
    ui.alert("ERROR GRAVE", "No se pudieron abrir las planillas necesarias.", ui.ButtonSet.OK);
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    return;
  }

  // Construir mapa de Asignación de Cupos
  const mapaAsignacion = construirMapaAsignacion(hojaAsignacion);
  let proximoCorrelativo = getProximoCorrelativoCE(hojaCE);

  // ========================================
  // FASE 3: PROCESAMIENTO CASO POR CASO
  // ========================================

  let contadorExito = 0;
  let contadorFallo = 0;
  const errores = [];
  const casosExitosos = [];

  filasParaProcesar.forEach((filaObjeto, index) => {
    const filaNum = filaObjeto.numero;
    const valoresFila = filaObjeto.datos;
    const idCaso = valoresFila[COLUMNAS_CUPOS.ID_CASO - 1];

    try {
      Logger.log(`[${index + 1}/${filasParaProcesar.length}] Procesando caso: ${idCaso}`);

      // PASO 1: Escribir en Supervisora
      const filaParaSupervisora = [
        idCaso,
        valoresFila[COLUMNAS_CUPOS.RUT - 1],
        valoresFila[COLUMNAS_CUPOS.DV - 1],
        valoresFila[COLUMNAS_CUPOS.NOMBRES - 1],
        valoresFila[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1],
        valoresFila[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1],
        "Agendar",
        valoresFila[COLUMNAS_CUPOS.ORIGEN - 1],
        valoresFila[COLUMNAS_CUPOS.FECHA_ENTRADA - 1],
        valoresFila[COLUMNAS_CUPOS.PRESTA_EST - 1],
        "Oficina de Red",
        valoresFila[COLUMNAS_CUPOS.COMENTARIO - 1],
        "",
        seleccion, // Agenda seleccionada
        valoresFila[COLUMNAS_CUPOS.FECHA_AGENDA - 1],
        valoresFila[COLUMNAS_CUPOS.HORA_AGENDA - 1],
        "Sin Gestión",
        fechaActual,
        "", "", "", "",
        valoresFila[COLUMNAS_CUPOS.EMAIL - 1],
        "",
        valoresFila[COLUMNAS_CUPOS.TELEFONO_FIJO - 1],
        valoresFila[COLUMNAS_CUPOS.TELEFONO_MOVIL - 1],
        "", "", "", "", "", "",
        valoresFila[COLUMNAS_CUPOS.FUENTE - 1]
      ];

      const filaInicioSupervisora = getPrimeraFilaVacia(hojaSupervisora, 1, 2);
      hojaSupervisora.getRange(filaInicioSupervisora, 1, 1, filaParaSupervisora.length)
        .setValues([filaParaSupervisora]);

      // PASO 2: Escribir en CambioEstado
      const filaParaCE = [
        proximoCorrelativo++,
        idCaso,
        3, // Estado "En campaña Agendar"
        fechaActual,
        "", "", "", "", "",
        seleccion
      ];
      hojaCE.appendRow(filaParaCE);

      // PASO 3: Actualizar datos y escribir en Seguimiento
      valoresFila[COLUMNAS_CUPOS.ESTADO_COD - 1] = 3;
      valoresFila[COLUMNAS_CUPOS.ESTADO_TEXTO - 1] = MAPAS.estados[3];
      valoresFila[COLUMNAS_CUPOS.FECHA_EN_CAMP_AGENDAR - 1] = fechaActual;
      valoresFila[COLUMNAS_CUPOS.TIPO_ACCION - 1] = "Agendar";
      valoresFila[COLUMNAS_CUPOS.AGENDA - 1] = seleccion;

      moverASeguimientoSinBorrar(valoresFila);

      // PASO 4: Marcar para eliminación
      casosExitosos.push({
        filaEspecialidad: filaNum,
        filaAsignacion: mapaAsignacion[idCaso] || null,
        idCaso: idCaso
      });

      contadorExito++;
      Logger.log(`  ✅ Caso ${idCaso} procesado exitosamente`);

    } catch (error) {
      contadorFallo++;
      errores.push(`Caso ${idCaso || 'sin ID'}: ${error.message}`);
      Logger.log(`  ❌ ERROR: ${error.message}`);

      // Desmarcar checkbox
      try {
        hoja.getRange(filaNum, COLUMNAS_CUPOS.CHECKBOX_ASIGNAR_AGENDA).uncheck();
      } catch (e) {
        Logger.log(`  ⚠️ No se pudo desmarcar checkbox`);
      }
    }
  });

  // ========================================
  // FASE 4: ELIMINACIÓN SEGURA
  // ========================================

  casosExitosos.sort((a, b) => b.filaEspecialidad - a.filaEspecialidad);

  casosExitosos.forEach(caso => {
    try {
      if (caso.filaAsignacion && caso.filaAsignacion > 1) {
        hojaAsignacion.deleteRow(caso.filaAsignacion);
        Logger.log(`  🗑️ Caso ${caso.idCaso}: eliminado de Asignación`);
      }
      hoja.deleteRow(caso.filaEspecialidad);
      Logger.log(`  🗑️ Caso ${caso.idCaso}: eliminado de Especialidad`);
    } catch (error) {
      Logger.log(`  ❌ ERROR al eliminar caso ${caso.idCaso}: ${error.message}`);
    }
  });

  // ========================================
  // FASE 5: REPORTE FINAL
  // ========================================

  SpreadsheetApp.getActiveSpreadsheet().toast('');

  let mensajeFinal = `✅ Proceso Completado\n\n${contadorExito} caso(s) procesado(s) y movido(s) a Supervisora y Seguimiento.`;

  if (contadorFallo > 0) {
    mensajeFinal += `\n\n⚠️ ${contadorFallo} caso(s) fallaron.`;
    mensajeFinal += `\nRevisa los logs para más detalles.`;
  }

  ui.alert("Proceso Completado", mensajeFinal, ui.ButtonSet.OK);
}

// --- FUNCIONES DE SOPORTE ORIGINALES (SIN GRANDES CAMBIOS) ---
// Estas funciones se mantienen como estaban en la versión original para asegurar la compatibilidad.

/**
 * Muestra el diálogo HTML personalizado para seleccionar una opción.
 */
function mostrarDialogoAsignarValor() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Elegir Agenda')
    .setWidth(400)
    .setHeight(220)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Asignar Agenda Abierta');
}

/**
 * Obtiene la lista de opciones para el desplegable del diálogo HTML.
 */
function getOpcionesParaDesplegable() {
  // Usa la planilla activa en el momento de la ejecución.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Agendas");

  if (sheet) {
    const lastRow = sheet.getLastRow();
    // Si la hoja no tiene datos, retorna un array vacío.
    if (lastRow < 1) return [];

    const range = sheet.getRange("L2:L" + lastRow);
    const values = range.getValues();

    return values.flat().filter(String);
  }

  return []; // Si no encuentra la hoja "Agendas" por alguna razón, retorna un array vacío.
}

// --- FUNCIÓN AÑADIDA PARA EL NUEVO MENÚ "ENVIAR A C.C." ---

/**
 * Procesa el envío a C.C. de las filas marcadas con el checkbox 'Enviar a C.C. Uno a Uno' (AO). Esta función es llamada desde el menú, no por un trigger. Valida cada fila antes de procesarla.
 * NUEVA VERSIÓN: Procesamiento secuencial caso por caso con límite de 10 casos.
 */
function procesarEnvioCC_lote() {
  const LIMITE_CASOS = 10;
  const ui = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Casos y Cupos");
  const fechaActual = new Date();
  const FILA_INICIO_DATOS = 2;
  const ANCHO_MAXIMO = COLUMNAS_CUPOS.ESTADO_DISTRIBUCION; // 61

  // ========================================
  // FASE 1: VALIDACIÓN INICIAL Y LECTURA
  // ========================================
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_INICIO_DATOS) {
    ui.alert("Sin Datos", "No hay datos para procesar.", ui.ButtonSet.OK);
    return;
  }

  // Leer todos los datos de una sola vez
  const rangoDatos = hoja.getRange(FILA_INICIO_DATOS, 1, ultimaFila - FILA_INICIO_DATOS + 1, ANCHO_MAXIMO);
  const todosLosValores = rangoDatos.getValues();

  // Identificar filas marcadas
  const filasParaProcesar = [];
  todosLosValores.forEach((fila, index) => {
    if (fila[COLUMNAS_CUPOS.CHECKBOX_ENVIAR_CC - 1] === true) {
      filasParaProcesar.push({
        numero: FILA_INICIO_DATOS + index,
        datos: fila
      });
    }
  });

  // Verificar que hay casos para procesar
  if (filasParaProcesar.length === 0) {
    ui.alert("Sin Cambios", "No se encontraron casillas 'Enviar a C.C. Uno a Uno' (AO) marcadas.", ui.ButtonSet.OK);
    return;
  }

  // VALIDAR LÍMITE DE CASOS
  if (filasParaProcesar.length > LIMITE_CASOS) {
    ui.alert(
      "Límite Excedido",
      `Solo puedes procesar ${LIMITE_CASOS} casos a la vez.\n\n` +
      `Actualmente tienes ${filasParaProcesar.length} casos marcados.\n\n` +
      `Por favor desmarca algunos casos e intenta nuevamente.\n\n`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Confirmación final
  const respuesta = ui.alert(
    'Confirmar Envío a C.C.',
    `Se procesarán ${filasParaProcesar.length} caso(s).\n\n` +
    `Este proceso:\n` +
    `• Enviará los casos a Supervisora\n` +
    `• Los moverá a Seguimiento\n` +
    `• Los eliminará de esta planilla y de Asignación de Cupos\n\n` +
    `¿Deseas continuar?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta !== ui.Button.OK) {
    ui.alert("Proceso Cancelado", "No se realizaron cambios.", ui.ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Procesando ${filasParaProcesar.length} caso(s)...`, 
    'Procesando', 
    -1
  );

  // ========================================
  // FASE 2: PREPARACIÓN DE RECURSOS
  // ========================================

  // Abrir planillas externas una sola vez
  let planillaSupervisora, hojaSupervisora;
  let planillaCE, hojaCE;
  let planillaAsignacion, hojaAsignacion;

  try {
    planillaSupervisora = SpreadsheetApp.openById(IDs.supervisora);
    hojaSupervisora = planillaSupervisora.getSheetByName("Casos");
    
    planillaCE = SpreadsheetApp.openById(IDs.cambioEstado);
    hojaCE = planillaCE.getSheetByName("2025");
    
    planillaAsignacion = SpreadsheetApp.openById(IDs.asignacionCupos);
    hojaAsignacion = planillaAsignacion.getSheetByName(NOMBRES_HOJAS.asignacionCupos);
  } catch (e) {
    Logger.log(`ERROR FATAL al abrir planillas: ${e.message}`);
    ui.alert(
      "Error Crítico",
      "No se pudieron abrir las planillas necesarias. El proceso se detendrá.\n\n" +
      `Error: ${e.message}`,
      ui.ButtonSet.OK
    );
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    return;
  }

  // Construir mapa de Asignación de Cupos para búsqueda rápida O(1)
  const mapaAsignacion = construirMapaAsignacion(hojaAsignacion);
  Logger.log(`Mapa de Asignación construido con ${Object.keys(mapaAsignacion).length} casos.`);

  // Obtener próximo correlativo para CambioEstado
  let proximoCorrelativo = getProximoCorrelativoCE(hojaCE);

  // ========================================
  // FASE 3: PROCESAMIENTO CASO POR CASO
  // ========================================

  let contadorExito = 0;
  let contadorFallo = 0;
  const errores = [];
  const casosExitosos = []; // Para eliminar al final

  // IMPORTANTE: Procesamos de arriba hacia abajo, pero guardamos las filas a eliminar
  filasParaProcesar.forEach((filaObjeto, index) => {
    const filaNum = filaObjeto.numero;
    const valoresFila = filaObjeto.datos;
    const idCaso = valoresFila[COLUMNAS_CUPOS.ID_CASO - 1];

    try {
      Logger.log(`\n[${index + 1}/${filasParaProcesar.length}] Procesando caso: ${idCaso} (Fila ${filaNum})`);

      // ============================================
      // PASO 1: VALIDACIONES COMPLETAS
      // ============================================
      validarCasoCompleto(valoresFila, filaNum, fechaActual);

      // ============================================
      // PASO 2: ESCRIBIR EN SUPERVISORA
      // ============================================
      const tipoAccion = valoresFila[COLUMNAS_CUPOS.TIPO_ACCION - 1];
      const filaParaSupervisora = construirFilaSupervisora(valoresFila, tipoAccion, fechaActual);
      
      const filaInicioSupervisora = getPrimeraFilaVacia(hojaSupervisora, 1, 2);
      hojaSupervisora.getRange(filaInicioSupervisora, 1, 1, filaParaSupervisora.length)
        .setValues([filaParaSupervisora]);
      
      Logger.log(`  ✓ Escrito en Supervisora (fila ${filaInicioSupervisora})`);

      // ============================================
      // PASO 3: ESCRIBIR EN SEGUIMIENTO
      // ============================================
      const esNotificar = (tipoAccion === "Notificar");
      const nuevoEstadoCod = esNotificar ? 4 : 3;
      const columnaFecha = esNotificar ? COLUMNAS_CUPOS.FECHA_EN_CAMP_NOTIFICAR : COLUMNAS_CUPOS.FECHA_EN_CAMP_AGENDAR;

      // Actualizar datos en memoria
      valoresFila[COLUMNAS_CUPOS.ESTADO_COD - 1] = nuevoEstadoCod;
      valoresFila[COLUMNAS_CUPOS.ESTADO_TEXTO - 1] = MAPAS.estados[nuevoEstadoCod];
      valoresFila[columnaFecha - 1] = fechaActual;

      moverASeguimientoSinBorrar(valoresFila);
      Logger.log(`  ✓ Copiado a Seguimiento`);

      // ============================================
      // PASO 4: ESCRIBIR EN CAMBIOESTADO
      // ============================================
      let fechaCita = "";
      let horaCita = "";
      if (nuevoEstadoCod === 4) {
        fechaCita = valoresFila[COLUMNAS_CUPOS.FECHA_AGENDA - 1];
        horaCita = valoresFila[COLUMNAS_CUPOS.HORA_AGENDA - 1];
      }

      const filaParaCE = [
        proximoCorrelativo++,
        idCaso,
        nuevoEstadoCod,
        fechaActual,
        "", "", "", "", "",
        valoresFila[COLUMNAS_CUPOS.AGENDA - 1],
        fechaCita || "",
        horaCita || ""
      ];
      hojaCE.appendRow(filaParaCE);
      Logger.log(`  ✓ Registrado en CambioEstado (correlativo ${proximoCorrelativo - 1})`);

      // ============================================
      // PASO 5: MARCAR PARA ELIMINACIÓN
      // ============================================
      // Guardamos la info del caso exitoso para eliminar al final
      casosExitosos.push({
        filaEspecialidad: filaNum,
        filaAsignacion: mapaAsignacion[idCaso] || null,
        idCaso: idCaso
      });

      contadorExito++;
      Logger.log(`  ✅ Caso ${idCaso} procesado exitosamente`);

    } catch (error) {
      // El caso falló, registramos el error y continuamos
      contadorFallo++;
      const mensajeError = `Fila ${filaNum} (ID: ${idCaso || 'sin ID'}): ${error.message}`;
      errores.push(mensajeError);
      Logger.log(`  ❌ ERROR: ${mensajeError}`);

      // Desmarcar checkbox para que no se re-procese
      try {
        hoja.getRange(filaNum, COLUMNAS_CUPOS.CHECKBOX_ENVIAR_CC).uncheck();
      } catch (e) {
        Logger.log(`  ⚠️ No se pudo desmarcar checkbox: ${e.message}`);
      }
    }
  });

  // ========================================
  // FASE 4: ELIMINACIÓN SEGURA (DE ABAJO HACIA ARRIBA)
  // ========================================
  
  Logger.log(`\n========== FASE DE ELIMINACIÓN ==========`);
  Logger.log(`Casos exitosos a eliminar: ${casosExitosos.length}`);

  // Ordenar de mayor a menor fila para eliminar de abajo hacia arriba
  casosExitosos.sort((a, b) => b.filaEspecialidad - a.filaEspecialidad);

  casosExitosos.forEach(caso => {
    try {
      // Eliminar de Asignación de Cupos
      if (caso.filaAsignacion && caso.filaAsignacion > 1) {
        hojaAsignacion.deleteRow(caso.filaAsignacion);
        Logger.log(`  🗑️ Caso ${caso.idCaso}: eliminado de Asignación (fila ${caso.filaAsignacion})`);
      } else {
        Logger.log(`  ⚠️ Caso ${caso.idCaso}: no encontrado en Asignación (no se eliminó)`);
      }

      // Eliminar de Especialidad
      hoja.deleteRow(caso.filaEspecialidad);
      Logger.log(`  🗑️ Caso ${caso.idCaso}: eliminado de Especialidad (fila ${caso.filaEspecialidad})`);

    } catch (error) {
      Logger.log(`  ❌ ERROR al eliminar caso ${caso.idCaso}: ${error.message}`);
      // No detenemos el proceso, solo registramos
    }
  });

  // ========================================
  // FASE 5: REPORTE FINAL
  // ========================================

  SpreadsheetApp.getActiveSpreadsheet().toast('');

  let mensajeFinal = `✅ Proceso Completado\n\n`;
  mensajeFinal += `${contadorExito} caso(s) procesado(s) exitosamente.\n`;

  if (contadorFallo > 0) {
    mensajeFinal += `\n⚠️ ${contadorFallo} caso(s) fallaron:\n\n`;
    // Mostrar máximo 10 errores en el alert
    const erroresMostrar = errores.slice(0, 10);
    mensajeFinal += erroresMostrar.join('\n');
    
    if (errores.length > 10) {
      mensajeFinal += `\n\n... y ${errores.length - 10} error(es) más.`;
    }
    
    mensajeFinal += `\n\n💡 Los casos fallidos fueron desmarcados.`;
    mensajeFinal += `\nRevisa los logs (Ver > Registros) para más detalles.`;
  }

  if (contadorExito === filasParaProcesar.length) {
    mensajeFinal += `\n\n🎉 ¡Todos los casos se procesaron correctamente!`;
  }

  ui.alert("Resultado del Proceso", mensajeFinal, ui.ButtonSet.OK);
  
  Logger.log(`\n========== RESUMEN FINAL ==========`);
  Logger.log(`Éxitos: ${contadorExito}`);
  Logger.log(`Fallos: ${contadorFallo}`);
  Logger.log(`Total procesado: ${filasParaProcesar.length}`);
}

// ============================================================================
// FUNCIONES AUXILIARES
// ============================================================================

/**
 * Construye un mapa {idCaso: numeroFila} de la hoja de Asignación de Cupos
 * para búsqueda O(1) en lugar de O(n) por cada caso.
 */
function construirMapaAsignacion(hojaAsignacion) {
  const mapa = {};
  
  try {
    const ultimaFila = hojaAsignacion.getLastRow();
    if (ultimaFila < 2) return mapa;

    // Leer toda la columna de ID_CASO de una sola vez
    const rango = hojaAsignacion.getRange(2, COLUMNAS_CUPOS.ID_CASO, ultimaFila - 1, 1);
    const valores = rango.getValues();

    valores.forEach((fila, index) => {
      const idCaso = fila[0];
      if (idCaso && String(idCaso).trim() !== "") {
        mapa[idCaso] = index + 2; // +2 porque empezamos en fila 2
      }
    });

    Logger.log(`Mapa de Asignación construido: ${Object.keys(mapa).length} casos indexados.`);
  } catch (error) {
    Logger.log(`ERROR al construir mapa de Asignación: ${error.message}`);
  }

  return mapa;
}

/**
 * Valida que un caso cumpla todos los requisitos antes de procesarlo.
 * Lanza una excepción si alguna validación falla.
 */
function validarCasoCompleto(valoresFila, filaNum, fechaActual) {
  // 1. Validar ID Caso
  const idCaso = valoresFila[COLUMNAS_CUPOS.ID_CASO - 1];
  if (!idCaso || String(idCaso).trim() === "") {
    throw new Error("No tiene ID de Caso");
  }

  // 2. Validar Tipo de Acción
  const tipoAccion = valoresFila[COLUMNAS_CUPOS.TIPO_ACCION - 1];
  if (tipoAccion !== "Agendar" && tipoAccion !== "Notificar") {
    throw new Error("Debe seleccionar 'Tipo de Acción' válido (Agendar o Notificar)");
  }

  // 3. Validar Agenda
  const agenda = valoresFila[COLUMNAS_CUPOS.AGENDA - 1];
  if (!agenda || String(agenda).trim() === "") {
    throw new Error("Debe seleccionar una 'Agenda'");
  }

  // 4. Validar Fecha NO sea pasada
  const fechaAgenda = valoresFila[COLUMNAS_CUPOS.FECHA_AGENDA - 1];
  if (fechaAgenda && fechaAgenda instanceof Date) {
    const horaAgenda = valoresFila[COLUMNAS_CUPOS.HORA_AGENDA - 1];
    
    let year = fechaAgenda.getFullYear();
    let month = fechaAgenda.getMonth();
    let day = fechaAgenda.getDate();
    let hour = 0;
    let minute = 0;

    if (horaAgenda && horaAgenda instanceof Date) {
      hour = horaAgenda.getHours();
      minute = horaAgenda.getMinutes();
    }

    const fechaHoraAgenda = new Date(year, month, day, hour, minute);
    
    if (fechaHoraAgenda < fechaActual) {
      const fechaStr = Utilities.formatDate(fechaHoraAgenda, 
        SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 
        'dd/MM/yyyy HH:mm');
      throw new Error(`La fecha de la agenda (${fechaStr}) ya pasó`);
    }
  }

  // Todas las validaciones pasaron
  return true;
}

/**
 * Construye el array de datos para enviar a Supervisora.
 */
function construirFilaSupervisora(valoresFila, tipoAccion, fechaActual) {
  return [
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
    "Sin Gestión",
    fechaActual,
    "", "", "", "",
    valoresFila[COLUMNAS_CUPOS.EMAIL - 1],
    "",
    valoresFila[COLUMNAS_CUPOS.TELEFONO_FIJO - 1],
    valoresFila[COLUMNAS_CUPOS.TELEFONO_MOVIL - 1],
    "", "", "", "", "", "",
    valoresFila[COLUMNAS_CUPOS.FUENTE - 1]
  ];
}