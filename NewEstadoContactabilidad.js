// New Estado Contactabilidad.gs

// ==========================================
// SECCIÓN 1: MENÚS Y TRIGGER
// ==========================================

function Menu_SoloActualizarEstados() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Actualización de Estados', 'Se actualizarán los estados en esta planilla.\n\nIMPORTANTE: Si se detectan cambios en los estados del caso, se actualizarán inmediatamente. La fila se mantendrá aquí para su visualización.', ui.ButtonSet.OK);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Procesando estados y citas...', 'Monitoreo', -1);
  SpreadsheetApp.flush();
  
  const resultado = NewContactabilidad_Actualizar_SoloVista();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('');
  ui.alert(`Actualización Finalizada.\n\nCasos con cambios de estado: ${resultado.casosConCambio}`);
}

function Menu_ArchivarYBorrarTerminos() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('Confirmar Archivado y Limpieza', 
    '⚠️ ATENCIÓN: Se procesarán los casos FINALIZADOS (Notificado, Fallecido, Solucionado, Rechazo, Posterga).\n\n' +
    '1. Se ENVIARÁN a Pre-Egresos o Bandeja de Revisión.\n' +
    '2. Se COPIARÁN a las planillas de Gestiones Realizadas.\n' +
    '3. Se ELIMINARÁN de las planillas activas.\n\n'+
    '¿Deseas continuar?', ui.ButtonSet.OK_CANCEL);

  if (respuesta === ui.Button.OK) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Archivando y limpiando...', 'Limpieza', -1);
    
    const resultado = NewContactabilidad_ProcesarTerminos_Completo();
    
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    ui.alert(`Proceso Finalizado.\n\nCasos archivados y eliminados: ${resultado.casosBorrados}`);
  }
}

function NewContactabilidad_TriggerNocturno() {
  console.log("Inicio Trigger: 1. Actualización");
  NewContactabilidad_Actualizar_SoloVista();
  
  console.log("Inicio Trigger: 2. Limpieza");
  NewContactabilidad_ProcesarTerminos_Completo();
}

// ==========================================
// SECCIÓN 2: LÓGICA CORE
// ==========================================

/**
 * CORE 1: Calcula estados, actualiza Supervisora y SINCRONIZA "NOTIFICADOS". NO BORRA NADA.
 */
function NewContactabilidad_Actualizar_SoloVista() {
  let casosConCambio = 0;

  const ESTADOS_TERMINO = ["NOTIFICADO", "SOLUCIONADO (Atendido (puede ser en HCUCh), Extrasistema, Recuperación espontanea)", "RECHAZO (rechaza, traslado coordinado)", "FALLECIDO", "POSTERGA"];

  const hojasDestino = abrirHojasDeDestino();
  const planillaSup = SpreadsheetApp.openById(IDs.supervisora);
  const hojaSup = planillaSup.getSheetByName(NOMBRES_HOJAS.supervisora);

  if (hojaSup.getLastRow() < 2) return { casosConCambio: 0 };

  const datosSup = hojaSup.getRange(2, 1, hojaSup.getLastRow() - 1, hojaSup.getLastColumn()).getValues();

  // Mapa de Seguimiento (Solo para Notificados en esta fase)
  const datosSeg = hojasDestino.seguimiento.getRange(2, 1, hojasDestino.seguimiento.getLastRow() - 1, hojasDestino.seguimiento.getLastColumn()).getValues();
  const mapaSeg = new Map();
  datosSeg.forEach((fila, idx) => {
    if (fila[0]) {
      const idLimpio = String(fila[0]).replace(/['"]/g, '').trim();
      mapaSeg.set(idLimpio, { fila: fila, rowIndex: idx + 2 });
    }
  });

  const casosPorAgente = new Map();
  datosSup.forEach((fila, index) => {
    const ag = fila[COLUMNAS_SUPERVISORA.AGENTE - 1];
    const id = fila[COLUMNAS_SUPERVISORA.ID_CASO - 1];
    if (ag && ag.trim()) {
      const idClean = String(id).replace(/['"]/g, '').trim();
      if (!casosPorAgente.has(ag.trim())) casosPorAgente.set(ag.trim(), []);
      casosPorAgente.get(ag.trim()).push({ id: idClean, idOriginal: id, rowIndex: index + 2, filaDatos: fila });
    }
  });

  const folder = DriveApp.getFolderById(IDs.carpetaAgentes);

  for (const [agente, casos] of casosPorAgente.entries()) {
    const files = folder.getFilesByName(agente);
    if (!files.hasNext()) continue;
    const ssAgente = SpreadsheetApp.open(files.next());
    const gestionesMap = obtenerTodasLasGestiones(ssAgente);

    for (const caso of casos) {
      const gestiones = gestionesMap.get(caso.id);
      if (!gestiones || gestiones.length === 0) continue;

      const estadoFinalObj = calcularEstado(gestiones);
      const estadoFinal = estadoFinalObj.estado;
      const bestGesture = estadoFinalObj.gestion;

      // Actualizar Supervisora
      const estadoAnt = caso.filaDatos[COLUMNAS_SUPERVISORA.ESTADO_CC - 1];
      const bandera = caso.filaDatos[COLUMNAS_SUPERVISORA.PROCESADO_PREEGRESO - 1];

      // Detección de cambio y sincronización
      if (estadoFinal !== estadoAnt) {
        casosConCambio++;
        hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.ESTADO_CC).setValue(estadoFinal);
        if (bandera === "PENDIENTE_BORRADO") {
          hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.PROCESADO_PREEGRESO).setValue("");
        }

        // Sincronización en VIVO de fechas hacia Seguimiento
        const infoSeg = mapaSeg.get(caso.id);
        if (infoSeg) {
          const fechaAccion = (bestGesture && bestGesture.fechaGestion) ? new Date(bestGesture.fechaGestion) : new Date();
          actualizarSeguimientoEnVivo(infoSeg.rowIndex, estadoFinal, hojasDestino.seguimiento, fechaAccion);
        }
      }

      // Si es NOTIFICADO, escribir datos de cita en Supervisora
      if (estadoFinal === "NOTIFICADO" && bestGesture) {
        hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.AGENDA).setValue(bestGesture.agenda);
        hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.DIA).setValue(bestGesture.fechaCita);
        hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.HORA).setValue(bestGesture.horaCita);
      }

      hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.CANT_GESTIONES).setValue(estadoFinalObj.cantidad);
      if (gestiones.length > 0 && gestiones[0].fechaGestion) {
        // Reordenamos para asegurar fecha inicio correcta
        gestiones.sort((a, b) => new Date(a.fechaGestion) - new Date(b.fechaGestion));
        hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.FECHA_INICIO_GESTION).setValue(gestiones[0].fechaGestion);
      }

      // SINCRONIZACIÓN (SOLO NOTIFICADOS) Los otros estados de término (Fallecido, etc.) se ignoran aquí y se procesan solo en el archivado.
      if (ESTADOS_TERMINO.includes(estadoFinal) && bandera !== "PENDIENTE_BORRADO") {

        if (estadoFinal === "NOTIFICADO") {
          const infoSeg = mapaSeg.get(caso.id);
          const fechaHoy = new Date();
          let datosCita = bestGesture ? { agenda: bestGesture.agenda, fechaCita: bestGesture.fechaCita, horaCita: bestGesture.horaCita } : null;

          // Ejecutamos la lógica de Notificado
          let exito = procesarLogicaNotificado(caso.id, infoSeg, fechaHoy, hojasDestino.cambioEstado, hojasDestino.seguimiento, datosCita);

          // SOLO si es Notificado y se procesó bien, marcamos para borrado.
          if (exito) {
            hojaSup.getRange(caso.rowIndex, COLUMNAS_SUPERVISORA.PROCESADO_PREEGRESO).setValue("PENDIENTE_BORRADO");
          }
          // CAMBIO CLAVE: Quitamos el "else { exito = true }"
          // Para el resto de estados, NO marcamos nada. Se quedan pendientes.
        }
      }
    }    
  }
  SpreadsheetApp.flush();
  return { casosConCambio };
}

/**
 * CORE 2: Limpieza (Enruta, Archiva y Borra).
 */
function NewContactabilidad_ProcesarTerminos_Completo() {
  let casosBorrados = 0;
  const ESTADOS_TERMINO = ["NOTIFICADO", "SOLUCIONADO (Atendido (puede ser en HCUCh), Extrasistema, Recuperación espontanea)", "RECHAZO (rechaza, traslado coordinado)", "FALLECIDO", "POSTERGA"];
  
  const hojasDestino = abrirHojasDeDestino();
  let hojaBandeja = null;
  try {
     const ssBandeja = SpreadsheetApp.openById(IDs.bandejaRevision);
     hojaBandeja = ssBandeja.getSheets()[0];
  } catch(e) { Logger.log("No se pudo abrir Bandeja Revisión: " + e.message); }

  const planillaSup = SpreadsheetApp.openById(IDs.supervisora);
  const hojaSup = planillaSup.getSheetByName(NOMBRES_HOJAS.supervisora);
  if (hojaSup.getLastRow() < 2) return { casosBorrados: 0 };

  const datosSup = hojaSup.getRange(2, 1, hojaSup.getLastRow() - 1, hojaSup.getLastColumn()).getValues();
  
  // Mapa Seguimiento
  const datosSeg = hojasDestino.seguimiento.getRange(2, 1, hojasDestino.seguimiento.getLastRow() - 1, hojasDestino.seguimiento.getLastColumn()).getValues();
  const mapaSeg = new Map();
  datosSeg.forEach((fila, idx) => { 
    if(fila[0]) {
        const idNorm = String(fila[0]).replace(/['"]/g, '').trim();
        mapaSeg.set(idNorm, {fila: fila, rowIndex: idx + 2}); 
    }
  });

  let filasBorrar = [];
  const folder = DriveApp.getFolderById(IDs.carpetaAgentes);
  let casosMap = new Map();
  
  datosSup.forEach((fila, index) => {
    const est = superTrim(fila[COLUMNAS_SUPERVISORA.ESTADO_CC - 1]);
    const ag = superTrim(fila[COLUMNAS_SUPERVISORA.AGENTE - 1]);
    const idRaw = String(fila[COLUMNAS_SUPERVISORA.ID_CASO - 1]);
    const idClean = idRaw.replace(/['"]/g, '').trim();

    if (ESTADOS_TERMINO.includes(est) && ag && idClean) {
        if (!casosMap.has(ag)) casosMap.set(ag, []);
        casosMap.get(ag).push({id: idClean, idOriginal: idRaw, filaDatos: fila, rowIndex: index + 2, estado: est});
    }
  });

  for (const [agente, casos] of casosMap.entries()) {
      const files = folder.getFilesByName(agente);
      let ssAgente = null;
      if (files.hasNext()) ssAgente = SpreadsheetApp.open(files.next());
      let gestionesTodas = ssAgente ? obtenerTodasLasGestiones(ssAgente) : null;

      for (const caso of casos) {
          try {
              const bandera = caso.filaDatos[COLUMNAS_SUPERVISORA.PROCESADO_PREEGRESO - 1];
              let listo = (bandera === "PENDIENTE_BORRADO");

              if (!listo) { 
                  const infoSeg = mapaSeg.get(caso.id);
                  const fechaHoy = new Date();
                  
                  if (caso.estado === "NOTIFICADO") {
                     listo = procesarLogicaNotificado(caso.id, infoSeg, fechaHoy, hojasDestino.cambioEstado, hojasDestino.seguimiento, null);
                  } else {
                     // Enrutamiento Inteligente (Directo o Bandeja)
                     listo = procesarLogicaEgresoOBandeja(caso.estado, caso.id, infoSeg, fechaHoy, hojasDestino, hojaBandeja, caso.filaDatos);
                  }
              }

          if (listo) {
              // 1. Archivar
              archivarCasoSupervisora(caso.filaDatos, hojasDestino.histSupervisora);
              // 2. Borrar Agente
              if (ssAgente && gestionesTodas) {
                  archivarYBorrarGestionesAgente_Directo(ssAgente, caso.id, agente, hojasDestino.histAgentes);
              }
              // 3. Marcar Borrado
              filasBorrar.push(caso.rowIndex);
              casosBorrados++;
          }
      } catch (error) {
            // Si falla un caso, lo registramos pero continuamos con el siguiente
            Logger.log(`Error procesando caso ${caso.id}: ${error.message}`);
      }
    }
  }
  filasBorrar.sort((a, b) => b - a);
  [...new Set(filasBorrar)].forEach(f => { try{hojaSup.deleteRow(f);}catch(e){} });
  return { casosBorrados };
}

// ==========================================
// SECCIÓN 3: HELPERS
// ==========================================

// HELPER PARA ENRUTAMIENTO

function procesarLogicaEgresoOBandeja(estado, id, infoSeg, fecha, hojas, hojaBandejaManual, filaSupervisora) {
    // Normalización adicional para "POSTERGACIONES"
    let causalID = MAPAS.causalPorEstado[estado];
    if (!causalID && (estado === "POSTERGACIONES" || String(estado).includes("POSTERGA"))) {
        causalID = 20; // Forzar Causal 20 si detecta palabra clave
    }
    
    // CASO A: ES PRE-EGRESO (CAUSAL 9 o 20)
    if (causalID) {
      // AQUÍ SI NECESITAMOS infoSeg OBLIGATORIAMENTE
        if (!infoSeg) {
            Logger.log(`[ADVERTENCIA] Caso ${id} (Causal ${causalID}) no encontrado en Seguimiento. No se puede generar Pre-Egreso.`);
            return false; // No podemos procesarlo sin datos maestros
        }

        const datos = infoSeg.fila;       
        let hojaDestino = null;
        let fila = [];
        let colsFormulas = [];
        let formulas = [];

        if (causalID === 9) {
            hojaDestino = hojas.preEgresos.hojas[NOMBRES_HOJAS_PREEGRESO.causal5y9];
            if(!hojaDestino) return false;            
            fila = [
                datos[COLUMNAS_CUPOS.ID_CASO - 1], 
                fecha, 
                datos[COLUMNAS_CUPOS.SERVICIO_SALUD - 1], 
                datos[COLUMNAS_CUPOS.RUT - 1], 
                datos[COLUMNAS_CUPOS.DV - 1], 
                datos[COLUMNAS_CUPOS.NOMBRES - 1], 
                datos[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1], 
                datos[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1], 
                datos[COLUMNAS_CUPOS.FECHA_NAC - 1], 
                datos[COLUMNAS_CUPOS.FECHA_DEF - 1], 
                datos[COLUMNAS_CUPOS.ESPECIALIDAD_CORREGIDA - 1] || datos[COLUMNAS_CUPOS.ESPECIALIDAD - 1], 
                datos[COLUMNAS_CUPOS.SOSPECHA_DIAGNOSTICA - 1], 
                causalID, 
                null, 
                null, 
                "Pendiente"
            ];
            colsFormulas = [13, 14];
            formulas = [
                '=IF(RC[-1]=9;"FALLECIMIENTO";IF(RC[-1]=5;"NO BENEFICIARIO";""))',
                '=IF(OR(RC[-2]=9;RC[-2]=5);"Marcos Batarce";"")'
            ];

        } else if (causalID === 20) {
            hojaDestino = hojas.preEgresos.hojas[NOMBRES_HOJAS_PREEGRESO.causal13_14_20];
            if(!hojaDestino) return false;
            fila = [
                datos[COLUMNAS_CUPOS.ID_CASO - 1], 
                fecha, 
                datos[COLUMNAS_CUPOS.SERVICIO_SALUD - 1], 
                datos[COLUMNAS_CUPOS.TIPO_LE - 1],
                datos[COLUMNAS_CUPOS.RUT - 1], 
                datos[COLUMNAS_CUPOS.DV - 1], 
                datos[COLUMNAS_CUPOS.NOMBRES - 1], 
                datos[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1], 
                datos[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1], 
                datos[COLUMNAS_CUPOS.FECHA_NAC - 1], 
                datos[COLUMNAS_CUPOS.ESPECIALIDAD_CORREGIDA - 1] || datos[COLUMNAS_CUPOS.ESPECIALIDAD - 1], 
                datos[COLUMNAS_CUPOS.FECHA_ENTRADA - 1], 
                datos[COLUMNAS_CUPOS.SOSPECHA_DIAGNOSTICA - 1], 
                datos[COLUMNAS_CUPOS.ID_LOCAL - 1], 
                causalID, 
                null, 
                null, 
                "Pendiente"
            ];
            colsFormulas = [15, 16];
            formulas = [
                '=IF(RC[-1]=13;"TRASLADO COORDINADO"; IF(RC[-1]=14; "NO PERTINENCIA";IF(RC[-1]=20; "POSTERGACIONES";"")))',
                '=IF(RC[-16]="";"";"María José Ariste")'
            ];
        }

        const fv = getPrimeraFilaVacia(hojaDestino);
        hojaDestino.getRange(fv, 1, 1, fila.length).setValues([fila]);

        if (colsFormulas.length > 0) {
            try {
                hojaDestino.getRange(fv, colsFormulas[0] + 1).setFormulaR1C1(formulas[0]); // Glosa
                hojaDestino.getRange(fv, colsFormulas[1] + 1).setFormulaR1C1(formulas[1]); // Responsable
            } catch (e) {
                Logger.log(`Error aplicando fórmulas en PreEgresos: ${e.message}`);
            }
        }

        // ACTUALIZAR SEGUIMIENTO Y CE
        hojas.seguimiento.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.ESTADO_COD).setValue(9);
        hojas.seguimiento.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue("Pre Egresado");
        hojas.seguimiento.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.FECHA_PRE_EGRESADO).setValue(fecha);

        const agenda = datos[COLUMNAS_CUPOS.AGENDA - 1] || "";
        const fechaCita = datos[COLUMNAS_CUPOS.FECHA_AGENDA - 1] || "";
        const horaCita = datos[COLUMNAS_CUPOS.HORA_AGENDA - 1] || "";        
        const corr = getProximoCorrelativoCE(hojas.cambioEstado);
        hojas.cambioEstado.appendRow([corr, id, 9, fecha, causalID, "", "", "", "", agenda, fechaCita, horaCita]);

    } else {
        // CASO B: ES RECHAZO O SOLUCIONADO -> VA A BANDEJA
        // CORRECCIÓN: No requerimos infoSeg para esto, usamos filaSupervisora
        const hojaBandeja = hojas.bandejaRevision || hojaBandejaManual;
        if (!hojaBandeja) return false;
        
        let filaBandeja = [
            filaSupervisora[COLUMNAS_SUPERVISORA.ID_CASO - 1],              // A. Id Caso
            fecha,                                                          // B. FECHA INGRESO BANDEJA (Nueva)
            filaSupervisora[COLUMNAS_SUPERVISORA.RUT - 1],                  // C. Rut
            filaSupervisora[COLUMNAS_SUPERVISORA.DV - 1],                   // D. dv
            filaSupervisora[COLUMNAS_SUPERVISORA.NOMBRES - 1],              // E. Nombres
            filaSupervisora[COLUMNAS_SUPERVISORA.PRIMER_APELLIDO - 1],      // F. Primer Apellido
            filaSupervisora[COLUMNAS_SUPERVISORA.SEGUNDO_APELLIDO - 1],     // G. Segundo Apellido
            filaSupervisora[COLUMNAS_SUPERVISORA.TIPO_ACCION - 1],          // H. Tipo de Acción
            filaSupervisora[COLUMNAS_SUPERVISORA.COMUNA - 1],               // I. Comuna
            filaSupervisora[COLUMNAS_SUPERVISORA.FECHA_ENTRADA - 1],        // J. Fecha de Entrada
            filaSupervisora[COLUMNAS_SUPERVISORA.PRESTACION - 1],           // K. Prestación
            filaSupervisora[COLUMNAS_SUPERVISORA.SOLICITADO_POR - 1],       // L. Solicitado por
            filaSupervisora[COLUMNAS_SUPERVISORA.COMENTARIO_ASIGNADOR - 1], // M. Comentario
            "",                                                             // N. Column 24
            filaSupervisora[COLUMNAS_SUPERVISORA.AGENDA - 1],               // O. Agenda
            filaSupervisora[COLUMNAS_SUPERVISORA.DIA - 1],                  // P. día
            filaSupervisora[COLUMNAS_SUPERVISORA.HORA - 1],                 // Q. hora
            filaSupervisora[COLUMNAS_SUPERVISORA.ESTADO_CC - 1],            // R. Estado Centro de Contacto
            filaSupervisora[COLUMNAS_SUPERVISORA.FECHA_RECEPCION_CC - 1],   // S. Fecha de Recepción en CC
            filaSupervisora[COLUMNAS_SUPERVISORA.FECHA_INICIO_GESTION - 1], // T. Fecha Inicio Gestión
            filaSupervisora[COLUMNAS_SUPERVISORA.CANT_GESTIONES - 1],       // U. Cant. de gestiones
            filaSupervisora[COLUMNAS_SUPERVISORA.AGENTE - 1],               // V. Agente
            "Pendiente",                                                    // W. Estado Revisión
            "",                                                             // X. Causal de Pre-Egreso
            false                                                           // Y. Enviar a Pre-Egreso
        ];
        
        const fv = getPrimeraFilaVacia(hojaBandeja);
        hojaBandeja.getRange(fv, 1, 1, filaBandeja.length).setValues([filaBandeja]);
    }
    
    return true;
}

/**
 * Actualiza en tiempo real la fecha y estado en Seguimiento.
 */
function actualizarSeguimientoEnVivo(rowIndex, estadoTexto, hojaSeg, fecha) {
  // Normalizar el estado a MAYÚSCULAS para comparación
  const estadoNormalizado = String(estadoTexto).toUpperCase().trim();
  // Mapeo manual de Texto Estado -> Columna Fecha en Seguimiento
  const MAPA_ESTADO_FECHA = {
    "EN CAMPAÑA AGENDAR": { cod: 3, col: COLUMNAS_CUPOS.FECHA_EN_CAMP_AGENDAR, textoOriginal: "En campaña Agendar" },
    "EN CAMPAÑA NOTIFICAR": { cod: 4, col: COLUMNAS_CUPOS.FECHA_EN_CAMP_NOTIFICAR, textoOriginal: "En campaña Notificar" },
    "OPERADOR ASIGNADO AGENDAR": { cod: 5, col: COLUMNAS_CUPOS.FECHA_OP_ASIGN_AGENDAR, textoOriginal: "Operador Asignado Agendar" },
    "OPERADOR ASIGNADO NOTIFICAR": { cod: 6, col: COLUMNAS_CUPOS.FECHA_OP_ASIGN_NOTIFICAR, textoOriginal: "Operador Asignado Notificar" },
    "NOTIFICADO": { cod: 7, col: COLUMNAS_CUPOS.FECHA_NOTIFICADO, textoOriginal: "Notificado" },
    "NO SE PRESENTA 1": { cod: 8, col: COLUMNAS_CUPOS.FECHA_NSP1, textoOriginal: "No se presenta 1" },
    "PRE EGRESADO": { cod: 9, col: COLUMNAS_CUPOS.FECHA_PRE_EGRESADO, textoOriginal: "Pre Egresado" },
    "GESTIÓN TERRITORIAL": { cod: 10, col: COLUMNAS_CUPOS.FECHA_GESTION_TERRITORIAL, textoOriginal: "Gestión territorial" }
  };

  const config = MAPA_ESTADO_FECHA[estadoNormalizado];  
  if (config) {
    // 1. Actualizar código de estado
    hojaSeg.getRange(rowIndex, COLUMNAS_CUPOS.ESTADO_COD).setValue(config.cod);
    // 2. Actualizar texto de estado
    hojaSeg.getRange(rowIndex, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue(config.textoOriginal);    
    // 3. Escribir fecha SOLO si está vacía
    const celdaFecha = hojaSeg.getRange(rowIndex, config.col);
    if (celdaFecha.getValue() === "") celdaFecha.setValue(fecha);
  } 
}

function obtenerTodasLasGestiones(agentSpreadsheet) {
  const gestionesMap = new Map();
  NOMBRES_HOJAS.pestañasAgente.forEach(tabName => {
      const sheet = agentSpreadsheet.getSheetByName(tabName);
      if (!sheet) return;
      // Forzamos la lectura hasta la columna 26 (Z) para asegurar que la Columna 24 (Fecha) sea leída
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      const data = sheet.getRange(1, 1, lastRow, 26).getValues(); 
      const isAgendar = (tabName === "Agendar");
      
      const colID = COLUMNAS_AGENTE.ID_CASO - 1;
      const colEstNum = (isAgendar ? COLUMNAS_AGENTE.AGENDAR.ESTADO_CONTACTO : COLUMNAS_AGENTE.NOTIFICAR.ESTADO_CONTACTO) - 1;
      const colEstAdh = (isAgendar ? COLUMNAS_AGENTE.AGENDAR.ESTADO_ADHERENCIA : COLUMNAS_AGENTE.NOTIFICAR.ESTADO_ADHERENCIA) - 1;
      const colFecha = 23;
      const colAgenda = 9; // J
      const colFechaCita = (isAgendar ? COLUMNAS_AGENTE.AGENDAR.FECHA_CITA : COLUMNAS_AGENTE.NOTIFICAR.FECHA_CITA) - 1;
      const colHoraCita = (isAgendar ? COLUMNAS_AGENTE.AGENDAR.HORA_CITA : COLUMNAS_AGENTE.NOTIFICAR.HORA_CITA) - 1;

      for (let k = 1; k < data.length; k++) {
         const idRaw = String(data[k][colID] || "");
         const idClean = idRaw.replace(/['"]/g, '').trim();

         if(idClean) {
            if(!gestionesMap.has(idClean)) gestionesMap.set(idClean, []);
            gestionesMap.get(idClean).push({
                estadoNumero: data[k][colEstNum],
                estadoAdherencia: data[k][colEstAdh],
                fechaGestion: data[k][colFecha],
                agenda: data[k][colAgenda],
                fechaCita: data[k][colFechaCita],
                horaCita: data[k][colHoraCita],
                filaDatos: data[k]
            });
         }
      }
  });
  return gestionesMap;
}

function calcularEstado(gestiones) {
  const PRIORIDAD_CONTACTO = ["CONTESTA", "NO CONTESTA", "CORREO ENVIADO", "VALIDADO", "NO CORRESPONDE", "SIN GESTIÓN"];
  const PRIORIDAD_ADHERENCIA = ["FALLECIDO", "NOTIFICADO", "SOLUCIONADO (Atendido (puede ser en HCUCh), Extrasistema, Recuperación espontanea)", "RECHAZO (rechaza, traslado coordinado)", "ACEPTA (quiere pero no hay cupos)", " EN ESPERA DE CUPO", "POSTERGA", "LLAMAR DESPUES"];  
  gestiones.sort((a, b) => new Date(a.fechaGestion) - new Date(b.fechaGestion));
  let cant = 0, bestContactIdx = 999, bestAdhIdx = 999, bestGesture = null;

  gestiones.forEach(g => {
      let c = superTrim(g.estadoNumero);
      Logger.log(`Leído: ${c}`);
      if(["CONTESTA", "NO CONTESTA", "CORREO ENVIADO", "VALIDADO","NO CORRESPONDE"].includes(c)) cant++;
      let idx = PRIORIDAD_CONTACTO.indexOf(c);
      if(idx!=-1 && idx < bestContactIdx) bestContactIdx = idx;
  });

  let estado = (bestContactIdx===999) ? "SIN GESTIÓN" : PRIORIDAD_CONTACTO[bestContactIdx];
  if(estado === "CONTESTA") {
      gestiones.forEach(g => {
          if(superTrim(g.estadoNumero)==="CONTESTA") {
              let adh = superTrim(g.estadoAdherencia);
              let idx = PRIORIDAD_ADHERENCIA.indexOf(adh);
              if(idx!=-1 && idx <= bestAdhIdx) { bestAdhIdx = idx; bestGesture = g; }
          }
      });
      if(bestAdhIdx!==999) estado = PRIORIDAD_ADHERENCIA[bestAdhIdx];
  }
  return { estado: estado, cantidad: cant, gestion: bestGesture };
}

// Lógica para actualizar Seguimiento y CE cuando hay Notificado
function procesarLogicaNotificado(id, infoSeg, fecha, hojaCE, hojaSeg, datosCita) {
    if(!infoSeg) return false;
    const row = infoSeg.rowIndex;    
    let ag = "", fc = "", hc = "";
    if(datosCita) { ag=datosCita.agenda; fc=datosCita.fechaCita; hc=datosCita.horaCita; }
    else { 
        ag = infoSeg.fila[COLUMNAS_CUPOS.AGENDA - 1];
        fc = infoSeg.fila[COLUMNAS_CUPOS.FECHA_AGENDA - 1]; 
        hc = infoSeg.fila[COLUMNAS_CUPOS.HORA_AGENDA - 1]; 
    }
    actualizarSeguimientoEnVivo(row, "NOTIFICADO", hojaSeg, fecha);    
    if(datosCita) {
        hojaSeg.getRange(row, COLUMNAS_CUPOS.AGENDA).setValue(ag);
        hojaSeg.getRange(row, COLUMNAS_CUPOS.FECHA_AGENDA).setValue(fc);
        hojaSeg.getRange(row, COLUMNAS_CUPOS.HORA_AGENDA).setValue(hc);
    }
    const correlativo = getProximoCorrelativoCE(hojaCE); 
    hojaCE.appendRow([correlativo, id, 7, fecha, "", "", "", "", "", ag, fc, hc]);
    return true;
}

function archivarCasoSupervisora(filaDatos, hojaHist) {
  try {
    const fv = getPrimeraFilaVacia(hojaHist);
    hojaHist.getRange(fv, 1, 1, filaDatos.length).setValues([filaDatos]);
  } catch (e) { Logger.log("Error archivar supervisora: " + e.message); }
}

// Helper específico para borrado seguro
function archivarYBorrarGestionesAgente_Directo(agentSpreadsheet, idCaso, nombreAgente, hojaHistAgentes) {
    let gestionesParaArchivar = [];
    let filasParaBorrar = { "Agendar": [], "Notificar": [] };
    const idLimpioCaso = String(idCaso).replace(/['"]/g, '').trim();

    NOMBRES_HOJAS.pestañasAgente.forEach(tabName => {
        const sheet = agentSpreadsheet.getSheetByName(tabName);
        if (!sheet) return;
        const lastRow = sheet.getLastRow();
        if(lastRow < 2) return;
        const data = sheet.getRange(1, 1, lastRow, 26).getValues();
        for (let k = 1; k < data.length; k++) {
            const idRow = String(data[k][0]).replace(/['"]/g, '').trim();
            if (idRow === idLimpioCaso) {
                gestionesParaArchivar.push({ tipoGestion: tabName, filaDatos: data[k] });
                filasParaBorrar[tabName].push(k + 1);
            }
        }
    });

    let filasHist = [];
    gestionesParaArchivar.forEach(g => {
        const filaH = construirFilaHistoricoAgente(g.filaDatos, g.tipoGestion, nombreAgente);
        if(filaH) filasHist.push(filaH);
    });
    
    if (filasHist.length > 0) {
        try {
           const fv = getPrimeraFilaVacia(hojaHistAgentes);
           hojaHistAgentes.getRange(fv, 1, filasHist.length, filasHist[0].length).setValues(filasHist);
        } catch(e) { Logger.log("Error hist agentes: " + e.message); }
    }

    for (const tab in filasParaBorrar) {
        const sheet = agentSpreadsheet.getSheetByName(tab);
        if(sheet && filasParaBorrar[tab].length > 0) {
            filasParaBorrar[tab].sort((a,b)=>b-a);            
            const frozenRows = sheet.getFrozenRows() || 1;
            const currentMaxRows = sheet.getMaxRows();
            const filasABorrar = filasParaBorrar[tab].length;            
            if ((currentMaxRows - filasABorrar) <= frozenRows) sheet.insertRowAfter(currentMaxRows);
            filasParaBorrar[tab].forEach(f => { try { sheet.deleteRow(f); } catch(e) { if(e.message.includes("eliminar todas")) sheet.getRange(f, 1, 1, sheet.getLastColumn()).clearContent(); } });
        }
    }
}

function abrirHojasDeDestino() {
  const ssPre = SpreadsheetApp.openById(IDs.preEgresos);
  const ssSeg = SpreadsheetApp.openById(IDs.seguimiento);
  const ssCE = SpreadsheetApp.openById(IDs.cambioEstado);
  const ssHS = SpreadsheetApp.openById(IDs.historicoSupervisora);
  const ssHA = SpreadsheetApp.openById(IDs.historicoAgentes);
  
  return {
    preEgresos: { 
      ss: ssPre, 
      hojas: {
        [NOMBRES_HOJAS_PREEGRESO.causal5y9]: ssPre.getSheetByName(NOMBRES_HOJAS_PREEGRESO.causal5y9),
        [NOMBRES_HOJAS_PREEGRESO.causal13_14_20]: ssPre.getSheetByName(NOMBRES_HOJAS_PREEGRESO.causal13_14_20)
      }
    },
    seguimiento: ssSeg.getSheetByName(NOMBRES_HOJAS.seguimiento),
    cambioEstado: ssCE.getSheetByName(NOMBRES_HOJAS.cambioEstado),
    histSupervisora: ssHS.getSheetByName("Casos"),
    histAgentes: ssHA.getSheetByName("Gestiones")
  };
}