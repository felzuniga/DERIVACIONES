// AsignarAgente.gs

/**
 * Procesa todos los casos marcados en la planilla Supervisora, verificando en Seguimiento para no re-procesar.
 */
function asignarAgentesEnLote(idPlanillaSupervisora) {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('Confirmar Asignación', 'Este proceso asignará todos los casos marcados a sus respectivos agentes. ¿Deseas continuar?', ui.ButtonSet.OK_CANCEL);
  if (respuesta !== ui.Button.OK) return;

  SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando asignación...', 'Procesando', -1);

  try {
    // --- 1. ABRIR TODAS LAS HOJAS NECESARIAS ---
    const planillaSupervisora = SpreadsheetApp.openById(idPlanillaSupervisora);
    const hojaSupervisora = planillaSupervisora.getSheetByName("Casos");
    const hojaSeguimiento = SpreadsheetApp.openById(IDs.seguimiento).getSheetByName("Casos y Cupos");

    // --- 2. CARGAR DATOS EN MEMORIA PARA EFICIENCIA ---
    // Cargar todos los casos de Supervisora
    const rangoSupervisora = hojaSupervisora.getRange(2, 1, hojaSupervisora.getLastRow() - 1, hojaSupervisora.getLastColumn());
    const valoresSupervisora = rangoSupervisora.getValues();
    
    // Cargar el estado de todos los casos de Seguimiento en un mapa para búsqueda rápida
    const datosSeguimiento = hojaSeguimiento.getRange(2, 1, hojaSeguimiento.getLastRow() - 1, COLUMNAS_CUPOS.ESTADO_COD).getValues();
    const mapaEstadoSeguimiento = new Map();
    datosSeguimiento.forEach(fila => {
      const idCaso = fila[COLUMNAS_CUPOS.ID_CASO - 1];
      const estadoCod = fila[COLUMNAS_CUPOS.ESTADO_COD - 1];
      if (idCaso) {
        mapaEstadoSeguimiento.set(idCaso.toString().trim(), estadoCod);
      }
    });

    // --- 3. FILTRAR Y AGRUPAR CASOS VÁLIDOS ---
    const casosPorAgente = {};
    let casosYaAsignados = 0;

    valoresSupervisora.forEach((fila, index) => {
      const numeroFilaReal = index + 2;
      const checkboxMarcado = fila[COLUMNAS_SUPERVISORA.CHECKBOX_ASIGNAR - 1];
      if (checkboxMarcado === true) {
        const idCaso = fila[COLUMNAS_SUPERVISORA.ID_CASO - 1].toString().trim();
        const estadoActualSeguimiento = mapaEstadoSeguimiento.get(idCaso);

        // Verificar si ya está asignado (estados 5 o +)
        if (estadoActualSeguimiento >= 5) {          
          casosYaAsignados++;
          Logger.log(`Caso ${idCaso} ya está en estado ${estadoActualSeguimiento} (asignado o finalizado). Omitiendo.`);

          // Desmarcamos el check para evitar confusiones futuras
        hojaSupervisora.getRange(numeroFilaReal, COLUMNAS_SUPERVISORA.CHECKBOX_ASIGNAR).uncheck();

        } else {
          const nombreAgente = fila[COLUMNAS_SUPERVISORA.AGENTE - 1];
          if (nombreAgente) {
            if (!casosPorAgente[nombreAgente]) {
              casosPorAgente[nombreAgente] = [];
            }
            casosPorAgente[nombreAgente].push({ datos: fila });
          }
        }
      }
    });

    // --- LÓGICA DE SALIDA ---
    if (Object.keys(casosPorAgente).length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('');
      let mensajeAlerta = "No hay nuevos casos para procesar. Debe seleccionar un Agente para iniciar el proceso.";
      if (casosYaAsignados > 0) {
        mensajeAlerta = `Se encontraron ${casosYaAsignados} caso(s) que ya habían sido asignados previamente. No hay nuevas asignaciones que realizar.`;
      }
      ui.alert("Nada que procesar", mensajeAlerta, ui.ButtonSet.OK);
      return;
    }
    
    // --- 4. PROCESAR LOS GRUPOS DE AGENTES ---
    const bloqueParaCE = [];
    const casosParaActualizarSeguimiento = [];
    let totalCasosAsignados = 0;
    const fechaActual = new Date();

    for (const nombreAgente in casosPorAgente) {
      const casos = casosPorAgente[nombreAgente];
      const PLANILLA_AGENTE = encontrarArchivoPorNombreEnCarpeta(IDs.carpetaAgentes, nombreAgente);
      if (!PLANILLA_AGENTE) {
        Logger.log(`Planilla para agente ${nombreAgente} no encontrada. Omitiendo ${casos.length} casos.`);
        continue;
      }

      const agentSpreadsheet = SpreadsheetApp.openById(PLANILLA_AGENTE.getId());
      const hojaAgenteAgendar = agentSpreadsheet.getSheetByName("Agendar");
      const hojaAgenteNotificar = agentSpreadsheet.getSheetByName("Notificar");

      let bloqueAgendar = [];
      let bloqueNotificar = [];

      casos.forEach(caso => {
        const valoresFila = caso.datos;
        const tipoAccion = valoresFila[COLUMNAS_SUPERVISORA.TIPO_ACCION - 1];
        const idCaso = valoresFila[COLUMNAS_SUPERVISORA.ID_CASO - 1];
        const filasParaAgente = construirFilaParaAgente(valoresFila); 
        
        if (tipoAccion === "Agendar" && hojaAgenteAgendar) {
          bloqueAgendar = bloqueAgendar.concat(filasParaAgente.agendar);
        } else if (tipoAccion === "Notificar" && hojaAgenteNotificar) {
          bloqueNotificar = bloqueNotificar.concat(filasParaAgente.notificar);
        }

        const nuevoEstadoCod = (tipoAccion === "Agendar") ? 5 : 6;

        // --- CAPTURAR FECHA/HORA PARA ESTADO 6 ---
        let fechaCita = "";
        let horaCita = "";
        
        // Si es "Operador Asignado Notificar" (Estado 6), capturamos la fecha/hora que la supervisora ingresó en las columnas O y P de la hoja "Supervisora"
        if (nuevoEstadoCod === 6) {
           fechaCita = valoresFila[COLUMNAS_SUPERVISORA.DIA - 1]; // Col O
           horaCita = valoresFila[COLUMNAS_SUPERVISORA.HORA - 1]; // Col P
        }
        // Para Estado 5 (Agendar), se quedan vacíos, lo cual es correcto.

        bloqueParaCE.push([
          , // A (Correlativo)
          idCaso, 
          nuevoEstadoCod, 
          fechaActual, 
          "", 
          "", 
          "", 
          "", 
          "", 
          valoresFila[COLUMNAS_SUPERVISORA.AGENDA - 1],
          fechaCita, // K (Fecha Cita)
          horaCita // L (Hora Cita)        
        ]);

        casosParaActualizarSeguimiento.push({ idCaso: idCaso, estado: nuevoEstadoCod, fecha: fechaActual });
        totalCasosAsignados++;
      });

      if (bloqueAgendar.length > 0) {
        hojaAgenteAgendar.getRange(getPrimeraFilaVacia(hojaAgenteAgendar), 1, bloqueAgendar.length, bloqueAgendar[0].length).setValues(bloqueAgendar);
      }
      if (bloqueNotificar.length > 0) {
        hojaAgenteNotificar.getRange(getPrimeraFilaVacia(hojaAgenteNotificar), 1, bloqueNotificar.length, bloqueNotificar[0].length).setValues(bloqueNotificar);
      }
    }
    
    // --- 5. ACTUALIZAR PLANILLAS EXTERNAS Y LIMPIAR ---
    actualizarCambioEstadoEnLote(bloqueParaCE);
    actualizarSeguimientoEnLote(casosParaActualizarSeguimiento);

    SpreadsheetApp.getActiveSpreadsheet().toast('');
    let mensajeFinal = `Se asignaron ${totalCasosAsignados} nuevos casos con éxito.`;
    if (casosYaAsignados > 0) {
      mensajeFinal += `\nSe omitieron ${casosYaAsignados} casos que ya estaban asignados.`;
    }
    ui.alert("Proceso Finalizado", mensajeFinal, ui.ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    Logger.log(`Error en asignarAgentesEnLote: ${e.toString()} \n ${e.stack}`);
    ui.alert(`Ocurrió un error en la asignación: ${e.message}`);
  }
}

// --- FUNCIONES AUXILIARES PARA EL PROCESO ---

function construirFilaParaAgente(valoresFila) {
  const filaBase = [
    valoresFila[COLUMNAS_SUPERVISORA.ID_CASO - 1], valoresFila[COLUMNAS_SUPERVISORA.RUT - 1], valoresFila[COLUMNAS_SUPERVISORA.DV - 1], 
    valoresFila[COLUMNAS_SUPERVISORA.NOMBRES - 1], valoresFila[COLUMNAS_SUPERVISORA.PRIMER_APELLIDO - 1], valoresFila[COLUMNAS_SUPERVISORA.SEGUNDO_APELLIDO - 1],
    valoresFila[COLUMNAS_SUPERVISORA.COMUNA - 1], valoresFila[COLUMNAS_SUPERVISORA.FECHA_ENTRADA - 1], valoresFila[COLUMNAS_SUPERVISORA.PRESTACION - 1], 
    valoresFila[COLUMNAS_SUPERVISORA.AGENDA - 1]
  ];
  
  const mediosDeContacto = new Set();
  for (let i = COLUMNAS_SUPERVISORA.MAIL_1 - 1; i <= COLUMNAS_SUPERVISORA.TELEFONO_8 - 1; i++) { 
    if (String(valoresFila[i] || '').trim()) mediosDeContacto.add(String(valoresFila[i]).trim());
  }

  const filasAgendar = [];
  const filasNotificar = [];

  mediosDeContacto.forEach(medio => {
    let agendar = filaBase.slice();
    agendar[10] = medio;
    agendar[11] = "SIN GESTIÓN";
    agendar[14] = valoresFila[COLUMNAS_SUPERVISORA.DIA - 1];
    agendar[15] = valoresFila[COLUMNAS_SUPERVISORA.HORA - 1];
    while(agendar.length < 23) agendar.push('');
    agendar[22] = valoresFila[COLUMNAS_SUPERVISORA.SOLICITADO_POR - 1];
    filasAgendar.push(agendar);
    
    let notificar = filaBase.slice();
    notificar[10] = valoresFila[COLUMNAS_SUPERVISORA.DIA - 1];
    notificar[11] = valoresFila[COLUMNAS_SUPERVISORA.HORA - 1];
    notificar[12] = medio;
    notificar[13] = "SIN GESTIÓN";
    while(notificar.length < 23) notificar.push('');
    notificar[22] = valoresFila[COLUMNAS_SUPERVISORA.SOLICITADO_POR - 1];
    filasNotificar.push(notificar);
  });
  
  return { agendar: filasAgendar, notificar: filasNotificar };
}

function actualizarCambioEstadoEnLote(bloque) {
  if (bloque.length === 0) return;
  const hojaCE = SpreadsheetApp.openById(IDs.cambioEstado).getSheetByName("2025");
  if (!hojaCE) return;
  let proximoCorrelativo = getProximoCorrelativoCE(hojaCE);
  bloque.forEach(fila => fila[0] = proximoCorrelativo++);
  hojaCE.getRange(hojaCE.getLastRow() + 1, 1, bloque.length, bloque[0].length).setValues(bloque);
}

/**
 * Actualiza el estado de los casos en la planilla de Seguimiento de forma individual y precisa.
 */
function actualizarSeguimientoEnLote(casos) {
  if (casos.length === 0) return;
  const hojaSeguimiento = SpreadsheetApp.openById(IDs.seguimiento).getSheetByName("Casos y Cupos");
  if (!hojaSeguimiento) return;

  // 1. Creamos un mapa de ID_Caso -> Fila para encontrar rápidamente dónde escribir.
  const idsEnSeguimiento = hojaSeguimiento.getRange(2, COLUMNAS_CUPOS.ID_CASO, hojaSeguimiento.getLastRow() - 1, 1).getValues();
  const mapaDeFilas = new Map();
  idsEnSeguimiento.forEach((fila, index) => {
    const idCaso = fila[0];
    if (idCaso) {
      mapaDeFilas.set(idCaso.toString().trim(), index + 2); // Guardamos el número de fila real
    }
  });

  // 2. Por cada caso a actualizar, vamos directamente a su fila y actualizamos solo las celdas necesarias.
  casos.forEach(caso => {
    const numeroFila = mapaDeFilas.get(caso.idCaso.toString().trim());
    if (numeroFila) {
      const nuevoEstadoCod = caso.estado;
      const columnaFecha = (nuevoEstadoCod === 5) ? COLUMNAS_CUPOS.FECHA_OP_ASIGN_AGENDAR : COLUMNAS_CUPOS.FECHA_OP_ASIGN_NOTIFICAR;

      hojaSeguimiento.getRange(numeroFila, COLUMNAS_CUPOS.ESTADO_COD).setValue(nuevoEstadoCod);
      hojaSeguimiento.getRange(numeroFila, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue(MAPAS.estados[nuevoEstadoCod]);
      hojaSeguimiento.getRange(numeroFila, columnaFecha).setValue(caso.fecha);
    }
  });
}