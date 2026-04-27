// Recibir Casos.gs
// Script para Procesar archivos de múltiples carpetas y enrutar los casos a la planilla correcta.

/**
 * Ejecuta el proceso de carga para todas las carpetas.
 */
function CargarTodosLosCasos() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('Confirmar Inicio de Carga', '¿Deseas iniciar el proceso de carga de archivos?', ui.ButtonSet.OK_CANCEL);
  
  if (respuesta !== ui.Button.OK) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Operación cancelada.');
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando proceso... por favor espera.', 'Cargando Casos', -1);
    
    let resumen = "";
    
    // 1. Procesar Asignación (Todo lo que hay aquí entra a Asignación)
    resumen += procesarCarpetaDeAsignacion(IDs.origenAsignacion);
    
    // 2. Procesar Pre-Egresos (Todo lo que hay aquí entra a Pre-Egresos)
    resumen += procesarCarpetaDePreEgresos(IDs.origenFallecidos, 9); // Causal 9: Fallecido
    resumen += procesarCarpetaDePreEgresos(IDs.origenPrevision, 5);  // Causal 5: Previsión

    SpreadsheetApp.getActiveSpreadsheet().toast('');
    ui.alert('Proceso de Carga Finalizado', resumen, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`Error en CargarTodosLosCasos: ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    ui.alert(`Ocurrió un error crítico. Detalles: ${e.message}`);
  }
}

/**
 * Procesa la carpeta de asignación. No valida reglas de negocio. Si el archivo está en esta carpeta, se asume que es un caso válido para asignación.
 */
function procesarCarpetaDeAsignacion(idCarpetaOrigen) {
  // Solo abrimos las hojas necesarias para Asignación
  const hojaAsignaCupo = SpreadsheetApp.openById(IDs.asignacionCupos).getSheetByName("Casos y Cupos");
  //const hojaSeguimiento = SpreadsheetApp.openById(IDs.seguimiento).getSheetByName("Casos y Cupos");
  const hojaCE = SpreadsheetApp.openById(IDs.cambioEstado).getSheetByName("2025");

  const carpetaOrigen = DriveApp.getFolderById(idCarpetaOrigen);
  const carpetaProcesados = DriveApp.getFolderById(IDs.destinoProcesados);
  const files = carpetaOrigen.getFiles();
  
  let totalArchivos = 0, totalAsignacion = 0;

  while (files.hasNext()) {
    const file = files.next();
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const datosSinEncabezado = spreadsheet.getSheets()[0].getDataRange().getValues().slice(1);

    if (datosSinEncabezado.length > 0) {
      let filasParaAsignacion = [];
      //let filasParaSeguimiento = []; // Copia exacta para seguimiento
      let filasParaCE = [];
      
      const fechaActual = new Date();
      let proximoCorrelativoCE = getProximoCorrelativoCE(hojaCE);

      datosSinEncabezado.forEach(filaOrigen => {
        let filaDestino = construirFilaBase(filaOrigen);

        filaDestino[COLUMNAS_CUPOS.FECHA_ASIGNACION_CUPOS - 1] = fechaActual;
        filaDestino[COLUMNAS_CUPOS.ESTADO_COD - 1] = 2; // Estado 2 = Asignación

        filasParaAsignacion.push(filaDestino);
        //filasParaSeguimiento.push(filaDestino); // También va a seguimiento

        filasParaCE.push([
            proximoCorrelativoCE++, 
            filaOrigen[COLUMNAS_CUPOS.ID_CASO - 1], 
            2, 
            fechaActual, 
            ""
        ]);
      });

      // 1. Escribir en Asignación de Cupos
      if (filasParaAsignacion.length > 0) {
        const filaInicioAsignacion = getPrimeraFilaVacia(hojaAsignaCupo);
        hojaAsignaCupo.getRange(filaInicioAsignacion, 1, filasParaAsignacion.length, filasParaAsignacion[0].length)
                      .setValues(filasParaAsignacion);
        aplicarFormulas(hojaAsignaCupo, filaInicioAsignacion, filasParaAsignacion.length);
        totalAsignacion += filasParaAsignacion.length;
      }

      // 2. Escribir en Seguimiento (Espejo)
      /*
      if (filasParaSeguimiento.length > 0) {
        const filaInicioSeguimiento = getPrimeraFilaVacia(hojaSeguimiento);
        hojaSeguimiento.getRange(filaInicioSeguimiento, 1, filasParaSeguimiento.length, filasParaSeguimiento[0].length)
                       .setValues(filasParaSeguimiento);
        aplicarFormulas(hojaSeguimiento, filaInicioSeguimiento, filasParaSeguimiento.length);
      }*/

      // 3. Escribir en Cambio de Estado
      if (filasParaCE.length > 0) {
        hojaCE.getRange(hojaCE.getLastRow() + 1, 1, filasParaCE.length, filasParaCE[0].length)
              .setValues(filasParaCE);
      }
    }

    file.moveTo(carpetaProcesados);
    totalArchivos++;
  }
  
  return `Carpeta ASIGNACION DE CUPOS:\n- Archivos procesados: ${totalArchivos}\n- Casos cargados: ${totalAsignacion}\n\n`;
}

/**
 * Procesa las carpetas de pre-egreso directo (Prevision o Fallecidos).
 */
function procesarCarpetaDePreEgresos(idCarpetaOrigen, causal) {
  const hojaPreEgresos = SpreadsheetApp.openById(IDs.preEgresos).getSheetByName("Causal 5 y 9");
  const hojaSeguimiento = SpreadsheetApp.openById(IDs.seguimiento).getSheetByName("Casos y Cupos");
  const hojaCE = SpreadsheetApp.openById(IDs.cambioEstado).getSheetByName("2025");

  const carpetaOrigen = DriveApp.getFolderById(idCarpetaOrigen);
  const carpetaProcesados = DriveApp.getFolderById(IDs.destinoProcesados);
  const files = carpetaOrigen.getFiles();
  
  let totalArchivos = 0, totalCasos = 0;

  while (files.hasNext()) {
    const file = files.next();
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const datosSinEncabezado = spreadsheet.getSheets()[0].getDataRange().getValues().slice(1);

    if (datosSinEncabezado.length > 0) {
      let filasParaPreEgreso = [], filasParaSeguimiento = [], filasParaCE = [];
      const fechaActual = new Date();
      let proximoCorrelativoCE = getProximoCorrelativoCE(hojaCE);

      datosSinEncabezado.forEach(filaOrigen => {
        const filaLimpia = limpiarFila(filaOrigen);

        const filaCompletaPreEgreso = [
          filaLimpia[COLUMNAS_CUPOS.ID_CASO - 1], 
          fechaActual, 
          MAPAS.ssalud[filaLimpia[COLUMNAS_CUPOS.SERVICIO_SALUD - 1]] || filaLimpia[COLUMNAS_CUPOS.SERVICIO_SALUD - 1],
          filaLimpia[COLUMNAS_CUPOS.RUT - 1], 
          filaLimpia[COLUMNAS_CUPOS.DV - 1], 
          filaLimpia[COLUMNAS_CUPOS.NOMBRES - 1],
          filaLimpia[COLUMNAS_CUPOS.PRIMER_APELLIDO - 1], 
          filaLimpia[COLUMNAS_CUPOS.SEGUNDO_APELLIDO - 1], 
          filaLimpia[COLUMNAS_CUPOS.FECHA_NAC - 1],
          filaLimpia[COLUMNAS_CUPOS.FECHA_DEF - 1] || "", 
          filaLimpia[COLUMNAS_CUPOS.ESPECIALIDAD - 1], 
          filaLimpia[COLUMNAS_CUPOS.SOSPECHA_DIAGNOSTICA - 1],
          causal,
          null, // N - Glosa (se llenará con fórmula)
          null, // O - Responsable (se llenará con fórmula)
          "Pendiente" // P - Estado
        ];
        filasParaPreEgreso.push(filaCompletaPreEgreso);

        // Registro en Cambio de Estado
        filasParaCE.push([proximoCorrelativoCE++, filaOrigen[COLUMNAS_CUPOS.ID_CASO - 1], 9, fechaActual, causal]);
        
        let filaDestinoSeguimiento = construirFilaBase(filaLimpia);
        filaDestinoSeguimiento[COLUMNAS_CUPOS.FECHA_PRE_EGRESADO - 1] = fechaActual;
        filaDestinoSeguimiento[COLUMNAS_CUPOS.ESTADO_COD - 1] = 9;
        filasParaSeguimiento.push(filaDestinoSeguimiento);
      });

      // 1. Escribir en PreEgresos
      if (filasParaPreEgreso.length > 0) {
        const filaInicioPreEgreso = getPrimeraFilaVacia(hojaPreEgresos);
        const numFilasNuevas = filasParaPreEgreso.length;

        hojaPreEgresos.getRange(filaInicioPreEgreso, 1, numFilasNuevas, filasParaPreEgreso[0].length)
                      .setValues(filasParaPreEgreso);

        try {
          const formulaGlosaR1C1 = '=IF(RC[-1]=9;"FALLECIMIENTO";IF(RC[-1]=5;"NO BENEFICIARIO";""))';
          hojaPreEgresos.getRange(filaInicioPreEgreso, 14, numFilasNuevas).setFormulaR1C1(formulaGlosaR1C1);
          
          const formulaRespR1C1 = '=IF(OR(RC[-2]=9;RC[-2]=5);"Marcos Batarce";"")';
          hojaPreEgresos.getRange(filaInicioPreEgreso, 15, numFilasNuevas).setFormulaR1C1(formulaRespR1C1);
        } catch (e) {
          Logger.log(`Error al aplicar fórmulas explícitas en PreEgresos (Directo): ${e.message}`);
        }

        totalCasos += filasParaPreEgreso.length;
      }

      // 2. Escribir en Seguimiento
      if (filasParaSeguimiento.length > 0) {
        const filaInicioSeguimiento = getPrimeraFilaVacia(hojaSeguimiento);
        hojaSeguimiento.getRange(filaInicioSeguimiento, 1, filasParaSeguimiento.length, filasParaSeguimiento[0].length)
                       .setValues(filasParaSeguimiento);
        aplicarFormulas(hojaSeguimiento, filaInicioSeguimiento, filasParaSeguimiento.length);
      }

      // 3. Escribir en Cambio de Estado
      if (filasParaCE.length > 0) {
        hojaCE.getRange(hojaCE.getLastRow() + 1, 1, filasParaCE.length, filasParaCE[0].length)
              .setValues(filasParaCE);
      }
    }
    
    file.moveTo(carpetaProcesados);
    totalArchivos++;
  }
  
  const nombreCarpeta = (idCarpetaOrigen === IDs.origenFallecidos) ? "FALLECIDOS" : "PREVISION";
  return `Carpeta ${nombreCarpeta}:\n- Archivos procesados: ${totalArchivos}\n- Casos a Pre-Egreso: ${totalCasos}\n\n`;
}

/**
 * FUNCIÓN DE APOYO PARA MAPEO. Construye la fila de 62 columnas base para Asignación/Seguimiento
 */
function construirFilaBase(filaOrigen) {
  let filaDestino = new Array(62).fill(null);
  const columnasALimpiar = [25, 26, 27, 28, 29]; // Y, Z, AA, AB, AC

  // Mapeo Directo 1-a-1 (Columnas A a AC, índices 0 a 28)
  for (let i = 0; i < 29; i++) {
    const valor = filaOrigen[i];
    const valorString = String(valor || '').trim().toLowerCase();

    if (columnasALimpiar.includes(i + 1) && (valor === 0 || valor === '0' || valorString === 'null')) {
      filaDestino[i] = null;
    } else {
      filaDestino[i] = valor;
    }
  }

  // Mapeo Específico de las columnas restantes (Indices basados en el orden del archivo origen vs destino)
  // Origen AD (29) -> Destino AD (PRESTA_EST)
  filaDestino[COLUMNAS_CUPOS.PRESTA_EST - 1] = filaOrigen[29];

  // Origen AE (30) -> Destino AF (FECHA_ENTRADA)
  filaDestino[COLUMNAS_CUPOS.FECHA_ENTRADA - 1] = filaOrigen[30];

  // Origen AF (31) -> Destino AG (ID_LOCAL)
  filaDestino[COLUMNAS_CUPOS.ID_LOCAL - 1] = filaOrigen[31];

  // Origen AG (32) -> Destino AH (FUENTE)
  filaDestino[COLUMNAS_CUPOS.FUENTE - 1] = filaOrigen[32];

  // Mapeo especial para el código de SSalud
  filaDestino[COLUMNAS_CUPOS.SSALUD_COD - 1] = filaOrigen[COLUMNAS_CUPOS.SERVICIO_SALUD - 1];

  return filaDestino;
}

/**
 * FUNCIÓN DE APOYO PARA INSERTAR FÓRMULAS
 */
function aplicarFormulas(sheet, filaInicio, numFilas) {
  const offsetSsalud = COLUMNAS_CUPOS.SSALUD_COD - COLUMNAS_CUPOS.SERVICIO_SALUD;
  sheet.getRange(filaInicio, COLUMNAS_CUPOS.SERVICIO_SALUD, numFilas).setFormulaR1C1(`=IF(RC[${offsetSsalud}]=9;"SSMN";IF(RC[${offsetSsalud}]=10;"SSMOC";""))`);
  
  sheet.getRange(filaInicio, COLUMNAS_CUPOS.FECHA_RECIBIDO, numFilas).setFormulaR1C1('=RC[-45]');
  
  const offsetEstado = COLUMNAS_CUPOS.ESTADO_COD - COLUMNAS_CUPOS.ESTADO_TEXTO;
  sheet.getRange(filaInicio, COLUMNAS_CUPOS.ESTADO_TEXTO, numFilas).setFormulaR1C1(`=IF(RC[${offsetEstado}]=7;"Notificado"; IF(RC[${offsetEstado}]=2;"Asignación de cupo"; IF(RC[${offsetEstado}]=9;"Pre Egresado";IF(RC[${offsetEstado}]=11;"Egresado"; IF(RC[${offsetEstado}]=6;"Operador Asignado Notificar"; IF(RC[${offsetEstado}]=4;"En campaña Notificar";IF(RC[${offsetEstado}]=3;"En campaña Agendar";IF(RC[${offsetEstado}]=5;"Operador Asignado Agendar";""))))))))`);
}
