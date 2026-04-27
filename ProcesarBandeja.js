// ProcesarBandeja.gs

/**
 * FUNCIÓN MANUAL: Procesa todos los casos marcados con check en la Bandeja de Revisión.
 * Enruta a la hoja correcta de PreEgresos y aplica las fórmulas específicas.
 */
function procesarCasosBandeja_manual() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No hay casos en la bandeja.");
    return;
  }

  // Leer datos: Columnas A hasta Y (1 a 25)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 25);
  const data = dataRange.getValues();
  
  let casosAProcesar = [];
  
  // 1. Filtrar filas marcadas
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const isChecked = (row[24] === true); // Columna Y
    
    if (isChecked) {
      const causalTexto = row[23]; // Columna X
      const idCaso = row[0];       // Columna A
      
      if (!causalTexto || String(causalTexto).trim() === "") {
        ui.alert(`Error en fila ${i + 2}: El caso ${idCaso} no tiene Causal seleccionada.`);
        return; 
      }
      
      const causalID = parseInt(String(causalTexto).split(" ")[0]);
      if (isNaN(causalID)) {
        ui.alert(`Error en fila ${i + 2}: Causal inválida.`);
        return;
      }

      casosAProcesar.push({ rowIndex: i + 2, id: idCaso, causalID: causalID, data: row });
    }
  }

  if (casosAProcesar.length === 0) {
    ui.alert("No hay casos marcados para enviar.");
    return;
  }

  const resp = ui.alert(`Confirmar Envío`, `Se enviarán ${casosAProcesar.length} casos a Pre-Egresos.\n\n¿Deseas continuar?`, ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;

  SpreadsheetApp.getActiveSpreadsheet().toast("Procesando...", "Enviando", -1);

  try {
    const ssPre = SpreadsheetApp.openById(IDs.preEgresos);
    const ssSeg = SpreadsheetApp.openById(IDs.seguimiento);
    const hojaSeg = ssSeg.getSheetByName(NOMBRES_HOJAS.seguimiento);
    const ssCE = SpreadsheetApp.openById(IDs.cambioEstado);
    const hojaCE = ssCE.getSheetByName("2025");
    
    let filasBorrar = [];
    let conteoExito = 0;

    // Mapa de Seguimiento para obtener datos completos (Necesario para llenar COLUMNAS_CUPOS)
    const dataSeg = hojaSeg.getRange(2, 1, hojaSeg.getLastRow()-1, hojaSeg.getLastColumn()).getValues();
    const mapSeg = new Map();
    dataSeg.forEach((r, idx) => { if(r[0]) mapSeg.set(String(r[0]).trim(), {fila: r, rowIndex: idx + 2}); });

    for (let caso of casosAProcesar) {
      const infoSeg = mapSeg.get(String(caso.id).trim());
      
      if (!infoSeg) {
        Logger.log(`Caso ${caso.id} no encontrado en Seguimiento. Saltando.`);
        continue;
      }

      const datos = infoSeg.fila;
      const fecha = new Date();
      const causalID = caso.causalID;
      const datosBandeja = caso.data; // Datos leídos de la bandeja (útil para Agenda/Cita si estuvieran ahí)

      let nombreHojaDestino = "";
      let fila = [];
      let colsFormulas = []; 
      let formulas = [];    

      // --- DEFINICIÓN DE GRUPOS Y FORMULAS ---

      // GRUPO 1: CAUSAL 1 y 10
      if (causalID === 1 || causalID === 10) {
          nombreHojaDestino = NOMBRES_HOJAS_PREEGRESO.causal1y10;
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
                null, // P (Glosa)
                null, // Q (Responsable)
                "Pendiente" // R
          ];
          colsFormulas = [15, 16]; // Col P, Q (0-based)
          formulas = [
              '=IF(RC[-1]=1;"ATENCIÓN REALIZADA"; IF(RC[-1]=10; "SOLICITUD DE INDICACIÓN DUPLICADA";""))',
              '=IF(RC[-2]=1;"Marcos Batarce";IF(RC[-2]=10;"María José Ariste";""))'
          ];
      }

      // GRUPO 2: CAUSAL 4, 6, 7, 8, 11
      else if ([4, 6, 7, 8, 11].includes(causalID)) {
          nombreHojaDestino = NOMBRES_HOJAS_PREEGRESO.causal4_6_7_8_11;
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
              '=IF(RC[-1]=7;"RECUPERACIÓN ESPONTÁNEA"; IF(RC[-1]=8; "INASISTENCIAS";IF(RC[-1]=11; "CONTACTO NO CORRESPONDE";IF(RC[-1]=6;"RENUNCIA O RECHAZO VOLUNTARIO";IF(RC[-1]=4;"ATENCIÓN OTORGADA EN EL EXTRA SISTEMA";"")))))',
              '=IF(OR(RC[-2]=4;RC[-2]=6;RC[-2]=7;RC[-2]=8;RC[-2]=11);"Rodrigo González"; "")'
          ];
      }

      // GRUPO 3: CAUSAL 13, 14, 20
      else if ([13, 14, 20].includes(causalID)) {
          nombreHojaDestino = NOMBRES_HOJAS_PREEGRESO.causal13_14_20;
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
              '=IF(OR(RC[-2]=13;RC[-2]=14;RC[-2]=20);"María José Ariste"; "")'
          ];
      }

      // ESCRITURA
      if (!nombreHojaDestino) {
          Logger.log(`Causal ${causalID} no mapeada.`);
          continue;
      }
      
      const hojaDestino = ssPre.getSheetByName(nombreHojaDestino);
      if (!hojaDestino) {
          ui.alert(`Error: Hoja "${nombreHojaDestino}" no encontrada en PreEgresos.`);
          continue;
      }

      // Escribir fila
      const fv = getPrimeraFilaVacia(hojaDestino);
      hojaDestino.getRange(fv, 1, 1, fila.length).setValues([fila]);

      // Aplicar Fórmulas
      if (colsFormulas.length > 0) {
          try {
              hojaDestino.getRange(fv, colsFormulas[0] + 1).setFormulaR1C1(formulas[0]); // Glosa
              hojaDestino.getRange(fv, colsFormulas[1] + 1).setFormulaR1C1(formulas[1]); // Responsable
          } catch (e) {
              Logger.log(`Error fórmulas PreEgresos: ${e.message}`);
          }
      }

      // ACTUALIZACIONES
      
      // Seguimiento
      hojaSeg.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.ESTADO_COD).setValue(9);
      hojaSeg.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.ESTADO_TEXTO).setValue("Pre Egresado");
      hojaSeg.getRange(infoSeg.rowIndex, COLUMNAS_CUPOS.FECHA_PRE_EGRESADO).setValue(fecha);

      // Cambio Estado (Recuperando cita de la Bandeja si existe, o de Seguimiento)
      const agenda = datosBandeja[14] || "";
      const fechaCita = datosBandeja[15] || "";
      const horaCita = datosBandeja[16] || "";
      
      const corr = getProximoCorrelativoCE(hojaCE);  
      hojaCE.appendRow([corr, caso.id, 9, fecha, causalID, "", "", "", "", agenda, fechaCita, horaCita]);

      filasBorrar.push(caso.rowIndex);
      conteoExito++;
    }

    // Borrar de Bandeja
    filasBorrar.sort((a, b) => b - a);
    filasBorrar.forEach(r => sheet.deleteRow(r));

    SpreadsheetApp.getActiveSpreadsheet().toast("");
    ui.alert(`Proceso finalizado. ${conteoExito} casos enviados a Pre-Egresos.`);

  } catch (e) {
    Logger.log(e);
    ui.alert("Error crítico: " + e.message);
  }
}
