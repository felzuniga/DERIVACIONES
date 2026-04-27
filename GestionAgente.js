// GestionAgente.gs

/**
 * Función central que procesa una edición realizada en una planilla de agente.
 */
function procesarEdicionAgente(eventInfo) {
  
  // 1. OBTENER DATOS
  const sheetName = eventInfo.sheetName;
  if (!sheetName) return;

  const valorEditado = eventInfo.valorEditadoTexto;
  const valorEstadoFono = eventInfo.valorEstadoFonoTexto;
  
  const source = SpreadsheetApp.openById(eventInfo.sourceId);

  // 1. Obtenemos la HOJA por su NOMBRE (ej: "Notificar")
  const sheet = source.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Error: No se pudo encontrar la hoja "${sheetName}"`);
    return;
  }
  // 2. Obtenemos el RANGO desde ESA hoja (ej: "N3" en "Notificar")
  const range = sheet.getRange(eventInfo.rangeA1Notation);

  const col = range.getColumn();
  const row = range.getRow();

  const checkMarcado = (eventInfo.value === true || eventInfo.value === "TRUE");

  // 2. DETERMINAR COLUMNAS
  const sheetNameNorm = sheetName.trim().toLowerCase();
  const isAgendar = (sheetNameNorm === "agendar");
  // 17/11: Se añaden nuevas columnas
  let COL_ESTADO_FONO, COL_ESTADO_ADHE, COL_CHECK_COPIA, COL_FECHA_GESTION, COL_FECHA_CITA, COL_HORA_CITA;

  if (isAgendar) {
    COL_ESTADO_FONO = 12;   // L
    COL_ESTADO_ADHE = 14;   // N
    COL_FECHA_CITA = 15;    // O
    COL_HORA_CITA = 16;     // P
    COL_CHECK_COPIA = 13;   // M
    COL_FECHA_GESTION = 24; // X
    Logger.log(`Detectada Hoja 'Agendar'. Columnas FORZADAS: Fono=12, Adhe=14, Copia=13, Fecha=24`);
  } else if (sheetNameNorm === "notificar") {
    COL_ESTADO_FONO = 14;   // N
    COL_ESTADO_ADHE = 16;   // P
    COL_FECHA_CITA = 11;    // K
    COL_HORA_CITA = 12;     // L
    COL_CHECK_COPIA = 15;   // O
    COL_FECHA_GESTION = 24; // X
    Logger.log(`Detectada Hoja 'Notificar'. Columnas FORZADAS: Fono=14, Adhe=16, Copia=15, Fecha=24`);
  } else {
    return;
  }

  // 3. SALIDA RÁPIDA
  const esColumnaRelevante = (col === COL_ESTADO_FONO || col === COL_ESTADO_ADHE || col === COL_CHECK_COPIA);
  if (row <= 1 || !esColumnaRelevante) return;

  // --- LÓGICA PRINCIPAL ---
  Logger.log("La columna es relevante. Iniciando lógica principal...");
  Logger.log(`   -> valorEditado (recibido): "${valorEditado}"`);
  Logger.log(`   -> valorEstadoFono (recibido): "${valorEstadoFono}"`);
  
  // -- Flujo 1: Se editó ESTADO FONO/MAIL --
  if (col === COL_ESTADO_FONO) {
    Logger.log("-> Flujo 1: Detectada edición en ESTADO FONO/MAIL.");
    
    if (valorEditado !== "") {
      Logger.log("   -> Valor no vacío. Escribiendo fecha...");
      sheet.getRange(row, COL_FECHA_GESTION).setValue(new Date());
    } else {
      Logger.log("   -> Valor vacío. Limpiando fecha...");
      sheet.getRange(row, COL_FECHA_GESTION).clearContent();
    }
    if (valorEditado !== "CONTESTA") {
      Logger.log("   -> Valor no es CONTESTA. Limpiando adherencia...");
      sheet.getRange(row, COL_ESTADO_ADHE).clearContent();
    }
    Logger.log("-> Flujo 1 Terminado.");
    SpreadsheetApp.flush();
    return;
  }

  // -- Flujo 2: Se editó ESTADO ADHERENCIA --
  if (col === COL_ESTADO_ADHE) {
    Logger.log("-> Flujo 2: Detectada edición en ESTADO ADHERENCIA.");

    // --- Flujo 2a: Validar "CONTESTA"
    if (valorEstadoFono !== "CONTESTA") {
      range.setValue("");
      SpreadsheetApp.getUi().alert("Solo puede editar ESTADO ADHERENCIA si ESTADO FONO_MAIL dice 'CONTESTA'");
      Logger.log("-> Flujo 2 Terminado (Fallo validación 'CONTESTA').");
      SpreadsheetApp.flush();
      return;
    }

    // --- FLUJO 2b: Validar "NOTIFICADO"
    if (valorEditado === "NOTIFICADO") {
      Logger.log("-> Valor es NOTIFICADO. Validando Fecha/Hora Cita...");

      // Leemos las celdas de Fecha y Hora Cita de esta fila
      const fechaCita = sheet.getRange(row, COL_FECHA_CITA).getValue();
      const horaCita = sheet.getRange(row, COL_HORA_CITA).getValue();

      // Si CUALQUIERA de las dos está vacía
      if (!fechaCita || !horaCita) {
        Logger.log("-> Validación fallida. Faltan Fecha/Hora Cita.");
        range.setValue(""); // Revertir el cambio (borra "NOTIFICADO")
        SpreadsheetApp.getUi().alert(
          "Acción Detenida: Faltan Datos",
          "Para marcar un caso como 'NOTIFICADO', primero debe ingresar la 'Fecha Cita' y 'Hora Cita' del agendamiento.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } else {
        Logger.log("-> Validación de Fecha/Hora Cita OK.");
      }
    }
    Logger.log("-> Flujo 2 Terminado.");
    SpreadsheetApp.flush();
    return;
  }

  // -- Flujo 3: Se marcó el CHECK de COPIA --
  if (col === COL_CHECK_COPIA && checkMarcado) {
    Logger.log("-> Flujo 3: Detectada edición en CHECK COPIA (marcado).");
    
    if (valorEstadoFono !== "NO CONTESTA" && valorEstadoFono !== "CORREO ENVIADO") {
      range.setValue(false);
      SpreadsheetApp.getUi().alert("Solo puede marcar 'Volver a gestionar' si el estado es 'NO CONTESTA' o 'CORREO ENVIADO'.");
      Logger.log("-> Flujo 3 Terminado (Condición no cumplida).");
      return;
    }
    
    Logger.log("-> Flujo 3: Condiciones cumplidas. Copiando...");
    try {
      const valoresFila = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      valoresFila[COL_ESTADO_FONO - 1] = "SIN GESTIÓN";
      valoresFila[COL_CHECK_COPIA - 1] = false;
      valoresFila[COL_ESTADO_ADHE - 1] = "";
      valoresFila[COL_FECHA_GESTION - 1] = "";

      const nuevaFila = getPrimeraFilaVacia(sheet);
      sheet.getRange(nuevaFila, 1, 1, valoresFila.length).setValues([valoresFila]);

      const rangoFilaEditada = sheet.getRange(row, 1, 1, sheet.getLastColumn());
      rangoFilaEditada.setFontLine("line-through");

      const protecciones = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      protecciones.forEach(p => { if (p.getRange().getRow() === row) p.remove(); });

      const proteccion = rangoFilaEditada.protect().setDescription(`Protegida por copia - Fila ${row}`);
      proteccion.removeEditors(proteccion.getEditors());
      proteccion.addEditor(Session.getEffectiveUser());
      proteccion.addEditors(ADMIN_EDITORS);
      Logger.log("-> Flujo 3 Terminado (Éxito).");
      
    } catch (err) {
      Logger.log("ERROR en Flujo 3 (Copiar/Proteger): " + err.message);
      range.setValue(false);
    }
    SpreadsheetApp.flush();
    return;
  }
}
