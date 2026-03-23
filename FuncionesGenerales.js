// Funciones Generales.gs

/**
 * Encuentra la primera fila vacía en la columna A
 */
function getPrimeraFilaVacia(sheet) {
  const columnaA = sheet.getRange("A:A").getValues();
  // Iteramos desde el final para encontrar la última fila con datos
  for (let i = columnaA.length - 1; i >= 0; i--) {
    if (columnaA[i][0] !== "") {
      return i + 2; // La siguiente fila vacía es la última con datos + 1 (i+1) + 1 = i+2
    }
  }
  return 1; // Si la hoja está completamente vacía, empieza en la fila 1.
}

/**
 * Busca una fila que contenga un valor específico en una columna determinada.
 * Intenta comparar los valores como NÚMEROS (parseFloat) para manejar cualquier variación de texto ("98765", 98765.0, "0098765", "98765.00").
 * Si no son números, los compara como texto.
 */
function findRowPorPatron(sheet, columnNumber, PatronBusqueda) {
  // 1. Validar entradas
  if (!sheet || !columnNumber || PatronBusqueda == null || PatronBusqueda === "") {
    Logger.log(`findRowPorPatron: Entradas inválidas.`);
    return 0; 
  }

  // 2. Intentar parsear el patrón de búsqueda
  const patronFloat = parseFloat(String(PatronBusqueda).trim());
  const patronEsNumero = !isNaN(patronFloat);
  const patronString = String(PatronBusqueda).trim();
  
  // 3. Obtener valores de la columna
  let columnValues;
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return 0; // Hoja vacía
    columnValues = sheet.getRange(1, columnNumber, lastRow, 1).getValues();
  } catch (e) {
    Logger.log(`findRowPorPatron: Error al leer rango: ${e}`);
    return 0; 
  }

  // 4. Iterar y comparar
  for (let i = 0; i < columnValues.length; i++) {
    const valorCelda = columnValues[i][0];
    
    if (valorCelda == null || valorCelda === "") {
      continue; // Saltar celdas vacías
    }
    
    // 5. Intentar parsear el valor de la celda
    const valorCeldaFloat = parseFloat(String(valorCelda).trim());
    const valorCeldaEsNumero = !isNaN(valorCeldaFloat);
    
    let coinciden = false;
    
    // 6. Lógica de comparación
    if (patronEsNumero && valorCeldaEsNumero) {
      // Si ambos son números, comparar los números
      // Esto soluciona 98765 == 98765.0 == 98765.00
      if (valorCeldaFloat === patronFloat) {
        coinciden = true;
      }
    } else {
      // Si uno o ambos no son números (ej. "ID-123"), comparar como texto
      const valorCeldaString = String(valorCelda).trim();
      if (valorCeldaString === patronString) {
        coinciden = true;
      }
    }

    // 7. Si coinciden, devolver la fila
    if (coinciden) {
      return i + 1; // Devolvemos el número de fila (1-based)
    }
  }

  // 8. Si no se encuentra, devolver 0 (número)
  Logger.log(`findRowPorPatron: Patrón "${patronString}" no encontrado en ${sheet.getName()}.`);
  return 0;
}

/**
 * Busca un archivo por su nombre dentro de una carpeta de Drive.
 */
function encontrarArchivoPorNombreEnCarpeta(ID_CARPETA, NOMBRE_ARCHIVO_BUSCADO) {
  try {
    const carpeta = DriveApp.getFolderById(ID_CARPETA);
    const archivos = carpeta.getFilesByName(NOMBRE_ARCHIVO_BUSCADO);
    if (archivos.hasNext()) {
      return archivos.next();
    } else {
      Logger.log(`Archivo "${NOMBRE_ARCHIVO_BUSCADO}" no encontrado.`);
      return null;
    }
  } catch (e) {
    Logger.log(`Error al buscar archivo: ${e.toString()}`);
    return null;
  }
}

/**
 * Elimina espacios normales y espacios no separables (CHAR(160)).
 */
function superTrim(str) {
  if (str == null) return "";
  // \s   = espacios normales
  // \u00A0 = espacio no separable
  return str.toString().replace(/[\s\u00A0]+/g, ' ').trim();
}

/**
 * Toma una fila de datos crudos y la limpia, omitiendo valores nulos o cero.
 */
function limpiarFila(fila) {
  return fila.map(valor => {
    // Si el valor es 0, null, o el texto "null", devuelve un string vacío.
    const valorString = String(valor || '').trim().toLowerCase();
    if (valor === 0 || valor === '0' || valorString === 'null') {
      return ""; // Devuelve un string vacío para que la celda quede en blanco.
    }
    return valor; // De lo contrario, devuelve el valor original.
  });
}

/**
 * Limpia un texto: lo convierte a mayúsculas y quita tildes/diacríticos.
 * @param {string} texto El texto a normalizar.
 * @returns {string} El texto normalizado.
 */
function normalizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
}

/**
 * Esta función es solo para forzar la solicitud de permisos 
 * para 'openById', que es lo que 'asignarValorPorCheckboxConInput' necesita.
 */
function TEST_ForzarPermisosCompletos() {
  try {
    // Intentamos abrir todas las hojas que el script necesita
    SpreadsheetApp.openById(IDs.supervisora);
    SpreadsheetApp.openById(IDs.cambioEstado);
    SpreadsheetApp.openById(IDs.seguimiento);
    SpreadsheetApp.openById(IDs.asignacionCupos);

    Logger.log("¡Permisos concedidos para openById!");
    SpreadsheetApp.getUi().alert("Prueba de permisos exitosa.");

  } catch (e) {
    Logger.log("Error en la prueba de permisos: " + e.message);
    SpreadsheetApp.getUi().alert("Error en la prueba de permisos: " + e.message);
  }
}

/**
 * Busca y devuelve el próximo número correlativo, ignorando encabezados y texto.
 * Esto previene la corrupción por la suma de strings ("Correlativo" + 1 = "Correlativo1").
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaCE - La hoja 'CambioEstado'.
 * @return {number} El próximo correlativo numérico.
 */
function getProximoCorrelativoCE(hojaCE) {
  // Columna A = Correlativo (índice 1)
  const COLUMNA_CORRELATIVO = 1;

  // Si la hoja está vacía (solo encabezado en Fila 1)
  const ultimaFila = hojaCE.getLastRow();
  if (ultimaFila <= 1) {
    return 1;
  }

  // Obtener todos los valores de la columna A, desde la fila 2 hasta la última.
  // getRange(fila inicial, columna inicial, num filas, num columnas)
  const rangoValores = hojaCE.getRange(2, COLUMNA_CORRELATIVO, ultimaFila - 1, 1).getValues();

  let maxCorrelativo = 0;

  // Iterar y encontrar el correlativo más grande
  for (let i = 0; i < rangoValores.length; i++) {
    // Intentar convertir el valor a un número. parseInt() ignora texto.
    const numero = parseInt(rangoValores[i][0], 10);

    // Verificar que sea un número válido (> 0) y que sea mayor que el máximo actual
    if (!isNaN(numero) && numero > maxCorrelativo) {
      maxCorrelativo = numero;
    }
  }

  return maxCorrelativo + 1;
}

// --- NUEVAS FUNCIONES PARA EL ARCHIVADO Y BORRADO ---

/**
 * Re-mapea una fila de Agente al layout unificado del Histórico.
 */
function construirFilaHistoricoAgente(filaDatos, tipoGestion, nombreAgente) {
  const MAPA = COLUMNAS_AGENTE;
  let filaHistorico = new Array(28).fill(""); 
  
  // Identificar si es Agendar o Notificar
  const isAgendar = (tipoGestion === "Agendar");

  try {
    // Copiar columnas comunes (A-J)
    for (let i = 0; i < 10; i++) filaHistorico[i] = filaDatos[i];

    // Mapeo de BLOQUE CENTRAL (K a P) y RESOLUCIÓN    
    if (isAgendar) {
      // LAYOUT: AGENDAR (El layout destino está basado en esta pestaña)
      filaHistorico[10] = filaDatos[MAPA.AGENDAR.ESTADO_CONTACTO - 2];   // K: Fono/Mail
      filaHistorico[11] = filaDatos[MAPA.AGENDAR.ESTADO_CONTACTO - 1];   // L: Estado Fono
      filaHistorico[12] = filaDatos[MAPA.AGENDAR.CHECK_COPIA - 1];       // M: Check Copia
      filaHistorico[13] = filaDatos[MAPA.AGENDAR.ESTADO_ADHERENCIA - 1]; // N: Adherencia
      filaHistorico[14] = filaDatos[MAPA.AGENDAR.FECHA_CITA - 1];        // O: Fecha Cita
      filaHistorico[15] = filaDatos[MAPA.AGENDAR.HORA_CITA - 1];         // P: Hora Cita      
    } else {
      // LAYOUT: NOTIFICAR (Necesita remapeo para coincidir con Agendar)
      filaHistorico[10] = filaDatos[MAPA.NOTIFICAR.ESTADO_CONTACTO - 2];
      filaHistorico[11] = filaDatos[MAPA.NOTIFICAR.ESTADO_CONTACTO - 1];
      filaHistorico[12] = filaDatos[MAPA.NOTIFICAR.CHECK_COPIA - 1];
      filaHistorico[13] = filaDatos[MAPA.NOTIFICAR.ESTADO_ADHERENCIA - 1];
      filaHistorico[14] = filaDatos[MAPA.NOTIFICAR.FECHA_CITA - 1];
      filaHistorico[15] = filaDatos[MAPA.NOTIFICAR.HORA_CITA - 1]; 
    }    
    // Copiar el resto de resolución (Q a W)
    for (let i = 16; i <= 22; i++) filaHistorico[i] = filaDatos[i];

    // MAPEO EXPLÍCITO DE COLUMNA X
    filaHistorico[23] = filaDatos[23];

    // Copiar Columna Y (24) si existe
    if (filaDatos.length > 24) {
        filaHistorico[24] = filaDatos[24];
    }

    // Añadir columnas extra (Z, AA)
    filaHistorico[25] = tipoGestion; // Columna Z
    filaHistorico[26] = nombreAgente; // Columna AA

    return filaHistorico;

  } catch (e) { 
    Logger.log(`Error re-mapeo Agente ${nombreAgente} (${tipoGestion}): ${e.message}. Datos originales: ${filaDatos}`);
    return null; 
  }
}

function archivarCasoSupervisora(filaDatos, hojaHistSupervisora) {
  try {
    const filaVacia = getPrimeraFilaVacia(hojaHistSupervisora);
    hojaHistSupervisora.getRange(filaVacia, 1, 1, filaDatos.length).setValues([filaDatos]);
  } catch (e) { Logger.log("Error archivando Supervisora: " + e.message); }
}

function archivarYBorrarGestionesAgente(agentSpreadsheet, idCaso, nombreAgente, gestiones, hojaHistAgentes) {
  let filasParaHistorico = [];
  let filasParaBorrar = { "Agendar": [], "Notificar": [] };

  gestiones.forEach(g => {
    const tipoGestion = g.tipoGestion.charAt(0).toUpperCase() + g.tipoGestion.slice(1);
    const filaH = construirFilaHistoricoAgente(g.filaDatos, tipoGestion, nombreAgente);
    if(filaH) filasParaHistorico.push(filaH);

    const sheet = agentSpreadsheet.getSheetByName(tipoGestion);
    if (sheet) {
        const filaEnAgente = findRowPorPatron(sheet, 1, idCaso); // Asume ID en Col 1
        if (filaEnAgente !== "0") filasParaBorrar[tipoGestion].push(parseInt(filaEnAgente));
    }
  });

  if (filasParaHistorico.length > 0) {
    try {
      const fv = getPrimeraFilaVacia(hojaHistAgentes);
      hojaHistAgentes.getRange(fv, 1, filasParaHistorico.length, filasParaHistorico[0].length).setValues(filasParaHistorico);
    } catch(e){}
  }

  for (const tab in filasParaBorrar) {
    const sheet = agentSpreadsheet.getSheetByName(tab);
    if(sheet) {
        const unicas = [...new Set(filasParaBorrar[tab])].sort((a,b)=>b-a);
        unicas.forEach(f => { try{ sheet.deleteRow(f); }catch(e){} });
    }
  }
}