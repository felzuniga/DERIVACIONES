// BusquedaLocal.gs

/**
 * BÚSQUEDA LOCAL DE CASOS
 * Script que se ejecuta en la planilla actual para buscar IDs.
 * 
 * INSTRUCCIONES:
 * 1. Crea una hoja llamada "Búsqueda" en la planilla donde quieras buscar
 * 2. En la columna A (desde fila 2), pega los IDs que quieres buscar
 * 3. Ejecuta la función desde el menú: Búsqueda > Buscar IDs en esta planilla
 * 4. Los resultados aparecerán en las columnas B, C, D
 */

/**
 * Crea el menú de búsqueda en la planilla actual.
 */
/*
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Búsqueda')
    .addItem('📋 Preparar Hoja de Búsqueda', 'prepararHojaBusqueda')
    .addItem('🔎 Buscar IDs en esta Planilla', 'buscarIDsEnPlanillaActual')
    .addItem('🧹 Limpiar Resultados', 'limpiarResultados')
    .addToUi();
}
*/
/**
 * Crea una hoja de búsqueda con el formato correcto.
 */
function prepararHojaBusqueda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Verificar si ya existe
  let hojaBusqueda = ss.getSheetByName("Búsqueda");
  
  if (hojaBusqueda) {
    const respuesta = ui.alert(
      'La hoja "Búsqueda" ya existe',
      '¿Deseas limpiarla y empezar de nuevo?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta === ui.Button.YES) {
      hojaBusqueda.clear();
    } else {
      return;
    }
  } else {
    hojaBusqueda = ss.insertSheet("Búsqueda");
  }
  
  // Configurar encabezados
  hojaBusqueda.getRange(1, 1, 1, 5).setValues([[
    "ID a Buscar",
    "Encontrado",
    "Ubicaciones",
    "Cantidad",
    "Links"
  ]]);
  
  // Formato de encabezados
  hojaBusqueda.getRange(1, 1, 1, 5)
    .setFontWeight("bold")
    .setBackground("#4CAF50")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  
  // Instrucciones en la celda A2
  hojaBusqueda.getRange(2, 1).setValue("Pega aquí los IDs a buscar (uno por fila) →");
  hojaBusqueda.getRange(2, 1).setFontStyle("italic").setFontColor("#666666");
  
  // Ajustar anchos de columna
  hojaBusqueda.setColumnWidth(1, 200); // ID
  hojaBusqueda.setColumnWidth(2, 100); // Encontrado
  hojaBusqueda.setColumnWidth(3, 300); // Ubicaciones
  hojaBusqueda.setColumnWidth(4, 80);  // Cantidad
  hojaBusqueda.setColumnWidth(5, 150); // Links
  
  // Congelar primera fila
  hojaBusqueda.setFrozenRows(1);
  
  // Activar la hoja
  ss.setActiveSheet(hojaBusqueda);
  
  ui.alert(
    'Hoja de Búsqueda Creada',
    'Ahora puedes:\n\n' +
    '1. Pegar los IDs en la columna A (desde la fila 2 o 3)\n' +
    '2. Ir al menú: Búsqueda > Buscar IDs en esta Planilla\n' +
    '3. Los resultados aparecerán automáticamente',
    ui.ButtonSet.OK
  );
}

/**
 * Busca los IDs en todas las hojas de la planilla actual.
 */
function buscarIDsEnPlanillaActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Verificar que existe la hoja de búsqueda
  const hojaBusqueda = ss.getSheetByName("Búsqueda");
  if (!hojaBusqueda) {
    ui.alert(
      'Hoja no encontrada',
      'Primero debes crear la hoja de búsqueda.\n\n' +
      'Ve a: Búsqueda > Preparar Hoja de Búsqueda',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Leer IDs a buscar desde la columna A
  const lastRow = hojaBusqueda.getLastRow();
  if (lastRow < 2) {
    ui.alert(
      'No hay IDs para buscar',
      'Por favor, pega los IDs en la columna A (desde la fila 2)',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const rangoIDs = hojaBusqueda.getRange(2, 1, lastRow - 1, 1);
  const datosIDs = rangoIDs.getValues();
  
  // Normalizar y filtrar IDs vacíos
  const idsBuscar = datosIDs
    .map(fila => normalizarID(fila[0]))
    .filter(id => id !== '');
  
  if (idsBuscar.length === 0) {
    ui.alert(
      'No hay IDs válidos',
      'Asegúrate de que la columna A tenga IDs válidos',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Confirmar búsqueda
  const respuesta = ui.alert(
    'Confirmar Búsqueda',
    `Se buscarán ${idsBuscar.length} IDs en todas las hojas de esta planilla.\n\n¿Continuar?`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuesta !== ui.Button.OK) return;
  
  // Iniciar búsqueda
  ss.toast('Buscando...', 'Procesando', -1);
  
  const resultados = buscarEnTodasLasHojas(ss, idsBuscar, hojaBusqueda);
  
  // Escribir resultados
  escribirResultados(hojaBusqueda, resultados, idsBuscar);
  
  ss.toast('');
  
  // Mostrar resumen
  const encontrados = resultados.filter(r => r.encontrado).length;
  const noEncontrados = idsBuscar.length - encontrados;
  
  ui.alert(
    'Búsqueda Completada',
    `Resultados:\n\n` +
    `✅ Encontrados: ${encontrados}\n` +
    `❌ No encontrados: ${noEncontrados}\n\n` +
    `Los resultados están en la hoja "Búsqueda"`,
    ui.ButtonSet.OK
  );
}

/**
 * Normaliza un ID eliminando apóstrofos, comillas y espacios.
 */
function normalizarID(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  
  return String(valor)
    .replace(/^['"]|['"]$/g, '')  // Elimina ' o " al inicio/final
    .trim();
}

/**
 * Busca los IDs en todas las hojas de la planilla (excepto "Búsqueda").
 */
function buscarEnTodasLasHojas(ss, idsBuscar, hojaBusqueda) {
  const hojas = ss.getSheets();
  const resultados = [];
  
  // Crear un mapa para búsqueda rápida
  const mapaIDs = new Set(idsBuscar);
  
  // Iterar por cada hoja
  hojas.forEach(hoja => {
    const nombreHoja = hoja.getName();
    
    // Saltar la hoja de búsqueda
    if (nombreHoja === "Búsqueda") return;
    
    Logger.log(`Buscando en hoja: ${nombreHoja}`);
    
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return; // Hoja vacía o solo encabezado
    
    try {
      // Leer columna A (asumimos que el ID está en columna A)
      const datos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
      
      datos.forEach((fila, index) => {
        const idEncontrado = normalizarID(fila[0]);
        
        if (idEncontrado !== '' && mapaIDs.has(idEncontrado)) {
          const numeroFila = index + 2;
          
          resultados.push({
            idBuscado: idEncontrado,
            encontrado: true,
            hoja: nombreHoja,
            fila: numeroFila,
            url: `#gid=${hoja.getSheetId()}&range=A${numeroFila}`
          });
        }
      });
    } catch (e) {
      Logger.log(`Error en hoja ${nombreHoja}: ${e.message}`);
    }
  });
  
  return resultados;
}

/**
 * Escribe los resultados en la hoja de búsqueda.
 * MODIFICADO: Los duplicados se muestran en la misma fila concatenados.
 */
function escribirResultados(hojaBusqueda, resultados, idsBuscar) {
  // Limpiar resultados anteriores (columnas B-E)
  const lastRow = hojaBusqueda.getLastRow();
  if (lastRow > 1) {
    hojaBusqueda.getRange(2, 2, lastRow - 1, 4).clearContent();
    hojaBusqueda.getRange(2, 2, lastRow - 1, 4).setBackground(null);
  }
  
  // Crear un mapa de resultados por ID
  const mapaResultados = {};
  resultados.forEach(r => {
    if (!mapaResultados[r.idBuscado]) {
      mapaResultados[r.idBuscado] = [];
    }
    mapaResultados[r.idBuscado].push(r);
  });
  
  // Preparar datos para escribir en batch
  const datosAEscribir = [];
  const formulas = [];
  const backgrounds = [];
  
  idsBuscar.forEach((id, index) => {
    const resultadosID = mapaResultados[id] || [];
    const filaActual = index + 2;
    
    if (resultadosID.length === 0) {
      // No encontrado
      datosAEscribir.push([
        "❌ No",
        "",
        "",
        ""
      ]);
      formulas.push("");
      backgrounds.push("#ffebee");
    } else {
      // Encontrado - concatenar todas las ubicaciones
      const ubicaciones = resultadosID.map(r => `${r.hoja} (Fila ${r.fila})`).join("\n");
      const cantidad = resultadosID.length;
      
      // Crear fórmula de hyperlinks múltiples
      let formulaLinks = "";
      if (resultadosID.length === 1) {
        formulaLinks = `=HYPERLINK("${resultadosID[0].url}"; "Ver →")`;
      } else {
        // Si hay múltiples, crear links separados por saltos de línea
        const links = resultadosID.map((r, i) => 
          `HYPERLINK("${r.url}"; "Ver ${i + 1}")`
        ).join("; CHAR(10); ");
        formulaLinks = `=${links}`;
      }
      
      datosAEscribir.push([
        cantidad === 1 ? "✅ Sí" : `✅ Sí (${cantidad}x)`,
        ubicaciones,
        cantidad,
        "" // Se llenará con la fórmula
      ]);
      formulas.push(formulaLinks);
      backgrounds.push("#e8f5e9");
    }
  });
  
  // Escribir todos los datos de una vez
  if (datosAEscribir.length > 0) {
    hojaBusqueda.getRange(2, 2, datosAEscribir.length, 4).setValues(datosAEscribir);
    
    // Aplicar fórmulas de links
    formulas.forEach((formula, index) => {
      if (formula) {
        hojaBusqueda.getRange(index + 2, 5).setFormula(formula);
      }
    });
    
    // Aplicar colores de fondo
    backgrounds.forEach((color, index) => {
      hojaBusqueda.getRange(index + 2, 2).setBackground(color);
    });
    
    // Formato final
    hojaBusqueda.getRange(2, 2, datosAEscribir.length, 1).setHorizontalAlignment("center");
    hojaBusqueda.getRange(2, 4, datosAEscribir.length, 1).setHorizontalAlignment("center");
    
    // Habilitar text wrapping para la columna de ubicaciones
    hojaBusqueda.getRange(2, 3, datosAEscribir.length, 1).setWrap(true);
  }
}

/**
 * Limpia los resultados de búsqueda (columnas B-E).
 */
function limpiarResultados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const hojaBusqueda = ss.getSheetByName("Búsqueda");
  if (!hojaBusqueda) {
    ui.alert('La hoja "Búsqueda" no existe.');
    return;
  }
  
  const respuesta = ui.alert(
    'Confirmar Limpieza',
    '¿Deseas limpiar los resultados de la búsqueda?\n\n' +
    '(Los IDs en la columna A se mantendrán)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuesta !== ui.Button.OK) return;
  
  const lastRow = hojaBusqueda.getLastRow();
  if (lastRow > 1) {
    hojaBusqueda.getRange(2, 2, lastRow - 1, 4).clearContent();
    hojaBusqueda.getRange(2, 2, lastRow - 1, 4).setBackground(null);
  }
  
  ui.alert('Resultados limpiados correctamente.');
}