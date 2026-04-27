// Distribucion.gs

/**
 * Función principal que se asigna al botón para distribuir los casos. Procesa la distribución de casos desde Asignación de Cupos a las planillas de especialidad, en lotes controlados para evitar timeouts.
 */
function distribuirCasosPorEspecialidad() {
  const ui = SpreadsheetApp.getUi();

  // --- CONFIGURACIÓN DEL LOTE ---
  const BATCH_SIZE = 100; // Número de casos a procesar por ejecución

  const respuesta = ui.alert('Confirmar Distribución',
                           `Este proceso distribuirá hasta ${BATCH_SIZE} casos a sus planillas por especialidad y los marcará como 'Distribuido' en esta hoja. ¿Deseas continuar?`,
                           ui.ButtonSet.OK_CANCEL);
  if (respuesta !== ui.Button.OK) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Operación cancelada.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Iniciando distribución (lote de ${BATCH_SIZE})...`, 'Procesando', -1);
  Logger.log(`--- INICIO DISTRIBUCIÓN LOTE ${BATCH_SIZE} ---`);

  try {
    const hojaFuente = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Casos y Cupos");
    const ultimaFilaConDatos = hojaFuente.getLastRow();

    // Si solo está el encabezado (o la hoja está vacía), no hay nada que procesar.
    if (ultimaFilaConDatos < 2) {
      SpreadsheetApp.getActiveSpreadsheet().toast(''); // Limpia el mensaje de "Procesando..."
      ui.alert('No hay datos', 'No se encontraron casos para distribuir en la planilla.', ui.ButtonSet.OK);
      return;
    }

    // Leemos solo las columnas necesarias para la decisión y los datos a mover
    const rangoLectura = hojaFuente.getRange(2, 1, ultimaFilaConDatos - 1, hojaFuente.getLastColumn());
    const todosLosValores = rangoLectura.getValues();

    const casosParaDistribuir = [];

    // 1. Filtrar primero: Encontrar todos los casos que SÍ necesitan moverse
    todosLosValores.forEach((fila, index) => {
      const numeroFilaReal = index + 2; // Número de fila original en la hoja
      const especialidadCorregida = fila[COLUMNAS_CUPOS.ESPECIALIDAD_CORREGIDA - 1];
      const prestacionOriginal = fila[COLUMNAS_CUPOS.PRESTA_EST - 1];

      // Leer la columna de estado para no reprocesar
      const estadoActual = fila[COLUMNAS_CUPOS.ESTADO_DISTRIBUCION - 1];

      let especialidadFinal = "";

      if (especialidadCorregida && especialidadCorregida.toString().trim() !== "") {
        especialidadFinal = especialidadCorregida.toString().trim();
      } else if (prestacionOriginal && prestacionOriginal.toString().trim() !== "") {
        especialidadFinal = homologarEspecialidad(prestacionOriginal.toString());
      }

      // Si encontramos una especialidad válida y mapeada, lo añadimos a la lista
      // Solo procesa las filas que el usuario seleccionó con "Enviar" en el menú desplegable de la columna BI.
      if (estadoActual === "Enviar" && especialidadFinal && CONFIG_ESPECIALIDADES[especialidadFinal]) {
        casosParaDistribuir.push({
          especialidad: especialidadFinal,
          datosFila: fila,
          numeroFilaOriginal: numeroFilaReal
        });
      }
    });

    const totalPendientes = casosParaDistribuir.length;
    if (totalPendientes === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('');
      ui.alert("Nada que Distribuir", "No se encontraron casos marcados como \"Enviar\".\n\nPor favor, seleccione \"Enviar\" en la columna BI de los casos que desea distribuir.", ui.ButtonSet.OK);
      return;
    }

    // 2. Tomar el lote actual
    const loteActual = casosParaDistribuir.slice(0, BATCH_SIZE);
    Logger.log(`Identificados ${totalPendientes} casos para distribuir. Procesando ${loteActual.length} en este lote.`);

    // 3. Agrupar el lote por especialidad
    const casosAgrupados = {};

    // Renombrado de 'filasParaEliminar' a 'filasParaMarcar'
    const filasParaMarcar = [];
    loteActual.forEach(caso => {
      if (!casosAgrupados[caso.especialidad]) {
        casosAgrupados[caso.especialidad] = [];
      }
      casosAgrupados[caso.especialidad].push(caso.datosFila);
      filasParaMarcar.push(caso.numeroFilaOriginal); // Guardamos la fila original para MARCARLA después
    });

    // 4. Escribir lotes en planillas de destino
    Logger.log("Iniciando escritura en planillas de especialidad...");
    let countEspecialidades = 0;
    for (const especialidad in casosAgrupados) {
      countEspecialidades++;
      Logger.log(`(${countEspecialidades}/${Object.keys(casosAgrupados).length}) Escribiendo ${casosAgrupados[especialidad].length} casos para ${especialidad}...`);
      try {
          const configDestino = CONFIG_ESPECIALIDADES[especialidad];
          const datosParaMover = casosAgrupados[especialidad];
          const hojaDestino = SpreadsheetApp.openById(configDestino.id).getSheetByName(configDestino.nombreHoja);
          const filaInicio = getPrimeraFilaVacia(hojaDestino);
          hojaDestino.getRange(filaInicio, 1, datosParaMover.length, datosParaMover[0].length).setValues(datosParaMover);
      } catch (writeError) {
          Logger.log(`Error escribiendo para ${especialidad}: ${writeError}. Omitiendo.`);
      }
    }
    Logger.log("Escritura completada.");

    // Reemplazo del bloque de Borrado por uno de Marcado
    // 5. Marcado
    Logger.log(`Iniciando marcado de ${filasParaMarcar.length} filas procesadas...`);
    filasParaMarcar.forEach((numFila, index) => {
      try {
        hojaFuente.getRange(numFila, COLUMNAS_CUPOS.ESTADO_DISTRIBUCION).setValue("Distribuido");
        // Loguear progreso cada 10 filas marcadas
        if ((index + 1) % 10 === 0) {
          Logger.log(`   ... ${index + 1} filas marcadas.`);
        }
      } catch (markError) {
         Logger.log(`Error marcando fila ${numFila}: ${markError}.`);
      }
    });
    Logger.log("Marcado completado.");

    /* CÓDIGO ANTIGUO (COMENTADO PARA REFERENCIA)
     // 5. Borrado
     // Ordenamos las filas a eliminar en orden descendente para evitar problemas con los índices.
     Logger.log(`Iniciando borrado de ${filasParaEliminar.length} filas procesadas...`);
     filasParaEliminar.sort((a, b) => b - a); // Orden descendente es crucial
     filasParaEliminar.forEach((numFila, index) => {
       hojaFuente.deleteRow(numFila);
       // Loguear progreso cada 10 filas borradas
       if ((index + 1) % 10 === 0) {
         Logger.log(`   ... ${index + 1} filas borradas.`);
       }
     });
     Logger.log("Borrado completado.");
    */

    // 6. Informar resultado
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    const restantes = totalPendientes - loteActual.length;
    let mensajeFinal = `Se distribuyeron ${loteActual.length} casos en este lote.`;
    if (restantes > 0) {
      mensajeFinal += `\nQuedan ${restantes} casos pendientes y marcados como "Enviar". Vuelve a ejecutar la función para procesar el siguiente lote.`;
    } else {
      mensajeFinal += `\n¡Distribución completada! No quedan casos pendientes y marcados como "Enviar.`;
    }
    ui.alert('Lote Procesado', mensajeFinal, ui.ButtonSet.OK);
    Logger.log(`--- FIN DISTRIBUCIÓN LOTE ---`);

  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('');
    Logger.log(`Error en distribuirCasosPorEspecialidad: ${e.toString()}\n${e.stack}`);
    ui.alert(`Ocurrió un error durante la distribución. Revisa los registros. Detalles: ${e.message}`);
  }
}

/**
 * Intenta encontrar una especialidad oficial a partir de un texto de prestación.
 * @param {string} textoPrestacion El texto de la columna PRESTA_EST.
 * @returns {string} La especialidad oficial encontrada, o una cadena vacía si no hay coincidencia.
 */
function homologarEspecialidad(textoPrestacion) {
  const textoNormalizado = normalizarTexto(textoPrestacion);
  if (!textoNormalizado) return "";

  for (const especialidadOficial in MAPA_HOMOLOGACION) {
    const palabrasClave = MAPA_HOMOLOGACION[especialidadOficial];
    for (const clave of palabrasClave) {
      if (textoNormalizado.includes(clave)) {
        return especialidadOficial; // Devuelve la especialidad oficial al encontrar la primera coincidencia
      }
    }
  }
  return ""; // No se encontró ninguna coincidencia
}
