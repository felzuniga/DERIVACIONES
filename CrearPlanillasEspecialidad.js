// CrearPlanillasEspecialidad.gs

/**
 * Lee la lista de especialidades y crea una planilla para cada una usando la plantilla definida en Configuracion.gs.
 */
function crearPlanillasEspecialidad() {
  const ui = SpreadsheetApp.getUi(); // Necesario para las alertas
  const respuesta = ui.alert('Confirmar Creación',
                           'Este proceso creará una planilla para cada especialidad listada en la hoja "Configuración" (desde A2), usando la plantilla especificada.\n\n¿Deseas continuar?',
                           ui.ButtonSet.OK_CANCEL);
  if (respuesta !== ui.Button.OK) {
      Logger.log("Creación cancelada por el usuario.");
      return; // Detiene si el usuario cancela
  }

  Logger.log("--- INICIO CREACIÓN PLANILLAS ESPECIALIDAD ---");

  try {
    const ssAsignacion = SpreadsheetApp.openById(IDs.asignacionCupos);
    const hojaConfig = ssAsignacion.getSheetByName("Configuración");
    const carpetaDestino = DriveApp.getFolderById(IDs.carpetaEspecialidades);
    const plantillaFile = DriveApp.getFileById(IDs.plantillaEspecialidad);

    if (!hojaConfig) throw new Error('No se encontró la hoja "Configuración".');
    if (!carpetaDestino) throw new Error(`Carpeta destino no encontrada (ID: ${IDs.carpetaEspecialidades}).`);
    if (!plantillaFile) throw new Error(`Plantilla no encontrada (ID: ${IDs.plantillaEspecialidad}).`);

    const ultimaFila = hojaConfig.getLastRow();
    if (ultimaFila < 2) { // Si no hay datos después del encabezado
        Logger.log('No se encontraron nombres de especialidad en la hoja "Configuración" (desde A2).');
        ui.alert('No hay datos', 'No se encontraron nombres de especialidad en la hoja "Configuración" a partir de la fila 2.', ui.ButtonSet.OK);
        return;
    }
    const listaEspecialidades = hojaConfig.getRange("A2:A" + ultimaFila).getValues()
                                   .flat()
                                   .filter(String)
                                   .map(nombre => nombre.trim());

    if (listaEspecialidades.length === 0) {
      Logger.log('La lista de especialidades (desde A2) está vacía.');
      ui.alert('Lista Vacía', 'No se encontraron nombres de especialidad válidos a partir de la fila 2.', ui.ButtonSet.OK);
      return;
    }

    Logger.log(`Se encontraron ${listaEspecialidades.length} especialidades (desde A2).`);

    let creadas = 0;
    let omitidas = 0;

    listaEspecialidades.forEach(nombreEspecialidad => {
      const nombreArchivo = `${nombreEspecialidad}`;
      const archivosExistentes = carpetaDestino.getFilesByName(nombreArchivo);
      if (archivosExistentes.hasNext()) {
        Logger.log(`   -> Omitido: Ya existe "${nombreArchivo}".`);
        omitidas++;
      } else {
        try {
          plantillaFile.makeCopy(nombreArchivo, carpetaDestino);
          Logger.log(`   -> CREADO: "${nombreArchivo}"`);
          creadas++;
        } catch (copyError) {
          Logger.log(`   -> ERROR al crear "${nombreArchivo}": ${copyError.message}`);
        }
      }
    });

    Logger.log(`--- FIN CREACIÓN PLANILLAS --- Creadas: ${creadas}, Omitidas: ${omitidas}.`);
    ui.alert('Proceso Completado', `Se crearon ${creadas} nuevas planillas de especialidad.\nSe omitieron ${omitidas} porque ya existían.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`Error fatal en crearPlanillasEspecialidad: ${e.toString()}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`Ocurrió un error: ${e.message}`);
  }
}
