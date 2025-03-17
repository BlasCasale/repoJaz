/**
 * Abre el formulario en un sidebar.
 */
function mostrarFormulario() {
  const html = HtmlService.createHtmlOutputFromFile('modificarExpReso')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Busca registros por resolución o expediente.
 * @param {string} str El valor a buscar.
 * @return {Array} Un array de objetos con los registros encontrados.
 */
function buscarResolucion(str) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasValidas = [
    "Certificación Fortalecimiento",
    "Certificación Autonomia Joven",
    "Certificación Operadores de Calle"
  ];

  const registrosEncontrados = [];

  hojasValidas.forEach((nombreHoja) => {
    const hoja = ss.getSheetByName(nombreHoja);
    const columnaBusqueda = str.slice(0, 2) === 'EX' ? 9 : 8; // Columna I (9) o H (8)
    const rango = hoja.getRange(4, columnaBusqueda, hoja.getLastRow() - 3).getValues();

    rango.forEach((fila, index) => {
      if (fila[0] === str) {
        registrosEncontrados.push({
          nombre: hoja.getRange(index + 4, 1).getValue(),
          apellido: hoja.getRange(index + 4, 2).getValue(),
          resolucion: fila[0],
          hoja: nombreHoja,
          esExpediente: str.slice(0, 2) === 'EX'
        });
      }
    });
  });

  return registrosEncontrados;
}

/**
 * Modifica los registros seleccionados.
 * @param {Array} registros Un array de objetos con los registros a modificar.
 * @param {string} nuevoValor El nuevo valor para los registros.
 * @return {string} Un mensaje indicando el resultado de la operación.
 */
function modificarRegistrosSeleccionados(registros, nuevoValor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    registros.forEach((registro) => {
      const hoja = ss.getSheetByName(registro.hoja);
      if (!hoja) {
        console.error("No se encontró la hoja:", registro.hoja);
        return;
      }

      const datosNombre = hoja.getRange('A4:A').getValues().flat();
      const datosApellido = hoja.getRange('B4:B').getValues().flat();

      const index = datosNombre.findIndex((nombre, i) => nombre.trim() === registro.nombre && datosApellido[i].trim() === registro.apellido);

      if (index === -1) {
        console.error("No se encontró el registro:", registro);
        return;
      }

      const columnaModificar = registro.esExpediente ? 9 : 8;
      hoja.getRange(index + 4, columnaModificar).setValue(nuevoValor);

      console.log(`Registro modificado: ${registro.nombre} ${registro.apellido} en la hoja ${registro.hoja}, columna ${columnaModificar}`);
    });

    return "Registros modificados correctamente.";
  } catch (error) {
    console.error("Error en la función modificarRegistrosSeleccionados:", error);
    return "Error al modificar registros: " + error.message;
  }
}