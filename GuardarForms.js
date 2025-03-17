function Limpiar() {
  const hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  const formulario = hojaActiva.getSheetByName("Formulario");

  // Celdas a limpiar
  const celdasALimpiar = ["C6", "C8", "C10", "C12", "C14", "C16", "C18", "C20", "C22", "C24", "C26", "C28", "C30", "C32"];
  
  // Limpiar las celdas
  celdasALimpiar.forEach(celda => formulario.getRange(celda).clearContent());
}


function Guardar() {
  const hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  const formulario = hojaActiva.getSheetByName("Formulario");
  const programa = formulario.getRange("C16").getValue();

  // Determinar la hoja de destino según el programa
  const hojasProgramas = {
    "Fortalecimiento": "Certificación Fortalecimiento",
    "Autonomia Joven": "Certificación Autonomia Joven",
    "Operadores de Calle": "Certificación Operadores de Calle"
  };

  const datosProgramas = hojaActiva.getSheetByName(hojasProgramas[programa]);
  if (!datosProgramas) {
    SpreadsheetApp.getUi().alert("Programa no válido. Verifica la celda C16.");
    return;
  }

  const datosRegistro = hojaActiva.getSheetByName("Registros") || hojaActiva.insertSheet("Registros");

  // Obtener los datos a guardar
  const datosPrograma = [
    formulario.getRange("C10").getValue(),
    formulario.getRange("C14").getValue(),
    formulario.getRange("C18").getValue(),
    formulario.getRange("C20").getValue(),
    formulario.getRange("C22").getValue(),
    formulario.getRange("C24").getValue(),
    formulario.getRange("C26").getValue()
  ];

  const nombre = formulario.getRange("C6").getValue();
  const apellido = formulario.getRange("C8").getValue();
  const mail = formulario.getRange("C12").getValue();
  const formato = "dd/MM/yyyy";
  const periodoInicio = new Date(formulario.getRange("C28").getValue());
  const periodoFin = new Date(formulario.getRange("C30").getValue());
  const periodo = Utilities.formatDate(periodoInicio, Session.getScriptTimeZone(), formato) + 
                  " - " + 
                  Utilities.formatDate(periodoFin, Session.getScriptTimeZone(), formato);
  const observaciones = formulario.getRange("C32").getValue();

  // Inyección de datos a la hoja de destino
  datosProgramas.appendRow([nombre, apellido, ...datosPrograma, periodo, observaciones]);
  datosRegistro.appendRow([nombre, apellido, programa, mail]);

  Limpiar(); // Ejecución de función para limpieza de celdas
}


function asignarValoresPorDefectoForm() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formulario");
  const tipoBeca = hoja.getRange("C18").getValue();
  const rangoBecas = hoja.getRange("G4:H8").getValues();  // Obtener el rango completo que incluye las columnas G y H

  let montoBeca = null;

  // Buscar el tipo de beca en el rango G4:G8
  for (let i = 0; i < rangoBecas.length; i++) {
    if (rangoBecas[i][0] === tipoBeca) {  // Verificar si el tipo de beca coincide
      montoBeca = rangoBecas[i][1];  // Obtener el monto de la columna H
      break;
    }
  }

  if (montoBeca !== null) {
    // Asignar el monto a la celda C20
    hoja.getRange("C20").setValue(montoBeca);
    // Escribir "Completa" en la celda C22
    hoja.getRange("C22").setValue("Completa");
  } else {
    Logger.log("Tipo de beca no encontrado en el rango especificado.");
  }
}
