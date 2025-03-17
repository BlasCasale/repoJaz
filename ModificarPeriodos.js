function modificarHoja(hoja, periodo) {
  const lastRow = hoja.getLastRow();
  const range = hoja.getRange(`J4:J${lastRow}`);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== '-') {
      values[i][0] = periodo;
    }
  }

  range.setValues(values);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaIomaName = "IOMA " + hoja.getName().split(' ')[1];
  const hojaIoma = ss.getSheetByName(hojaIomaName);

  if (hojaIoma) {
    hojaIoma.getRange('A4:Z4').clearContent();
    hojaIoma.getRange('A4').setValue(periodo);
  } else {
    SpreadsheetApp.getUi().alert(`No se encontró la hoja ${hojaIomaName}.`);
  }
}


function modificarHojaActiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();
  const hojasValidas = [
    "Certificación Fortalecimiento",
    "Certificación Autonomia Joven",
    "Certificación Operadores de Calle"
  ];

  if (hojasValidas.includes(hojaActiva.getName())) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Ingrese el periodo (formato: DD/MM/YYYY al DD/MM/YYYY):").getResponseText().trim();

    if (!/^\d{2}\/\d{2}\/\d{4} al \d{2}\/\d{2}\/\d{4}$/.test(response)) {
      ui.alert("La entrada no es válida. Por favor, ingrese en el formato DD/MM/YYYY al DD/MM/YYYY.");
      return;
    }

    modificarHoja(hojaActiva, response);
  } else {
    SpreadsheetApp.getUi().alert("La función no puede ejecutarse, ya que no es una hoja válida.");
  }
}

function modificarTodasLasHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasValidas = [
    "Certificación Fortalecimiento",
    "Certificación Autonomia Joven",
    "Certificación Operadores de Calle"
  ];

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Ingrese el periodo (formato: DD/MM/YYYY al DD/MM/YYYY):").getResponseText().trim();

  if (!/^\d{2}\/\d{2}\/\d{4} al \d{2}\/\d{2}\/\d{4}$/.test(response)) {
    ui.alert("La entrada no es válida. Por favor, ingrese en el formato DD/MM/YYYY al DD/MM/YYYY.");
    return;
  }

  hojasValidas.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja) {
      modificarHoja(hoja, response);
    }
  });
}

