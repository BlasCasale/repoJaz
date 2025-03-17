const moverDatos = (sheet, e, col, row, dato, rowMenor, rowMayor, colMenor, colMayor) => {
  if ((row >= rowMenor && row <= rowMayor) && (col >= colMenor && col <= colMayor)) {
    var oldValue;
    if (col != colMayor) oldValue = sheet.getRange(row, col + 1).getValue();
    if (oldValue !== undefined && (col >= colMenor && col <= colMayor)) {
      var nextCol = col + 1;
      if (row % 3 === 0) {
        if (nextCol <= colMayor) {
          sheet.getRange(row, nextCol).setValue(dato);
        }
      } else {
        var h4 = sheet.getRange('H4');
        var h5 = sheet.getRange('H5');
        var h7 = sheet.getRange('H7');
        var h8 = sheet.getRange('H8');

        if (e.range.getA1Notation() === h4.getA1Notation()) {
          if (nextCol <= colMayor) {
            h5.setValue(h4.getValue());
            sheet.getRange(row, nextCol).setValue(dato);
            sheet.getRange(row + 1, nextCol).setValue(dato);
          }
        } else if (e.range.getA1Notation() === h5.getA1Notation()) {
          if (nextCol <= colMayor) {
            h4.setValue(h5.getValue());
            sheet.getRange(row, nextCol).setValue(dato);
            sheet.getRange(row - 1, nextCol).setValue(dato);
          }
        } else if (e.range.getA1Notation() === h7.getA1Notation()) {
          if (nextCol <= colMayor) {
            h8.setValue(h7.getValue());
            sheet.getRange(row, nextCol).setValue(dato);
            sheet.getRange(row + 1, nextCol).setValue(dato);
          }
        } else if (e.range.getA1Notation() === h8.getA1Notation()) {
          if (nextCol <= colMayor) {
            h7.setValue(h8.getValue());
            sheet.getRange(row, nextCol).setValue(dato);
            sheet.getRange(row - 1, nextCol).setValue(dato);
          }
        }
      }
    }
    col += 1;
    if (col <= colMayor) moverDatos(sheet, e, col, row, oldValue, rowMenor, rowMayor, colMenor, colMayor);
  }
}

const actualizarRegistros = (beca, montoNuevo) => {
  const libro = SpreadsheetApp.getActiveSpreadsheet()
  const hojas = ['Certificación Fortalecimiento', 'Certificación Autonomia Joven', 'Certificación Operadores de Calle']

  hojas.forEach((hoja) => {
    const actual = libro.getSheetByName(hoja)

    const becados = actual.getRange('E4:E').getValues()
    const observaciones = actual.getRange('K4:K').getValues()
    Logger.log(observaciones)

    becados.forEach((becado, index) => {
      if (becado[0] === beca && !observaciones[index][0].toLowerCase().includes('baja') && !observaciones[index][0].toLowerCase().includes('limitación')) {
        actual.getRange(`F${index + 4}`).setValue(montoNuevo)
      }
    })
  })
}

function aplicarFormato(sheet) {
  const rango = sheet.getRange('H3:O8')
  rango.setNumberFormat('"$"#,##0.00')
}

function onEdit(e) {
  const hojaActiva = e.source.getActiveSheet();
  const rangoEditado = e.range;
  const row = rangoEditado.getRow();
  const col = rangoEditado.getColumn();

  // ----------------------- Lógica para la hoja "Formulario" -----------------------
  if (hojaActiva.getName() === "Formulario") {
    // Lógica de la primera función
    if (rangoEditado.getA1Notation() === "C18") {
      asignarValoresPorDefectoForm();
    }

    // Lógica de la segunda función
    moverDatos(hojaActiva, e, col, row, e.oldValue, 3, 9, 8, 15); // Desde H3:O3 y H8:O8

    if ((row == 4 || row == 7) && col == 8) {
      const tipo1 = hojaActiva.getRange(row, col - 1).getValue();
      const tipo2 = hojaActiva.getRange(row + 1, col - 1).getValue();
      actualizarRegistros(tipo1, e.value);
      actualizarRegistros(tipo2, e.value);
    } else if ((row == 5 || row == 8) && col == 8) {
      const tipo1 = hojaActiva.getRange(row, col - 1).getValue();
      const tipo2 = hojaActiva.getRange(row - 1, col - 1).getValue();
      actualizarRegistros(tipo1, e.value);
      actualizarRegistros(tipo2, e.value);
    } else if (row == 6 && col == 8) {
      const tipo = hojaActiva.getRange(row, col - 1).getValue();
      actualizarRegistros(tipo, e.value);
    }

    aplicarFormato(hojaActiva);
  }

  // ----------------------- Lógica para las hojas de certificación -----------------------
  const hojasCertificacion = ['Certificación Fortalecimiento', 'Certificación Autonomia Joven', 'Certificación Operadores de Calle'];
  if (hojasCertificacion.includes(hojaActiva.getName()) && col === 11) { // Columna K (número 11)
    const valorEditado = rangoEditado.getValue().toString().toLowerCase();

    // Verificar si incluye "baja" o "limitacion" con variantes
    if (valorEditado.includes("baja") || valorEditado.includes("limitacion") || 
        valorEditado.includes("limitación")) {
      
      // Aplicar formato gris a toda la fila desde la columna 1 hasta la columna K
      const rangoFila = hojaActiva.getRange(row, 1, 1, 11); // Desde columna 1 hasta 11
      rangoFila.setBackground("#d3d3d3"); // Formato gris
      
      // Editar la columna F con un 0
      hojaActiva.getRange(row, 6).setValue(0); // Columna F (número 6)

      // Agregar un guion en la columna J
      hojaActiva.getRange(row, 10).setValue("-"); // Columna J (número 10)
    }
  }
}
  // let proxVal
  // while ((row >= 3 && row <= 9) && (col >= 8 && col <= 15)) {
  //   proxVal = range.getRange(row, col + 1).getValue()
  //   var oldValue = e.oldValue
  //   if (oldValue !== undefined && (col >= 8 && col <= 15)) {
  //     var nextCol = col + 1;
  //     sheet.getRange(row, nextCol).setValue(oldValue);
  //   }
  //   col = col + 1
  // }
  // if ((row >= 3 && row <= 9) && (col >= 8 && col <= 15)) { // Rango H3:O9
  //   var oldValue = e.oldValue;
  //   if (oldValue !== undefined) {
  //     var nextCol = col + 1;
  //     sheet.getRange(row, nextCol).setValue(oldValue);
  //   }
  // }