// function procesarLiquidaciones() { 
//   const ss = SpreadsheetApp.getActiveSpreadsheet(); 
//   const hojaActiva = ss.getActiveSheet(); 
//   const hojasValidas = ["Liquidación Fortalecimiento", "Liquidación Autonomia Joven", "Liquidación Operadores de Calle"]; 
  
//   if (hojasValidas.includes(hojaActiva.getName())) { 
//       const hojaFormulario = ss.getSheetByName("Formulario"); 
//       const hojaProximosVencimientos = ss.getSheetByName('Proximos Vencimientos'); 
//       const hojaVencimientosbaja = ss.getSheetByName('Vencimientos baja'); 
//       const titulo = hojaActiva.getRange("A1").getValue().trim().toLowerCase(); 
//       const programa = identificarPrograma(titulo); 
//       const hojaPrograma = ss.getSheetByName("Certificación " + programa); 
//       const retorna = incorporarColumnasRetroactivos(hojaActiva, hojaPrograma); 
//       const hayRetroactivos = retorna.respuesta; 
//       const columnasOrigen = retorna.columnasOrigen; 
//       let soloPeriodo = null; 
//       let fechaRetroactiva = null; 
      
//       if (hayRetroactivos === "yes") { 
//           soloPeriodo = Browser.msgBox("¿Se contempla los meses retroactivos en el periodo?", Browser.Buttons.YES_NO); 
//           if (soloPeriodo === "no") { 
//               fechaRetroactiva = obtenerMesRetroactivo();
//           } 
//       } 
      
//       const {totalColumna5o8, totalColumna6o9, totalColumna7o10} = procesarFilas(hojaActiva, hojaPrograma, columnasOrigen, hojaFormulario, hojaProximosVencimientos, hojaVencimientosbaja, programa, hayRetroactivos, soloPeriodo, fechaRetroactiva); 
      
//       copiarFormatos(hojaActiva); 
//       agregarFilaTotales(hojaActiva, totalColumna5o8, totalColumna6o9, totalColumna7o10, hayRetroactivos); 
//       agregarDatosIoma(ss, programa, totalColumna5o8, totalColumna6o9, totalColumna7o10); 
//   } else { 
//       Logger.log("No puede ejecutarse la funcion, ya que no es una hoja valida"); 
//   } 
// }
// function obtenerMesRetroactivo() { 
//   const ui = SpreadsheetApp.getUi(); 
//   const respuesta = ui.prompt("¿A partir de qué mes corresponde el aumento? (Formato: aaaa-mm-dd)").getResponseText(); 
//   // Verificar si la fecha es válida en el formato aaaa-mm-dd 
//   if (!/^\d{4}-\d{2}-\d{2}$/.test(respuesta)) { 
//       ui.alert("La entrada no es una fecha válida. Por favor, ingrese la fecha en el formato aaaa-mm-dd."); 
//       return null; 
//   } 
//   // Convertir la fecha manualmente para evitar problemas de zona horaria 
//   const [year, month, day] = respuesta.split('-').map(Number); 
//   const fechaRetroactiva = new Date(Date.UTC(year, month - 1, day)); 
//   //Logger.log(`Corresponde al mes retroactivo: ${fechaRetroactiva.toISOString().split('T')[0]}`); 
//   return fechaRetroactiva; 
// }

// function procesarFilas(hojaActiva, hojaPrograma, columnasOrigen, hojaFormulario, hojaProximosVencimientos, hojaVencimientosbaja, programa, respuestaUsuario, soloPeriodo, fechaRetroactiva) { 
//   const data = hojaPrograma.getDataRange().getValues(); 
//   const currentDate = new Date(); 
//   let totalColumna5o8 = 0; 
//   let totalColumna6o9 = 0; 
//   let totalColumna7o10 = 0; 
//   let idsProximosVencimientos = new Set(); 
//   let idsVencimientosBaja = new Set(); 
//   const dataProximosVencimientos = hojaProximosVencimientos.getDataRange().getValues(); 
//   const dataVencimientosBaja = hojaVencimientosbaja.getDataRange().getValues(); 
//   dataProximosVencimientos.forEach(row => { 
//       const idUnicoPersona = row[1] + "-" + row[2]; 
//       idsProximosVencimientos.add(idUnicoPersona); 
//   }); 
//   dataVencimientosBaja.forEach(row => { 
//       const idUnicoPersona = row[1] + "-" + row[2]; 
//       idsVencimientosBaja.add(idUnicoPersona); 
//   }); 
//   data.slice(3).forEach((row, i) => { 
//       const columnaK = row[10].toString().toLowerCase(); 
//       const fechaColumnaD = new Date(row[3]); 
//       const diferenciaMeses = (currentDate.getFullYear() - fechaColumnaD.getFullYear()) * 12 + (currentDate.getMonth() - fechaColumnaD.getMonth()); 
//       let filaDatos = columnasOrigen.map((col, idx) => row[col]); 
//       const idUnicoPersona = row[1] + "-" + row[2]; 
//       console.log("filaDatos inicial: ", filaDatos); 
//       if (columnaK.includes('baja')) { 
//           if (diferenciaMeses === 11 && !idsVencimientosBaja.has(idUnicoPersona)) { 
//               agregarFilaAVencimientos(hojaVencimientosbaja, filaDatos, row, programa, idUnicoPersona, idsVencimientosBaja); 
//           } else if (diferenciaMeses !== 11 && idsVencimientosBaja.has(idUnicoPersona)) { 
//               eliminarFila(hojaVencimientosbaja, idUnicoPersona); 
//           } 
//       } else { 
//           if (diferenciaMeses === 11 && !idsProximosVencimientos.has(idUnicoPersona)) { 
//               agregarFilaAVencimientos(hojaProximosVencimientos, filaDatos, row, programa, idUnicoPersona, idsProximosVencimientos); 
//           } else if (diferenciaMeses !== 11 && idsProximosVencimientos.has(idUnicoPersona)) { 
//               eliminarFila(hojaProximosVencimientos, idUnicoPersona); 
//           } 
//           console.log("filaDatos antes de procesarMonto: ", filaDatos); 
//           ({fila: filaDatos, totalColumna5o8, totalColumna6o9, totalColumna7o10} = procesarMonto(filaDatos, row, hojaFormulario, respuestaUsuario, soloPeriodo, totalColumna5o8, totalColumna6o9, totalColumna7o10, fechaRetroactiva)); 
//           console.log("filaDatos después de procesarMonto: ", filaDatos) 
//           hojaActiva.appendRow(filaDatos); 
//       } 
//   }); 
//   return {totalColumna5o8, totalColumna6o9, totalColumna7o10}; 
// }

// function agregarFilaAVencimientos(hoja, fila, row, programa, idPersona, ids) { 
//   fila[2] = row[2]; 
//   fila[3] = row[3]; 
//   fila[4] = row[4]; 
//   fila[5] = programa; 
//   hoja.appendRow(fila); 
//   ids.add(idPersona); 
// }

// function procesarMonto(filaHojaLiquidacion, row, hojaFormulario, hayRetroactivo, soloPeriodo, totalColumna5o8, totalColumna6o9, totalColumna7o10, fechaRetroactiva) { 
//   let montoALiquidar; 
//   if (hayRetroactivo === "no") { 
//       montoALiquidar = calcularMontoSinRetroactivo(row, hojaFormulario); 
//       filaHojaLiquidacion[4] = montoALiquidar.monto; 
//       filaHojaLiquidacion[5] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2)); 
//       filaHojaLiquidacion[6] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2)); 
//       filaHojaLiquidacion[7] = parseFloat((montoALiquidar.monto - filaHojaLiquidacion[5]).toFixed(2)); 
//       filaHojaLiquidacion[9] = montoALiquidar.detalles; 
//       totalColumna5o8 += filaHojaLiquidacion[5]; 
//       totalColumna6o9 += filaHojaLiquidacion[6]; 
//       totalColumna7o10 += filaHojaLiquidacion[7]; 
//   } else if (hayRetroactivo === "yes") { 
//       filaHojaLiquidacion[4] = buscarValorBecaAnterior(filaHojaLiquidacion[3], hojaFormulario); 
//       filaHojaLiquidacion[5] = row[5]; 
//       filaHojaLiquidacion[6] = filaHojaLiquidacion[5] - filaHojaLiquidacion[4]; 
//       if (soloPeriodo === "yes") { 
//           montoALiquidar = calcularMontoRetroactivoSoloPeriodo(row, hojaFormulario, filaHojaLiquidacion[6]); 
//           filaHojaLiquidacion[7] = montoALiquidar.monto; 
//       } else if (soloPeriodo === "no") { 
//           montoALiquidar = calcularRetroactivoCompleto(row, fechaRetroactiva, filaHojaLiquidacion[6], hojaFormulario); 
//           filaHojaLiquidacion[7] = montoALiquidar.monto; 
//       } 
//       filaHojaLiquidacion[8] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2)); 
//       filaHojaLiquidacion[9] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2)); 
//       filaHojaLiquidacion[10] = parseFloat((montoALiquidar.monto - filaHojaLiquidacion[9]).toFixed(2)); 
//       filaHojaLiquidacion[12] = montoALiquidar.detalles 
//       totalColumna5o8 += filaHojaLiquidacion[8]; 
//       totalColumna6o9 += filaHojaLiquidacion[9]; 
//       totalColumna7o10 += filaHojaLiquidacion[10]; 
//   } 
//   return {fila: filaHojaLiquidacion, totalColumna5o8, totalColumna6o9, totalColumna7o10}; 
// }
// function agregarFilaTotales(hoja, totalColumna5o8, totalColumna6o9, totalColumna7o10, respuesta) {
//   const ultimaFila = hoja.getLastRow() + 1;
//   hoja.getRange(ultimaFila, 1).setValue("TOTAL");
//   hoja.getRange(ultimaFila, 1).setFontWeight("bold").setFontSize(11);

//   if (respuesta === "no") {
//     hoja.getRange(ultimaFila, 1, 1, 5).mergeAcross();
//     hoja.getRange(ultimaFila, 6).setValue(totalColumna5o8);
//     hoja.getRange(ultimaFila, 7).setValue(totalColumna6o9);
//     hoja.getRange(ultimaFila, 8).setValue(totalColumna7o10);
//     hoja.getRange(ultimaFila, 6).setFontWeight("bold").setFontSize(11);
//     hoja.getRange(ultimaFila, 7).setFontWeight("bold").setFontSize(11);
//     hoja.getRange(ultimaFila, 8).setFontWeight("bold").setFontSize(11);
//   } else if (respuesta === "yes") {
//     hoja.getRange(ultimaFila, 1, 1, 8).mergeAcross();
//     hoja.getRange(ultimaFila, 9).setValue(totalColumna5o8);
//     hoja.getRange(ultimaFila, 10).setValue(totalColumna6o9);
//     hoja.getRange(ultimaFila, 11).setValue(totalColumna7o10);
//     hoja.getRange(ultimaFila, 9).setFontWeight("bold").setFontSize(11);
//     hoja.getRange(ultimaFila, 10).setFontWeight("bold").setFontSize(11);
//     hoja.getRange(ultimaFila, 11).setFontWeight("bold").setFontSize(11);
//   }
// }

// function agregarDatosIoma(ss, programa, totalColumna5o8, totalColumna6o9, totalColumna7o10) {
//   const hojaIoma = ss.getSheetByName("IOMA " + programa);
//   const numberFormat = '[$$]#,##0.00';
//   hojaIoma.getRange(4, 2).setValue(totalColumna7o10).setNumberFormat(numberFormat).setFontWeight("bold");
//   hojaIoma.getRange(4, 3).setValue(totalColumna6o9).setNumberFormat(numberFormat);
//   hojaIoma.getRange(4, 4).setValue(totalColumna6o9).setNumberFormat(numberFormat);
//   hojaIoma.getRange(4, 5).setValue(totalColumna6o9 + totalColumna5o8).setNumberFormat(numberFormat).setFontWeight("bold");
// }

// function eliminarFila(hoja, idUnicoPersona) {
//   const data = hoja.getDataRange().getValues();
//   for (let i = 0; i < data.length; i++) {
//     const id = data[i][1] + "-" + data[i][2];
//     if (id === idUnicoPersona) {
//       hoja.deleteRow(i + 1);
//       break;
//     }
//   }
// }

// function obtenerMesReferencia(fecha) {
//   const dia = fecha.getDate();
//   const mes = fecha.getMonth();
//   const anio = fecha.getFullYear();

//   if (dia >= 15) {
//     return new Date(anio, mes, 15);
//   } else {
//     return new Date(anio, mes - 1, 15);
//   }
// }

// function incorporarColumnasRetroactivos(hojaActiva ,hojaPrograma) {
//   const hayRetroactivos = Browser.msgBox("¿Hay retroactivos a calcular?", Browser.Buttons.YES_NO);
//   hojaActiva.getRange("A4:M" + hojaPrograma.getLastRow()).clearContent();

//   if (hayRetroactivos === "yes") {
//     const respuesta1 = Browser.msgBox("¿Debo agregar columnas?",Browser.Buttons.YES_NO);
//     if (respuesta1 === "yes") {
//       agregarColumnasRetroactivas(hojaActiva);
//     }
//     const columnasOrigen = [0, 1, 3, 4, 5];
//     return {columnasOrigen: columnasOrigen, respuesta: "yes"}
//   } else if (hayRetroactivos === "no") {
//     const respuesta2 = Browser.msgBox("¿Debo eliminar columnas?", Browser.Buttons.YES_NO);
//     if (respuesta2 === "yes") {
//       eliminarColumnas(hojaActiva, [5, 7, 8]);
//     }
//     hojaActiva.getRange("E3").setValue("Monto");
//     hojaActiva.getRange("A4:L" + hojaPrograma.getLastRow()).clearContent();
//     const columnasOrigen = [0, 1, 3, 4, 5];
//     return {columnasOrigen: columnasOrigen, respuesta: "no"}
//   } else {
//     Browser.msgBox("Respuesta no válida. Por favor, ingrese 'Si' o 'No'.");
//   }
// }

// function agregarColumnasRetroactivas(hoja) {
//   hoja.insertColumnBefore(5);
//   hoja.getRange("E3").setValue("Monto de la beca Anterior");
//   hoja.insertColumnBefore(7);
//   hoja.getRange("F3").setValue("Monto Actualizado por");
//   hoja.insertColumnBefore(8);
//   hoja.getRange("G3").setValue("Diferencia incremento");
//   hoja.getRange("H3").setValue("Incremento Diferencial Retroactivo");
// }

// function eliminarColumnas(hoja, columnas) {
//   for (let i = columnas.length - 1; i >= 0; i--) {
//     try {
//       hoja.deleteColumn(columnas[i]);
//     } catch (e) {
//       // Logger.log("La columna " + columnas[i] + " no existe.");
//     }
//   }
// }

// function buscarValorBecaAnterior(tipoBeca, hojaFormulario) {
//   const rangoTipos = hojaFormulario.getRange("G13:G17").getValues().flat();
//   const rangoDatos = hojaFormulario.getRange("I13:I17").getValues().flat();
//   const index = rangoTipos.findIndex(tipo => tipo.trim() === tipoBeca);
//   return index !== -1 ? rangoDatos[index] : "Tipo de beca no encontrado";
// }

// function convertirFecha(fechaTexto) {
//   const partes = fechaTexto.split('/');
//   if (partes.length === 3) {
//     const [dia, mes, anio] = partes.map(Number);
//     return new Date(anio, mes - 1, dia);
//   }
//   return new Date(fechaTexto);
// }

// function separarFechasPeriodo(filaCertificacion) {
//   const periodo = filaCertificacion[9];
//   const fechas = periodo.split(' al ');
//   if (fechas.length < 2) {
//     return 0;
//   }
//   const fechaInicio = convertirFecha(fechas[0].trim());
//   const fechaFin = convertirFecha(fechas[1].trim());

//   if (isNaN(fechaInicio) || isNaN(fechaFin)) {
//     return 0;
//   }
//   return { inicio: fechaInicio, fin: fechaFin }
// }

// function fechaMayor(fecha1, fecha2) {
//   let date1 = new Date(fecha1);
//   let date2 = new Date(fecha2);

//   if (isNaN(date2.getTime())) {
//     const [dia2, mes2, año2] = fecha2.split('/').map(Number);
//     date2 = new Date(año2, mes2 - 1, dia2);
//   }

//   if (isNaN(date1.getTime()) || isNaN(date2.getTime())) {
//     throw new Error("Una o ambas fechas no son válidas.");
//   }

//   return date1 >= date2 ? fecha1 : fecha2;
// }
// function calcularRetroactivoCompleto(filaCertificacion, dato, diferencia, hojaFormulario) {
//   const fecha = separarFechasPeriodo(filaCertificacion);
//   let diasHabiles;
//   let montoMensual = 0;
//   let montoTotal = 0;
//   let detallesCalculo = "";
//   let resultado = fechaMayor(dato, filaCertificacion[3]);
//   const mesRetroactivo = obtenerMesesEntreFechas(resultado, fecha.inicio);

//   mesRetroactivo.forEach(mesR => {
//     if (mesR.fin < fecha.inicio) {
//       diasHabiles = calcularDiasHabiles(mesR.inicio, mesR.fin);
//       if (diasHabiles > 0) {
//         montoMensual = (diferencia / 24) * diasHabiles;
//         montoTotal += montoMensual;
//         detallesCalculo += `${mesR.nombre}: ${montoMensual.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
//       }
//     }
//   });

//   const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);
//   meses.forEach(mes => {
//     diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
//     if (diasHabiles > 0) {
//       const tipoBeca = filaCertificacion[4];
//       montoMensual = obtenerMontoMensual(hojaFormulario, tipoBeca, mes.nombre);
//       const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
//       montoTotal += montoMensualCalculado;
//       detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
//     }
//   });

//   return { monto: montoTotal, detalles: detallesCalculo };
// }

// function calcularMontoSinRetroactivo(filaCertificacion, dato) {
//   const fecha = separarFechasPeriodo(filaCertificacion);
//   let diasHabiles;
//   let montoMensual = 0;
//   let montoTotal = 0;
//   let detallesCalculo = "";
//   const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);

//   meses.forEach(mes => {
//     diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
//     const tipoBeca = filaCertificacion[4];
//     montoMensual = obtenerMontoMensual(dato, tipoBeca, mes.nombre);
//     const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
//     montoTotal += montoMensualCalculado;
//     detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
//   });

//   return { monto: montoTotal, detalles: detallesCalculo };
// }

// function calcularMontoRetroactivoSoloPeriodo(filaCertificacion, dato, diferencia) {
//   const fecha = separarFechasPeriodo(filaCertificacion);
//   let diasHabiles;
//   let montoMensual = 0;
//   let montoTotal = 0;
//   let detallesCalculo = "";
//   const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);

//   meses.forEach(mes => {
//     diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
//     montoMensual = diferencia;
//     const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
//     montoTotal += montoMensualCalculado;
//     detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
//   });

//   return { monto: montoTotal, detalles: detallesCalculo };
// }

// function obtenerMesesEntreFechas(fechaInicio, fechaFin) {
//   const meses = [];
//   let fechaActual = new Date(fechaInicio);

//   while (fechaActual <= fechaFin) {
//     let inicioMes = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), 1);
//     let finMes = new Date(fechaActual.getFullYear(), fechaActual.getMonth() + 1, 0);

//     if (inicioMes < fechaInicio) {
//       inicioMes = new Date(fechaInicio);
//     }
//     if (finMes > fechaFin) {
//       finMes = new Date(fechaFin);
//     }

//     meses.push({ inicio: inicioMes, fin: finMes, nombre: obtenerNombreMes(fechaActual.getMonth()) });
//     fechaActual.setMonth(fechaActual.getMonth() + 1);
//     fechaActual.setDate(1);
//   }

//   return meses;
// }

// function obtenerNombreMes(mes) {
//   const nombresMeses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
//   return nombresMeses[mes];
// }

// function calcularDiasHabiles(fechaInicio, fechaFin) {
//   let diasHabiles = 0;
//   const fechaActual = new Date(fechaInicio);

//   while (fechaActual <= fechaFin) {
//     const diaSemana = fechaActual.getDay();
//     if (diaSemana >= 1 && diaSemana <= 6) {
//       diasHabiles++;
//     }
//     fechaActual.setDate(fechaActual.getDate() + 1);
//   }
//   return Math.min(diasHabiles, 24);
// }

// function obtenerMontoMensual(hojaFormulario, tipoBeca, mes) {
//   const datosFormulario = hojaFormulario.getDataRange().getValues();
//   const monthNames = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
//   const mesIndex = monthNames.indexOf(mes.toLowerCase());

//   if (mesIndex === -1) {
//     return 0;
//   }

//   for (let i = 3; i < 8; i++) {
//     const tipoBecaActual = datosFormulario[i][6];
//     if (tipoBecaActual && tipoBecaActual.toLowerCase() === tipoBeca.toLowerCase()) {
//       for (let j = mesIndex; j >= 0; j--) {
//         const mesActual = monthNames[j];
//         for (let k = 7; k <= 14; k++) {
//           const mesCelda = datosFormulario[2][k];
//           if (mesCelda && mesCelda.toLowerCase() === mesActual) {
//             return datosFormulario[i][k];
//           }
//         }
//       }
//       for (let j = 14; j >= 7; j--) {
//         if (j == mesIndex) continue;
//         const mesAnterior = monthNames[j];
//         for (let k = 7; k <= 14; k++) {
//           const mesCelda = datosFormulario[2][k];
//           if (mesCelda && mesCelda.toLowerCase() === mesAnterior) {
//             return datosFormulario[i][k];
//           }
//         }
//       }
//     }
//   }
//   return 0;
// }

// function identificarPrograma(titulo) {
//   const programas = {
//     "fortalecimiento": "Fortalecimiento",
//     "operadores de calle": "Operadores de Calle",
//     "autonomia joven": "Autonomia Joven"
//   };
//   for (const key in programas) {
//     if (titulo.includes(key)) {
//       return programas[key];
//     }
//   }
//   return "Programa no identificado";
// }

// function buscarValoresBeca(tipoBeca, hojaFormulario) {
//   const rangoTipos = hojaFormulario.getRange("G4:G8").getValues().flat();
//   const rangoValoresAnteriores = hojaFormulario.getRange("I4:O8").getValues().flat();
//   const rangoValoresActuales = hojaFormulario.getRange("H4:H8").getValues().flat();
//   const index = rangoTipos.findIndex(tipo => tipo.trim() === tipoBeca);
//   if (index !== -1) {
//     return { montoAnterior: rangoValoresAnteriores[index], montoActual: rangoValoresActuales[index] };
//   }
//   return { montoAnterior: "Tipo de beca no encontrado", montoActual: "Tipo de beca no encontrado" };
// }