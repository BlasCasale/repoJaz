function procesarLiquidaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();
  const hojasValidas = ["Liquidación Fortalecimiento", "Liquidación Autonomia Joven", "Liquidación Operadores de Calle"];

  Logger.log("Nombre de la hoja activa: " + hojaActiva.getName());

  if (hojasValidas.includes(hojaActiva.getName())) {
    const hojaFormulario = ss.getSheetByName("Formulario");
    const hojaProximosVencimientos = ss.getSheetByName('Proximos Vencimientos');
    const titulo = hojaActiva.getRange("A1").getValue().trim().toLowerCase();
    const programa = identificarPrograma(titulo);
    const hojaCertificacion = ss.getSheetByName("Certificación " + programa);
    const hojaIOMA = ss.getSheetByName("IOMA " + programa);

    const periodoTotal = hojaIOMA.getRange("A4").getValue().trim();

    Logger.log("Periodo total: " + periodoTotal);

    const confirmacion = Browser.msgBox("¿Es " + periodoTotal + " el periodo total a liquidar?", Browser.Buttons.OK_CANCEL);
    if (confirmacion === "cancel") {
      Logger.log("Confirmación de periodo total cancelada");
      Browser.msgBox("Modifique el periodo desde el Menu personalizado, posicionado en la hoja Certificación, antes de realizar la liquidación");
      return;
    } else {
      Logger.log("Confirmación de periodo total aceptada");

      const respuestaRetroactivos = Browser.msgBox("¿Debo calcular retroactivos?", Browser.Buttons.YES_NO);
      const calcularRetroactivos = respuestaRetroactivos === "yes";

      Logger.log("Valor de calcularRetroactivos: " + calcularRetroactivos);

      let periodoRetroactivo = null;
      if (calcularRetroactivos) {
        Logger.log("Se deben calcular retroactivos");
        periodoRetroactivo = Browser.inputBox("Ingrese el periodo de retroactivos");
        Logger.log("Periodo de retroactivos ingresado: " + periodoRetroactivo);
      } else {
        Logger.log("No se deben calcular retroactivos");
      }

      // Calcular el período no retroactivo
      let periodoNoRetroactivo = null;
      if (calcularRetroactivos && periodoRetroactivo && periodoTotal) {
        periodoNoRetroactivo = calcularPeriodoNoRetroactivo(periodoRetroactivo, periodoTotal);
        Logger.log("Periodo no retroactivo calculado: " + periodoNoRetroactivo);
      } else {
        Logger.log("No se calculará el período no retroactivo");
      }

      agregarColumnas(hojaActiva, hojaFormulario, calcularRetroactivos, periodoRetroactivo, periodoNoRetroactivo);

      const { totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar } = procesarFilas(hojaActiva, hojaCertificacion, hojaFormulario, hojaProximosVencimientos, programa, calcularRetroactivos, periodoRetroactivo, periodoNoRetroactivo);

      // copiarFormatos(hojaActiva);
      agregarFilaTotales(hojaActiva, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar, calcularRetroactivos);
      agregarDatosIoma(ss, programa, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar);
    }
  } else {
    Logger.log("No puede ejecutarse la función, ya que no es una hoja válida");
  }
}

function agregarColumnas(hojaActiva, hojaFormulario, calcularRetroactivos, periodoRetroactivo, periodoNoRetroactivo) {
  hojaActiva.getRange("A3:Z300" + hojaFormulario.getLastRow()).clearContent();

  const meses = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ]

  // Agregar encabezados para las columnas A a D
  hojaActiva.getRange("A3").setValue("Apellido");
  hojaActiva.getRange("B3").setValue("Nombre");
  hojaActiva.getRange("C3").setValue("Fecha de Alta");
  hojaActiva.getRange("D3").setValue("Tipo de Beca");

  if (calcularRetroactivos) {
    const fechasPeriodo = separarFechasPeriodo(periodoRetroactivo);
    if (fechasPeriodo === 0) {
      Browser.msgBox("Formato de fecha retroactiva incorrecto. Modifique el periodo retroactivo.");
      return;
    }
    const meses = obtenerMesesEntreFechas(fechasPeriodo.inicio, fechasPeriodo.fin);
    const mesesEnFormulario = hojaFormulario.getRange("H3:O3").getValues()[0];
    const resosEnFormulario = hojaFormulario.getRange("H9:O9").getValues()[0];

    let columnaActual = 5; // Empezamos desde la columna E
    // Agregar la columna "Monto ANTERIOR por {MES} {RESO}"
    const primerMes = meses[0].nombre;
    let indicePrimerMes = -1;

    // Buscar el índice del primer mes en mesesEnFormulario
    for (let i = 0; i < mesesEnFormulario.length; i++) {
      if (mesesEnFormulario[i].toLowerCase() === primerMes.toLowerCase()) {
        indicePrimerMes = i;
        break;
      }
    }

    let resoAnterior = "";
    let mesAnterior = "";

    if (indicePrimerMes !== -1 && indicePrimerMes < resosEnFormulario.length - 1) {
      // Obtener la resolución y el mes de la columna siguiente
      resoAnterior = resosEnFormulario[indicePrimerMes + 1];
      mesAnterior = mesesEnFormulario[indicePrimerMes + 1];
    } else if (indicePrimerMes !== -1) {
      resoAnterior = resosEnFormulario[indicePrimerMes];
      mesAnterior = mesesEnFormulario[indicePrimerMes];
    }
    hojaActiva.getRange(3, columnaActual).setValue("Monto ANTERIOR " + mesAnterior + " por " + resoAnterior);
    columnaActual++;

    // Agregar montos actualizados
    meses.forEach(mes => {
      for (let i = mesesEnFormulario.length - 1; i >= 0; i--) {
        if (mes.nombre.toLowerCase() === mesesEnFormulario[i].toLowerCase()) {
          hojaActiva.getRange(3, columnaActual).setValue("Monto actualizado " + mes.nombre + " por " + resosEnFormulario[i]);
          columnaActual++;
          break;
        }
      }
    });
    // *** Lógica modificada para agregar columnas de diferencia ***
    let textoDiferencia = "";
    let columnasDiferencia = []; // Array para almacenar las columnas de diferencia

    meses.forEach(mes => {
      let columnaMontoActualizado = null;
      for (let i = 5; i < columnaActual; i++) { // Buscar columna de monto actualizado
        if (hojaActiva.getRange(3, i).getValue().includes("Monto actualizado " + mes.nombre)) {
          columnaMontoActualizado = i;
          break;
        }
      }

      if (columnaMontoActualizado) {
        if (textoDiferencia) {
          textoDiferencia += " y " + mes.nombre;
        } else {
          textoDiferencia = "DIFERENCIA monto actualizado " + mes.nombre;
        }
      } else {
        // Si no hay columna de monto actualizado, agregar el mes al texto de diferencia del mes anterior
        if (columnasDiferencia.length > 0) {
          // En lugar de añadir "y", se verifica si ya existe "y", y si no existe se agrega.
          if (!columnasDiferencia[columnasDiferencia.length - 1].texto.includes(" y ")) {
            columnasDiferencia[columnasDiferencia.length - 1].texto += " y " + mes.nombre;
          } else {
            columnasDiferencia[columnasDiferencia.length - 1].texto += " y " + mes.nombre;
          }

        } else {
          textoDiferencia = "DIFERENCIA monto actualizado " + mes.nombre;
        }
      }

      if (textoDiferencia) {
        columnasDiferencia.push({ mes: mes.nombre, texto: textoDiferencia });
        textoDiferencia = "";
      }
    });

    // Agregar columnas de diferencia al final
    columnasDiferencia.forEach(diff => {
      hojaActiva.getRange(3, columnaActual).setValue(diff.texto);
      columnaActual++;
    });
    // Agregar la columna TOTAL DIFERENCIAS {mes}-{mes} + Liquidacion del mes
    const ultimoMes = meses[meses.length - 1].nombre;
    let totalDiferencias = "TOTAL DIFERENCIAS ";

    if (meses.length === 1) {
      totalDiferencias += ultimoMes; // Mostrar solo el mes si hay uno solo
    } else {
      totalDiferencias += primerMes + "-" + ultimoMes;
    }

    let mesesNoRetroactivos = []; // Declarar la variable 

    if (periodoNoRetroactivo) {
      // Analizar el periodoNoRetroactivo y obtener los meses
      const [inicioNoRetroactivoStr, finNoRetroactivoStr] = periodoNoRetroactivo.split(" al ");
      const inicioNoRetroactivo = parseDate(inicioNoRetroactivoStr);
      const finNoRetroactivo = parseDate(finNoRetroactivoStr);
      const mesesNoRetroactivos = obtenerMesesEntreFechas(inicioNoRetroactivo, finNoRetroactivo);

      // Agregar los nombres de los meses al título
      totalDiferencias += " + Liquidación de ";
      mesesNoRetroactivos.forEach((mes, index) => {
        if (mesesNoRetroactivos.length === 1) {
          totalDiferencias += mes.nombre;
        } else if (index === mesesNoRetroactivos.length - 1) {
          totalDiferencias += "y " + mes.nombre;
        } else {
          totalDiferencias += mes.nombre;
          if (index < mesesNoRetroactivos.length - 2) {
            totalDiferencias += ", ";
          } else if (index < mesesNoRetroactivos.length - 1) {
            totalDiferencias += " ";
          }
        }
      });
    }

    // Verificar si el mes de inicio del período retroactivo es el mismo que el mes de inicio del período total
    if (primerMes === mesesNoRetroactivos[0]?.nombre) {
      // Si son iguales, no agregar el mes de inicio nuevamente
      totalDiferencias = "TOTAL DIFERENCIAS " + ultimoMes;
    }

    hojaActiva.getRange(3, columnaActual).setValue(totalDiferencias);
    columnaActual++;

    // Agregar encabezados para las últimas 4 columnas
    hojaActiva.getRange(3, columnaActual).setValue("Monto IOMA Personal");
    hojaActiva.getRange(3, columnaActual + 1).setValue("Monto IOMA Patronal");
    hojaActiva.getRange(3, columnaActual + 3).setValue("OBSERVACIONES");

    // Crear el mapeo de columnas y pasarlo a procesarFilas

    const mapeoMeses = buscarMesesEnColumnas()



  } else {
    if (!calcularRetroactivos) {
      // Agregar encabezado para la columna E si no hay retroactivos
      hojaActiva.getRange("E3").setValue("Monto Actualizado ");

      // Agregar encabezados para las últimas 4 columnas
      hojaActiva.getRange("F3").setValue("Monto IOMA Personal");
      hojaActiva.getRange("G3").setValue("Monto IOMA Patronal");
      hojaActiva.getRange("H3").setValue("Importe a Liquidar");
      hojaActiva.getRange("I3").setValue("OBSERVACIONES");
    }
  }
}

function buscarMesesEnColumnas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  const meses = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ]

  const encabezados = sheet.getRange('F3:Z3').getValues()[0]

  const mapeoColumnasMeses = []

  encabezados.forEach((encabezado) => {
    const encabezadoPartido = encabezado.split(' ')
    encabezadoPartido.forEach((parte) => {
      const resultado = meses.find((mes) => mes.toLowerCase() === parte.toLowerCase())
      if (resultado) {
        mapeoColumnasMeses.push(resultado.toLowerCase())
        return
      }
    })
  })

  return mapeoColumnasMeses
}

function calcularPeriodoNoRetroactivo(periodoRetroactivo, periodoTotal) {
  try {
    // Analizar las cadenas de período
    const [inicioRetroactivoStr, finRetroactivoStr] = periodoRetroactivo.split(" al ");
    const [inicioTotalStr, finTotalStr] = periodoTotal.split(" al ");

    // Convertir las fechas a objetos Date
    const inicioRetroactivo = parseDate(inicioRetroactivoStr);
    const finRetroactivo = parseDate(finRetroactivoStr);
    const inicioTotal = parseDate(inicioTotalStr);
    const finTotal = parseDate(finTotalStr);

    // Verificar si el período retroactivo es mayor que el período total
    if (finRetroactivo > finTotal) {
      Browser.msgBox("Error: El período retroactivo no puede ser mayor que el período total.");
      return 0; // Código de error para detener la ejecución
    }

    // Verificar si los períodos son iguales
    if (periodoRetroactivo === periodoTotal) {
      return null;
    }

    // Calcular el período no retroactivo
    const inicioNoRetroactivo = new Date(finRetroactivo.getFullYear(), finRetroactivo.getMonth() + 1, 1);

    // Verificar si el inicio del período no retroactivo es después del fin del período total
    if (inicioNoRetroactivo > finTotal) {
      return null;
    }

    // Formatear el resultado
    const inicioNoRetroactivoStr = formatDate(inicioNoRetroactivo);
    const finNoRetroactivoStr = formatDate(finTotal);

    return inicioNoRetroactivoStr + " al " + finNoRetroactivoStr;
  } catch (e) {
    Browser.msgBox("Error inesperado al calcular el período no retroactivo.");
    Logger.log("Error al calcular el período no retroactivo: " + e);
    return 0; // Código de error en caso de error inesperado
  }
}
function parseDate(dateStr) {
  const [dia, mes, anio] = dateStr.split("/");
  return new Date(anio, mes - 1, dia);
}
function formatDate(date) {
  const dia = date.getDate().toString().padStart(2, "0");
  const mes = (date.getMonth() + 1).toString().padStart(2, "0");
  const anio = date.getFullYear();
  return dia + "/" + mes + "/" + anio;
}
function separarFechasPeriodo(periodo) {

  const fechas = periodo.split(' al ');
  if (fechas.length < 2) {
    return 0;
  }
  const fechaInicio = convertirFecha(fechas[0].trim());
  const fechaFin = convertirFecha(fechas[1].trim());

  if (isNaN(fechaInicio) || isNaN(fechaFin)) {
    return 0;
  }
  return { inicio: fechaInicio, fin: fechaFin }
}
function identificarPrograma(titulo) {
  const programas = {
    "fortalecimiento": "Fortalecimiento",
    "operadores de calle": "Operadores de Calle",
    "autonomia joven": "Autonomia Joven"
  };
  for (const key in programas) {
    if (titulo.includes(key)) {
      return programas[key];
    }
  }
  return "Programa no identificado";
}
function obtenerMesesEntreFechas(fechaInicio, fechaFin) {
  const meses = [];
  let fechaActual = new Date(fechaInicio);

  while (fechaActual <= fechaFin) {
    let inicioMes = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), 1);
    let finMes = new Date(fechaActual.getFullYear(), fechaActual.getMonth() + 1, 0);

    if (inicioMes < fechaInicio) {
      inicioMes = new Date(fechaInicio);
    }
    if (finMes > fechaFin) {
      finMes = new Date(fechaFin);
    }

    meses.push({ inicio: inicioMes, fin: finMes, nombre: obtenerNombreMes(fechaActual.getMonth()) });
    fechaActual.setMonth(fechaActual.getMonth() + 1);
    fechaActual.setDate(1);
  }

  return meses;
}


function agregarFilaAVencimientos(hoja, fila, row, programa, idPersona, ids) {
  fila[2] = row[2];
  fila[3] = row[3];
  fila[4] = row[4];
  fila[5] = programa;
  hoja.appendRow(fila);
  ids.add(idPersona);
}
function eliminarFila(hoja, idUnicoPersona) {
  const data = hoja.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const id = data[i][1] + "-" + data[i][2];
    if (id === idUnicoPersona) {
      hoja.deleteRow(i + 1);
      break;
    }
  }
}

function procesarFilas(hojaActiva, hojaCertificacion, hojaFormulario, hojaProximosVencimientos, programa, hayRetroactivos, periodoRetroactivo, periodoNoRetroactivo) {
  const data = hojaCertificacion.getDataRange().getValues();
  const currentDate = new Date();
  let totalIomaPersonal = 0;
  let totalIomaPatronal = 0;
  let totalImporteALiquidar = 0;
  let idsProximosVencimientos = new Set();
  const dataProximosVencimientos = hojaProximosVencimientos.getDataRange().getValues();

  dataProximosVencimientos.forEach(row => {
    const idUnicoPersona = row[1] + "-" + row[2];
    idsProximosVencimientos.add(idUnicoPersona);
  });

  data.slice(3).forEach((row, i) => {
    const columnaK = row[10].toString().toLowerCase(); // Columna "Observación"
    const fechaColumnaD = new Date(row[3]); // Columna "Fecha de Alta"
    const diferenciaMeses = (currentDate.getFullYear() - fechaColumnaD.getFullYear()) * 12 + (currentDate.getMonth() - fechaColumnaD.getMonth());
    const idUnicoPersona = row[1] + "-" + row[2]; // Columna "Apellido" y "Nombre"
    const periodoLiquidacionParticular = row[9]; // Columna "Periodo"

    // Crear filaDatos con las columnas necesarias
    let filaDatos = [
      row[0], // Apellido
      row[1], // Nombre
      row[3], // Fecha de Alta
      row[4], // Tipo de Beca
      // ... (agregar aquí las demás columnas que necesites)
    ];

    // Lógica para próximos vencimientos
    if (!columnaK.includes('baja') && !columnaK.includes('limitacion')) {
      if (diferenciaMeses === 11 && !idsProximosVencimientos.has(idUnicoPersona)) {
        agregarFilaAVencimientos(hojaProximosVencimientos, filaDatos, row, programa, idUnicoPersona, idsProximosVencimientos);
      } else if (diferenciaMeses !== 11 && idsProximosVencimientos.has(idUnicoPersona)) {
        eliminarFila(hojaProximosVencimientos, idUnicoPersona);
      }
    }
    // Lógica para procesar montos
    if (hayRetroactivos) {
      // Si hay retroactivos, evaluar bajas y limitaciones

      ({ fila: filaDatos, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar } = procesarMonto(filaDatos, row, hojaFormulario, hayRetroactivos, periodoRetroactivo, periodoNoRetroactivo, periodoLiquidacionParticular, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar));
      hojaActiva.appendRow(filaDatos);
    } else {
      // Si no hay retroactivos, excluir bajas y limitaciones
      if (!columnaK.includes('baja') && !columnaK.includes('limitacion' || 'limitación')) {
        ({ fila: filaDatos, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar } = procesarMonto(filaDatos, row, hojaFormulario, hayRetroactivos, null, null, periodoLiquidacionParticular, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar));
        hojaActiva.appendRow(filaDatos);
      }
    }
  });

  return { totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar };
}
function procesarMonto(filaHojaLiquidacion, row, hojaFormulario, hayRetroactivo, periodoRetroactivo, periodoNoRetroactivo, periodoLiquidacionParticular, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar) {
  let montoALiquidar;
  if (!hayRetroactivo) {
    montoALiquidar = calcularMontoSinRetroactivo(row, hojaFormulario);
    filaHojaLiquidacion[4] = montoALiquidar.monto;
    filaHojaLiquidacion[5] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2));
    filaHojaLiquidacion[6] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2));
    filaHojaLiquidacion[7] = parseFloat((montoALiquidar.monto - filaHojaLiquidacion[5]).toFixed(2));
    filaHojaLiquidacion[9] = montoALiquidar.detalles;
    totalIomaPersonal += filaHojaLiquidacion[5];
    totalIomaPatronal += filaHojaLiquidacion[6];
    totalImporteALiquidar += filaHojaLiquidacion[7];
  } else {



    //   filaHojaLiquidacion[4] = buscarValorBecaAnterior(filaHojaLiquidacion[3], hojaFormulario);
    //   filaHojaLiquidacion[5] = row[5];
    //   filaHojaLiquidacion[6] = filaHojaLiquidacion[5] - filaHojaLiquidacion[4];
    //   if (soloPeriodo === "yes") {
    //     montoALiquidar = calcularMontoRetroactivoSoloPeriodo(row, hojaFormulario, filaHojaLiquidacion[6]);
    //     filaHojaLiquidacion[7] = montoALiquidar.monto;
    //   } else if (soloPeriodo === "no") {
    //     montoALiquidar = calcularRetroactivoCompleto(row, periodoRetroactivo, filaHojaLiquidacion[6], hojaFormulario);
    //     filaHojaLiquidacion[7] = montoALiquidar.monto;
    //   }
    //   filaHojaLiquidacion[8] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2));
    //   filaHojaLiquidacion[9] = parseFloat((montoALiquidar.monto * 4.8 / 100).toFixed(2));
    //   filaHojaLiquidacion[10] = parseFloat((montoALiquidar.monto - filaHojaLiquidacion[9]).toFixed(2));
    //   filaHojaLiquidacion[12] = montoALiquidar.detalles
    //   totalIomaPersonal += filaHojaLiquidacion[8];
    //   totalIomaPatronal += filaHojaLiquidacion[9];
    //   totalImporteALiquidar += filaHojaLiquidacion[10];
  }

  return { fila: filaHojaLiquidacion, totalIomaPersonal, totalIomaPatronal, totalImporteALiquidar };
}

function calcularRetroactivoCompleto(filaCertificacion, dato, diferencia, hojaFormulario) {
  const fecha = separarFechasPeriodo(filaCertificacion[9]);
  let diasHabiles;
  let montoMensual = 0;
  let montoTotal = 0;
  let detallesCalculo = "";
  let resultado = fechaMayor(dato, filaCertificacion[3]);
  const mesRetroactivo = obtenerMesesEntreFechas(resultado, fecha.inicio);

  mesRetroactivo.forEach(mesR => {
    if (mesR.fin < fecha.inicio) {
      diasHabiles = calcularDiasHabiles(mesR.inicio, mesR.fin);
      if (diasHabiles > 0) {
        montoMensual = (diferencia / 24) * diasHabiles;
        montoTotal += montoMensual;
        detallesCalculo += `${mesR.nombre}: ${montoMensual.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
      }
    }
  });

  const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);
  meses.forEach(mes => {
    diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
    if (diasHabiles > 0) {
      const tipoBeca = filaCertificacion[4];
      montoMensual = obtenerMontoMensual(hojaFormulario, tipoBeca, mes.nombre);
      const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
      montoTotal += montoMensualCalculado;
      detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
    }
  });

  return { monto: montoTotal, detalles: detallesCalculo };
}

function calcularMontoSinRetroactivo(filaCertificacion, dato) {
  const fecha = separarFechasPeriodo(filaCertificacion[9]);
  let diasHabiles;
  let montoMensual = 0;
  let montoTotal = 0;
  let detallesCalculo = "";
  const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);

  meses.forEach(mes => {
    diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
    const tipoBeca = filaCertificacion[4];
    montoMensual = obtenerMontoMensual(dato, tipoBeca, mes.nombre);
    const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
    montoTotal += montoMensualCalculado;
    detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
  });

  return { monto: montoTotal, detalles: detallesCalculo };
}

function calcularMontoRetroactivoSoloPeriodo(filaCertificacion, dato, diferencia) {
  const fecha = separarFechasPeriodo(filaCertificacion[9]);
  let diasHabiles;
  let montoMensual = 0;
  let montoTotal = 0;
  let detallesCalculo = "";
  const meses = obtenerMesesEntreFechas(fecha.inicio, fecha.fin);

  meses.forEach(mes => {
    diasHabiles = calcularDiasHabiles(mes.inicio, mes.fin);
    montoMensual = diferencia;
    const montoMensualCalculado = (montoMensual / 24) * diasHabiles;
    montoTotal += montoMensualCalculado;
    detallesCalculo += `${mes.nombre}: ${montoMensualCalculado.toFixed(2)} (días hábiles: ${diasHabiles})\n`;
  });

  return { monto: montoTotal, detalles: detallesCalculo };
}

function obtenerNombreMes(mes) {
  const nombresMeses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return nombresMeses[mes];
}

function calcularDiasHabiles(fechaInicio, fechaFin) {
  let diasHabiles = 0;
  const fechaActual = new Date(fechaInicio);

  while (fechaActual <= fechaFin) {
    const diaSemana = fechaActual.getDay();
    if (diaSemana >= 1 && diaSemana <= 6) {
      diasHabiles++;
    }
    fechaActual.setDate(fechaActual.getDate() + 1);
  }
  return Math.min(diasHabiles, 24);
}

function obtenerMontoMensual(hojaFormulario, tipoBeca, mes) {
  const datosFormulario = hojaFormulario.getDataRange().getValues();
  const monthNames = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const mesIndex = monthNames.indexOf(mes.toLowerCase());

  if (mesIndex === -1) {
    return 0;
  }

  for (let i = 3; i < 8; i++) {
    const tipoBecaActual = datosFormulario[i][6];
    if (tipoBecaActual && tipoBecaActual.toLowerCase() === tipoBeca.toLowerCase()) {
      for (let j = mesIndex; j >= 0; j--) {
        const mesActual = monthNames[j];
        for (let k = 7; k <= 14; k++) {
          const mesCelda = datosFormulario[2][k];
          if (mesCelda && mesCelda.toLowerCase() === mesActual) {
            return datosFormulario[i][k];
          }
        }
      }
      for (let j = 14; j >= 7; j--) {
        if (j == mesIndex) continue;

        const mesAnterior = monthNames[j];
        for (let k = 7; k <= 14; k++) {
          const mesCelda = datosFormulario[2][k];
          if (mesCelda && mesCelda.toLowerCase() === mesAnterior) {
            return datosFormulario[i][k];
          }
        }
      }
    }
  }
  return 0;
}

function buscarValoresBeca(tipoBeca, hojaFormulario) {
  const rangoTipos = hojaFormulario.getRange("G4:G8").getValues().flat();
  const rangoValoresAnteriores = hojaFormulario.getRange("I4:O8").getValues().flat();
  const rangoValoresActuales = hojaFormulario.getRange("H4:H8").getValues().flat();
  const index = rangoTipos.findIndex(tipo => tipo.trim() === tipoBeca);
  if (index !== -1) {
    return { montoAnterior: rangoValoresAnteriores[index], montoActual: rangoValoresActuales[index] };
  }
  return { montoAnterior: "Tipo de beca no encontrado", montoActual: "Tipo de beca no encontrado" };
}

function obtenerMesReferencia(fecha) {
  const dia = fecha.getDate();
  const mes = fecha.getMonth();
  const anio = fecha.getFullYear();

  if (dia >= 15) {
    return new Date(anio, mes, 15);
  } else {
    return new Date(anio, mes - 1, 15);
  }
}

function agregarFilaTotales(hoja, totalColumna5o8, totalColumna6o9, totalColumna7o10, respuesta) {
  const ultimaFila = hoja.getLastRow() + 1;
  hoja.getRange(ultimaFila, 1).setValue("TOTAL");
  hoja.getRange(ultimaFila, 1).setFontWeight("bold").setFontSize(11);

  if (!respuesta) {
    hoja.getRange(ultimaFila, 1, 1, 5).mergeAcross();
    hoja.getRange(ultimaFila, 6).setValue(totalColumna5o8);
    hoja.getRange(ultimaFila, 7).setValue(totalColumna6o9);
    hoja.getRange(ultimaFila, 8).setValue(totalColumna7o10);
    hoja.getRange(ultimaFila, 6).setFontWeight("bold").setFontSize(11);
    hoja.getRange(ultimaFila, 7).setFontWeight("bold").setFontSize(11);
    hoja.getRange(ultimaFila, 8).setFontWeight("bold").setFontSize(11);
  } else {
    const ultimaColumna = hoja.getLastColumn() - 1
    hoja.getRange(ultimaFila, 1, 1, 8).mergeAcross();
    hoja.getRange(ultimaFila, ultimaColumna - 3).setValue(totalColumna5o8);
    hoja.getRange(ultimaFila, ultimaColumna - 2).setValue(totalColumna6o9);
    hoja.getRange(ultimaFila, ultimaColumna - 1).setValue(totalColumna7o10);
    hoja.getRange(ultimaFila, ultimaColumna - 3).setFontWeight("bold").setFontSize(11);
    hoja.getRange(ultimaFila, ultimaColumna - 2).setFontWeight("bold").setFontSize(11);
    hoja.getRange(ultimaFila, ultimaColumna - 1).setFontWeight("bold").setFontSize(11);
  }
}
function agregarDatosIoma(ss, programa, totalColumna5o8, totalColumna6o9, totalColumna7o10) {
  const hojaIoma = ss.getSheetByName("IOMA " + programa);
  const numberFormat = '[$$]#,##0.00';
  hojaIoma.getRange(4, 2).setValue(totalColumna7o10).setNumberFormat(numberFormat).setFontWeight("bold");
  hojaIoma.getRange(4, 3).setValue(totalColumna6o9).setNumberFormat(numberFormat);
  hojaIoma.getRange(4, 4).setValue(totalColumna6o9).setNumberFormat(numberFormat);
  hojaIoma.getRange(4, 5).setValue(totalColumna6o9 + totalColumna5o8).setNumberFormat(numberFormat).setFontWeight("bold");
}

// funcion a futuro
// function recuperarInfoFormulario () {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Formulario')
//   const vector = []
//   const valores = sheet.getRange('H3:O9')
//   let i = 0

//   while (i < 8 && valores[i][0]) {
//     const info = {
//       mes: valores[i][0],
//       administrativo: valores[i][1],
//       profesional: valores[i][3],
//       coordinador: valores[i][4],
//       resolucion: valores[i][6]
//     }
//     vector.push(info)
//     i++
//   }
//   return vector
// }

// posible desuso

function buscarValorBecaAnterior(tipoBeca, hojaFormulario) {
  const rangoTipos = hojaFormulario.getRange("G13:G17").getValues().flat();
  const rangoDatos = hojaFormulario.getRange("I13:I17").getValues().flat();
  const index = rangoTipos.findIndex(tipo => tipo.trim() === tipoBeca);
  return index !== -1 ? rangoDatos[index] : "Tipo de beca no encontrado";
}

function convertirFecha(fechaTexto) {
  const partes = fechaTexto.split('/');
  if (partes.length === 3) {
    const [dia, mes, anio] = partes.map(Number);
    return new Date(anio, mes - 1, dia);
  }
  return new Date(fechaTexto);
}



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