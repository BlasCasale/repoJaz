function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú personalizado')
    .addSubMenu(ui.createMenu('Modificar Periodo/s')
      .addItem('Modificar Hoja Actual', 'modificarHojaActiva')
      .addItem('Modificar Todas las Hojas', 'modificarTodasLasHojas'))
    .addItem('Calcular Liquidacion', 'procesarLiquidaciones')
    .addItem('Modificar Exp o Reso', 'mostrarFormulario')
    .addToUi();
}