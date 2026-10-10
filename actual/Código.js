/**
 * RECOGIDA LECHE — creación de la estructura de la hoja.
 * Pegar en Extensiones > Apps Script y ejecutar crearEstructura() una sola vez.
 * Si una pestaña ya existe, no la toca (no borra datos).
 */
function crearEstructura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const HOJAS = {
    CONFIG: {
      cols: ['CLAVE', 'VALOR', 'NOTA'],
      filas: [
        ['CARGADOR_NOMBRE', 'LACTEOS CUQUERELLA S.L.', 'Sale en el DeCA'],
        ['CARGADOR_NIF', 'B13516620', 'Sale en el DeCA'],
        ['CARGADOR_DIRECCION', 'Autovía de Andalucía, km 172,5 · 13200 Manzanares (Ciudad Real)', 'Tomada de la web: confirmar que es la que debe salir'],
        ['DESTINO_DESCARGA', 'Manzanares', 'Destino de la mercancía en el DeCA'],
        ['MERCANCIA', 'Leche cruda', ''],
        ['FACTOR_KG_POR_LITRO', 1.03, 'Para pasar litros a peso en el DeCA'],
        ['TEMP_AVISO', 6, '°C. Aviso naranja (recogida en días alternos)'],
        ['TEMP_ALERTA', 8, '°C. Aviso rojo'],
        ['ARRASTRE_AVISO', 50, 'Litros. Aviso de arrastre alto'],
        ['CARPETA_DECA_ID', '1EcPBGUBmIGrO1bEAxeXWtkWlznBQGO3u', 'Carpeta RECOGIDA LECHE/DECA'],
        ['CARPETA_EXPORTACION_ID', '1qoFDcSbndVMs9u4lIU0_OofFxRvTwJts', 'Carpeta RECOGIDA LECHE/EXPORTACION PROGRAMA']
      ]
    },
    CONDUCTORES: {
      cols: ['CONDUCTOR_ID', 'NOMBRE', 'PIN', 'RUTAS', 'CAMION_HABITUAL', 'TRANSPORTISTA', 'NIF_TRANSPORTISTA', 'ACTIVO'],
      filas: [
        ['JAVIER', 'Javier', '', '02', 'CAMION_JAVIER', 'PENDIENTE: empresa de transporte', '', 'SI'],
        ['FELIPE', 'Felipe', '', '03,04', 'CAMION_FELIPE', 'PENDIENTE: empresa de transporte', '', 'SI'],
        ['JESUS', 'Jesús', '', '05,06', 'CAMION_JESUS', 'PENDIENTE: Jesús (autónomo), nombre completo', '', 'SI']
      ]
    },
    CAMIONES: {
      cols: ['CAMION_ID', 'MATRICULA', 'DESCRIPCION', 'TITULAR', 'N_COMPARTIMENTOS', 'ACTIVO', 'NOTAS'],
      filas: [
        ['CAMION_JAVIER', 'PENDIENTE', 'Cisterna 4 compartimentos', 'Empresa de transporte', 4, 'SI', ''],
        ['CAMION_FELIPE', 'PENDIENTE', 'Cisterna 3 compartimentos', 'Empresa de transporte', 3, 'SI', ''],
        ['CAMION_JESUS', 'PENDIENTE', 'Cisterna 3 compartimentos', 'Jesús (autónomo)', 3, 'SI', ''],
        ['CAMION_RESERVA', 'PENDIENTE', 'Camión de sustitución', '', '', 'NO', 'Rellenar y activar cuando haga falta']
      ]
    },
    COMPARTIMENTOS: {
      cols: ['CAMION_ID', 'ORDEN', 'LETRA', 'ESPECIES', 'MEZCLA_ADMITIDA', 'NOTA'],
      filas: [
        ['CAMION_JAVIER', 1, 'O', 'O', '', ''],
        ['CAMION_JAVIER', 2, 'C', 'C,V', '', 'Cabra + vaca de poca cantidad'],
        ['CAMION_JAVIER', 3, 'D', 'D', '', ''],
        ['CAMION_JAVIER', 4, 'd', 'D', '', 'D.O. (2)'],
        ['CAMION_FELIPE', 1, 'V', 'V', '', 'Vaca pura'],
        ['CAMION_FELIPE', 2, 'C', 'C,V', '', 'Cabra + vaca'],
        ['CAMION_FELIPE', 3, 'D', 'D', 'O', 'PENDIENTE: hoy se mezcla oveja con D.O.'],
        ['CAMION_JESUS', 1, 'D', 'D', '', ''],
        ['CAMION_JESUS', 2, 'O', 'O', '', ''],
        ['CAMION_JESUS', 3, 'C', 'C', '', '']
      ]
    },
    RUTAS: {
      cols: ['RUTA', 'CONDUCTOR_ID', 'DENOMINACION', 'PREFIJO_SI_CONJUNTA', 'ACTIVA'],
      filas: [
        ['02', 'JAVIER', 'RUTA 02 MUNERA', '02', 'SI'],
        ['03', 'FELIPE', 'RUTA 03 FELIPE', '03', 'SI'],
        ['04', 'FELIPE', 'RUTA 04 FELIPE', '03', 'SI'],
        ['05', 'JESUS', 'RUTA 05 JESUS', 'PENDIENTE (oficina)', 'SI'],
        ['06', 'JESUS', 'RUTA 06 JESUS', 'PENDIENTE (oficina)', 'SI']
      ]
    },
    GANADEROS: {
      cols: ['CODIGO', 'NOMBRE', 'RUTA', 'ORDEN', 'ESPECIE', 'DESTINO_HABITUAL', 'POBLACION', 'ULTIMOS_LITROS', 'ACTIVO', 'NOTAS'],
      filas: []
    },
    VIAJES: {
      cols: ['VIAJE_ID', 'FECHA', 'CONDUCTOR_ID', 'RUTAS', 'PREFIJO_DESTINO', 'CAMION_ID', 'MATRICULA', 'CAMION_HABITUAL', 'TRANSPORTISTA', 'NIF_TRANSPORTISTA',
             'DECA_NUM', 'DECA_URL', 'HORA_INICIO', 'LITROS_ESTIMADOS', 'PESO_ESTIMADO_KG',
             'HORA_FIN', 'LITROS_REALES', 'PESO_REAL_KG', 'ESTADO', 'EXPORTADO'],
      filas: []
    },
    RECOGIDAS: {
      cols: ['REC_ID', 'VIAJE_ID', 'FECHA', 'HORA', 'CONDUCTOR_ID', 'RUTA', 'ORDEN', 'CODIGO', 'ESPECIE',
             'LITROS', 'TEMPERATURA', 'MUESTRA', 'DESTINO', 'LITROS_DESTINO', 'DESTINO_2', 'LITROS_DESTINO_2',
             'NO_RECOGIDO', 'INCIDENCIA_TEMP', 'LATITUD', 'LONGITUD', 'RECIBIDO_EN'],
      filas: []
    }
  };

  // Columnas que deben guardarse como texto (códigos con ceros delante: 02, 020543, 010)
  const TEXTO = ['MATRICULA', 'CAMION_ID', 'CAMION_HABITUAL', 'RUTA', 'RUTAS', 'CODIGO', 'ORDEN', 'MUESTRA', 'PIN', 'PREFIJO_DESTINO', 'PREFIJO_SI_CONJUNTA', 'LETRA', 'DESTINO', 'DESTINO_2', 'DESTINO_HABITUAL'];

  Object.keys(HOJAS).forEach(function (nombre) {
    if (ss.getSheetByName(nombre)) return;
    const def = HOJAS[nombre];
    const sh = ss.insertSheet(nombre);
    def.cols.forEach(function (c, i) {
      if (TEXTO.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    if (def.filas.length) sh.getRange(2, 1, def.filas.length, def.cols.length).setValues(def.filas);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, def.cols.length);
  });

  // Quita la pestaña vacía inicial si sigue existiendo
  const inicial = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (inicial && ss.getSheets().length > 1 && inicial.getLastRow() === 0) ss.deleteSheet(inicial);

  SpreadsheetApp.getUi().alert('Estructura creada: ' + Object.keys(HOJAS).join(', '));
}