/**
 * Visitas Landing CCA — registra cada visita a la landing en la hoja de LEADS.
 * Proyecto independiente: NO toca el código del CRM ni el formulario de la landing.
 *
 * Crea/usa dos hojas dentro del archivo del CRM:
 *   - "VISITAS LANDING": una fila por visita (se actualiza mientras la persona navega)
 *   - "RESUMEN VISITAS": totales, conversión y tablas por asesor, origen y dispositivo
 *
 * Pasos: pegar este código → ejecutar configurar() una vez → Implementar como aplicación web
 * (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario) → copiar la URL /exec.
 */

var ID_HOJA = '1XZGj-uyAGVoi2uX7dA_qG-hzlFxkmKYME8i9Fo_4CbM';
var HOJA_VISITAS = 'VISITAS LANDING';
var HOJA_RESUMEN = 'RESUMEN VISITAS';
var ZONA = 'America/La_Paz';

var COLUMNAS = [
  'ID VISITA', 'FECHA', 'HORA', 'VISITANTE', 'Nº DE VISITA', 'CÓDIGO ASESOR', 'ASESOR', 'LLEGÓ POR LINK DEL ASESOR',
  'SEDE', 'DISPOSITIVO', 'NAVEGADOR / APP', 'ORIGEN', 'CAMPAÑA', 'TIEMPO EN LA PÁGINA (seg)', 'SCROLL MÁX (%)',
  'ABRIÓ A SOFÍA', 'CURSOS VISTOS', 'CURSOS CONSULTADOS A SOFÍA', 'SE REGISTRÓ', 'CLIC EN WHATSAPP',
  'TIPO DE VISITANTE', 'ÚLTIMA ACTUALIZACIÓN'
];

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.accion === 'visita' && p.datos) {
    try { guardarVisita_(JSON.parse(p.datos)); } catch (err) { console.error(err); }
  }
  return ContentService.createTextOutput('ok');
}

function guardarVisita_(d) {
  if (!d || !d.visitaId) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = obtenerHojaVisitas_();
    var ahora = new Date();
    var celda = hoja.getRange('A:A').createTextFinder(String(d.visitaId)).matchEntireCell(true).findNext();
    var tipo = d.registro && d.registro !== 'NO' ? 'Lead (se registró)'
      : (d.whatsapp && d.whatsapp !== 'NO' ? 'Contactó por WhatsApp'
      : (d.abrioSofia === 'SÍ' ? 'Interesado (habló con Sofía)'
      : (Number(d.segundos) < 10 ? 'Rebote (menos de 10 seg)' : 'Solo miró la oferta')));
    var fila = [
      d.visitaId,
      celda ? hoja.getRange(celda.getRow(), 2).getValue() : Utilities.formatDate(ahora, ZONA, 'dd/MM/yyyy'),
      celda ? hoja.getRange(celda.getRow(), 3).getValue() : Utilities.formatDate(ahora, ZONA, 'HH:mm'),
      d.visitante || '', Number(d.numVisita) || 1, d.asesorCod || '', d.asesor || '', d.porLink || '',
      d.sede || '', d.dispositivo || '', d.navegador || '', d.origen || '', d.campana || '',
      Number(d.segundos) || 0, Number(d.scroll) || 0, d.abrioSofia || 'NO',
      d.vistos || '', d.consultados || '', d.registro || 'NO', d.whatsapp || 'NO',
      tipo, Utilities.formatDate(ahora, ZONA, 'dd/MM/yyyy HH:mm:ss')
    ];
    if (celda) hoja.getRange(celda.getRow(), 1, 1, fila.length).setValues([fila]);
    else hoja.appendRow(fila);
  } finally {
    lock.releaseLock();
  }
}

function obtenerHojaVisitas_() {
  var libro = SpreadsheetApp.openById(ID_HOJA);
  var hoja = libro.getSheetByName(HOJA_VISITAS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_VISITAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS])
      .setFontWeight('bold').setBackground('#1E3F73').setFontColor('#FFFFFF').setWrap(true);
    hoja.setFrozenRows(1);
    hoja.getRange('B:C').setNumberFormat('@');
    hoja.setColumnWidths(1, COLUMNAS.length, 130);
    hoja.setColumnWidth(17, 260); hoja.setColumnWidth(18, 260);
  }
  return hoja;
}

/** Ejecutar UNA vez: crea las hojas y el resumen con fórmulas. */
function configurar() {
  var libro = SpreadsheetApp.openById(ID_HOJA);
  obtenerHojaVisitas_();
  var r = libro.getSheetByName(HOJA_RESUMEN) || libro.insertSheet(HOJA_RESUMEN);
  r.clear();
  var V = "'" + HOJA_VISITAS + "'!";
  var filas = [
    ['RESUMEN DE VISITAS A LA LANDING', ''],
    ['Visitas totales', '=COUNTA(' + V + 'A2:A)'],
    ['Personas distintas (visitantes únicos)', '=COUNTUNIQUE(' + V + 'D2:D)'],
    ['Visitas de hoy', '=COUNTIF(' + V + 'B2:B,TEXT(TODAY(),"dd/mm/yyyy"))'],
    ['Personas que volvieron (2 o más visitas)', '=COUNTUNIQUE(IFERROR(FILTER(' + V + 'D2:D,' + V + 'E2:E>=2)))'],
    ['Abrieron a Sofía', '=COUNTIF(' + V + 'P2:P,"SÍ")'],
    ['Se registraron (formulario o Sofía)', '=COUNTIF(' + V + 'S2:S,"<>NO")-COUNTBLANK(' + V + 'S2:S)'],
    ['Hicieron clic en WhatsApp', '=COUNTIF(' + V + 'T2:T,"SÍ*")'],
    ['Entraron y NO se registraron', '=B2-B7'],
    ['Conversión (registros / visitas)', '=IFERROR(B7/B2,0)'],
    ['Tiempo promedio en la página (seg)', '=IFERROR(ROUND(AVERAGE(' + V + 'N2:N)),0)']
  ];
  // etiquetas como texto y fórmulas con setFormulas (formato en inglés, válido en cualquier configuración regional)
  r.getRange(1, 1, filas.length, 1).setValues(filas.map(function (f) { return [f[0]]; }));
  r.getRange(2, 2, filas.length - 1, 1).setFormulas(filas.slice(1).map(function (f) { return [f[1]]; }));
  r.getRange('B10').setNumberFormat('0.0%');
  r.getRange('A1').setFontWeight('bold').setFontSize(14).setFontColor('#1E3F73');
  r.getRange('A2:A11').setFontWeight('bold');
  r.getRange('D1').setValue('VISITAS POR ASESOR').setFontWeight('bold');
  r.getRange('D2').setFormula('=IFERROR(QUERY(' + V + 'A1:V,"select G, count(A), avg(N) where A is not null group by G order by count(A) desc label G \'Asesor\', count(A) \'Visitas\', avg(N) \'Seg. promedio\'",1),"Sin datos")');
  r.getRange('H1').setValue('VISITAS POR ORIGEN').setFontWeight('bold');
  r.getRange('H2').setFormula('=IFERROR(QUERY(' + V + 'A1:V,"select L, count(A) where A is not null group by L order by count(A) desc label L \'Origen\', count(A) \'Visitas\'",1),"Sin datos")');
  r.getRange('K1').setValue('VISITAS POR DISPOSITIVO').setFontWeight('bold');
  r.getRange('K2').setFormula('=IFERROR(QUERY(' + V + 'A1:V,"select J, count(A) where A is not null group by J label J \'Dispositivo\', count(A) \'Visitas\'",1),"Sin datos")');
  r.getRange('N1').setValue('TIPO DE VISITANTE').setFontWeight('bold');
  r.getRange('N2').setFormula('=IFERROR(QUERY(' + V + 'A1:V,"select U, count(A) where A is not null group by U order by count(A) desc label U \'Tipo\', count(A) \'Visitas\'",1),"Sin datos")');
  r.setColumnWidth(1, 300);
  SpreadsheetApp.flush();
}

/** Prueba manual: crea una visita de ejemplo (puedes borrar la fila después). */
function probar() {
  guardarVisita_({ visitaId: 'PRUEBA-1', visitante: 'PRUEBA', numVisita: 1, asesorCod: 'RSCF', asesor: 'Ramiro Santos Condori Flores',
    porLink: 'SÍ', sede: 'sucre', dispositivo: 'PC', navegador: 'Chrome', origen: 'directo', segundos: 42, scroll: 60,
    abrioSofia: 'SÍ', vistos: 'Diseño Gráfico', consultados: 'Diseño Gráfico', registro: 'NO', whatsapp: 'NO' });
}
