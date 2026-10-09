/**
 * RECOGIDA LECHE — servidor de la app de los conductores.
 * Va en el mismo proyecto de Apps Script que Setup.gs (hoja RECOGIDA LECHE).
 * Se publica como Aplicación web: Ejecutar como "Yo" y acceso "Cualquier usuario".
 *
 * La app envía POST con texto JSON: {accion, pin, ...}. Todas las acciones exigen PIN
 * y solo devuelven/escriben datos de las rutas de ese conductor.
 */
const TZ = 'Europe/Madrid';
const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  if (e && e.parameter && e.parameter.tipo) return puente_(e.parameter);
  return json_({ ok: true, app: 'RECOGIDA LECHE', hora: ahora_('dd/MM/yyyy HH:mm') });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const cond = auth_(req.pin, req.dev);
    const roles = roles_(cond);
    // Cada rol tiene sus acciones; quien tiene varios roles (p. ej. OFICINA y CALIDAD) tiene la suma. El login es el del rol principal.
    const POR_ROL = {
      OFICINA: { login: loginOficina, oficinaViajes: oficinaViajes, historicoViajes: historicoViajes, historicoCalidad: historicoCalidad, corregir: corregir, recepViajes: recepViajes, guardarRecepcion: guardarRecepcion, guardarAgua: guardarAgua, guardarRecepCalidad: guardarRecepCalidad,
        anularRecepcion: anularRecepcion, guardarObsDia: guardarObsDia, guardarVenta: guardarVenta, borrarVenta: borrarVenta, guardarMezcla: guardarMezcla },
      CALIDAD: { login: loginCalidad, recepViajes: recepViajes, historicoCalidad: historicoCalidad, guardarMezcla: guardarMezcla, verificarDia: verificarDia, guardarObsDia: guardarObsDia, guardarVenta: guardarVenta, borrarVenta: borrarVenta },
      RECEPCION: { login: loginRecepcion, recepViajes: recepViajes, guardarRecepcion: guardarRecepcion, guardarAgua: guardarAgua, anularRecepcion: anularRecepcion, guardarRecepCalidad: guardarRecepCalidad },
      CONDUCTOR: { login: login, iniciarViaje: iniciarViaje, guardar: guardar, trasvase: trasvase, cerrarViaje: cerrarViaje, historial: historial, anularViaje: anularViaje }
    };
    const acciones = {};
    roles.slice().reverse().forEach(function (r) { Object.keys(POR_ROL[r]).forEach(function (k) { acciones[k] = POR_ROL[r][k]; }); });   // el primero (principal) gana
    const fn = acciones[req.accion];
    if (!fn) throw new Error('Acción desconocida: ' + req.accion);
    // Firma de la recepción: en la pantalla compartida de fábrica cada persona firma con su PIN al guardar
    let quien = cond;
    if (req.firma && /^(guardarRecepcion|anularRecepcion|guardarAgua|guardarRecepCalidad)$/.test(req.accion)) {
      quien = tabla_(HOJA_US_()).filas.filter(function (r) { return String(r.PIN).trim() === String(req.firma).trim() && String(r.ACTIVO).toUpperCase() !== 'NO' && (esRecepcion_(r) || esOficina_(r)); })[0];
      if (!quien) throw new Error('PIN de firma incorrecto: no es de nadie de recepción.');
    }
    const res = fn(req, quien);
    if (quien !== cond && res && res.ok) res.firmado = String(quien.NOMBRE);
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ───────────── Acciones ───────────── */

/** Datos que necesita el móvil: conductor, sus rutas, sus ganaderos, camiones y configuración. */
function login(req, c) {
  columnas_(HOJA_US_(), ['ROL'], 'ACTIVO');
  columnas_('COMPARTIMENTOS', ['RUTA'], 'CAPACIDAD');
  const rutasC = split_(c.RUTAS).map(pad2_);
  if (!rutasC.length) throw new Error('Este PIN no tiene rutas de recogida ni un rol de oficina, recepción o calidad. Revisa la columna ROL en la hoja.');
  const rutas = {};
  tabla_('RUTAS').filas.forEach(function (r) {
    const k = pad2_(r.RUTA);
    if (rutasC.indexOf(k) < 0 || String(r.ACTIVA).toUpperCase() === 'NO') return;
    const p = String(r.PREFIJO_SI_CONJUNTA || '').trim();
    rutas[k] = { nombre: String(r.DENOMINACION || 'RUTA ' + k), pref: /^\d{1,2}$/.test(p) ? pad2_(p) : '' };
  });

  const ganaderos = tabla_('GANADEROS').filas
    .filter(function (g) { return g.CODIGO !== '' && rutas[pad2_(g.RUTA)] && String(g.ACTIVO).toUpperCase() !== 'NO'; })
    // Un ganadero puede tener varias especies en la misma fila (ESPECIE = "CO", "C,O", "CD"…):
    // se convierte en una parada por especie, con el mismo código, ruta y orden.
    .reduce(function (acc, g) {
      const esps = especies_(g.ESPECIE); if (!esps.length) esps.push('');
      const ult = porEspecie_(g.ULTIMOS_LITROS, esps), dl = destinosPor_(g.DESTINO_HABITUAL, esps);
      esps.forEach(function (e) {
        acc.push({
          cod: String(g.CODIGO).trim(), nom: String(g.NOMBRE).trim(), ruta: pad2_(g.RUTA),
          ord: String(g.ORDEN).trim(), esp: e, dl: dl[e] || '', pue: String(g.POBLACION || '').trim(),
          ult: Number(ult[e]) || 0
        });
      });
      return acc;
    }, [])
    .sort(function (a, b) { return a.ruta === b.ruta ? (Number(a.ord) || 0) - (Number(b.ord) || 0) : (a.ruta < b.ruta ? -1 : 1); });

  // Camiones que se le ofrecen: su habitual, los que ya ha usado y los activos que no son habituales de otro
  const conductores = tabla_(HOJA_US_()).filas;
  const habituales = conductores.map(function (x) { return String(x.CAMION_HABITUAL); });
  const usados = {};
  usados[String(c.CAMION_HABITUAL)] = true;
  tabla_('VIAJES').filas.forEach(function (v) { if (v.CONDUCTOR_ID === c.CONDUCTOR_ID && v.CAMION_ID) usados[String(v.CAMION_ID)] = true; });
  const comps = tabla_('COMPARTIMENTOS').filas;
  const camiones = tabla_('CAMIONES').filas
    .filter(function (k) {
      const id = String(k.CAMION_ID);
      if (usados[id]) return true;
      return String(k.ACTIVO).toUpperCase() === 'SI' && habituales.indexOf(id) < 0;
    })
    .map(function (k) { return camionObj_(k, comps); })
    .filter(function (k) { return k.letras.length; });

  const cfg = cfg_();
  return {
    ok: true,
    conductor: {
      id: c.CONDUCTOR_ID, nombre: String(c.NOMBRE), habitual: String(c.CAMION_HABITUAL),
      transportista: String(c.TRANSPORTISTA), nif: String(c.NIF_TRANSPORTISTA)
    },
    rutas: rutas, ganaderos: ganaderos, camiones: camiones, hoy: hoy_(rutasC),
    config: {
      tempAviso: Number(cfg.TEMP_AVISO) || 6, tempAlerta: Number(cfg.TEMP_ALERTA) || 8,
      arrastreAviso: Number(cfg.ARRASTRE_AVISO) || 50, factorKg: Number(cfg.FACTOR_KG_POR_LITRO) || 1.03,
      cargador: String(cfg.CARGADOR_NOMBRE), cargadorNif: String(cfg.CARGADOR_NIF),
      direccion: String(cfg.CARGADOR_DIRECCION), destino: String(cfg.DESTINO_DESCARGA), mercancia: String(cfg.MERCANCIA)
    }
  };
}

/** Crea el viaje, el PDF del DeCA (público por enlace) y devuelve su URL de descarga directa. */
function iniciarViaje(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = req.viaje;
    const ya = buscar_('VIAJES', 'VIAJE_ID', v.id);
    if (ya) return { ok: true, viaje: viajeResp_(ya) };            // reintento: no duplicar
    comprobarRutas_(c, v.rutas);
    // ¿Ya hay hoy un viaje de estas rutas? Se avisa antes de sacar un segundo DeCA (evita repetir la ruta por error)
    const h = hoy_(v.rutas);
    // Una ruta cerrada hoy no se vuelve a hacer (la descarga parcial va dentro del mismo viaje).
    // Durante las pruebas se puede permitir con CONFIG → PERMITIR_REPETIR_RUTA = SI.
    const cerrados = h.viajes.filter(function (x) { return /^CERRADO$/.test(String(x.estado).trim()); });
    if (cerrados.length && String(cfg_().PERMITIR_REPETIR_RUTA || '').toUpperCase() !== 'SI') {
      const x = cerrados[0];
      throw new Error('La ruta ' + split_(x.rutas).map(function (r) { return +r; }).join(' y ') + ' ya se hizo hoy y está cerrada (DeCA ' + x.deca + ', ' + x.ini + '–' + x.fin + '). No se puede repetir. Si hay un error, avisa a la oficina.');
    }
    if (h.viajes.length && !v.confirmado) return { ok: true, aviso: h.viajes, hoy: h };
    const cam = resolverCamion_(v.camion, c);
    const cfg = cfg_();
    const litros = Math.round(Number(v.estLitros) || 0);
    if (litros <= 0) throw new Error('Indica los litros estimados del viaje: el DeCA no puede salir con 0 kg.');
    const kg = Math.round(litros * (Number(cfg.FACTOR_KG_POR_LITRO) || 1.03));
    let num = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-' + c.CONDUCTOR_ID + '-' + Utilities.formatDate(new Date(), TZ, 'HHmm');
    const usados = tabla_('VIAJES').filas.map(function (x) { return String(x.DECA_NUM); });
    if (usados.indexOf(num) >= 0) { let k = 2; while (usados.indexOf(num + '-' + k) >= 0) k++; num = num + '-' + k; }
    const pdf = crearDeca_({
      num: num, cfg: cfg, cond: c, matricula: cam.matricula, letraQ: cam.letraQ, rutas: v.rutas, kg: kg, litros: litros,
      nota: v.nota || ''
    });
    const fila = {
      VIAJE_ID: v.id, FECHA: ahora_('dd/MM/yyyy'), CONDUCTOR_ID: c.CONDUCTOR_ID, RUTAS: v.rutas.join(','),
      PREFIJO_DESTINO: pad2_(v.pref), CAMION_ID: cam.id, MATRICULA: cam.matricula, LETRA_Q: cam.letraQ, CAMION_HABITUAL: String(c.CAMION_HABITUAL),
      TRANSPORTISTA: String(c.TRANSPORTISTA), NIF_TRANSPORTISTA: String(c.NIF_TRANSPORTISTA),
      DECA_NUM: num, DECA_URL: pdf.url, HORA_INICIO: ahora_('HH:mm'), LITROS_ESTIMADOS: litros, PESO_ESTIMADO_KG: kg,
      ESTADO: v.nota ? 'EN CURSO (' + v.nota + ')' : 'EN CURSO', EXPORTADO: 'NO'
    };
    anadir_('VIAJES', fila);
    return { ok: true, viaje: viajeResp_(fila), camion: cam, hoy: h };
  } finally { lock.releaseLock(); }
}

/** Guarda o actualiza recogidas (por REC_ID). Actualiza últimos litros y destino habitual del ganadero. */
function guardar(req, c) {
  const todas = req.recs || [];
  let recs = todas;
  if (!recs.length) return { ok: true, ids: [] };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const viajes = tabla_('VIAJES').filas;
    const mios = {};
    viajes.forEach(function (v) { if (v.CONDUCTOR_ID === c.CONDUCTOR_ID) mios[v.VIAJE_ID] = true; });

    columnas_('RECOGIDAS', ['DESTINO_3', 'LITROS_DESTINO_3'], 'LITROS_DESTINO_2');
    const t = tabla_('RECOGIDAS');
    const idx = {}; t.filas.forEach(function (r) { idx[r.REC_ID] = r._fila; });
    const nuevas = [];
    const recibido = ahora_('dd/MM/yyyy HH:mm:ss');
    // Recogidas de un viaje que ya no existe (borrado de la hoja, pruebas…): no bloquean el móvil.
    // Se guardan aparte en SIN_VIAJE para que la oficina las revise, y se dan por recibidas.
    const huerfanas = recs.filter(function (r) { return !mios[r.viajeId]; });
    if (huerfanas.length) guardarSinViaje_(huerfanas, c);
    recs = recs.filter(function (r) { return mios[r.viajeId]; });
    recs.forEach(function (r) {
      // Hasta tres destinos: el principal, el arrastre de manguera y lo que no cupo (reparto)
      const a = r.arr ? Number(r.arr.l) || 0 : 0, rp = r.rep ? Number(r.rep.l) || 0 : 0;
      const obj = {
        REC_ID: r.recId, VIAJE_ID: r.viajeId, FECHA: r.fecha, HORA: r.hora, CONDUCTOR_ID: c.CONDUCTOR_ID,
        RUTA: pad2_(r.ruta), ORDEN: r.ord, CODIGO: r.cod, ESPECIE: r.esp,
        LITROS: r.sin ? '' : r.litros, TEMPERATURA: r.sin ? '' : r.temp, MUESTRA: r.sin ? '' : r.muestra,
        DESTINO: r.sin ? '' : r.dest, LITROS_DESTINO: r.sin ? '' : (r.litros - a - rp),
        DESTINO_2: a ? r.arr.dest : (rp ? r.rep.dest : ''), LITROS_DESTINO_2: a || rp || '',
        DESTINO_3: a && rp ? r.rep.dest : '', LITROS_DESTINO_3: a && rp ? rp : '',
        NO_RECOGIDO: r.sin ? 'SI' : '',
        INCIDENCIA_TEMP: r.inc ? 'SI' : '' ,
        LATITUD: r.lat || '', LONGITUD: r.lon || '', RECIBIDO_EN: recibido
      };
      if (r.mezcla) obj.INCIDENCIA_TEMP = (obj.INCIDENCIA_TEMP ? 'SI · ' : '') + 'MEZCLA';
      const row = t.cab.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
      if (idx[r.recId]) t.sh.getRange(idx[r.recId], 1, 1, row.length).setValues([row]);
      else nuevas.push(row);
    });
    if (nuevas.length) {
      const f0 = t.sh.getLastRow() + 1;
      t.cab.forEach(function (h, j) { if (COLS_TEXTO.indexOf(h) >= 0) t.sh.getRange(f0, j + 1, nuevas.length, 1).setNumberFormat('@'); });
      t.sh.getRange(f0, 1, nuevas.length, t.cab.length).setValues(nuevas);
    }

    // Últimos litros y destino habitual (la letra) para el día siguiente
    const g = tabla_('GANADEROS');
    const cUlt = g.cab.indexOf('ULTIMOS_LITROS') + 1, cDest = g.cab.indexOf('DESTINO_HABITUAL') + 1;
    // Clave código|especie. Si la fila tiene varias especies, se guarda por especie: "O:880 C:310"
    const filaCod = {}; g.filas.forEach(function (x) {
      const cod = String(x.CODIGO).trim();
      especies_(x.ESPECIE).forEach(function (e) { filaCod[cod + '|' + e] = x; });
      if (!filaCod[cod]) filaCod[cod] = x;
    });
    recs.forEach(function (r) {
      const x = filaCod[r.cod + '|' + especie_(r.esp)] || filaCod[r.cod]; if (!x || r.sin) return;
      const esps = especies_(x.ESPECIE), e = especie_(r.esp);
      if (esps.length <= 1) {
        if (cUlt) g.sh.getRange(x._fila, cUlt).setValue(r.litros);
        if (cDest && !r.mezcla && !r.noHabitual && r.letra) g.sh.getRange(x._fila, cDest).setValue(r.letra);
        return;
      }
      if (cUlt) {
        const m = porEspecie_(x.ULTIMOS_LITROS, esps); m[e] = r.litros;
        x.ULTIMOS_LITROS = juntar_(m, esps); g.sh.getRange(x._fila, cUlt).setValue(x.ULTIMOS_LITROS);
      }
      if (cDest && !r.mezcla && !r.noHabitual && r.letra) {
        const d = destinosPor_(x.DESTINO_HABITUAL, esps); d[e] = r.letra;
        x.DESTINO_HABITUAL = juntar_(d, esps); g.sh.getRange(x._fila, cDest).setValue(x.DESTINO_HABITUAL);
      }
    });
    return { ok: true, ids: todas.map(function (r) { return r.recId; }), sinViaje: todas.length - recs.length };
  } finally { lock.releaseLock(); }
}

/** Cierra el viaje del camión averiado y abre uno nuevo (con su DeCA) en el camión de destino. */
function trasvase(req, c) {
  const vieja = buscar_('VIAJES', 'VIAJE_ID', req.viajeId);
  if (!vieja || vieja.CONDUCTOR_ID !== c.CONDUCTOR_ID) throw new Error('Viaje no encontrado');
  const litros = Math.round(Number(req.litrosHasta) || 0);
  actualizar_('VIAJES', vieja._fila, {
    HORA_FIN: ahora_('HH:mm'), LITROS_REALES: litros,
    PESO_REAL_KG: Math.round(litros * (Number(cfg_().FACTOR_KG_POR_LITRO) || 1.03)),
    ESTADO: 'CERRADO POR TRASVASE → ' + req.nuevo.id
  });
  req.nuevo.nota = 'trasvase desde ' + vieja.MATRICULA;
  req.nuevo.confirmado = true;   // es la continuación del mismo viaje: no avisar de ruta repetida
  return iniciarViaje({ viaje: req.nuevo }, c);
}

/** Cierra el viaje: guarda lo pendiente, anota litros reales y deja el fichero para el programa de gestión. */
function cerrarViaje(req, c) {
  if (req.recs && req.recs.length) guardar(req, c);
  const v = buscar_('VIAJES', 'VIAJE_ID', req.viajeId);
  if (!v || v.CONDUCTOR_ID !== c.CONDUCTOR_ID) {
    // Cierre de un viaje que ya no existe (borrado de la hoja): se apunta en SIN_VIAJE y se da por hecho,
    // para que el móvil no se quede bloqueado reintentándolo. Sus recogidas (si venían) ya fueron a SIN_VIAJE en guardar().
    guardarSinViaje_([{ viajeId: req.viajeId, recId: 'CIERRE', fecha: ahora_('dd/MM/yyyy'), hora: ahora_('HH:mm'), ruta: '', cod: '', esp: '',
      litros: Math.round(Number(req.litros) || 0), dest: 'CIERRE DE VIAJE' }], c);
    return { ok: true, sinViaje: true };
  }
  const cfg = cfg_();
  const litros = Math.round(Number(req.litros) || 0);
  const ids = (req.viajes && req.viajes.length ? req.viajes : [req.viajeId]);

  if (/^CERRADO$/.test(String(v.ESTADO).trim()) && /^R\w{7}\.TXT$/i.test(String(v.EXPORTADO))) return { ok: true, exportado: v.EXPORTADO };   // reintento
  const exportado = exportar_(ids, null);
  actualizar_('VIAJES', v._fila, {
    HORA_FIN: ahora_('HH:mm'), LITROS_REALES: litros,
    PESO_REAL_KG: Math.round(litros * (Number(cfg.FACTOR_KG_POR_LITRO) || 1.03)),
    ESTADO: 'CERRADO', EXPORTADO: exportado
  });
  return { ok: true, exportado: exportado };
}

/** Fichero para el programa de gestión con todas las recogidas de esos viajes.
 *  nombre = null → nombre nuevo del día (R + aammdd + nº). Con nombre → se reescribe ese mismo fichero (correcciones). */
function lineas_(ids) {
  return tabla_('RECOGIDAS').filas.filter(function (r) { return ids.indexOf(r.VIAJE_ID) >= 0 && r.NO_RECOGIDO !== 'SI'; })
    .map(function (r) {
      return [r.FECHA, pad2_(r.RUTA), r.ORDEN, r.CODIGO, r.ESPECIE, r.LITROS, num1_(r.TEMPERATURA),
              r.DESTINO, r.LITROS_DESTINO, r.DESTINO_2, r.LITROS_DESTINO_2, r.DESTINO_3 || '', r.LITROS_DESTINO_3 || '', r.MUESTRA].join(';');
    });
}
function exportar_(ids, nombre) {
  const cfg = cfg_();
  const lineas = lineas_(ids);
  if (!cfg.CARPETA_EXPORTACION_ID || !lineas.length) return nombre || 'NO';
  const carpetaExp = DriveApp.getFolderById(cfg.CARPETA_EXPORTACION_ID);
  const texto = lineas.join('\r\n') + '\r\n';
  if (nombre && nombre !== 'NO') {
    const it = carpetaExp.getFilesByName(nombre);
    if (it.hasNext()) { it.next().setContent(texto); return nombre; }
    carpetaExp.createFile(nombre, texto, MimeType.PLAIN_TEXT); return nombre;
  }
  // Nombre corto MS-DOS (8.3): R + aammdd + nº del día (1-9, luego A-Z) → R2609281.TXT
  const base = 'R' + Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  const SEQ = '123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (let i = 0; i < SEQ.length; i++) {
    const n = base + SEQ.charAt(i) + '.TXT';
    if (!carpetaExp.getFilesByName(n).hasNext()) { carpetaExp.createFile(n, texto, MimeType.PLAIN_TEXT); return n; }
  }
  throw new Error('Demasiados ficheros de exportación hoy');
}
/** Todos los viajes encadenados por trasvase con este, y el último (el que tiene el fichero). */
function cadena_(viajeId, viajes) {
  viajes = viajes || tabla_('VIAJES').filas;
  const porId = {}; viajes.forEach(function (v) { porId[v.VIAJE_ID] = v; });
  const sig = function (v) { const m = String(v.ESTADO).match(/TRASVASE\s*→\s*(\S+)/); return m ? porId[m[1]] : null; };
  let fin = porId[viajeId]; if (!fin) return null;
  while (sig(fin)) fin = sig(fin);
  const ids = [fin.VIAJE_ID];
  let cambio = true;
  while (cambio) { cambio = false; viajes.forEach(function (v) { const n = sig(v); if (n && ids.indexOf(n.VIAJE_ID) >= 0 && ids.indexOf(v.VIAJE_ID) < 0) { ids.push(v.VIAJE_ID); cambio = true; } }); }
  return { ids: ids, fin: fin };
}

/* ───────────── Oficina (PIN de ROL OFICINA): ver todas las rutas y corregir viajes cerrados ───────────── */

function loginOficina(req, c) {
  const cfg = cfg_();
  return {
    ok: true, rol: 'OFICINA', roles: roles_(c), usuario: { id: c.CONDUCTOR_ID, nombre: String(c.NOMBRE) },
    config: { tempAviso: Number(cfg.TEMP_AVISO) || 6, tempAlerta: Number(cfg.TEMP_ALERTA) || 8 }
  };
}
/** Viajes de todos los conductores de los últimos días, con cada recogida y lo que falta en los que están en curso. */
function oficinaViajes(req, c) {
  const dias = Math.min(Math.max(Number(req.dias) || 15, 1), 45);
  const lim = new Date(); lim.setHours(0, 0, 0, 0); lim.setDate(lim.getDate() - (dias - 1));
  const fecha = function (s) { const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null; };
  const todos = tabla_('VIAJES').filas;
  const solo = Array.isArray(req.ids) && req.ids.length ? req.ids.map(String) : null;   // un viaje concreto del histórico, sea de la fecha que sea
  const viajes = todos.filter(function (v) { if (solo) return solo.indexOf(String(v.VIAJE_ID)) >= 0; const d = fecha(v.FECHA); return d && d >= lim && String(v.ESTADO).indexOf('ANULADO') !== 0; });
  const conds = {}; tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = x; });
  const gans = tabla_('GANADEROS').filas;
  const nombres = {}; gans.forEach(function (g) { const k = String(g.CODIGO).trim(); if (!nombres[k]) nombres[k] = g.NOMBRE; });
  const camTab = tabla_('CAMIONES').filas, compTab = tabla_('COMPARTIMENTOS').filas;
  const porId = {};
  viajes.forEach(function (v) {
    const k = camTab.filter(function (x) { return String(x.CAMION_ID) === String(v.CAMION_ID); })[0];
    const letras = k ? letrasRuta_(camionObj_(k, compTab).letras, v.PREFIJO_DESTINO).map(function (x) { return pad2_(v.PREFIJO_DESTINO) .replace(/^0/, '') + x.l; }) : [];
    porId[v.VIAJE_ID] = { id: v.VIAJE_ID, fecha: v.FECHA, deca: v.DECA_NUM, decaUrl: String(v.DECA_URL || ''), matricula: v.MATRICULA, rutas: v.RUTAS, estado: v.ESTADO,
      ini: v.HORA_INICIO, fin: v.HORA_FIN, conductor: v.CONDUCTOR_ID, nombre: conds[v.CONDUCTOR_ID] ? String(conds[v.CONDUCTOR_ID].NOMBRE) : v.CONDUCTOR_ID,
      exportado: v.EXPORTADO, comps: letras, recs: [] };
  });
  tabla_('RECOGIDAS').filas.forEach(function (r) {
    const v = porId[r.VIAJE_ID]; if (!v) return;
    v.recs.push({ id: r.REC_ID, ruta: r.RUTA, ord: r.ORDEN, cod: r.CODIGO, nom: nombres[String(r.CODIGO).trim()] || '', esp: especie_(r.ESPECIE),
      l: numES_(r.LITROS), t: r.TEMPERATURA, mue: r.MUESTRA, d1: r.DESTINO, l1: numES_(r.LITROS_DESTINO), d2: r.DESTINO_2, l2: numES_(r.LITROS_DESTINO_2),
      d3: r.DESTINO_3 || '', l3: numES_(r.LITROS_DESTINO_3), sin: String(r.NO_RECOGIDO).toUpperCase() === 'SI', hora: r.HORA,
      lat: r.LATITUD, lon: r.LONGITUD, cor: r.CORREGIDO || '' });
  });
  const lista = Object.keys(porId).map(function (k) { return porId[k]; });
  lista.forEach(function (v) {
    v.recs.sort(function (a, b) { return a.ruta === b.ruta ? (Number(a.ord) || 0) - (Number(b.ord) || 0) : (a.ruta < b.ruta ? -1 : 1); });
    // Ganaderos de sus rutas que aún no tienen recogida (para ver lo que falta y para añadir uno olvidado)
    const hechos = {}; v.recs.forEach(function (r) { hechos[String(r.cod).trim() + '|' + r.esp] = true; });
    const rs = split_(v.rutas).map(pad2_), falta = [];
    gans.forEach(function (g) {
      if (rs.indexOf(pad2_(g.RUTA)) < 0 || String(g.ACTIVO).toUpperCase() === 'NO') return;
      especies_(g.ESPECIE).forEach(function (e) { const k = String(g.CODIGO).trim() + '|' + e;
        if (!hechos[k]) falta.push({ cod: String(g.CODIGO).trim(), nom: String(g.NOMBRE), esp: e, ruta: pad2_(g.RUTA), ord: String(g.ORDEN).trim() }); });
    });
    falta.sort(function (a, b) { return a.ruta === b.ruta ? (Number(a.ord) || 0) - (Number(b.ord) || 0) : (a.ruta < b.ruta ? -1 : 1); });
    v.falta = falta;
  });
  lista.sort(function (a, b) { return (fecha(b.fecha) - fecha(a.fecha)) || (String(a.ini) < String(b.ini) ? -1 : 1); });
  return { ok: true, viajes: lista, hasta: ahora_('dd/MM/yyyy HH:mm') };
}
/** Histórico de viajes para la oficina: todos los de VIAJES entre dos fechas (por defecto, el último año), sin las recogidas
 *  (rápido), con su DeCA. El detalle de uno se pide aparte con oficinaViajes({ids:[id]}). */
/** Histórico del registro de calidad (Recepción 2): un registro por camión y viaje, entre dos fechas (por defecto, el último año). */
function historicoCalidad(req, c) {
  hoja_('RECEPCION_CALIDAD', CAB_RCAL); hoja_('MEZCLAS', CAB_MEZCLA); hoja_('PARTE_DIA', CAB_PARTE);
  const hoy = new Date(), haceUnAno = new Date(); haceUnAno.setFullYear(hoy.getFullYear() - 1);
  const d = fechaNum_(req.desde) || fechaNum_(Utilities.formatDate(haceUnAno, TZ, 'dd/MM/yyyy')), h = fechaNum_(req.hasta) || fechaNum_(Utilities.formatDate(hoy, TZ, 'dd/MM/yyyy'));
  const mz = {}; tabla_('MEZCLAS').filas.forEach(function (r) { mz[r.VIAJE_ID + '|' + r.COMPARTIMENTO] = String(r.MEZCLA); });
  const ver = {}; tabla_('PARTE_DIA').filas.forEach(function (r) { if (String(r.VERIFICADO_POR || '').trim()) ver[String(r.FECHA)] = String(r.VERIFICADO_POR); });
  const regs = {}, orden = [];
  tabla_('RECEPCION_CALIDAD').filas.forEach(function (r) {
    const id = String(r.VIAJE_ID); if (!/^Q/.test(id)) return;
    const f = fechaNum_(r.FECHA); if (!f || f < d || f > h) return;
    if (!regs[id]) { regs[id] = { id: id, fecha: String(r.FECHA), f: f, matricula: String(r.MATRICULA || ''), conductor: String(r.CONDUCTOR || ''), usuario: String(r.USUARIO || ''),
      limpieza: String(r.LIMPIEZA_CISTERNA || ''), filtro: String(r.LIMPIEZA_FILTRO || ''), obs: String(r.OBSERVACIONES || ''), verificado: ver[String(r.FECHA)] || '', comps: [], litros: 0 }; orden.push(id); }
    const g = regs[id], l = numES_(r.LITROS);
    g.comps.push({ n: String(r.N_COMP_CISTERNA || ''), esp: String(r.ESPECIE), l: l, temp: String(r.TEMPERATURA || ''), ph: String(r.PH || ''), dornic: String(r.DORNIC || ''),
      visual: String(r.VISUAL || ''), mezcla: mz[id + '|' + r.ID] || '' });
    g.litros += l;
  });
  const lista = orden.map(function (id) { return regs[id]; }).sort(function (a, b) { return b.f - a.f; });
  return { ok: true, registros: lista };
}
function historicoViajes(req, c) {
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const haceUnAno = new Date(hoy); haceUnAno.setFullYear(hoy.getFullYear() - 1);
  const d = fechaNum_(req.desde) || fechaNum_(Utilities.formatDate(haceUnAno, TZ, 'dd/MM/yyyy')), h = fechaNum_(req.hasta) || fechaNum_(Utilities.formatDate(hoy, TZ, 'dd/MM/yyyy'));
  const conds = {}; tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = String(x.NOMBRE); });
  const lista = tabla_('VIAJES').filas.filter(function (v) { const f = fechaNum_(v.FECHA); return f && f >= d && f <= h; }).map(function (v) {
    return { id: v.VIAJE_ID, fecha: v.FECHA, f: fechaNum_(v.FECHA), conductor: v.CONDUCTOR_ID, nombre: conds[v.CONDUCTOR_ID] || v.CONDUCTOR_ID, rutas: v.RUTAS, matricula: v.MATRICULA,
      deca: v.DECA_NUM, decaUrl: String(v.DECA_URL || ''), ini: v.HORA_INICIO, fin: v.HORA_FIN, litros: numES_(v.LITROS_REALES), estado: String(v.ESTADO), exportado: v.EXPORTADO };
  });
  lista.sort(function (a, b) { return (b.f - a.f) || (String(b.ini) < String(a.ini) ? -1 : 1); });
  return { ok: true, viajes: lista };
}
/** Corrección de la oficina en un viaje CERRADO: cambia o añade una recogida, lo apunta en CAMBIOS (como un tachón
 *  en la cartilla: quién, cuándo, antes, después y motivo) y rehace el fichero de ese viaje con el mismo nombre. */
function corregir(req, c) {
  const motivo = String(req.motivo || '').trim();
  if (!motivo) throw new Error('Escribe el motivo de la corrección.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    columnas_('RECOGIDAS', ['CORREGIDO'], 'RECIBIDO_EN');
    const viajes = tabla_('VIAJES').filas;
    const v = viajes.filter(function (x) { return x.VIAJE_ID === req.viajeId; })[0];
    if (!v) throw new Error('Viaje no encontrado');
    if (!/^CERRADO/.test(String(v.ESTADO))) throw new Error('Solo se corrigen viajes cerrados. Este está ' + String(v.ESTADO).toLowerCase() + '.');
    const d = req.datos || {};
    const sin = !!d.sin, litros = sin ? 0 : Math.round(Number(d.litros) || 0);
    const l2 = sin ? 0 : Math.round(Number(d.l2) || 0), l3 = sin ? 0 : Math.round(Number(d.l3) || 0);
    if (!sin && !(litros > 0)) throw new Error('Faltan los litros.');
    if (!sin && !String(d.d1 || '').trim()) throw new Error('Falta el destino.');
    if (l2 + l3 >= litros && !sin) throw new Error('Lo que va a otros destinos tiene que ser menos que el total.');
    if ((l2 && !String(d.d2 || '').trim()) || (l3 && !String(d.d3 || '').trim())) throw new Error('Falta el destino de los otros litros.');
    const nuevo = {
      LITROS: sin ? '' : litros, TEMPERATURA: sin || d.temp === '' || d.temp == null ? '' : Number(String(d.temp).replace(',', '.')), MUESTRA: sin ? '' : String(d.muestra || ''),
      DESTINO: sin ? '' : String(d.d1).trim(), LITROS_DESTINO: sin ? '' : litros - l2 - l3,
      DESTINO_2: l2 ? String(d.d2).trim() : '', LITROS_DESTINO_2: l2 || '', DESTINO_3: l3 ? String(d.d3).trim() : '', LITROS_DESTINO_3: l3 || '',
      NO_RECOGIDO: sin ? 'SI' : ''
    };
    const quien = String(c.NOMBRE), cuando = ahora_('dd/MM/yyyy HH:mm');
    const t = tabla_('RECOGIDAS');
    let fila = req.recId ? t.filas.filter(function (r) { return r.REC_ID === req.recId && r.VIAJE_ID === v.VIAJE_ID; })[0] : null;
    const cambios = [];
    if (fila) {
      Object.keys(nuevo).forEach(function (k) {
        const antes = String(fila[k] == null ? '' : fila[k]).trim(), despues = String(nuevo[k]).trim();
        if (antes.replace(',', '.') !== despues.replace(',', '.')) cambios.push([k, antes, despues]);
      });
      if (!cambios.length) return { ok: true, sinCambios: true };
      nuevo.CORREGIDO = quien + ' ' + cuando;
      actualizar_('RECOGIDAS', fila._fila, nuevo);
    } else {
      // Ganadero que se olvidó apuntar
      const cod = String(req.cod || '').trim(), esp = especie_(req.esp);
      if (!cod || !esp) throw new Error('Falta el ganadero.');
      const g = tabla_('GANADEROS').filas.filter(function (x) { return String(x.CODIGO).trim() === cod; })[0];
      const recId = v.VIAJE_ID + '-' + cod + '-' + esp;
      if (t.filas.some(function (r) { return r.REC_ID === recId; })) throw new Error('Ese ganadero ya está en el viaje: corrígelo en vez de añadirlo.');
      const obj = Object.assign({ REC_ID: recId, VIAJE_ID: v.VIAJE_ID, FECHA: v.FECHA, HORA: '', CONDUCTOR_ID: v.CONDUCTOR_ID,
        RUTA: pad2_(req.ruta || (g && g.RUTA) || ''), ORDEN: String(req.ord || (g && g.ORDEN) || ''), CODIGO: cod, ESPECIE: esp,
        INCIDENCIA_TEMP: '', LATITUD: '', LONGITUD: '', RECIBIDO_EN: cuando, CORREGIDO: quien + ' ' + cuando + ' (añadido)' }, nuevo);
      anadir_('RECOGIDAS', obj);
      cambios.push(['AÑADIDO', '', (sin ? 'no recogido' : litros + ' l → ' + nuevo.DESTINO)]);
      fila = { REC_ID: recId, CODIGO: cod, ESPECIE: esp };
    }
    // Registro de cambios
    let sh = SS.getSheetByName('CAMBIOS');
    if (!sh) {
      sh = SS.insertSheet('CAMBIOS');
      sh.getRange(1, 1, 1, 11).setValues([['FECHA_HORA', 'USUARIO', 'VIAJE_ID', 'DECA_NUM', 'FECHA_VIAJE', 'REC_ID', 'CODIGO', 'CAMPO', 'ANTES', 'DESPUES', 'MOTIVO']])
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
      sh.setFrozenRows(1);
    }
    const filas = cambios.map(function (x) { return [cuando, quien, v.VIAJE_ID, v.DECA_NUM, v.FECHA, fila.REC_ID, fila.CODIGO, x[0], x[1], x[2], motivo]; });
    const f0 = sh.getLastRow() + 1;
    sh.getRange(f0, 1, filas.length, 11).setNumberFormat('@').setValues(filas);
    // Litros reales del viaje y fichero para el programa, con el mismo nombre
    const cad = cadena_(v.VIAJE_ID, viajes);
    const total = tabla_('RECOGIDAS').filas.filter(function (r) { return cad.ids.indexOf(r.VIAJE_ID) >= 0 && r.NO_RECOGIDO !== 'SI'; })
      .reduce(function (a, r) { return a + numES_(r.LITROS); }, 0);
    const cfg = cfg_();
    actualizar_('VIAJES', cad.fin._fila, { LITROS_REALES: total, PESO_REAL_KG: Math.round(total * (Number(cfg.FACTOR_KG_POR_LITRO) || 1.03)) });
    const nombre = exportar_(cad.ids, /^R\w{7}\.TXT$/i.test(String(cad.fin.EXPORTADO)) ? String(cad.fin.EXPORTADO) : null);
    if (nombre !== cad.fin.EXPORTADO) actualizar_('VIAJES', cad.fin._fila, { EXPORTADO: nombre });
    return { ok: true, cambios: cambios.length, fichero: nombre };
  } finally { lock.releaseLock(); }
}

/** Anula un viaje recién iniciado por error: solo si está en curso y no tiene ningún ganadero apuntado.
 *  No genera fichero para el programa y no bloquea la ruta. El PDF del DeCA queda renombrado como ANULADO. */
function anularViaje(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = buscar_('VIAJES', 'VIAJE_ID', req.viajeId);
    if (!v || v.CONDUCTOR_ID !== c.CONDUCTOR_ID) throw new Error('Viaje no encontrado');
    if (String(v.ESTADO).indexOf('ANULADO') === 0) return { ok: true };
    if (String(v.ESTADO).indexOf('EN CURSO') !== 0) throw new Error('Este viaje ya está cerrado: no se puede anular.');
    const n = tabla_('RECOGIDAS').filas.filter(function (r) { return r.VIAJE_ID === v.VIAJE_ID; }).length;
    if (n) throw new Error('Este viaje ya tiene ganaderos apuntados: no se puede anular. Termínalo o avisa a la oficina.');
    actualizar_('VIAJES', v._fila, { ESTADO: 'ANULADO', HORA_FIN: ahora_('HH:mm'), LITROS_REALES: 0, PESO_REAL_KG: 0, EXPORTADO: 'NO' });
    try {
      const m = String(v.DECA_URL).match(/id=([\w-]+)/);
      if (m) { const f = DriveApp.getFileById(m[1]); f.setName('ANULADO - ' + f.getName()); }
    } catch (e) {}
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** Lo recogido hoy en estas rutas (cualquier conductor): viajes y, por ganadero|especie, litros y hora. */
function hoy_(rutas) {
  const hoy = ahora_('dd/MM/yyyy'), rs = (rutas || []).map(pad2_);
  const viajes = tabla_('VIAJES').filas.filter(function (v) {
    return v.FECHA === hoy && String(v.ESTADO).indexOf('ANULADO') !== 0 && split_(v.RUTAS).map(pad2_).some(function (r) { return rs.indexOf(r) >= 0; });
  });
  const ids = {}; viajes.forEach(function (v) { ids[v.VIAJE_ID] = v; });
  const recs = {}, litros = {};
  tabla_('RECOGIDAS').filas.forEach(function (r) {
    const v = ids[r.VIAJE_ID]; if (!v || String(r.NO_RECOGIDO).toUpperCase() === 'SI') return;
    const l = numES_(r.LITROS); litros[r.VIAJE_ID] = (litros[r.VIAJE_ID] || 0) + l;
    recs[String(r.CODIGO).trim() + '|' + especie_(r.ESPECIE)] = { l: l, hora: r.HORA, viaje: r.VIAJE_ID, deca: v.DECA_NUM };
  });
  return {
    viajes: viajes.map(function (v) { return { id: v.VIAJE_ID, deca: v.DECA_NUM, rutas: v.RUTAS, ini: v.HORA_INICIO, fin: v.HORA_FIN,
      estado: v.ESTADO, conductor: v.CONDUCTOR_ID, matricula: v.MATRICULA, litros: litros[v.VIAJE_ID] || 0 }; }),
    recs: recs
  };
}

/** Consulta (solo lectura) de los últimos días del conductor: por día y viaje, cada ganadero con sus litros y destinos. */
function historial(req, c) {
  const dias = Math.min(Math.max(Number(req.dias) || 15, 1), 45);
  const lim = new Date(); lim.setHours(0, 0, 0, 0); lim.setDate(lim.getDate() - (dias - 1));
  const fecha = function (s) { const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null; };
  const viajes = tabla_('VIAJES').filas.filter(function (v) { const d = fecha(v.FECHA); return v.CONDUCTOR_ID === c.CONDUCTOR_ID && d && d >= lim; });
  const porId = {}; viajes.forEach(function (v) { porId[v.VIAJE_ID] = { id: v.VIAJE_ID, fecha: v.FECHA, deca: v.DECA_NUM, matricula: v.MATRICULA,
    rutas: v.RUTAS, estado: v.ESTADO, ini: v.HORA_INICIO, fin: v.HORA_FIN, recs: [] }; });
  const nombres = {}; tabla_('GANADEROS').filas.forEach(function (g) { const k = String(g.CODIGO).trim(); if (!nombres[k]) nombres[k] = g.NOMBRE; });
  tabla_('RECOGIDAS').filas.forEach(function (r) {
    const v = porId[r.VIAJE_ID]; if (!v) return;
    v.recs.push({ ruta: r.RUTA, ord: r.ORDEN, cod: r.CODIGO, nom: nombres[String(r.CODIGO).trim()] || '', esp: especie_(r.ESPECIE),
      l: numES_(r.LITROS), t: r.TEMPERATURA, mue: r.MUESTRA, d1: r.DESTINO, l1: numES_(r.LITROS_DESTINO), d2: r.DESTINO_2, l2: numES_(r.LITROS_DESTINO_2), d3: r.DESTINO_3 || '', l3: numES_(r.LITROS_DESTINO_3),
      sin: String(r.NO_RECOGIDO).toUpperCase() === 'SI', inc: r.INCIDENCIA_TEMP, hora: r.HORA, cor: r.CORREGIDO || '' });
  });
  const lista = Object.keys(porId).map(function (k) { return porId[k]; }).filter(function (v) { return v.recs.length; });
  lista.forEach(function (v) { v.recs.sort(function (a, b) { return a.ruta === b.ruta ? (Number(a.ord) || 0) - (Number(b.ord) || 0) : (a.ruta < b.ruta ? -1 : 1); }); });
  const dias_ = {};
  lista.forEach(function (v) { (dias_[v.fecha] = dias_[v.fecha] || { fecha: v.fecha, viajes: [] }).viajes.push(v); });
  const out = Object.keys(dias_).map(function (k) { return dias_[k]; })
    .sort(function (a, b) { return fecha(b.fecha) - fecha(a.fecha); });
  out.forEach(function (d) { d.viajes.sort(function (a, b) { return String(a.ini) < String(b.ini) ? -1 : 1; }); });
  return { ok: true, dias: out, hasta: ahora_('dd/MM/yyyy HH:mm') };
}
/** "2.005" → 2005 · "4,5" → 4.5 · "" → 0 */
function numES_(x) {
  let t = String(x == null ? '' : x).trim(); if (!t) return 0;
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, '');
  return Number(t.replace(',', '.')) || 0;
}

/* ───────────── DeCA ───────────── */

/** DeCA electrónico (Orden FOM/2861/2012 y Resolución DGTCF de 5/6/2026, obligatoria desde el 5/10/2026):
 *  PDF nativo, guardado en Drive con enlace público de descarga directa (HTTPS, sin credenciales) y, dentro del propio PDF,
 *  el código QR con esa dirección. Como la dirección depende del archivo, primero se crea el PDF y luego se rehace con el QR
 *  (hace falta el servicio avanzado «Drive API» activado en Apps Script y el archivo Qr.gs; si falta algo, el PDF queda sin QR y se avisa). */
function crearDeca_(d) {
  const cfg = d.cfg;
  const raiz = DriveApp.getFolderById(cfg.CARPETA_DECA_ID);
  const mes = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const it = raiz.getFoldersByName(mes);
  const carpeta = it.hasNext() ? it.next() : raiz.createFolder(mes);
  const emitido = ahora_('dd/MM/yyyy HH:mm');
  const f = carpeta.createFile(htmlDeca_(d, emitido, '', ''));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?export=download&id=' + f.getId();
  try {
    if (typeof Drive === 'undefined') throw new Error('Falta activar el servicio avanzado Drive API (Servicios + → Drive API)');
    const qr = qrPngDataUri_(url, 4);
    Drive.Files.update({}, f.getId(), htmlDeca_(d, emitido, url, qr));
  } catch (e) { Logger.log('DeCA ' + d.num + ' sin QR: ' + e.message); }
  return { id: f.getId(), url: url };
}
/** PDF del DeCA. url/qr vacíos = primera pasada (aún no se sabe la dirección del archivo). */
function htmlDeca_(d, emitido, url, qr) {
  const cfg = d.cfg;
  const fila = function (k, v) { return '<tr><td style="width:34%;background:#f6eef0"><b>' + k + '</b></td><td>' + esc_(v) + '</td></tr>'; };
  const pueblos = origenRutas_(d.rutas);
  const html =
    '<html><body style="font-family:Arial,sans-serif;font-size:10.5pt;color:#111">' +
    '<table style="width:100%;border-bottom:4px solid #862133;margin-bottom:8px"><tr><td style="vertical-align:top">' +
    '<div style="font-size:9pt;color:#862133"><b>' + esc_(cfg.CARGADOR_NOMBRE) + '</b></div>' +
    '<div style="font-size:17pt;font-weight:bold;color:#862133">Documento de Control Administrativo (DeCA)</div>' +
    '<div style="font-size:9pt">Transporte público de mercancías por carretera · Orden FOM/2861/2012</div>' +
    '<div style="font-size:12pt;margin-top:6px"><b>Nº ' + esc_(d.num) + '</b> · versión 1</div></td>' +
    (qr ? '<td style="width:150px;text-align:right;vertical-align:top"><img src="' + qr + '" width="140" height="140"></td>' : '') + '</tr></table>' +
    '<table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;width:100%">' +
    fila('Fecha del transporte', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy')) +
    fila('Fecha y hora de emisión', emitido) +
    fila('a) Cargador contractual', cfg.CARGADOR_NOMBRE + ' · NIF ' + cfg.CARGADOR_NIF + ' · ' + cfg.CARGADOR_DIRECCION) +
    fila('b) Transportista efectivo', String(d.cond.TRANSPORTISTA) + ' · NIF ' + String(d.cond.NIF_TRANSPORTISTA)) +
    fila('c) Origen', 'Recogida en explotaciones ganaderas de la ruta ' + d.rutas.map(function (r) { return Number(r); }).join(' y ') + (pueblos ? ' (' + pueblos + ')' : '')) +
    fila('c) Destino', cfg.CARGADOR_NOMBRE + ' · NIF ' + cfg.CARGADOR_NIF + ' · ' + cfg.CARGADOR_DIRECCION) +
    fila('d) Mercancía · naturaleza', cfg.MERCANCIA || 'Leche cruda') +
    fila('d) Mercancía · peso bruto (estimado al inicio)', d.kg.toLocaleString('es-ES') + ' kg (' + d.litros.toLocaleString('es-ES') + ' l)') +
    fila('d) Mercancía · bultos', '1 (cisterna)') +
    fila('e) Autorización especial de circulación', 'No requerida') +
    fila('f) Fecha del transporte', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy')) +
    fila('g) Vehículo (matrícula)', d.matricula) +
    (d.letraQ ? fila('g) Cisterna (código Letra Q)', d.letraQ) : '') +
    fila('g) Conductor', String(d.cond.NOMBRE || '')) +
    fila('h) Observaciones', d.nota || 'Sin observaciones') +
    '</table>' +
    (url ? '<p style="font-size:8.5pt;margin-top:8px">Documento accesible en:<br><b>' + esc_(url) + '</b></p>' : '') +
    '<p style="font-size:8pt;color:#555">Documento electrónico emitido antes del inicio del servicio. El peso real se registra al finalizar la recogida. ' +
    'Generado por la aplicación de recogida de leche de ' + esc_(cfg.CARGADOR_NOMBRE) + ' (Resolución DGTCF de 5 de junio de 2026).</p>' +
    '</body></html>';
  return Utilities.newBlob(html, 'text/html', 'deca.html').getAs(MimeType.PDF).setName('DeCA ' + d.num + '.pdf');
}
/** Poblaciones de las explotaciones de esas rutas, para el lugar de origen (máx. 8). */
function origenRutas_(rutas) {
  try {
    const rs = (rutas || []).map(pad2_), vistos = {}, lista = [];
    tabla_('GANADEROS').filas.forEach(function (g) {
      const p = String(g.POBLACION || '').trim(); if (!p || rs.indexOf(pad2_(g.RUTA)) < 0 || String(g.ACTIVO).toUpperCase() === 'NO' || vistos[p.toUpperCase()]) return;
      vistos[p.toUpperCase()] = 1; lista.push(p);
    });
    return lista.length > 8 ? lista.slice(0, 8).join(', ') + '…' : lista.join(', ');
  } catch (e) { return ''; }
}

/** Guarda tal cual las recogidas cuyo viaje ya no existe (hoja SIN_VIAJE), para no perderlas ni bloquear el móvil. */
function guardarSinViaje_(recs, c) {
  let sh = SS.getSheetByName('SIN_VIAJE');
  const cab = ['RECIBIDO_EN', 'CONDUCTOR_ID', 'VIAJE_ID', 'REC_ID', 'FECHA', 'HORA', 'RUTA', 'CODIGO', 'ESPECIE', 'LITROS', 'TEMPERATURA', 'MUESTRA', 'DESTINO', 'DATOS'];
  if (!sh) {
    sh = SS.insertSheet('SIN_VIAJE');
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    sh.getRange(2, 1, sh.getMaxRows() - 1, cab.length).setNumberFormat('@');
  }
  const ahora = ahora_('dd/MM/yyyy HH:mm:ss');
  const filas = recs.map(function (r) {
    return [ahora, c.CONDUCTOR_ID, r.viajeId, r.recId, r.fecha || '', r.hora || '', pad2_(r.ruta), r.cod, r.esp, r.sin ? 'NO RECOGIDO' : r.litros, r.temp || '', r.muestra || '', r.dest || '', JSON.stringify(r)];
  });
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, cab.length).setValues(filas);
}

/* ───────────── Recepción en fábrica (PIN de ROL RECEPCION; la oficina también puede) ─────────────
   Por camión: cada compartimento con su parcial de contador, controles (T, pH, ºD, visual) y reparto a silos
   (principal + arrastres / adelanto a mezcla). Al final, el total del contador del camión, limpieza y filtro.
   Hoja RECEPCIONES: una fila por cada reparto compartimento → silo. Hoja RECEPCION_CAMION: una por camión.
   Fichero para el programa: D (descarga) + aammdd + nº (D2610011.TXT), en la misma carpeta de exportación. */

const CAB_RECEP = ['REC_ID', 'FECHA', 'HORA', 'VIAJE_ID', 'DECA_NUM', 'MATRICULA', 'CONDUCTOR', 'RUTA', 'ORDEN_DESCARGA', 'COMPARTIMENTO', 'ESPECIE',
  'LITROS_DECLARADOS', 'LITROS_CONTADOR', 'DEPOSITO', 'LITROS_DEPOSITO', 'TIPO', 'TEMPERATURA', 'PH', 'DORNIC', 'VISUAL', 'USUARIO', 'N_COMP_CISTERNA'];
const CAB_AGUA = ['FECHA', 'CODIGO', 'DEPOSITO', 'LITROS', 'DESDE_VIAJE', 'USUARIO', 'HORA', 'FICHERO'];
const CAB_RCAL = ['FECHA', 'HORA', 'VIAJE_ID', 'DECA_NUM', 'MATRICULA', 'CONDUCTOR', 'RUTA', 'ORDEN', 'ID', 'ESPECIE', 'N_COMP_CISTERNA', 'LITROS', 'DEPOSITOS',
  'TEMPERATURA', 'PH', 'DORNIC', 'VISUAL', 'LIMPIEZA_CISTERNA', 'LIMPIEZA_FILTRO', 'OBSERVACIONES', 'USUARIO', 'ALTA'];
const CAB_MEZCLA = ['FECHA', 'VIAJE_ID', 'COMPARTIMENTO', 'MEZCLA', 'USUARIO', 'HORA'];
const CAB_PARTE = ['FECHA', 'OBSERVACIONES', 'OBS_USUARIO', 'VERIFICADO_POR', 'HORA_VERIFICACION'];
const CAB_VENTA = ['ID', 'FECHA', 'HORA', 'DEPOSITO', 'ESPECIE', 'LITROS', 'CLIENTE', 'USUARIO'];
const CAB_RECEP_CAM = ['VIAJE_ID', 'FECHA', 'HORA', 'DECA_NUM', 'MATRICULA', 'CONDUCTOR', 'RUTA', 'TOTAL_CONTADOR', 'SUMA_PARCIALES', 'DIFERENCIA',
  'LIMPIEZA_CISTERNA', 'FILTRO', 'OBSERVACIONES', 'USUARIO', 'FICHERO', 'LITROS_CARTILLA', 'DIF_CARTILLA', 'LITROS_CISTERNA', 'DIF_CISTERNA'];

function hoja_(nombre, cab, filas) {
  let sh = SS.getSheetByName(nombre);
  if (!sh) {
    sh = SS.insertSheet(nombre);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    sh.setFrozenRows(1);
    if (filas && filas.length) sh.getRange(2, 1, filas.length, cab.length).setValues(filas);
  }
  return sh;
}
function depositos_() {
  hoja_('DEPOSITOS', ['DEPOSITO', 'CAPACIDAD', 'USO', 'ACTIVO', 'NOTA'], [
    ['1', 10000, 'Cabra', 'SI', 'Silo pequeño: cabra u otras cantidades pequeñas'],
    ['2', 25000, 'Mezcla', 'SI', 'Mezcla (cabra, vaca y oveja)'],
    ['3', 25000, 'Oveja + D.O.', 'SI', 'Nunca cabra ni vaca'],
    ['4', 25000, 'Comodín', 'SI', 'Fines de semana, excedentes, cisternas enteras'],
    ['5', 25000, 'Comodín', 'SI', '']]);
  columnas_('DEPOSITOS', ['LETRA_Q'], 'CAPACIDAD');
  return tabla_('DEPOSITOS').filas.filter(function (d) { return String(d.ACTIVO).toUpperCase() !== 'NO' && String(d.DEPOSITO).trim(); })
    .map(function (d) { return { id: String(d.DEPOSITO).trim().replace(/^S/i, ''), cap: Number(String(d.CAPACIDAD).replace(/\./g, '')) || 0, uso: String(d.USO || ''), letraQ: String(d.LETRA_Q || '').trim() }; });
}
function loginRecepcion(req, c) {
  return { ok: true, rol: 'RECEPCION', roles: roles_(c), usuario: { id: c.CONDUCTOR_ID, nombre: String(c.NOMBRE) }, depositos: depositos_() };
}
function loginCalidad(req, c) {
  return { ok: true, rol: 'CALIDAD', roles: roles_(c), usuario: { id: c.CONDUCTOR_ID, nombre: String(c.NOMBRE) }, depositos: depositos_() };
}
/** Calidad: mezcla C (correcto) / I (incorrecto) de un compartimento recibido. Vacío = sin verificar. */
function guardarMezcla(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = hoja_('MEZCLAS', CAB_MEZCLA);
    hoja_('RECEPCION_CALIDAD', CAB_RCAL); hoja_('RECEPCION_CAMION', CAB_RECEP_CAM);
    const cam = buscar_('RECEPCION_CALIDAD', 'VIAJE_ID', req.viajeId) || buscar_('RECEPCION_CAMION', 'VIAJE_ID', req.viajeId);
    if (!cam) throw new Error('Ese camión no está recibido.');
    const comp = String(req.comp || ''), m = String(req.mezcla || '').toUpperCase();
    if (m && !/^[CI]$/.test(m)) throw new Error('Mezcla: C o I.');
    const ya = tabla_('MEZCLAS').filas.filter(function (r) { return r.VIAJE_ID === req.viajeId && r.COMPARTIMENTO === comp; })[0];
    const fila = { FECHA: cam.FECHA, VIAJE_ID: req.viajeId, COMPARTIMENTO: comp, MEZCLA: m, USUARIO: String(c.NOMBRE), HORA: ahora_('HH:mm') };
    if (ya) actualizar_('MEZCLAS', ya._fila, fila);
    else { sh.getRange(sh.getLastRow() + 1, 1, 1, CAB_MEZCLA.length).setNumberFormat('@').setValues([CAB_MEZCLA.map(function (h) { return String(fila[h]); })]); }
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function parteDia_(fecha, obj) {
  const sh = hoja_('PARTE_DIA', CAB_PARTE), ya = tabla_('PARTE_DIA').filas.filter(function (r) { return r.FECHA === fecha; })[0];
  if (ya) actualizar_('PARTE_DIA', ya._fila, obj);
  else { const f = Object.assign({ FECHA: fecha }, obj); sh.getRange(sh.getLastRow() + 1, 1, 1, CAB_PARTE.length).setNumberFormat('@').setValues([CAB_PARTE.map(function (h) { return f[h] !== undefined ? String(f[h]) : ''; })]); }
}
/** Calidad firma la verificación del parte de recepción de un día (equivale a la firma en verde). */
function verificarDia(req, c) {
  const fecha = String(req.fecha || ahora_('dd/MM/yyyy'));
  parteDia_(fecha, req.quitar ? { VERIFICADO_POR: '', HORA_VERIFICACION: '' } : { VERIFICADO_POR: String(c.NOMBRE), HORA_VERIFICACION: ahora_('dd/MM/yyyy HH:mm') });
  return { ok: true };
}
/** Acciones correctivas / observaciones del día (pie del parte). */
function guardarObsDia(req, c) {
  parteDia_(String(req.fecha || ahora_('dd/MM/yyyy')), { OBSERVACIONES: String(req.obs || '').trim(), OBS_USUARIO: String(c.NOMBRE) });
  return { ok: true };
}
/** Leche vendida (salida de un depósito a un cliente). */
function guardarVenta(req, c) {
  const v = req.venta || {}, l = Math.round(Number(v.l) || 0);
  if (!(l > 0)) throw new Error('Faltan los litros.');
  if (!String(v.cliente || '').trim()) throw new Error('Falta a quién se vende.');
  if (!/^[DOCV]$/.test(String(v.esp || ''))) throw new Error('Falta la especie.');
  const sh = hoja_('VENTAS_LECHE', CAB_VENTA), f = { ID: 'VL' + Date.now().toString(36).toUpperCase(), FECHA: String(req.fecha || ahora_('dd/MM/yyyy')), HORA: ahora_('HH:mm'),
    DEPOSITO: String(v.dep || ''), ESPECIE: v.esp, LITROS: l, CLIENTE: String(v.cliente).trim(), USUARIO: String(c.NOMBRE) };
  sh.getRange(sh.getLastRow() + 1, 1, 1, CAB_VENTA.length).setNumberFormat('@').setValues([CAB_VENTA.map(function (h) { return String(f[h]); })]);
  return { ok: true, id: f.ID };
}
function borrarVenta(req, c) {
  const ya = buscar_('VENTAS_LECHE', 'ID', req.id); if (!ya) throw new Error('Esa venta no existe.');
  SS.getSheetByName('VENTAS_LECHE').deleteRow(ya._fila); return { ok: true };
}
/** Camiones de hoy con sus compartimentos, lo declarado por el recogedor en cada uno y la recepción ya hecha. */
function recepViajes(req, c) {
  hoja_('RECEPCIONES', CAB_RECEP); hoja_('RECEPCION_CAMION', CAB_RECEP_CAM); hoja_('AGUA', CAB_AGUA);
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const hoyTxt = Utilities.formatDate(hoy, TZ, 'dd/MM/yyyy');
  // Otro día (oficina y calidad, para revisar o imprimir el parte): solo lo que se recibió ese día
  const otroDia = req && req.fecha && fechaNum_(req.fecha) && String(req.fecha) !== hoyTxt ? String(req.fecha) : '';
  const fechas = [otroDia || hoyTxt];
  const conds = {}; tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = String(x.NOMBRE); });
  const camTab = tabla_('CAMIONES').filas, compTab = tabla_('COMPARTIMENTOS').filas;
  // Hoy: todos los viajes. Días anteriores (hasta RECEP_DIAS, 4 por defecto; p. ej. el fin de semana que se descarga el lunes):
  // los que aún no se han recibido y los recibidos hoy. Nunca antes de RECEP_DESDE (arranque de la recepción con la app).
  const cfg0 = cfg_(), diasAtras = String(cfg0.RECEP_DIAS || '') !== '' && Number(cfg0.RECEP_DIAS) >= 0 ? Number(cfg0.RECEP_DIAS) : 4;
  const limite = new Date(hoy); limite.setDate(limite.getDate() - diasAtras);
  const desde = Math.max(fechaNum_(cfg0.RECEP_DESDE || '02/10/2026'), fechaNum_(Utilities.formatDate(limite, TZ, 'dd/MM/yyyy')));
  const recibido = {}; tabla_('RECEPCION_CAMION').filas.forEach(function (r) { recibido[r.VIAJE_ID] = String(r.FECHA); });
  const viajes = tabla_('VIAJES').filas.filter(function (v) {
    if (String(v.ESTADO).indexOf('ANULADO') === 0 || String(v.ESTADO).indexOf('TRASVASE') >= 0) return false;
    if (otroDia) return recibido[v.VIAJE_ID] === otroDia || (v.FECHA === otroDia && !recibido[v.VIAJE_ID]);   // recibidos ese día y los de ese día aún sin recibir
    if (fechas.indexOf(v.FECHA) >= 0) return true;
    const f = fechaNum_(v.FECHA); if (!f || f < desde) return false;
    return !recibido[v.VIAJE_ID] || recibido[v.VIAJE_ID] === fechas[0];
  });
  const ids = {}; viajes.forEach(function (v) { ids[v.VIAJE_ID] = true; });
  // Lo declarado por compartimento (sumando destinos 1, 2 y 3 de cada recogida)
  const decl = {}, cart = {};   // decl: «litros cisterna» por compartimento (destinos); cart: «litros cartilla» por especie
  tabla_('RECOGIDAS').filas.forEach(function (r) {
    if (!ids[r.VIAJE_ID] || String(r.NO_RECOGIDO).toUpperCase() === 'SI') return;
    const e = String(r.ESPECIE || '').trim().toUpperCase(); if (e) { cart[r.VIAJE_ID] = cart[r.VIAJE_ID] || {}; cart[r.VIAJE_ID][e] = (cart[r.VIAJE_ID][e] || 0) + numES_(r.LITROS); }
    [[r.DESTINO, r.LITROS_DESTINO], [r.DESTINO_2, r.LITROS_DESTINO_2], [r.DESTINO_3, r.LITROS_DESTINO_3]].forEach(function (d) {
      if (!String(d[0] || '').trim()) return;
      const k = r.VIAJE_ID + '|' + String(d[0]).trim(); decl[k] = (decl[k] || 0) + numES_(d[1]);
    });
  });
  const hechas = {}; tabla_('RECEPCIONES').filas.forEach(function (r) { (hechas[r.VIAJE_ID] = hechas[r.VIAJE_ID] || []).push(r); });
  const camiones = {}; tabla_('RECEPCION_CAMION').filas.forEach(function (r) { camiones[r.VIAJE_ID] = r; });
  const salida = viajes.map(function (v) {
    const k = camTab.filter(function (x) { return String(x.CAMION_ID) === String(v.CAMION_ID); })[0];
    const pref = String(+pad2_(v.PREFIJO_DESTINO) || '');
    const comps = (k ? letrasRuta_(camionObj_(k, compTab).letras, v.PREFIJO_DESTINO) : []).map(function (x) {
      const id = pref + x.l; return { id: id, n: x.n, esp: x.esp, cap: x.cap, decl: Math.round(decl[v.VIAJE_ID + '|' + id] || 0) };
    });
    // Destinos que el recogedor usó y no están en el camión (por si acaso)
    Object.keys(decl).forEach(function (kk) { const p = kk.split('|'); if (p[0] === v.VIAJE_ID && !comps.some(function (x) { return x.id === p[1]; })) comps.push({ id: p[1], esp: [], cap: 0, decl: Math.round(decl[kk]) }); });
    return { id: v.VIAJE_ID, fecha: v.FECHA, deca: v.DECA_NUM, matricula: v.MATRICULA, conductor: conds[v.CONDUCTOR_ID] || v.CONDUCTOR_ID,
      rutas: v.RUTAS, ruta: pad2_(v.PREFIJO_DESTINO), estado: v.ESTADO, ini: v.HORA_INICIO, fin: v.HORA_FIN, comps: comps, cart: cart[v.VIAJE_ID] || {},
      nCompCam: k ? (Number(k.N_COMPARTIMENTOS) || Math.max.apply(null, [0].concat(camionObj_(k, compTab).letras.map(function (x) { return x.n; })))) : 0,
      recep: recepObj_(hechas[v.VIAJE_ID], camiones[v.VIAJE_ID]) };
  });
  // Entradas sin viaje de la app (cisterna de otro proveedor, agua…) de hoy
  Object.keys(camiones).forEach(function (id) {
    const r = camiones[id]; if (ids[id] || fechas.indexOf(r.FECHA) < 0) return;
    salida.push({ id: id, fecha: r.FECHA, deca: r.DECA_NUM, matricula: r.MATRICULA, conductor: r.CONDUCTOR, rutas: r.RUTA, ruta: r.RUTA, estado: 'OTRA ENTRADA',
      ini: r.HORA, fin: '', comps: [], otra: true, recep: recepObj_(hechas[id], r) });
  });
  const cfg = cfg_();
  hoja_('MEZCLAS', CAB_MEZCLA); hoja_('PARTE_DIA', CAB_PARTE); hoja_('VENTAS_LECHE', CAB_VENTA);
  // Cisternas de otros proveedores ya usadas (para elegirlas de una lista en «Otra entrada»)
  const espX = {}; tabla_('RECEPCIONES').filas.forEach(function (r) { if (/^X/.test(r.VIAJE_ID) && !espX[r.VIAJE_ID] && r.ESPECIE !== 'AGUA') espX[r.VIAJE_ID] = String(r.ESPECIE); });
  const vistos = {}, externos = [];
  tabla_('RECEPCION_CAMION').filas.filter(function (r) { return /^X/.test(r.VIAJE_ID) && String(r.CONDUCTOR).trim(); }).reverse().forEach(function (r) {
    const k = (String(r.CONDUCTOR).trim() + '|' + String(r.MATRICULA).trim()).toUpperCase(); if (vistos[k] || externos.length >= 40) return; vistos[k] = 1;
    externos.push({ prov: String(r.CONDUCTOR).trim(), mat: String(r.MATRICULA || '').trim(), esp: espX[r.VIAJE_ID] || '' });
  });
  // Registro de calidad por especie (Recepción 2): no va al programa de gestión
  hoja_('RECEPCION_CALIDAD', CAB_RCAL);
  const ids2 = {}; salida.forEach(function (x) { ids2[x.id] = 1; });
  // Último nº de compartimento usado por cada camión (matrícula) y especie: se propone la próxima vez
  const compsCal = {}, calTodo = tabla_('RECEPCION_CALIDAD').filas;
  calTodo.forEach(function (r) { const m = String(r.MATRICULA || '').toUpperCase().replace(/[^A-Z0-9]/g, ''), n = String(r.N_COMP_CISTERNA || '').trim();
    if (m && n) (compsCal[m] = compsCal[m] || {})[String(r.ESPECIE).toUpperCase()] = n; });
  const calidad = {}; calTodo.forEach(function (r) {
    if (!ids2[r.VIAJE_ID] && fechas.indexOf(String(r.FECHA)) < 0) return;
    const k = calidad[r.VIAJE_ID] = calidad[r.VIAJE_ID] || { fecha: String(r.FECHA), hora: String(r.HORA), usuario: String(r.USUARIO), limpieza: String(r.LIMPIEZA_CISTERNA || ''),
      filtro: String(r.LIMPIEZA_FILTRO || ''), obs: String(r.OBSERVACIONES || ''), matricula: String(r.MATRICULA || ''), conductor: String(r.CONDUCTOR || ''),
      alta: String(r.ALTA || ''), comps: [] };
    k.comps.push({ id: String(r.ID), esp: String(r.ESPECIE), ord: Number(r.ORDEN) || 0, nComp: String(r.N_COMP_CISTERNA || ''), contador: numES_(r.LITROS), temp: r.TEMPERATURA, ph: r.PH,
      dornic: r.DORNIC, visual: String(r.VISUAL || ''), reps: String(r.DEPOSITOS || '').split('|').filter(String).map(function (t) { const q = t.split(':'); return { dep: q[0], l: numES_(q[1]), tipo: q[2] || 'PRINCIPAL' }; }) });
  });
  Object.keys(calidad).forEach(function (id) { calidad[id].comps.sort(function (a, b) { return a.ord - b.ord; }); });
  // Recepción 2: camiones activos con sus compartimentos físicos (nº y capacidad) y el último conductor que lo llevó
  const ultCond = {}; tabla_('VIAJES').filas.forEach(function (v) { if (v.CAMION_ID) ultCond[v.CAMION_ID] = conds[v.CONDUCTOR_ID] || v.CONDUCTOR_ID; });
  const camionesAct = camTab.filter(function (k) { return String(k.ACTIVO).toUpperCase() !== 'NO' && String(k.MATRICULA || '').trim() && !/PENDIENTE/i.test(String(k.MATRICULA)); }).map(function (k) {
    const o = camionObj_(k, compTab), porN = {};
    o.letras.forEach(function (x) { if (x.n) porN[x.n] = Math.max(porN[x.n] || 0, x.cap || 0); });
    const n = Math.max(Number(k.N_COMPARTIMENTOS) || 0, Object.keys(porN).length ? Math.max.apply(null, Object.keys(porN).map(Number)) : 0);
    return { id: o.id, matricula: o.matricula, conductor: ultCond[o.id] || String(k.TITULAR || ''), desc: o.desc, nComp: n,
      comps: Array.from({ length: n }, function (_, i) { return { n: i + 1, cap: porN[i + 1] || 0 }; }) };
  });
  const mezclas = {}; tabla_('MEZCLAS').filas.forEach(function (r) { if (r.FECHA === fechas[0]) mezclas[r.VIAJE_ID + '|' + r.COMPARTIMENTO] = String(r.MEZCLA); });
  const pd = tabla_('PARTE_DIA').filas.filter(function (r) { return r.FECHA === fechas[0]; })[0] || {};
  return { ok: true, fecha: fechas[0], hoy: !otroDia, externos: externos, viajes: salida, depositos: depositos_(), hasta: ahora_('HH:mm'), calidad: calidad, compsCal: compsCal, camiones: camionesAct,
    mezclas: mezclas, parte: { obs: String(pd.OBSERVACIONES || ''), verificado: String(pd.VERIFICADO_POR || ''), horaVerif: String(pd.HORA_VERIFICACION || '') },
    ventas: tabla_('VENTAS_LECHE').filas.filter(function (r) { return r.FECHA === fechas[0]; }).map(function (r) { return { id: r.ID, dep: String(r.DEPOSITO), esp: String(r.ESPECIE), l: numES_(r.LITROS), cliente: String(r.CLIENTE), hora: String(r.HORA), usuario: String(r.USUARIO) }; }),
    aguaDef: { arranque: Number(cfg.AGUA_ARRANQUE) || 70, final: Number(cfg.AGUA_FINAL) || 70 },
    aguaHoy: tabla_('AGUA').filas.filter(function (r) { return r.FECHA === fechas[0]; })
      .map(function (r) { return { cod: String(r.CODIGO), dep: String(r.DEPOSITO), l: numES_(r.LITROS), desde: String(r.DESDE_VIAJE || '') }; }) };
}
function recepObj_(filas, cam) {
  if (!filas && !cam) return null;
  const comps = {}, orden = [];
  const agua = [];
  (filas || []).forEach(function (r) {
    const id = String(r.COMPARTIMENTO);
    if (String(r.ESPECIE).toUpperCase() === 'AGUA' && /^AGUA/.test(String(r.TIPO))) { agua.push({ dep: String(r.DEPOSITO), l: numES_(r.LITROS_DEPOSITO), tipo: String(r.TIPO).replace(/^AGUA\s*/, '') }); return; }
    if (!comps[id]) { comps[id] = { id: id, esp: String(r.ESPECIE), ord: Number(r.ORDEN_DESCARGA) || 0, nComp: String(r.N_COMP_CISTERNA || ''), contador: numES_(r.LITROS_CONTADOR), temp: r.TEMPERATURA, ph: r.PH, dornic: r.DORNIC, visual: r.VISUAL, reps: [] }; orden.push(id); }
    comps[id].reps.push({ dep: String(r.DEPOSITO), l: numES_(r.LITROS_DEPOSITO), tipo: String(r.TIPO) });
  });
  return { comps: orden.map(function (id) { return comps[id]; }), agua: agua, total: cam ? numES_(cam.TOTAL_CONTADOR) : 0, limpieza: cam ? cam.LIMPIEZA_CISTERNA : '',
    filtro: cam ? cam.FILTRO : '', obs: cam ? cam.OBSERVACIONES : '', hora: cam ? cam.HORA : '', usuario: cam ? cam.USUARIO : '', fichero: cam ? cam.FICHERO : '' };
}
/** Guarda (o rehace) la recepción completa de un camión y su fichero E…TXT. */
function guardarRecepcion(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const r = req.recep || {};
    const comps = (r.comps || []).filter(function (x) { return Number(x.contador) > 0; });
    if (!comps.length) throw new Error('No hay ningún compartimento con litros de contador.');
    comps.forEach(function (x) {
      const suma = (x.reps || []).reduce(function (a, p) { return a + (Number(p.l) || 0); }, 0);
      if (Math.round(suma) !== Math.round(Number(x.contador))) throw new Error('En ' + x.id + ' el reparto a depósitos (' + suma + ' l) no suma lo del contador (' + x.contador + ' l).');
      (x.reps || []).forEach(function (p) { if (!String(p.dep || '').trim()) throw new Error('Falta el depósito de destino en ' + x.id + '.'); });
      const distintos = {}; (x.reps || []).forEach(function (p) { distintos[String(p.dep).trim()] = 1; });
      if (Object.keys(distintos).length > 3) throw new Error('En ' + x.id + ' hay más de 3 depósitos: el programa admite como máximo 3 por compartimento.');
    });
    if (/^X/.test(String(r.viajeId || ''))) {   // cisterna externa: cada especie es un solo destino (VC, VV…) con máx. 3 depósitos
      const porEsp = {}; comps.forEach(function (x) { (x.reps || []).forEach(function (p) { const e = 'V' + String(x.esp).toUpperCase(); (porEsp[e] = porEsp[e] || {})[String(p.dep).trim()] = 1; }); });
      Object.keys(porEsp).forEach(function (e) { if (Object.keys(porEsp[e]).length > 3) throw new Error(e + ' va a más de 3 depósitos entre todos sus compartimentos: el programa admite como máximo 3.'); });
    }
    const shR = hoja_('RECEPCIONES', CAB_RECEP), shC = hoja_('RECEPCION_CAMION', CAB_RECEP_CAM);
    columnas_('RECEPCION_CAMION', ['LITROS_CARTILLA', 'DIF_CARTILLA', 'LITROS_CISTERNA', 'DIF_CISTERNA'], 'FICHERO');
    columnas_('RECEPCIONES', ['N_COMP_CISTERNA'], 'USUARIO');
    // Nº de compartimento de la cisterna (posición física: 1 junto a la cabina). Distinto del orden de descarga. No va al fichero.
    const nUsados = {}; comps.forEach(function (x) { const n = String(x.nComp || '').trim(); if (!n) return;
      if (!/^\d{1,2}$/.test(n)) throw new Error('Nº de compartimento de la cisterna no válido en ' + x.id + '.');
      if (nUsados[n]) throw new Error('El compartimento nº ' + n + ' de la cisterna está puesto en ' + nUsados[n] + ' y en ' + x.id + '.'); nUsados[n] = x.id; });
    const cabR = tabla_('RECEPCIONES').cab;
    let v = buscar_('VIAJES', 'VIAJE_ID', r.viajeId);
    const otra = !v;
    if (otra) {
      if (!/^X/.test(String(r.viajeId || ''))) throw new Error('Viaje no encontrado');
      v = { VIAJE_ID: r.viajeId, DECA_NUM: r.deca || '', MATRICULA: r.matricula || '', CONDUCTOR_ID: r.proveedor || 'OTRA ENTRADA', PREFIJO_DESTINO: r.ruta || '' };
    }
    const conds = {}; tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = String(x.NOMBRE); });
    const conductor = otra ? v.CONDUCTOR_ID : (conds[v.CONDUCTOR_ID] || v.CONDUCTOR_ID);
    const prev = buscar_('RECEPCION_CAMION', 'VIAJE_ID', r.viajeId);
    const fecha = prev ? prev.FECHA : fechaAtrasada_(r.fecha, c), hora = prev ? prev.HORA : ahora_('HH:mm');
    // Se borran las filas anteriores de este camión y se escriben de nuevo (la recepción se guarda entera)
    const tR = tabla_('RECEPCIONES');
    tR.filas.filter(function (x) { return x.VIAJE_ID === r.viajeId; }).map(function (x) { return x._fila; }).sort(function (a, b) { return b - a; })
      .forEach(function (f) { shR.deleteRow(f); });
    const filas = [];
    comps.forEach(function (x, i) {
      (x.reps || []).forEach(function (p, j) {
        const o = { REC_ID: r.viajeId + '-' + x.id + '-' + (j + 1), FECHA: fecha, HORA: hora, VIAJE_ID: r.viajeId, DECA_NUM: v.DECA_NUM, MATRICULA: v.MATRICULA,
          CONDUCTOR: conductor, RUTA: pad2_(v.PREFIJO_DESTINO), ORDEN_DESCARGA: i + 1, COMPARTIMENTO: x.id, ESPECIE: x.esp,
          LITROS_DECLARADOS: x.decl || '', LITROS_CONTADOR: x.contador, DEPOSITO: p.dep, LITROS_DEPOSITO: p.l, TIPO: p.tipo || 'PRINCIPAL',
          TEMPERATURA: x.temp === '' || x.temp == null ? '' : x.temp, PH: x.ph || '', DORNIC: x.dornic || '', VISUAL: x.visual || '', USUARIO: String(c.NOMBRE),
          N_COMP_CISTERNA: String(x.nComp || '').trim() };
        filas.push(cabR.map(function (h) { return o[h] !== undefined ? o[h] : ''; }));
      });
    });
    // Agua de arranque y de arrastre final que entra al depósito (no la cuenta el total del camión)
    (r.agua || []).filter(function (a) { return Number(a.l) > 0 && String(a.dep || '').trim(); }).forEach(function (a, k) {
      const o = { REC_ID: r.viajeId + '-AGUA-' + (k + 1), FECHA: fecha, HORA: hora, VIAJE_ID: r.viajeId, DECA_NUM: v.DECA_NUM, MATRICULA: v.MATRICULA,
        CONDUCTOR: conductor, RUTA: pad2_(v.PREFIJO_DESTINO), ORDEN_DESCARGA: a.tipo === 'ARRANQUE' ? 0 : 99, COMPARTIMENTO: 'AG1', ESPECIE: 'AGUA',
        LITROS_DECLARADOS: '', LITROS_CONTADOR: Math.round(Number(a.l)), DEPOSITO: a.dep, LITROS_DEPOSITO: Math.round(Number(a.l)), TIPO: 'AGUA ' + (a.tipo || 'ARRASTRE'),
        TEMPERATURA: '', PH: '', DORNIC: '', VISUAL: '', USUARIO: String(c.NOMBRE) };
      filas.push(cabR.map(function (h) { return o[h] !== undefined ? o[h] : ''; }));
    });
    const f0 = shR.getLastRow() + 1;
    ['RUTA', 'COMPARTIMENTO', 'DECA_NUM', 'MATRICULA', 'DEPOSITO', 'FECHA', 'HORA'].forEach(function (h) { if (cabR.indexOf(h) >= 0) shR.getRange(f0, cabR.indexOf(h) + 1, filas.length, 1).setNumberFormat('@'); });
    shR.getRange(f0, 1, filas.length, cabR.length).setValues(filas);
    const suma = comps.reduce(function (a, x) { return a + (Number(x.contador) || 0); }, 0), total = Math.round(Number(r.total) || 0);
    const cam = { VIAJE_ID: r.viajeId, FECHA: fecha, HORA: hora, DECA_NUM: v.DECA_NUM, MATRICULA: v.MATRICULA, CONDUCTOR: conductor, RUTA: pad2_(v.PREFIJO_DESTINO),
      TOTAL_CONTADOR: total, SUMA_PARCIALES: suma, DIFERENCIA: total ? total - suma : '', LIMPIEZA_CISTERNA: r.limpieza || '', FILTRO: r.filtro || '',
      OBSERVACIONES: r.obs || '', USUARIO: String(c.NOMBRE), FICHERO: prev ? prev.FICHERO : '' };
    // Cuadre como en el «Registro litros contador»: cartilla = por especie; cisterna = por destinos (compartimentos)
    const q = r.cuadre || {}, cuenta = total || suma;
    if (Number(q.cartilla) > 0) { cam.LITROS_CARTILLA = Math.round(q.cartilla); cam.DIF_CARTILLA = cuenta - Math.round(q.cartilla); }
    if (Number(q.cisterna) > 0) { cam.LITROS_CISTERNA = Math.round(q.cisterna); cam.DIF_CISTERNA = cuenta - Math.round(q.cisterna); }
    cam.FICHERO = nombreDia_(fecha);
    if (prev) actualizar_('RECEPCION_CAMION', prev._fila, cam);
    else { const t = tabla_('RECEPCION_CAMION'); shC.getRange(shC.getLastRow() + 1, 1, 1, t.cab.length).setValues([t.cab.map(function (h) { return cam[h] !== undefined ? String(cam[h]) : ''; })]); }
    exportarDia_(fecha);
    return { ok: true, fichero: cam.FICHERO, diferencia: cam.DIFERENCIA };
  } finally { lock.releaseLock(); }
}
/** RECEPCIÓN 2 · registro de CALIDAD por especie (FOR PR 7.-02): T, pH, ºD, visual, nº de compartimento de la cisterna, litros,
 *  depósitos, limpieza de cisterna y de filtro. Se guarda en RECEPCION_CALIDAD y NO toca el fichero del programa de gestión
 *  (eso sigue saliendo solo de la recepción por destinos). Se guarda entero cada vez (se borran las filas anteriores del camión). */
function guardarRecepCalidad(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const r = req.recep || {}, id = String(r.viajeId || '');
    if (!id) throw new Error('Falta el camión.');
    const comps = (r.comps || []).filter(function (x) { return Number(x.contador) > 0; });
    if (!comps.length) throw new Error('No hay ninguna leche con litros.');
    const nUsados = {};
    comps.forEach(function (x) {
      if (!(x.reps || []).length || (x.reps || []).some(function (p) { return !String(p.dep || '').trim(); })) throw new Error('Falta el depósito de ' + x.id + '.');
      String(x.nComp || '').split(',').filter(String).forEach(function (n) { if (nUsados[n]) throw new Error('El compartimento ' + n + ' de la cisterna está en ' + nUsados[n] + ' y en ' + x.id + '.'); nUsados[n] = x.id; });
    });
    const sh = hoja_('RECEPCION_CALIDAD', CAB_RCAL);
    // ALTA: momento real en que se registró por primera vez (fecha y hora con segundos). Ordena el parte como se fueron metiendo.
    if (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String).indexOf('ALTA') < 0) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue('ALTA').setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    }
    const t = tabla_('RECEPCION_CALIDAD');
    const prev = t.filas.filter(function (x) { return x.VIAJE_ID === id; });
    hoja_('RECEPCION_CAMION', CAB_RECEP_CAM);
    const v = buscar_('VIAJES', 'VIAJE_ID', id) || {}, cam = buscar_('RECEPCION_CAMION', 'VIAJE_ID', id) || {};
    const conds = {}; tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = String(x.NOMBRE); });
    const fecha = prev.length ? String(prev[0].FECHA) : (cam.FECHA ? String(cam.FECHA) : fechaAtrasada_(r.fecha, c)), hora = prev.length ? String(prev[0].HORA) : ahora_('HH:mm');
    const alta = prev.length ? String(prev[0].ALTA || '') : ahora_('yyyy-MM-dd HH:mm:ss');
    prev.map(function (x) { return x._fila; }).sort(function (a, b) { return b - a; }).forEach(function (f) { sh.deleteRow(f); });
    const filas = comps.map(function (x, i) {
      const o = { FECHA: fecha, HORA: hora, VIAJE_ID: id, DECA_NUM: v.DECA_NUM || cam.DECA_NUM || '', MATRICULA: v.MATRICULA || cam.MATRICULA || r.matricula || '',
        CONDUCTOR: conds[v.CONDUCTOR_ID] || cam.CONDUCTOR || r.conductor || '', RUTA: v.PREFIJO_DESTINO ? pad2_(v.PREFIJO_DESTINO) : (cam.RUTA || ''), ORDEN: Number(x.ord) || i + 1,
        ID: String(x.id), ESPECIE: String(x.esp), N_COMP_CISTERNA: String(x.nComp || '').replace(/\s/g, ''), LITROS: Math.round(Number(x.contador)),
        DEPOSITOS: (x.reps || []).map(function (p) { return String(p.dep).trim() + ':' + Math.round(Number(p.l) || 0) + ':' + (p.tipo || 'PRINCIPAL'); }).join('|'),
        TEMPERATURA: x.temp === '' || x.temp == null ? '' : x.temp, PH: x.ph || '', DORNIC: x.dornic || '', VISUAL: x.visual || '',
        LIMPIEZA_CISTERNA: r.limpieza || '', LIMPIEZA_FILTRO: r.filtro || '', OBSERVACIONES: r.obs || '', USUARIO: String(c.NOMBRE), ALTA: alta };
      return t.cab.map(function (h) { return o[h] !== undefined ? o[h] : ''; });
    });
    const f0 = sh.getLastRow() + 1;
    sh.getRange(f0, 1, filas.length, t.cab.length).setNumberFormat('@').setValues(filas.map(function (f) { return f.map(String); }));
    return { ok: true, fecha: fecha };
  } finally { lock.releaseLock(); }
}
/** Fecha con la que se guarda una recepción o el agua: hoy, salvo que OFICINA esté registrando un día anterior que se quedó
 *  sin hacer (por ejemplo, un camión del sábado que no se descargó en la app). Nunca un día futuro ni de hace más de 15 días. */
function fechaAtrasada_(f, c) {
  const hoy = ahora_('dd/MM/yyyy');
  if (!f || String(f) === hoy || !esOficina_(c)) return hoy;
  const n = fechaNum_(f), h = fechaNum_(hoy);
  if (!n) throw new Error('Fecha no válida: ' + f);
  if (n > h) throw new Error('No se puede registrar una recepción en un día futuro.');
  const lim = new Date(); lim.setDate(lim.getDate() - 15);
  if (n < fechaNum_(Utilities.formatDate(lim, TZ, 'dd/MM/yyyy'))) throw new Error('Solo se pueden registrar recepciones de los últimos 15 días.');
  return String(f);
}
/** Anula por completo la recepción de un camión (prueba, error de día, camión equivocado…): borra sus líneas de RECEPCIONES
 *  y RECEPCION_CAMION, lo apunta en CAMBIOS y rehace el fichero del día sin ese camión. Si era un viaje de la app,
 *  vuelve a «Por recibir». */
function anularRecepcion(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const id = String(req.viajeId || ''), motivo = String(req.motivo || '').trim();
    if (!motivo) throw new Error('Escribe el motivo de la anulación.');
    const cam = buscar_('RECEPCION_CAMION', 'VIAJE_ID', id);
    if (!cam) throw new Error('Esa recepción no está guardada (puede que ya se haya anulado).');
    const shR = hoja_('RECEPCIONES', CAB_RECEP), shC = hoja_('RECEPCION_CAMION', CAB_RECEP_CAM);
    const filas = tabla_('RECEPCIONES').filas.filter(function (r) { return r.VIAJE_ID === id; });
    const porDep = {}; filas.forEach(function (r) { const d = String(r.DEPOSITO).replace(/^S/i, ''); porDep[d] = (porDep[d] || 0) + numES_(r.LITROS_DEPOSITO); });
    const antes = Object.keys(porDep).sort().map(function (d) { return 'dep ' + d + ': ' + porDep[d] + ' l'; }).join(', ');
    filas.map(function (r) { return r._fila; }).sort(function (a, b) { return b - a; }).forEach(function (f) { shR.deleteRow(f); });
    shC.deleteRow(cam._fila);
    // Fichero del día para el programa: se rehace sin este camión
    const fichero = exportarDia_(String(cam.FECHA));
    const shK = hoja_('CAMBIOS', ['FECHA_HORA', 'USUARIO', 'VIAJE_ID', 'DECA_NUM', 'FECHA_VIAJE', 'REC_ID', 'CODIGO', 'CAMPO', 'ANTES', 'DESPUES', 'MOTIVO']);
    shK.getRange(shK.getLastRow() + 1, 1, 1, 11).setNumberFormat('@').setValues([[ahora_('dd/MM/yyyy HH:mm'), String(c.NOMBRE), id, String(cam.DECA_NUM || ''), String(cam.FECHA),
      fichero, String(cam.MATRICULA || '') + ' ' + String(cam.CONDUCTOR || ''), 'RECEPCIÓN ANULADA', antes, 'ANULADA', motivo]]);
    return { ok: true, fichero: fichero, revisarAgua: tabla_('AGUA').filas.some(function (r) { return String(r.FECHA) === String(cam.FECHA); }) };
  } finally { lock.releaseLock(); }
}
/** Agua de arrastre del día (AG1/AG2) repartida por depósito. Se guarda entera cada vez y rehace el fichero del día. */
function guardarAgua(req, c) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ['AG1', 'AG2'].forEach(function (cod) { const ds = {}; (req.agua || []).forEach(function (a) { if (a.cod === cod && Number(a.l) > 0) ds[String(a.dep).trim()] = 1; });
      if (Object.keys(ds).length > 3) throw new Error(cod + ' va a ' + Object.keys(ds).length + ' depósitos: el programa admite como máximo 3. Junta el agua en 3.'); });
    const sh = hoja_('AGUA', CAB_AGUA), hoy = fechaAtrasada_(req.fecha, c);
    const t = tabla_('AGUA'), previas = t.filas.filter(function (r) { return r.FECHA === hoy; });
    previas.map(function (r) { return r._fila; }).sort(function (a, b) { return b - a; }).forEach(function (f) { sh.deleteRow(f); });
    const filas = (req.agua || []).filter(function (a) { return /^AG[12]$/.test(a.cod) && Number(a.l) > 0 && String(a.dep).trim(); });
    const nombre = nombreDia_(hoy);
    if (filas.length) {
      const valores = filas.map(function (a) { return [hoy, a.cod, String(a.dep), Math.round(Number(a.l)), a.desde || '', String(c.NOMBRE), ahora_('HH:mm'), nombre]; });
      const f0 = sh.getLastRow() + 1;
      sh.getRange(f0, 1, valores.length, 1).setNumberFormat('@'); sh.getRange(f0, 3, valores.length, 1).setNumberFormat('@');
      sh.getRange(f0, 1, valores.length, CAB_AGUA.length).setValues(valores);
    }
    exportarDia_(hoy);
    return { ok: true, fichero: nombre };
  } finally { lock.releaseLock(); }
}
/** UN SOLO FICHERO POR DÍA para el programa: D + aammdd + 0 + .TXT (p. ej. D2610070.TXT), con todos los camiones recibidos ese día
 *  (por orden de llegada) y al final el agua de arranque/arrastre (AG1, AG2) como un compartimento más. Se rehace entero cada
 *  vez que se guarda o anula una recepción o se guarda el agua de ese día, así que el programa siempre lo machaca.
 *  Formato de cada línea:
 *  FECHA;RUTA;ORDEN;COMPARTIMENTO;ESPECIE;LITROS_CONTADOR;DEP1;LIT1;DEP2;LIT2;DEP3;LIT3;TEMP;PH;DORNIC;MATRICULA;DECA;CONDUCTOR
 *  Agua: FECHA;A;0;AG1;AGUA;LITROS_TOTAL;DEP1;LIT1;DEP2;LIT2;DEP3;LIT3;;;;;;
 *  Si el día se queda vacío (todo anulado) lleva una sola línea ANULADA. */
function nombreDia_(fecha) {
  const p = String(fecha).split('/');
  return 'D' + p[2].slice(-2) + ('0' + p[1]).slice(-2) + ('0' + p[0]).slice(-2) + '0.TXT';   // el 0 final: 8 caracteres, como espera la copia al servidor
}
/** Rehace el fichero del día de hoy (o de la fecha dd/MM/yyyy que se ponga). Para ejecutarlo a mano desde el editor. */
function rehacerFicheroDia(fecha) {
  const f = fecha || ahora_('dd/MM/yyyy'), n = exportarDia_(f);
  Logger.log('Fichero rehecho: ' + n); return n;
}
function exportarDia_(fecha) {
  const nombre = nombreDia_(fecha);
  const cfg = cfg_(); if (!cfg.CARPETA_EXPORTACION_ID) return nombre;
  const recs = tabla_('RECEPCIONES').filas.filter(function (r) { return String(r.FECHA) === String(fecha) && String(r.ESPECIE).toUpperCase() !== 'AGUA'; });
  const horas = {}; tabla_('RECEPCION_CAMION').filas.forEach(function (x) { if (String(x.FECHA) === String(fecha)) horas[x.VIAJE_ID] = String(x.HORA || ''); });
  const viajes = [], por = {};
  recs.forEach(function (r) { const v = String(r.VIAJE_ID); if (!por[v]) { por[v] = []; viajes.push(v); } por[v].push(r); });
  const pos = {}; viajes.forEach(function (v, i) { pos[v] = i; });
  viajes.sort(function (a, b) { const ha = horas[a] || '99:99', hb = horas[b] || '99:99'; return ha < hb ? -1 : ha > hb ? 1 : pos[a] - pos[b]; });
  let lineas = [];
  viajes.forEach(function (v) { lineas = lineas.concat(lineasViaje_(v, por[v])); });
  hoja_('AGUA', CAB_AGUA);
  lineas = lineas.concat(lineasAgua_(tabla_('AGUA').filas.filter(function (a) { return String(a.FECHA) === String(fecha); })
    .map(function (a) { return { cod: String(a.CODIGO), dep: String(a.DEPOSITO), l: a.LITROS }; }), fecha));
  if (!lineas.length) lineas = [[fecha, '', 0, 'ANULADA', '', 0, '', '', '', '', '', '', '', '', '', '', '', ''].join(';')];
  const carpeta = DriveApp.getFolderById(cfg.CARPETA_EXPORTACION_ID);
  escribirD_(carpeta, nombre, lineas.join('\r\n') + '\r\n');
  // Quita el fichero con el nombre corto que se usó el 07/10/2026 (D261007.TXT), que la copia al servidor no recoge
  const corto = carpeta.getFilesByName(nombre.replace('0.TXT', '.TXT')); while (corto.hasNext()) corto.next().setTrashed(true);
  return nombre;
}
/** Líneas del agua: una por código (AG1/AG2) con hasta 3 depósitos. */
function lineasAgua_(filas, fecha) {
  const cods = [], por = {};
  filas.forEach(function (a) { if (!/^AG[12]$/.test(a.cod) || !(numES_(a.l) > 0)) return; if (!por[a.cod]) { por[a.cod] = { deps: [], l: {} }; cods.push(a.cod); } const g = por[a.cod], d = String(a.dep).replace(/^S/i, ''); if (g.deps.indexOf(d) < 0) g.deps.push(d); g.l[d] = (g.l[d] || 0) + Math.round(numES_(a.l)); });
  cods.sort();
  return cods.map(function (c) {
    const g = por[c], tot = g.deps.reduce(function (s, d) { return s + g.l[d]; }, 0), d = [0, 1, 2].map(function (n) { const x = g.deps[n]; return x ? [x, g.l[x]] : ['', '']; });
    return [fecha, 'A', 0, c, 'AGUA', tot, d[0][0], d[0][1], d[1][0], d[1][1], d[2][0], d[2][1], '', '', '', '', '', ''].join(';');
  });
}
/** Ficheros de descarga en fábrica. Los antiguos E… se renombran a D… la primera vez que se rehacen. */
function escribirD_(carpeta, nombre, texto) {
  const viejo = nombre, nuevo = String(nombre).replace(/^E/i, 'D');
  let it = carpeta.getFilesByName(nuevo);
  if (it.hasNext()) { it.next().setContent(texto); return nuevo; }
  it = carpeta.getFilesByName(viejo);
  if (viejo !== nuevo && it.hasNext()) { const f = it.next(); f.setName(nuevo); f.setContent(texto); return nuevo; }
  carpeta.createFile(nuevo, texto, MimeType.PLAIN_TEXT); return nuevo;
}
/** Líneas de un camión, una por compartimento con hasta 3 depósitos (como la «W» del programa).
 *  El depósito principal va primero; arrastres y repartos después (si van al mismo depósito se suman). Otras entradas: RUTA=EX. */
function lineasViaje_(viajeId, filas) {
  // Cisterna externa: todos sus compartimentos de una especie son UN destino (VC, VV, VO, VD), como la muestra única que va al
  // laboratorio. Litros = suma de los parciales; T, pH y ºD = media ponderada por litros; depósitos sumados (máx. 3).
  const externo = /^X/.test(String(viajeId));
  const grupos = {}, orden = [];
  filas.forEach(function (r) {
    const comp = String(r.COMPARTIMENTO), k = externo ? 'V' + String(r.ESPECIE).toUpperCase() : comp;
    if (!grupos[k]) { grupos[k] = { r: r, deps: [], l: {}, comps: {}, cont: 0, m: { t: [0, 0], ph: [0, 0], d: [0, 0] } }; orden.push(k); }
    const g = grupos[k], d = String(r.DEPOSITO).replace(/^S/i, '');
    if (g.deps.indexOf(d) < 0) g.deps.push(d);
    g.l[d] = (g.l[d] || 0) + numES_(r.LITROS_DEPOSITO);
    if (!g.comps[comp]) {   // cada compartimento físico cuenta una vez (sus filas repiten el contador)
      g.comps[comp] = 1; const lc = numES_(r.LITROS_CONTADOR); g.cont += lc;
      [['t', r.TEMPERATURA], ['ph', r.PH], ['d', r.DORNIC]].forEach(function (x) { const n = Number(String(x[1]).replace(',', '.')); if (String(x[1]).trim() !== '' && !isNaN(n)) { let v = n; if (x[0] === 'ph') while (v > 14) v = v / 10; g.m[x[0]][0] += v * lc; g.m[x[0]][1] += lc; } });
    }
  });
  const media = function (m) { return m[1] ? m[0] / m[1] : ''; };
  return orden.map(function (k) {
    const g = grupos[k], r = g.r, d = [0, 1, 2].map(function (n) { const x = g.deps[n]; return x ? [x, Math.round(g.l[x])] : ['', '']; });
    return [r.FECHA, externo ? 'EX' : r.RUTA, r.ORDEN_DESCARGA, k, r.ESPECIE, Math.round(g.cont),
      d[0][0], d[0][1], d[1][0], d[1][1], d[2][0], d[2][1], num1_(media(g.m.t)), ph_(media(g.m.ph)), num1_(media(g.m.d)), r.MATRICULA, r.DECA_NUM, r.CONDUCTOR].join(';');
  });
}

/* ───────────── Camiones ───────────── */

function camionObj_(k, comps) {
  const id = String(k.CAMION_ID);
  const letras = comps.filter(function (x) { return String(x.CAMION_ID) === id; })
    .sort(function (a, b) { return (Number(a.ORDEN) || 0) - (Number(b.ORDEN) || 0); })
    .map(function (x) { return { n: Number(x.ORDEN) || 0, l: String(x.LETRA).trim(), esp: split_(x.ESPECIES).map(especie_), mezcla: split_(x.MEZCLA_ADMITIDA).map(especie_), cap: Number(String(x.CAPACIDAD || '').replace(/\./g, '')) || 0,
      ruta: String(x.RUTA || '').trim() ? pad2_(x.RUTA) : '' }; });
  return { id: id, matricula: String(k.MATRICULA), desc: String(k.DESCRIPCION || ''), letraQ: String(k.CODIGO_LETRA_Q || '').trim(), letras: letras };
}
/** Compartimentos que valen para una ruta: los que tienen esa RUTA en COMPARTIMENTOS; si no hay ninguno, los que tienen RUTA vacía.
 *  Así un mismo camión puede ser 3V·3C·3D en la ruta 3 y 4V·4v·4O en la ruta 4. */
function letrasRuta_(letras, ruta) {
  const r = pad2_(ruta), propias = letras.filter(function (x) { return x.ruta === r; });
  return propias.length ? propias : letras.filter(function (x) { return !x.ruta; });
}

/** Si el conductor da de alta un camión nuevo desde el móvil, se añade a CAMIONES y COMPARTIMENTOS. */
function resolverCamion_(cam, c) {
  const comps = tabla_('COMPARTIMENTOS').filas;
  if (typeof cam === 'string') {
    const k = buscar_('CAMIONES', 'CAMION_ID', cam);
    if (!k) throw new Error('Camión no encontrado: ' + cam);
    return camionObj_(k, comps);
  }
  const mat = String(cam.matricula || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (mat.length < 5) throw new Error('Matrícula no válida');
  const id = 'CAM_' + mat;
  const ya = buscar_('CAMIONES', 'CAMION_ID', id);
  if (!ya) {
    anadir_('CAMIONES', {
      CAMION_ID: id, MATRICULA: cam.matricula, DESCRIPCION: 'Cisterna ' + cam.letras.length + ' compartimentos',
      TITULAR: String(c.TRANSPORTISTA), N_COMPARTIMENTOS: cam.letras.length, ACTIVO: 'SI', CODIGO_LETRA_Q: String(cam.letraQ || '').trim(),
      NOTAS: 'Alta del conductor ' + c.NOMBRE + ' el ' + ahora_('dd/MM/yyyy') + ' · REVISAR'
    });
    const t = tabla_('COMPARTIMENTOS');
    const filas = cam.letras.map(function (x, i) {
      const o = { CAMION_ID: id, ORDEN: i + 1, LETRA: x.l, ESPECIES: x.esp.join(','), MEZCLA_ADMITIDA: (x.mezcla || []).join(','), NOTA: 'Alta del conductor' };
      return t.cab.map(function (h) { return o[h] !== undefined ? o[h] : ''; });
    });
    t.sh.getRange(t.sh.getLastRow() + 1, 1, filas.length, t.cab.length).setValues(filas);
    return camionObj_(buscar_('CAMIONES', 'CAMION_ID', id), tabla_('COMPARTIMENTOS').filas);
  }
  return camionObj_(ya, comps);
}

/* ───────────── Utilidades de hoja ───────────── */

function tabla_(nombre) {
  const sh = SS.getSheetByName(nombre);
  if (!sh) throw new Error('Falta la pestaña ' + nombre + '. Ejecuta crearEstructura.');
  const v = sh.getDataRange().getDisplayValues();
  const cab = v.shift().map(String);
  const filas = v.map(function (r, i) { const o = { _fila: i + 2 }; cab.forEach(function (h, j) { o[h] = r[j]; }); return o; });
  return { sh: sh, cab: cab, filas: filas };
}
function buscar_(hoja, col, valor) {
  return tabla_(hoja).filas.filter(function (r) { return String(r[col]) === String(valor); })[0] || null;
}
const COLS_TEXTO = ['RUTAS', 'RUTA', 'PREFIJO_DESTINO', 'CAMION_ID', 'CAMION_HABITUAL', 'MATRICULA', 'CODIGO', 'ORDEN', 'MUESTRA', 'LETRA_Q', 'CODIGO_LETRA_Q', 'DECA_NUM', 'LETRA', 'DESTINO', 'DESTINO_2', 'DESTINO_3'];
/** Añade columnas que falten (en hojas creadas con una versión anterior), detrás de 'despues'. */
function columnas_(hoja, cols, despues) {
  const sh = SS.getSheetByName(hoja);
  let cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  cols.forEach(function (c) {
    if (cab.indexOf(c) >= 0) return;
    const tras = cab.indexOf(despues);
    const col = tras >= 0 ? tras + 2 : cab.length + 1;
    if (tras >= 0 && col <= cab.length) sh.insertColumnBefore(col); 
    sh.getRange(1, col).setValue(c).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    if (COLS_TEXTO.indexOf(c) >= 0) sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    despues = c;
  });
}
function anadir_(hoja, obj) {
  const sh = SS.getSheetByName(hoja);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const fila = sh.getLastRow() + 1;
  cab.forEach(function (h, j) { if (COLS_TEXTO.indexOf(h) >= 0) sh.getRange(fila, j + 1).setNumberFormat('@'); });
  sh.getRange(fila, 1, 1, cab.length).setValues([cab.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; })]);
}
function actualizar_(hoja, fila, obj) {
  const sh = SS.getSheetByName(hoja);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  Object.keys(obj).forEach(function (k) { const j = cab.indexOf(k); if (j >= 0) sh.getRange(fila, j + 1).setValue(obj[k]); });
}
function cfg_() { const o = {}; tabla_('CONFIG').filas.forEach(function (r) { o[r.CLAVE] = r.VALOR; }); return o; }
/** PIN: tras 5 intentos fallidos desde el mismo móvil, 15 minutos bloqueado. */
function auth_(pin, dev) {
  pin = String(pin || '').trim();
  if (!pin) throw new Error('Falta el PIN');
  const cache = CacheService.getScriptCache(), k = dev ? 'fallos_' + String(dev).slice(0, 60) : null;
  const fallos = k ? Number(cache.get(k) || 0) : 0;
  if (fallos >= 5) throw new Error('Demasiados intentos con un PIN incorrecto. Espera 15 minutos o llama a la oficina.');
  const c = tabla_(HOJA_US_()).filas.filter(function (r) {
    return String(r.PIN).trim() === pin && String(r.ACTIVO).toUpperCase() !== 'NO';
  })[0];
  if (!c) { if (k) cache.put(k, String(fallos + 1), 900); throw new Error('PIN incorrecto'); }
  if (fallos) cache.remove(k);
  return c;
}
/** Hoja de usuarios: USUARIOS si existe (nueva), si no CONDUCTORES (la de siempre). Mismas columnas. */
function HOJA_US_() { return SS.getSheetByName('USUARIOS') ? 'USUARIOS' : 'CONDUCTORES'; }
/** Roles de una persona, ordenados por importancia: OFICINA, CALIDAD, RECEPCION, CONDUCTOR.
 *  Se leen de la columna ROL escrita como sea («Oficina», «recepcionista», «calidad, oficina», «Recogedor», «Gerente» = todos…);
 *  por compatibilidad, también de RUTAS = OFICINA / RECEPCION / CALIDAD. Sin rol y con rutas → CONDUCTOR. */
function roles_(c) {
  const txt = (String(c.ROL || '') + ' ' + (/^\s*\d/.test(String(c.RUTAS || '')) ? '' : String(c.RUTAS || ''))).toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
  const r = [];
  const todo = /GEREN|DIRECC|DIRECTOR|TODO/.test(txt);   // GERENTE: todo lo de oficina, calidad y recepción
  if (todo || /OFIC|ADMIN/.test(txt)) r.push('OFICINA');
  if (todo || /CALID/.test(txt)) r.push('CALIDAD');
  if (todo || /RECEP/.test(txt)) r.push('RECEPCION');
  if (/CONDUC|RECOG|CHOF|TRANSP|CAMION/.test(txt) || (!r.length && split_(c.RUTAS).length)) r.push('CONDUCTOR');
  if (!r.length) r.push('CONDUCTOR');
  return r;
}
function esRecepcion_(c) { return roles_(c).indexOf('RECEPCION') >= 0; }
function esCalidad_(c) { return roles_(c).indexOf('CALIDAD') >= 0; }
function esOficina_(c) { return roles_(c).indexOf('OFICINA') >= 0; }
function comprobarRutas_(c, rutas) {
  const mias = split_(c.RUTAS).map(pad2_);
  (rutas || []).forEach(function (r) { if (mias.indexOf(pad2_(r)) < 0) throw new Error('La ruta ' + r + ' no es de este conductor'); });
}
function viajeResp_(f) {
  return { id: f.VIAJE_ID, decaNum: f.DECA_NUM, decaUrl: f.DECA_URL, matricula: f.MATRICULA, letraQ: f.LETRA_Q || '', camionId: f.CAMION_ID, hora: f.HORA_INICIO };
}
function split_(s) { return String(s || '').split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(String); }
function pad2_(x) { x = String(x || '').trim(); return /^\d$/.test(x) ? '0' + x : x; }
function up_(x) { return String(x).toUpperCase(); }
/** Acepta lo que escriba la oficina: D, DO, D.O., OVEJA DO, O, OVEJA, C, CABRA, V, VACA… y devuelve D, O, C o V. */
function especie_(x) {
  const t = String(x || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z]/g, '');
  if (!t) return '';
  if (t === 'D' || t === 'DO' || t.indexOf('DO') >= 0 && t.indexOf('OVEJA') >= 0 || t.indexOf('DENOMINACION') >= 0 || t === 'MANCHEGA') return 'D';
  if (t === 'O' || t.indexOf('OVEJA') === 0) return 'O';
  if (t === 'C' || t.indexOf('CABRA') === 0) return 'C';
  if (t === 'V' || t.indexOf('VACA') === 0) return 'V';
  return t.charAt(0);
}
/** Varias especies en una celda: "CO", "C,O", "C O", "CD", "C y D", "Oveja y cabra"… → ['C','O'].
 *  "DO" o "D.O." solos siguen siendo oveja D.O. (no D + O). */
function especies_(x) {
  let t = String(x || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  if (!t) return [];
  if (/^(DO|D\.O\.?|OVEJA\s*D\.?O\.?)$/.test(t)) return ['D'];
  t = t.replace(/D\.O\.?/g, ' D ').replace(/DENOMINACION[A-Z ]*/g, ' D ').replace(/MANCHEGA/g, ' D ')
       .replace(/OVEJA\s+D\b/g, ' D ').replace(/OVEJA/g, ' O ').replace(/CABRA/g, ' C ').replace(/VACA/g, ' V ')
       .replace(/\bY\b/g, ' ');
  const out = [];
  t.split(/[^A-Z]+/).filter(String).forEach(function (tok) {
    const letras = /^[DOCV]+$/.test(tok) ? (tok === 'DO' ? ['D'] : tok.split('')) : [especie_(tok)];
    letras.forEach(function (l) { if (l && out.indexOf(l) < 0) out.push(l); });
  });
  return out;
}
/** "O:880 C:310" → {O:880, C:310}. Un número suelto se asigna a la primera especie. */
function porEspecie_(v, esps) {
  const m = {}, t = String(v == null ? '' : v).trim();
  if (!t) return m;
  if (t.indexOf(':') < 0) { m[esps[0]] = t; return m; }
  t.split(/[\s,;]+/).forEach(function (p) { const q = p.split(':'); if (q.length === 2) m[especie_(q[0])] = q[1]; });
  return m;
}
/** Destino habitual por especie. Acepta "O:O C:C", "5O 5C" o una sola letra. */
function destinosPor_(v, esps) {
  const m = {}, t = String(v || '').trim();
  if (!t) return m;
  if (t.indexOf(':') >= 0) {
    t.split(/[\s,;]+/).forEach(function (p) { const q = p.split(':'); if (q.length === 2) m[especie_(q[0])] = letraDestino_(q[1]); });
    return m;
  }
  const toks = t.split(/[\s,;\/]+/).filter(String).map(letraDestino_);
  if (esps.length === 1) { m[esps[0]] = toks[0] || ''; return m; }
  esps.forEach(function (e) {
    const hit = toks.filter(function (l) { return l.toUpperCase() === e || (e === 'D' && l === 'd'); })[0];
    if (hit) m[e] = hit;
  });
  return m;
}
function juntar_(m, esps) { return esps.filter(function (e) { return m[e] !== undefined && m[e] !== ''; }).map(function (e) { return e + ':' + m[e]; }).join(' '); }
/** Destino habitual: solo la letra del compartimento (5D → D, 2d → d, DO → D). */
function letraDestino_(x) {
  const t = String(x || '').trim().replace(/^\d+/, '');
  if (!t) return '';
  if (/^[a-z]$/.test(t)) return t;              // minúscula = segundo depósito de esa especie (d, v, o…)
  const u = t.toUpperCase();
  if (u === 'DO' || u === 'D.O.' || u === 'D.O') return 'D';
  if (u === 'D2' || u === 'DO2') return 'd';
  return u.charAt(0);
}
function num1_(x) { const n = Number(String(x).replace(',', '.')); return isNaN(n) || x === '' ? '' : n.toFixed(1); }
/** pH siempre con punto y 2 decimales (6,8 → 6.80). Si se tecleó sin coma (680) se entiende 6.80. */
function ph_(x) { let n = Number(String(x).replace(',', '.')); if (x === '' || x == null || isNaN(n)) return ''; while (n > 14) n = n / 10; return n.toFixed(2); }
/** dd/MM/yyyy → número aaaammdd para comparar fechas (0 si no es una fecha). */
function fechaNum_(x) { const m = String(x || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? +(m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2)) : 0; }
function ahora_(f) { return Utilities.formatDate(new Date(), TZ, f); }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Desde el editor: rehace con QR los DeCA de los viajes de hoy que se emitieron antes de tener el QR.
 *  Mismo archivo y misma dirección (el QR del móvil sigue valiendo); se conserva la hora de emisión original. */
function decasHoyConQr() {
  const hoy = ahora_('dd/MM/yyyy'), cfg = cfg_(), conds = {};
  tabla_(HOJA_US_()).filas.forEach(function (x) { conds[x.CONDUCTOR_ID] = x; });
  tabla_('VIAJES').filas.filter(function (v) { return v.FECHA === hoy && /[?&]id=/.test(String(v.DECA_URL)); }).forEach(function (v) {
    const id = String(v.DECA_URL).match(/[?&]id=([\w-]+)/)[1], url = 'https://drive.google.com/uc?export=download&id=' + id;
    const c = Object.assign({}, conds[v.CONDUCTOR_ID] || {}, { TRANSPORTISTA: v.TRANSPORTISTA, NIF_TRANSPORTISTA: v.NIF_TRANSPORTISTA });
    const d = { num: v.DECA_NUM, cfg: cfg, cond: c, matricula: v.MATRICULA, letraQ: v.LETRA_Q, rutas: split_(v.RUTAS),
      kg: numES_(v.PESO_ESTIMADO_KG), litros: numES_(v.LITROS_ESTIMADOS) };
    Drive.Files.update({}, id, htmlDeca_(d, hoy + ' ' + v.HORA_INICIO, url, qrPngDataUri_(url, 4)));
    Logger.log('Con QR: ' + v.DECA_NUM + ' · ' + url);
  });
}
/** Para probar desde el editor: crea un DeCA de prueba y muestra su enlace en el registro. */
function probarDeca() {
  const c = tabla_(HOJA_US_()).filas[0];
  const r = crearDeca_({ num: 'PRUEBA-' + ahora_('HHmmss'), cfg: cfg_(), cond: c, matricula: '0000-TEST', rutas: ['05'], kg: 1030, litros: 1000, nota: 'Prueba desde el editor' });
  Logger.log(r.url);
}

/** Ejecutar UNA vez en hojas creadas antes de añadir Letra Q: añade las columnas que falten al final. */
function anadirColumnasLetraQ() {
  [['CAMIONES', 'CODIGO_LETRA_Q'], ['VIAJES', 'LETRA_Q']].forEach(function (x) {
    const sh = SS.getSheetByName(x[0]);
    const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (cab.indexOf(x[1]) >= 0) return;
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(x[1]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
    sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  aviso_('Columnas Letra Q añadidas en CAMIONES y VIAJES.');
}

/** Ejecutar UNA vez: añade la columna CAPACIDAD (litros) a COMPARTIMENTOS.
 *  Rellena la cisterna de Jesús (D 5000, O 3500, C 1500) si esas celdas están vacías. */
function anadirCapacidad() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COMPARTIMENTOS');
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let col = cab.indexOf('CAPACIDAD') + 1;
  if (!col) {
    col = cab.length + 1;
    sh.getRange(1, col).setValue('CAPACIDAD').setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d5875');
  }
  const cId = cab.indexOf('CAMION_ID'), cL = cab.indexOf('LETRA');
  const JESUS = { D: 5000, O: 3500, C: 1500 };
  const n = sh.getLastRow() - 1; let hechos = 0;
  if (n > 0) {
    const datos = sh.getRange(2, 1, n, col).getValues();
    datos.forEach(function (r, i) {
      const cap = JESUS[String(r[cL]).trim()];
      if (String(r[cId]) === 'CAMION_JESUS' && cap && r[col - 1] === '') { sh.getRange(i + 2, col).setValue(cap); hechos++; }
    });
  }
  aviso_('Columna CAPACIDAD lista en COMPARTIMENTOS. Rellenadas ' + hechos + ' de la cisterna de Jesús. Revisa el resto de camiones.');
}

/** Mensaje al terminar una función del editor (si no hay hoja abierta, va al registro). */
function aviso_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); } }

/* ───────────── Arreglos a mano (se ejecutan desde el editor: elegir la función y pulsar Ejecutar) ───────────── */

/** 03/10/2026: Felipe hizo la ruta 4, la cerró y empezó la ruta 3 con el mismo camión. Debía ser un solo viaje 03,04.
 *  Los compartimentos físicos son los mismos con otro nombre: 4V = 3V (1º), 4v = 3C (2º), 4O = 3D (3º). */
function unirFelipe0310() { return unirViajes('VMURX4QZJ-FELIPE', 'VMUS1LSXQ-FELIPE', { '4V': '3V', '4v': '3C', '4O': '3D' }); }

/** Une el viaje A dentro del B (mismo camión, mismo día): las recogidas de A pasan a B con los destinos traducidos por
 *  «mapa», B queda con las dos rutas y su fichero R se rehace con todo; A queda ANULADO · UNIDO A B y su fichero R se
 *  manda a la papelera de Drive. Todo queda apuntado en CAMBIOS. No se puede si alguno ya tiene la recepción guardada. */
function unirViajes(idA, idB, mapa) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const vA = buscar_('VIAJES', 'VIAJE_ID', idA), vB = buscar_('VIAJES', 'VIAJE_ID', idB);
    if (!vA || !vB) throw new Error('No encuentro los dos viajes.');
    if (vA.CAMION_ID !== vB.CAMION_ID || vA.FECHA !== vB.FECHA) throw new Error('No son del mismo camión y día.');
    if (/ANULADO|TRASVASE/.test(vA.ESTADO)) throw new Error('El viaje ' + idA + ' ya está ' + vA.ESTADO);
    hoja_('RECEPCION_CAMION', CAB_RECEP_CAM);
    if (buscar_('RECEPCION_CAMION', 'VIAJE_ID', idA) || buscar_('RECEPCION_CAMION', 'VIAJE_ID', idB)) throw new Error('Alguno ya tiene la recepción guardada: anúlala primero.');
    const t = tabla_('RECOGIDAS'), deA = t.filas.filter(function (r) { return r.VIAJE_ID === idA; });
    const idsB = {}; t.filas.forEach(function (r) { if (r.VIAJE_ID === idB) idsB[r.REC_ID] = 1; });
    // Primero se comprueba todo; después se cambia
    deA.forEach(function (r) {
      ['DESTINO', 'DESTINO_2', 'DESTINO_3'].forEach(function (c) { const d = String(r[c] || '').trim(); if (d && !mapa.hasOwnProperty(d)) throw new Error('Destino sin traducir: ' + d + ' (' + r.CODIGO + ')'); });
      if (idsB[String(r.REC_ID).replace(idA, idB)]) throw new Error('El ganadero ' + r.CODIGO + ' ' + r.ESPECIE + ' ya está en el viaje ' + idB);
    });
    const cuando = ahora_('dd/MM/yyyy HH:mm'), quien = 'Arreglo ' + cuando;
    const shK = hoja_('CAMBIOS', ['FECHA_HORA', 'USUARIO', 'VIAJE_ID', 'DECA_NUM', 'FECHA_VIAJE', 'REC_ID', 'CODIGO', 'CAMPO', 'ANTES', 'DESPUES', 'MOTIVO']);
    const log = [], motivo = 'Viajes ' + vA.RUTAS + ' y ' + vB.RUTAS + ' hechos por separado: unidos en uno';
    deA.forEach(function (r) {
      const nuevoId = String(r.REC_ID).replace(idA, idB), upd = { VIAJE_ID: idB, REC_ID: nuevoId, CORREGIDO: (r.CORREGIDO ? r.CORREGIDO + ' · ' : '') + 'unido a ' + idB + ' ' + cuando };
      log.push([cuando, quien, idB, vB.DECA_NUM, vB.FECHA, nuevoId, r.CODIGO, 'VIAJE_ID', idA, idB, motivo]);
      ['DESTINO', 'DESTINO_2', 'DESTINO_3'].forEach(function (c) { const d = String(r[c] || '').trim(); if (d) { upd[c] = mapa[d]; log.push([cuando, quien, idB, vB.DECA_NUM, vB.FECHA, nuevoId, r.CODIGO, c, d, mapa[d], motivo]); } });
      actualizar_('RECOGIDAS', r._fila, upd);
    });
    const rutas = split_(vA.RUTAS).concat(split_(vB.RUTAS)).map(pad2_).filter(function (x, i, a) { return a.indexOf(x) === i; }).sort().join(',');
    const total = tabla_('RECOGIDAS').filas.filter(function (r) { return r.VIAJE_ID === idB && r.NO_RECOGIDO !== 'SI'; }).reduce(function (a, r) { return a + numES_(r.LITROS); }, 0);
    const cfg = cfg_();
    const nombre = exportar_([idB], /^R\w{7}\.TXT$/i.test(String(vB.EXPORTADO)) ? String(vB.EXPORTADO) : null);
    actualizar_('VIAJES', vB._fila, { RUTAS: rutas, HORA_INICIO: vA.HORA_INICIO, LITROS_REALES: total, PESO_REAL_KG: Math.round(total * (Number(cfg.FACTOR_KG_POR_LITRO) || 1.03)), EXPORTADO: nombre });
    actualizar_('VIAJES', vA._fila, { ESTADO: 'ANULADO · UNIDO A ' + idB, LITROS_REALES: 0, PESO_REAL_KG: 0, EXPORTADO: 'UNIDO A ' + nombre + ' (antes ' + vA.EXPORTADO + ')' });
    log.push([cuando, quien, idA, vA.DECA_NUM, vA.FECHA, '', '', 'VIAJE', vA.ESTADO + ' · ' + vA.EXPORTADO, 'UNIDO A ' + idB + ' · ' + nombre, motivo]);
    shK.getRange(shK.getLastRow() + 1, 1, log.length, 11).setNumberFormat('@').setValues(log);
    // El fichero R del viaje A ya no vale: a la papelera de Drive (si ya se copió al servidor, hay que quitarlo de allí a mano)
    let papelera = '';
    if (cfg.CARPETA_EXPORTACION_ID && /^R\w{7}\.TXT$/i.test(String(vA.EXPORTADO))) {
      const it = DriveApp.getFolderById(cfg.CARPETA_EXPORTACION_ID).getFilesByName(String(vA.EXPORTADO));
      while (it.hasNext()) { it.next().setTrashed(true); papelera = String(vA.EXPORTADO); }
    }
    const res = { ok: true, recogidasMovidas: deA.length, viaje: idB, rutas: rutas, litros: total, fichero: nombre, ficheroPapelera: papelera };
    Logger.log(JSON.stringify(res));
    return res;
  } finally { lock.releaseLock(); }
}

/* ───────────── Puente al servidor (ordenador del despacho, Windows 7) ─────────────
 * El ordenador PIDE los ficheros por HTTPS (GET con clave) y los deja en el servidor:
 *   ?tipo=pendientes&clave=…              → texto: OK y una línea por fichero «NOMBRE;CARPETA;VERSION»
 *   ?tipo=fichero&nombre=…&clave=…        → el contenido del fichero tal cual
 *   ?tipo=entregado&nombre=…&version=…&clave=…&equipo=…  → lo apunta en la pestaña PUENTE
 * R…TXT van a RECLECHE y D…TXT a DESLECHE. Un fichero vuelve a estar pendiente si se rehace (por ejemplo el
 * D…0.TXT del día, o un R corregido por la oficina), y el puente lo machaca en el servidor.
 * La clave NO está en el código: está en las propiedades del proyecto (la crea instalarPuente). */
const CAB_PUENTE = ['NOMBRE', 'CARPETA', 'VERSION', 'ENTREGADO', 'EQUIPO'];
const RE_PUENTE = /^[RD]\d{6}[0-9A-Z]\.TXT$/i;

function puente_(p) {
  const txt = function (t) { return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT); };
  try {
    const clave = PropertiesService.getScriptProperties().getProperty('PUENTE_CLAVE');
    if (!clave || String(p.clave || '') !== clave) return txt('ERROR;clave no válida');
    const cfg = cfg_(); if (!cfg.CARPETA_EXPORTACION_ID) return txt('ERROR;falta CARPETA_EXPORTACION_ID en CONFIG');
    const carpeta = DriveApp.getFolderById(cfg.CARPETA_EXPORTACION_ID);
    if (p.tipo === 'pendientes') {
      const hechos = {}; tabla_(hojaPuente_().getName()).filas.forEach(function (r) { hechos[String(r.NOMBRE).toUpperCase()] = String(r.VERSION); });
      const lineas = ficherosPuente_(carpeta).filter(function (f) { return hechos[f.nombre.toUpperCase()] !== f.version; })
        .map(function (f) { return [f.nombre, f.destino, f.version].join(';'); });
      return txt(['OK'].concat(lineas).join('\r\n'));
    }
    const nombre = String(p.nombre || '').toUpperCase();
    if (!RE_PUENTE.test(nombre)) return txt('ERROR;nombre no válido');
    if (p.tipo === 'fichero') {
      const it = carpeta.getFilesByName(nombre);
      if (!it.hasNext()) return txt('ERROR;no existe ' + nombre);
      return txt(it.next().getBlob().getDataAsString('UTF-8'));
    }
    if (p.tipo === 'entregado') {
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        const sh = hojaPuente_(), t = tabla_(sh.getName()), ya = t.filas.filter(function (r) { return String(r.NOMBRE).toUpperCase() === nombre; })[0];
        const o = { NOMBRE: nombre, CARPETA: nombre.charAt(0) === 'R' ? 'RECLECHE' : 'DESLECHE', VERSION: String(p.version || ''), ENTREGADO: ahora_('dd/MM/yyyy HH:mm'), EQUIPO: String(p.equipo || '') };
        if (ya) actualizar_(sh.getName(), ya._fila, o);
        else { const f0 = sh.getLastRow() + 1; sh.getRange(f0, 1, 1, CAB_PUENTE.length).setNumberFormat('@').setValues([CAB_PUENTE.map(function (h) { return o[h]; })]); }
      } finally { lock.releaseLock(); }
      return txt('OK');
    }
    return txt('ERROR;tipo no válido');
  } catch (err) { return txt('ERROR;' + (err && err.message || err)); }
}
function hojaPuente_() { return hoja_('PUENTE', CAB_PUENTE); }
/** Ficheros R/D de la carpeta de exportación tocados en los últimos 45 días, con su versión (fecha de modificación). */
function ficherosPuente_(carpeta) {
  const desde = new Date(Date.now() - 45 * 864e5), out = [];
  const it = carpeta.searchFiles('modifiedDate > "' + Utilities.formatDate(desde, 'UTC', "yyyy-MM-dd'T'HH:mm:ss") + '" and trashed = false');
  while (it.hasNext()) {
    const f = it.next(), n = String(f.getName()).toUpperCase();
    if (!RE_PUENTE.test(n)) continue;
    out.push({ nombre: n, destino: n.charAt(0) === 'R' ? 'RECLECHE' : 'DESLECHE', version: String(f.getLastUpdated().getTime()) });
  }
  return out.sort(function (a, b) { return a.version - b.version; });
}
/** EJECUTAR UNA VEZ desde el editor al montar el puente: crea la clave (si no existe), marca como YA ENTREGADOS todos
 *  los ficheros que hay ahora (el puente antiguo ya los dejó en el servidor) y escribe en el registro la clave. */
function instalarPuente() {
  const pr = PropertiesService.getScriptProperties();
  let clave = pr.getProperty('PUENTE_CLAVE');
  if (!clave) { clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20); pr.setProperty('PUENTE_CLAVE', clave); }
  const carpeta = DriveApp.getFolderById(cfg_().CARPETA_EXPORTACION_ID), sh = hojaPuente_();
  const hechos = {}; tabla_(sh.getName()).filas.forEach(function (r) { hechos[String(r.NOMBRE).toUpperCase()] = 1; });
  const nuevos = ficherosPuente_(carpeta).filter(function (f) { return !hechos[f.nombre]; })
    .map(function (f) { return [f.nombre, f.destino, f.version, ahora_('dd/MM/yyyy HH:mm'), 'INICIO (puente antiguo)']; });
  if (nuevos.length) { const f0 = sh.getLastRow() + 1; sh.getRange(f0, 1, nuevos.length, CAB_PUENTE.length).setNumberFormat('@').setValues(nuevos); }
  Logger.log('Ficheros marcados como ya entregados: ' + nuevos.length);
  Logger.log('CLAVE DEL PUENTE: ' + clave);
  return clave;
}

