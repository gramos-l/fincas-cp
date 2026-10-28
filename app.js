/* Fincas C&P: mapa de fincas de boda a 20-24 km de Ayala 71 (Madrid), fichas, lista, filtros y lectura del Excel. */
(function () {
  'use strict';

  /* ================== AJUSTES ================== */
  const DATA_XLSX = 'Fincas_bodas_20km_Ayala71.xlsx';   // si subes una versión nueva al repositorio, el mapa la lee
  const DATA_JSON = 'fincas.json';                      // copia de respaldo de los datos
  const SHEETS = { 'Fincas 20 km': 'dentro', 'Justo fuera 20-24 km': 'fuera', 'Otros espacios': 'otro' };
  const SHEETJS = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  const GEO_KEY = 'fincas-geo-v1';
  const CUSTOM_XLSX_KEY = 'fincas-custom-xlsx-v1';   // tu propio Excel, guardado en este dispositivo
  const PIXEL_RATIO_MAX = 1.5;   // más bajo = más fluido; súbelo a 2 si prefieres más nitidez

  const ORIGEN = { center: [-3.676147, 40.427407], zoom: 10.2, pitch: 0, bearing: 0 };   // Calle de Ayala 71
  const BOUNDS = [[-4.25, 39.95], [-3.10, 40.90]];   // margen holgado para los 20-24 km + rutas

  // Cabeceras del Excel (hojas "Fincas 20 km", "Justo fuera 20-24 km", "Otros espacios")
  const COL = {
    name: 'Nombre', tipo: 'Tipo de espacio', zona: 'Zona / municipio', addr: 'Dirección',
    dist: 'Distancia en línea recta (km)', precio: 'Precio orientativo', precioInv: 'Precio menú desde (€/invitado)',
    precio120: 'Coste menú orientativo para 120 invitados (€)', aforoMin: 'Aforo mín.', aforoMax: 'Aforo máx.',
    aforoDet: 'Aforo (detalle)', catering: 'Catering (quién lo da)', cateringLibre: '¿Catering libre?',
    salones: 'Salones / espacios', unoVarios: '¿Un salón o varios?', terraza: 'Terraza / jardín',
    novios: 'Cuarto o zona de novios', ceremonia: 'Ceremonia en la finca', alojamiento: 'Alojamiento',
    parking: 'Parking', horario: 'Horario / nº eventos', notas: 'Observaciones',
    googleNota: 'Google (nota)', googleResenas: 'Google (nº reseñas)', bodasNet: 'bodas.net (nota y opiniones)',
    tel: 'Teléfono', src: 'Fuente principal', foto: 'Foto', lat: 'Latitud', lon: 'Longitud',
    tiempoCoche: 'Tiempo estimado en coche (min)', insta: 'Instagram (@usuario)', categoria: 'Categoría'
  };

  /* ================== UTILIDADES ================== */
  const $ = id => document.getElementById(id);
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => String(s == null ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const safeUrl = u => /^https?:\/\//i.test(String(u || '').trim()) ? String(u).trim() : '';
  function fmtVal(v) {
    if (v == null) return '';
    return typeof v === 'number' ? v : String(v).trim();
  }
  const toNum = v => { const n = parseFloat(String(v).replace(',', '.').replace(/[^\d.\-]/g, '')); return isFinite(n) ? n : null; };
  const eur = n => (Math.round(n * 100) / 100).toString().replace('.', ',') + ' €';

  /* ================== ESTILO DEL MAPA (mínimo, rápido) ================== */
  const EXTRUDE_FROM = 15, EXTRUDE_FULL = 16.5, MAX_PITCH_AUTO = 50, FLAT_UNTIL = 14.8, FULL_3D_AT = 16.8;
  const H_EXPR = ['coalesce', ['get', 'render_height'], 8];
  const MINH_EXPR = ['coalesce', ['get', 'render_min_height'], 0];

  function makeStyle(t) {
    const dk = t === 'dark';
    // estética de boda: calles y edificios en tonos rosa empolvado / salvia / dorado
    const C = dk
      ? { bg: '#241c22', green: '#38422f', water: '#1f2a30', casing: '#160f14', minor: '#3a2c30', major: '#4a3540', motor: '#5c3e38', foot: '#2e2228',
          cols: [0, '#342830', 40, '#4a3540', 100, '#705248', 150, '#a37848', 220, '#cda45e'] }
      : { bg: '#faf4ec', green: '#c6d0b8', water: '#bcc9cc', casing: '#e3cfd0', minor: '#fffaf5', major: '#fff6ef', motor: '#e8b8a0', foot: '#f0e2de',
          cols: [0, '#f5e6e0', 40, '#e8c2b8', 100, '#d9a28c', 150, '#c9a24a', 220, '#a8792e'] };
    const wd = b => ['interpolate', ['exponential', 1.5], ['zoom'], 10, b * 0.15, 14, b, 18, b * 8];
    const h = H_EXPR, minH = MINH_EXPR;
    const S = 'ofm';
    const accent = (getComputedStyle(document.documentElement).getPropertyValue('--accent') || '#b8697d').trim();
    return {
      version: 8, name: 'fincas-cp',
      projection: { type: 'globe' },   // converge con mercator a partir de zoom ~12; solo se nota al alejar del todo
      glyphs: 'https://tiles.openfreemap.org/fonts/{fontstack}/{range}.pbf',
      sources: { ofm: { type: 'vector', url: 'https://tiles.openfreemap.org/planet' }, fincas: { type: 'geojson', data: geojson(), cluster: true, clusterRadius: 46, clusterMaxZoom: 13 },
        ruta: { type: 'geojson', data: rutaGeo() }, origen: { type: 'geojson', data: origenGeo() } },
      light: { anchor: 'viewport', position: [1.4, 215, 35], color: dk ? '#dba878' : '#fff2ec', intensity: dk ? 0.35 : 0.42 },
      // atmósfera del globo en la intro: visible muy alejado, se desvanece al acercarse (se nota poco pasado zoom 5-7)
      sky: dk
        ? { 'sky-color': '#0a0e16', 'horizon-color': '#2a3550', 'fog-color': '#0a0e16', 'atmosphere-blend': ['interpolate', ['linear'], ['zoom'], 0, 1, 5, 1, 7, 0] }
        : { 'sky-color': '#5fa8e8', 'horizon-color': '#cfe8fb', 'fog-color': '#eeece6', 'atmosphere-blend': ['interpolate', ['linear'], ['zoom'], 0, 1, 5, 1, 7, 0] },
      layers: [
        // transparente a zoom muy bajo: así se ve el cielo/atmósfera del globo en vez de un fondo plano;
        // opaco desde zoom ~6, mucho antes de que la app opere nunca por debajo del zoom mínimo
        { id: 'fondo', type: 'background', paint: { 'background-color': C.bg, 'background-opacity': ['interpolate', ['linear'], ['zoom'], 0, 0, 3, 0, 6, 1] } },
        { id: 'cobertura', type: 'fill', source: S, 'source-layer': 'landcover', filter: ['in', 'class', 'grass', 'wood', 'scrub'], paint: { 'fill-color': C.green, 'fill-opacity': 0.7 } },
        { id: 'parques', type: 'fill', source: S, 'source-layer': 'park', paint: { 'fill-color': C.green } },
        { id: 'agua', type: 'fill', source: S, 'source-layer': 'water', paint: { 'fill-color': C.water } },
        { id: 'rios', type: 'line', source: S, 'source-layer': 'waterway', minzoom: 11, paint: { 'line-color': C.water, 'line-width': wd(1.5) } },
        { id: 'edificios-planos', type: 'fill', source: S, 'source-layer': 'building', minzoom: 13, maxzoom: EXTRUDE_FROM, paint: { 'fill-color': C.foot } },
        { id: 'calles-borde', type: 'line', source: S, 'source-layer': 'transportation', minzoom: 10, filter: ['in', 'class', 'primary', 'secondary', 'motorway', 'trunk'],
          layout: { 'line-cap': 'round', 'line-join': 'round' }, paint: { 'line-color': C.casing, 'line-width': wd(3.6) } },
        { id: 'calles-menores', type: 'line', source: S, 'source-layer': 'transportation', minzoom: 12.5, filter: ['in', 'class', 'minor', 'tertiary'],
          layout: { 'line-cap': 'round', 'line-join': 'round' }, paint: { 'line-color': C.minor, 'line-width': wd(1.1) } },
        { id: 'calles-principales', type: 'line', source: S, 'source-layer': 'transportation', minzoom: 10, filter: ['in', 'class', 'primary', 'secondary'],
          layout: { 'line-cap': 'round', 'line-join': 'round' }, paint: { 'line-color': C.major, 'line-width': wd(2.4) } },
        { id: 'autovias', type: 'line', source: S, 'source-layer': 'transportation', minzoom: 9, filter: ['in', 'class', 'motorway', 'trunk'],
          layout: { 'line-cap': 'round', 'line-join': 'round' }, paint: { 'line-color': C.motor, 'line-width': wd(3) } },
        { id: 'edificios-3d', type: 'fill-extrusion', source: S, 'source-layer': 'building', minzoom: EXTRUDE_FROM,
          paint: {
            'fill-extrusion-height': hx(h),
            'fill-extrusion-base': hx(minH),
            'fill-extrusion-color': ['interpolate', ['linear'], h].concat(C.cols),
            'fill-extrusion-opacity': 1, 'fill-extrusion-vertical-gradient': false
          } },
        // nombres de calle: aparecen a partir de cierto zoom para orientarse, las principales un poco antes que las menores
        { id: 'calles-nombres-principales', type: 'symbol', source: S, 'source-layer': 'transportation_name', minzoom: 12,
          filter: ['in', 'class', 'motorway', 'trunk', 'primary', 'secondary'],
          layout: { 'symbol-placement': 'line', 'text-field': ['coalesce', ['get', 'name:es'], ['get', 'name']],
            'text-font': ['Noto Sans Regular'], 'text-size': ['interpolate', ['linear'], ['zoom'], 12, 11, 17, 13.5],
            'text-rotation-alignment': 'map', 'text-pitch-alignment': 'viewport', 'symbol-spacing': 300, 'text-letter-spacing': 0.02 },
          paint: { 'text-color': dk ? '#f5ebe6' : '#3a2f30', 'text-halo-color': dk ? '#241c22' : '#faf4ec', 'text-halo-width': 1.3 } },
        { id: 'calles-nombres-menores', type: 'symbol', source: S, 'source-layer': 'transportation_name', minzoom: 14.5,
          filter: ['in', 'class', 'minor', 'tertiary', 'service'],
          layout: { 'symbol-placement': 'line', 'text-field': ['coalesce', ['get', 'name:es'], ['get', 'name']],
            'text-font': ['Noto Sans Regular'], 'text-size': ['interpolate', ['linear'], ['zoom'], 14.5, 10.5, 18, 13],
            'text-rotation-alignment': 'map', 'text-pitch-alignment': 'viewport', 'symbol-spacing': 250, 'text-letter-spacing': 0.02 },
          paint: { 'text-color': dk ? '#d8c9c6' : '#5a4a48', 'text-halo-color': dk ? '#241c22' : '#faf4ec', 'text-halo-width': 1.2 } },
        { id: 'ruta-halo', type: 'line', source: 'ruta', layout: { 'line-cap': 'round', 'line-join': 'round' },
          paint: { 'line-color': dk ? '#0a0e16' : '#fff4e0', 'line-width': 7, 'line-opacity': 0.6 } },
        { id: 'ruta', type: 'line', source: 'ruta', layout: { 'line-cap': 'round', 'line-join': 'round' },
          paint: { 'line-color': accent, 'line-width': 3.5 } },
        { id: 'origen', type: 'circle', source: 'origen',
          paint: { 'circle-radius': 6, 'circle-color': accent, 'circle-stroke-color': '#fff', 'circle-stroke-width': 2 } },
        { id: 'pin-clusters', type: 'symbol', source: 'fincas', filter: ['has', 'point_count'],
          layout: { 'icon-image': ['concat', 'cl-', ['to-string', ['get', 'point_count']]], 'icon-allow-overlap': true, 'icon-ignore-placement': true } },
        { id: 'pins', type: 'symbol', source: 'fincas', filter: UNCLUSTERED,
          layout: { 'icon-image': PIN_IMG, 'icon-anchor': 'bottom', 'icon-allow-overlap': true, 'icon-ignore-placement': true } },
        { id: 'pin-sel', type: 'symbol', source: 'fincas', filter: selFilter(sel ? sel.idx : null),
          layout: { 'icon-image': PIN_IMG, 'icon-anchor': 'bottom', 'icon-size': 1.35, 'icon-allow-overlap': true, 'icon-ignore-placement': true } }
      ]
    };
  }

  const theme = 'light';   // sin modo oscuro en este proyecto, siempre claro
  let auto = true, userMoved = false, hScale = 1;
  // altura de extrusión: crece con el zoom y se multiplica por el control «Altura de edificios»
  const hx = v => ['interpolate', ['linear'], ['zoom'], EXTRUDE_FROM, 0, EXTRUDE_FULL, ['*', v, hScale]];
  let places = [], filter = 'all', sel = null;

  /* ---- Ruta real desde Ayala 71 (OSRM, servidor público de demostración: gratis, sin clave, sin SLA) ---- */
  const OSRM = 'https://router.project-osrm.org/route/v1/driving/';
  let rutaFeat = null, rutaReqId = 0;
  const rutaGeo = () => ({ type: 'FeatureCollection', features: rutaFeat ? [rutaFeat] : [] });
  const origenGeo = () => ({ type: 'FeatureCollection', features: [{ type: 'Feature', properties: {}, geometry: { type: 'Point', coordinates: ORIGEN.center } }] });

  // las rutas calculadas se guardan en el móvil: se piden a OSRM una sola vez por finca, nunca más
  const RUTA_KEY = 'fincas-rutas-v1';
  let rutaCache = {};
  try { rutaCache = JSON.parse(localStorage.getItem(RUTA_KEY) || '{}'); } catch (e) { rutaCache = {}; }
  const saveRutaCache = () => { try { localStorage.setItem(RUTA_KEY, JSON.stringify(rutaCache)); } catch (e) { /* sin almacenamiento */ } };
  const rutaKeyOf = p => norm(p.name) + '|' + norm(p.addr);
  async function fetchRuta(p) {
    const k = rutaKeyOf(p);
    if (rutaCache[k]) return rutaCache[k];
    const url = OSRM + ORIGEN.center[0] + ',' + ORIGEN.center[1] + ';' + p.lon + ',' + p.lat + '?overview=full&geometries=geojson';
    const r = await fetch(url);
    if (!r.ok) throw new Error('http ' + r.status);
    const j = await r.json();
    const ruta = j.routes && j.routes[0];
    if (!ruta) throw new Error('sin ruta');
    const data = { distance: ruta.distance, duration: ruta.duration, geometry: ruta.geometry };
    rutaCache[k] = data; saveRutaCache();
    return data;
  }
  function clearRuta() {
    rutaReqId++; rutaFeat = null;
    const src = map.getSource('ruta'); if (src) src.setData(rutaGeo());
    const info = $('rutaInfo'); if (info) info.textContent = '';
  }
  async function drawRuta(p) {
    const myId = ++rutaReqId;
    rutaFeat = null;   // quita la ruta anterior al momento; se dibuja la nueva en cuanto llegue
    const src0 = map.getSource('ruta'); if (src0) src0.setData(rutaGeo());
    const info0 = $('rutaInfo'); if (info0) info0.textContent = 'Calculando ruta…';
    if (!p.located) return;
    try {
      const ruta = await fetchRuta(p);
      if (myId !== rutaReqId) return;   // la ficha cambió mientras llegaba la respuesta
      rutaFeat = { type: 'Feature', properties: {}, geometry: ruta.geometry };
      const src = map.getSource('ruta'); if (src) src.setData(rutaGeo());
      const info = $('rutaInfo');
      if (info) {
        const km = (ruta.distance / 1000).toFixed(1).replace('.', ',');
        const min = Math.round(ruta.duration / 60);
        info.textContent = km + ' km en coche · ' + min + ' min (ruta real, sin tráfico)';
      }
      fitRuta();   // encuadra Ayala 71 + la finca + la carretera real entre ambos, no solo la chincheta
    } catch (e) {
      if (myId === rutaReqId) { const info = $('rutaInfo'); if (info) info.textContent = 'No se pudo calcular la ruta ahora mismo.'; }
    }
  }
  // Encuadra toda la ruta dibujada (no solo el destino), dejando hueco para la ficha abierta encima/al lado
  function fitRuta() {
    if (!rutaFeat) return;
    const b = new maplibregl.LngLatBounds();
    rutaFeat.geometry.coordinates.forEach(c => b.extend(c));
    const wide = window.innerWidth >= 900;
    map.fitBounds(b, {
      padding: wide ? { top: 70, bottom: 60, left: 420, right: 40 } : { top: 70, bottom: Math.round(window.innerHeight * 0.58), left: 30, right: 30 },
      maxZoom: 15, duration: 900
    });
  }

  /* ---- Precarga para uso sin cobertura: teselas del entorno de cada finca + sus rutas.
     Se hace una sola vez (se guarda en el móvil que ya se hizo); las siguientes aperturas no vuelven a pedir nada. ---- */
  const PRECACHE_KEY = 'fincas-precache-v1';
  const PRECACHE_RADIUS_M = 650, PRECACHE_ZMIN = 12, PRECACHE_ZMAX = 17;
  let tileTemplate = null;
  async function getTileTemplate() {
    if (tileTemplate) return tileTemplate;
    const r = await fetch('https://tiles.openfreemap.org/planet');
    const j = await r.json();
    tileTemplate = j.tiles[0];
    return tileTemplate;
  }
  function lonLatToTile(lon, lat, z) {
    const n = Math.pow(2, z);
    const x = Math.floor((lon + 180) / 360 * n);
    const rad = lat * Math.PI / 180;
    const y = Math.floor((1 - Math.log(Math.tan(rad) + 1 / Math.cos(rad)) / Math.PI) / 2 * n);
    return [Math.max(0, Math.min(n - 1, x)), Math.max(0, Math.min(n - 1, y))];
  }
  function tilesAround(lat, lon, radiusM, zMin, zMax) {
    const out = [];
    for (let z = zMin; z <= zMax; z++) {
      const dLat = radiusM / 111320, dLon = radiusM / (111320 * Math.cos(lat * Math.PI / 180));
      const [x0, y0] = lonLatToTile(lon - dLon, lat + dLat, z);
      const [x1, y1] = lonLatToTile(lon + dLon, lat - dLat, z);
      for (let x = x0; x <= x1; x++) for (let y = y0; y <= y1; y++) out.push([z, x, y]);
    }
    return out;
  }
  async function precacheOffline(list) {
    let done = [];
    try { done = JSON.parse(localStorage.getItem(PRECACHE_KEY) || '[]'); } catch (e) { done = []; }
    const doneSet = new Set(done);
    const locs = list.filter(p => p.located && !doneSet.has(rutaKeyOf(p)));
    if (!locs.length) return;   // ya se precargó todo (por nombre+dirección) en una visita anterior
    const info = $('precacheInfo');
    if (info) info.style.display = 'flex';
    try {
      const tpl = await getTileTemplate();
      for (let i = 0; i < locs.length; i++) {
        const p = locs[i];
        if (info) info.textContent = 'Precargando para sin cobertura… finca ' + (i + 1) + '/' + locs.length;
        await fetchRuta(p).catch(() => {});   // guarda la ruta en caché; si falla, no bloquea la precarga de teselas
        const tiles = tilesAround(p.lat, p.lon, PRECACHE_RADIUS_M, PRECACHE_ZMIN, PRECACHE_ZMAX);
        const BATCH = 8;
        for (let j = 0; j < tiles.length; j += BATCH) {
          await Promise.all(tiles.slice(j, j + BATCH).map(([z, x, y]) => {
            const url = tpl.replace('{z}', z).replace('{x}', x).replace('{y}', y);
            return fetch(url).catch(() => {});   // el service worker guarda cada tesela al vuelo
          }));
        }
        doneSet.add(rutaKeyOf(p));
        try { localStorage.setItem(PRECACHE_KEY, JSON.stringify(Array.from(doneSet))); } catch (e) { /* sin almacenamiento */ }
      }
    } catch (e) { /* sin red o servidor caído: se reintenta en la próxima visita desde donde se quedó */ }
    if (info) info.style.display = 'none';
  }

  /* ---- Estados de cada finca: de un vistazo por el color de la chincheta ---- */
  const ESTADOS = ['pendiente', 'esperando', 'descartada', 'visitada'];
  const ESTADO_LABEL = { pendiente: 'Pendiente', esperando: 'Esperando confirmación', descartada: 'Descartada', visitada: 'Visitada ✓' };
  const ESTADO_VAR = { pendiente: '--pin-pendiente', esperando: '--pin-esperando', descartada: '--pin-descartada', visitada: '--pin-visitada' };
  const ESTADO_FALLBACK = { pendiente: '#b8697d', esperando: '#d9a93f', descartada: '#b2473f', visitada: '#7c8c6e' };
  const estadoColor = e => 'var(' + (ESTADO_VAR[e] || ESTADO_VAR.pendiente) + ')';

  /* ---- Chinchetas dibujadas por el propio mapa (GPU): mucho más fluido que elementos HTML ---- */
  function geojson() {
    return { type: 'FeatureCollection', features: places.filter(p => p.located && match(p)).map(p => ({
      type: 'Feature', properties: { idx: p.idx, estado: p.estado, approx: p.approx },
      geometry: { type: 'Point', coordinates: [p.lon, p.lat] } })) };
  }
  const UNCLUSTERED = ['!', ['has', 'point_count']];
  const selFilter = i => ['all', UNCLUSTERED, ['==', ['get', 'idx'], i == null ? -1 : i]];
  const PIN_IMG = ['concat', 'pin-', ['get', 'estado'], ['case', ['get', 'approx'], '-a', '']];
  const PIN_PATH = 'M14 1.8C7.2 1.8 2 7 2 13.4 2 21.8 14 34.2 14 34.2S26 21.8 26 13.4C26 7 20.8 1.8 14 1.8z';
  function pinImage(id) {
    // id tiene forma "pin-<estado>" o "pin-<estado>-a" (ubicación aproximada)
    const dashed = /-a$/.test(id);
    const estado = id.replace(/^pin-/, '').replace(/-a$/, '');
    const css = getComputedStyle(document.documentElement);
    const fill = (css.getPropertyValue(ESTADO_VAR[estado] || ESTADO_VAR.pendiente) || ESTADO_FALLBACK[estado] || ESTADO_FALLBACK.pendiente).trim();
    const S = 2, c = document.createElement('canvas');
    c.width = 28 * S; c.height = 36 * S;
    const x = c.getContext('2d'); x.scale(S, S);
    const path = new Path2D(PIN_PATH);
    x.shadowColor = 'rgba(0,0,0,.35)'; x.shadowBlur = 2; x.shadowOffsetY = 1;
    x.fillStyle = fill; x.fill(path);
    x.shadowColor = 'transparent'; x.shadowBlur = 0; x.shadowOffsetY = 0;
    x.lineWidth = 1.6; x.strokeStyle = '#fff'; if (dashed) x.setLineDash([3, 2]); x.stroke(path);
    x.setLineDash([]); x.beginPath(); x.arc(14, 13.6, 4.6, 0, Math.PI * 2); x.fillStyle = '#fff'; x.fill();
    return x.getImageData(0, 0, c.width, c.height);
  }
  function clusterImage(n) {
    const S = 2, sz = 38, c = document.createElement('canvas');
    c.width = c.height = sz * S;
    const x = c.getContext('2d'); x.scale(S, S);
    const accent = (getComputedStyle(document.documentElement).getPropertyValue('--accent') || '#b8697d').trim();
    x.beginPath(); x.arc(sz / 2, sz / 2, sz / 2 - 2, 0, Math.PI * 2);
    x.shadowColor = 'rgba(0,0,0,.3)'; x.shadowBlur = 2; x.shadowOffsetY = 1;
    x.fillStyle = accent; x.fill();
    x.shadowColor = 'transparent'; x.shadowBlur = 0; x.shadowOffsetY = 0;
    x.lineWidth = 2.5; x.strokeStyle = '#fff'; x.stroke();
    x.fillStyle = '#fff'; x.font = '700 15px -apple-system, "Segoe UI", Roboto, sans-serif'; x.textAlign = 'center'; x.textBaseline = 'middle';
    x.fillText(String(n), sz / 2, sz / 2 + 1);
    return x.getImageData(0, 0, c.width, c.height);
  }
  function ensureImage(id) {
    if (map.hasImage(id)) return;
    if (id.indexOf('pin-') === 0) map.addImage(id, pinImage(id), { pixelRatio: 2 });
    else if (id.indexOf('cl-') === 0) map.addImage(id, clusterImage(id.slice(3)), { pixelRatio: 2 });
  }

  // Arranca viendo la Tierra desde lejos; playGlobeIntro() baja hasta Ayala 71 y solo entonces
  // fija minZoom/maxBounds (si se fijaran ya, el mapa se autocorregiría y no se vería el globo).
  const GLOBE_START_ZOOM = 1.4;
  const GLOBE_START_BEARING = -42;   // el giro hacia 0° durante la intro da sensación de "encontrar" el punto
  const map = new maplibregl.Map({
    container: 'map', style: makeStyle(theme),
    center: ORIGEN.center, zoom: GLOBE_START_ZOOM, pitch: 0, bearing: GLOBE_START_BEARING, maxPitch: 60,
    maxZoom: 19, fadeDuration: 0,
    renderWorldCopies: false, refreshExpiredTiles: false,
    pixelRatio: Math.min(window.devicePixelRatio || 1, PIXEL_RATIO_MAX),
    attributionControl: false
  });
  map.addControl(new maplibregl.AttributionControl({
    compact: true,
    customAttribution: '<a href="https://openfreemap.org" target="_blank" rel="noopener">OpenFreeMap</a> © <a href="https://www.openmaptiles.org/" target="_blank" rel="noopener">OpenMapTiles</a> Datos de <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener">OpenStreetMap</a>'
  }));
  map.addControl(new maplibregl.NavigationControl({ visualizePitch: true }), 'top-right');
  map.addControl(new maplibregl.ScaleControl({ unit: 'metric' }), 'bottom-left');

  function targetPitch(z) {
    const t = Math.min(1, Math.max(0, (z - FLAT_UNTIL) / (FULL_3D_AT - FLAT_UNTIL)));
    return t * t * (3 - 2 * t) * MAX_PITCH_AUTO;
  }
  // La inclinación se ajusta al terminar el zoom (no en cada fotograma): así el gesto va suave
  map.on('zoomend', () => {
    if (!auto) return;
    const t = targetPitch(map.getZoom());
    if (Math.abs(t - map.getPitch()) > 1.5) map.easeTo({ pitch: t, duration: 450 });
  });
  map.on('movestart', e => { if (e.originalEvent) userMoved = true; });
  const autoBox = $('auto');
  map.on('pitchstart', e => { if (e.originalEvent && auto) { auto = false; autoBox.checked = false; } });
  autoBox.addEventListener('change', () => {
    auto = autoBox.checked;
    if (auto) map.easeTo({ pitch: targetPitch(map.getZoom()), duration: 600 });
  });

  const hud = $('hud');
  const modeEl = $('mode'), statsEl = $('stats');
  let hudT = 0, lastMode = null;
  function updateHud(force) {
    const now = performance.now();
    if (!force && now - hudT < 150) return;
    hudT = now;
    const z = map.getZoom(), p = map.getPitch();
    const is3d = z >= EXTRUDE_FROM && p > 4;
    if (is3d !== lastMode) { lastMode = is3d; modeEl.textContent = is3d ? '3D' : '2D'; modeEl.classList.toggle('is3d', is3d); }
    if (!hud.classList.contains('collapsed')) statsEl.textContent = 'zoom ' + z.toFixed(1) + ' · inclinación ' + Math.round(p) + '°';
  }
  map.on('move', () => updateHud(false));
  map.on('moveend', () => updateHud(true));
  map.on('load', () => updateHud(true));

  const busy = $('busy');
  let busyOn = false;
  map.on('dataloading', e => { if (!busyOn && e.dataType === 'source' && map.getZoom() >= EXTRUDE_FROM - 0.5) { busyOn = true; busy.classList.add('on'); } });
  map.on('idle', () => { if (busyOn) { busyOn = false; busy.classList.remove('on'); } });

  /* ---- Control de altura de edificios ---- */
  const hs = $('hscale'), hsOut = $('hscaleOut');
  hs.addEventListener('input', () => { hsOut.textContent = '×' + (+hs.value).toFixed(1).replace('.', ','); });
  hs.addEventListener('change', () => {          // se aplica al soltar: recalcular las alturas es costoso
    hScale = +hs.value;
    if (map.getLayer('edificios-3d')) {
      map.setPaintProperty('edificios-3d', 'fill-extrusion-height', hx(H_EXPR));
      map.setPaintProperty('edificios-3d', 'fill-extrusion-base', hx(MINH_EXPR));
    }
  });

  /* ---- Panel de opciones ---- */
  if (window.innerWidth >= 700) hud.classList.remove('collapsed');
  $('toggleHud').addEventListener('click', () => {
    const c = hud.classList.toggle('collapsed');
    $('toggleHud').setAttribute('aria-expanded', String(!c));
  });
  $('reset').addEventListener('click', () => {
    closeSheets();
    map.flyTo({ center: ORIGEN.center, zoom: ORIGEN.zoom, pitch: 0, bearing: 0, speed: 1, curve: 1.3 });
  });

  const msg = $('msg');
  let errors = 0;
  map.on('error', () => {
    errors++;
    if (errors >= 3 && !map.loaded()) { msg.style.display = 'block'; msg.textContent = 'No se pueden cargar los mapas. Comprueba tu conexión.'; }
  });
  map.on('idle', () => { msg.style.display = 'none'; });

  /* ================== DATOS ================== */

  function toFinca(raw, idx, categoria) {
    const nm = {};
    Object.keys(raw).forEach(k => { nm[norm(k)] = raw[k]; });
    const g = key => fmtVal(nm[norm(COL[key])]);
    const p = {
      idx: idx,
      name: String(g('name') || ''), tipo: String(g('tipo') || ''), zona: String(g('zona') || ''), addr: String(g('addr') || ''),
      dist: toNum(g('dist')), precio: String(g('precio') || ''), precioInv: toNum(g('precioInv')), precio120: toNum(g('precio120')),
      aforoMin: toNum(g('aforoMin')), aforoMax: toNum(g('aforoMax')), aforoDet: String(g('aforoDet') || ''),
      catering: String(g('catering') || ''), cateringLibre: String(g('cateringLibre') || ''),
      salones: String(g('salones') || ''), unoVarios: String(g('unoVarios') || ''), terraza: String(g('terraza') || ''),
      novios: String(g('novios') || ''), ceremonia: String(g('ceremonia') || ''), alojamiento: String(g('alojamiento') || ''),
      parking: String(g('parking') || ''), horario: String(g('horario') || ''), notas: String(g('notas') || ''),
      googleNota: toNum(g('googleNota')), googleResenas: toNum(g('googleResenas')), bodasNet: String(g('bodasNet') || ''),
      tel: String(g('tel') || ''), src: String(g('src') || ''), foto: String(g('foto') || ''), tiempoCoche: toNum(g('tiempoCoche')), insta: String(g('insta') || '').replace(/^@/, ''),
      categoria: categoria || String(g('categoria') || 'dentro'),
      estado: 'pendiente',
      lat: toNum(g('lat')), lon: toNum(g('lon')), located: false, approx: false, failed: false
    };
    if (p.lat != null && p.lon != null && p.lat > 39.5 && p.lat < 41.2 && p.lon > -4.6 && p.lon < -3) p.located = true;
    else { p.lat = null; p.lon = null; }
    return p;
  }

  function loadSheetJS() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = SHEETJS; s.onload = res; s.onerror = () => rej(new Error('No se pudo cargar el lector de Excel'));
      document.head.appendChild(s);
    });
  }

  function catKey(nombreHoja) {
    return SHEETS[nombreHoja] || 'dentro';
  }

  async function parseWorkbook(buf) {
    await loadSheetJS();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    let idx = 0, out = [];
    Object.keys(SHEETS).forEach(nombreHoja => {
      const ws = wb.Sheets[nombreHoja];
      if (!ws) return;
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
      rows.forEach(r => { const f = toFinca(r, idx, catKey(nombreHoja)); if (f.name) { out.push(f); idx++; } });
    });
    return out;
  }

  function bufToB64(buf) {
    let bin = '';
    const bytes = new Uint8Array(buf), chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return btoa(bin);
  }
  function b64ToBuf(b64) {
    const bin = atob(b64), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }
  function getCustomXlsx() { try { return JSON.parse(localStorage.getItem(CUSTOM_XLSX_KEY) || 'null'); } catch (e) { return null; } }
  function clearCustomXlsx() { try { localStorage.removeItem(CUSTOM_XLSX_KEY); } catch (e) { /* sin almacenamiento */ } }

  async function loadData() {
    const custom = getCustomXlsx();
    if (custom) {
      try {
        const list = await parseWorkbook(b64ToBuf(custom.data));
        if (list.length) return { list: list, from: 'tu Excel', name: custom.name };
      } catch (e) { /* el Excel guardado no se pudo leer: se descarta y se sigue con el del repositorio */ }
      clearCustomXlsx();
    }
    try {
      const r = await fetch(DATA_XLSX, { cache: 'no-cache' });
      if (r.ok) { const list = await parseWorkbook(await r.arrayBuffer()); if (list.length) return { list: list, from: 'Excel' }; }
    } catch (e) { /* si no hay Excel subido, se usa la copia JSON */ }
    const r = await fetch(DATA_JSON, { cache: 'no-cache' });
    if (!r.ok) throw new Error('No se encuentran los datos');
    const rows = await r.json();
    return { list: rows.map((x, i) => toFinca(x, i)).filter(p => p.name), from: 'copia' };
  }

  /* ================== LOCALIZACIÓN (si el Excel no trae coordenadas) ================== */
  let geoCache = {};
  try { geoCache = JSON.parse(localStorage.getItem(GEO_KEY) || '{}'); } catch (e) { geoCache = {}; }
  const saveGeo = () => { try { localStorage.setItem(GEO_KEY, JSON.stringify(geoCache)); } catch (e) { /* sin almacenamiento */ } };
  const geoKey = p => norm(p.name) + '|' + norm(p.addr);

  function cleanAddr(s) {
    return String(s || '').replace(/\(.*?\)/g, '').replace(/\s+/g, ' ').replace(/[,\s]+$/, '').trim();
  }
  async function nominatim(q) {
    const url = 'https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=es&q=' + encodeURIComponent(q);
    const r = await fetch(url);
    if (!r.ok) throw new Error('http ' + r.status);
    const j = await r.json();
    if (!j.length) return null;
    const x = j[0], cat = x.category || x.class, t = x.type;
    const approx = cat === 'highway' || cat === 'boundary' || (cat === 'place' && t !== 'house' && t !== 'building');
    return { lat: parseFloat(x.lat), lon: parseFloat(x.lon), approx: approx };
  }
  async function geocodePlace(p) {
    const cleaned = cleanAddr(p.addr), hasNum = /\d/.test(cleaned);
    const nameClean = p.name.replace(/\(.*?\)/g, '').trim();
    const qs = [];
    if (cleaned && hasNum) qs.push(cleaned + ', España');
    qs.push(nameClean + ', ' + (p.zona || 'Madrid') + ', España');
    for (let i = 0; i < qs.length; i++) {
      const res = await nominatim(qs[i]);
      if (res) return res;
      await sleep(1100);
    }
    return null;
  }

  const geoBox = $('geo'), geoText = $('geoText'), geoRetry = $('geoRetry');
  let geoRunning = false;
  async function geocodeAll() {
    if (geoRunning) return;
    const todo = places.filter(p => !p.located);
    if (!todo.length) { geoBox.classList.remove('on'); return; }
    geoRunning = true; geoRetry.hidden = true;
    let done = 0, netErrors = 0;
    geoBox.classList.add('on'); geoText.textContent = 'Localizando fincas… 0/' + todo.length;
    for (const p of todo) {
      const k = geoKey(p);
      let hit = geoCache[k];
      if (!hit) {
        try { hit = await geocodePlace(p); netErrors = 0; if (hit) { geoCache[k] = hit; saveGeo(); } await sleep(1100); }
        catch (e) { netErrors++; hit = null; if (netErrors >= 3) break; await sleep(1500); }
      }
      if (hit) { p.lat = hit.lat; p.lon = hit.lon; p.approx = !!hit.approx; p.located = true; syncPins(); }
      else p.failed = true;
      done++;
      geoText.textContent = 'Localizando fincas… ' + done + '/' + todo.length;
      renderList();
    }
    geoRunning = false;
    const pending = places.filter(p => !p.located).length;
    if (pending) { geoText.textContent = pending + ' sin ubicar. Puedes añadir Latitud y Longitud en el Excel.'; geoRetry.hidden = false; }
    else geoBox.classList.remove('on');
    fitAll();
  }
  geoRetry.addEventListener('click', () => { places.forEach(p => { p.failed = false; }); geocodeAll(); });

  function fitAll() {
    if (userMoved) return;
    const pts = places.filter(p => p.located);
    if (!pts.length) return;
    const b = new maplibregl.LngLatBounds();
    pts.forEach(p => b.extend([p.lon, p.lat]));
    map.fitBounds(b, { padding: { top: 90, bottom: 90, left: 40, right: 40 }, maxZoom: 13, duration: 900 });
  }

  /* ================== CHINCHETAS (capa del mapa) ================== */
  function match(p) { return filter === 'all' || p.categoria === filter; }
  map.on('style.load', () => ESTADOS.concat(ESTADOS.map(e => e + '-a')).map(e => 'pin-' + e).forEach(ensureImage));
  map.on('styleimagemissing', e => ensureImage(e.id));
  let syncT = 0;
  function syncPins() {
    clearTimeout(syncT);
    syncT = setTimeout(() => { const src = map.getSource('fincas'); if (src) src.setData(geojson()); }, 80);
  }
  function setSel(p) { if (map.getLayer('pin-sel')) map.setFilter('pin-sel', selFilter(p ? p.idx : null)); }
  ['pins', 'pin-clusters'].forEach(id => {
    map.on('mouseenter', id, () => { map.getCanvas().style.cursor = 'pointer'; });
    map.on('mouseleave', id, () => { map.getCanvas().style.cursor = ''; });
  });

  /* ================== FICHA ================== */
  const card = $('card'), listEl = $('list');
  function closeSheets() {
    card.classList.remove('open'); listEl.classList.remove('open');
    document.body.classList.remove('sheet-open');
    if (sel) { sel = null; setSel(null); clearRuta(); }
  }
  function openSheet(el) {
    [card, listEl].forEach(s => { s.classList.remove('open', 'peek'); s.style.maxHeight = ''; });
    el.classList.add('open'); document.body.classList.add('sheet-open'); el.scrollTop = 0;
    const cb = el.querySelector('.close'); if (cb) cb.focus({ preventScroll: true });   // el foco pasa a la hoja (teclado y lector de pantalla)
  }

  // Al minimizar, la hoja se recorta justo debajo de #rutaInfo (distancia/tiempo) si existe,
  // o debajo del título si no — así "bajar el nombre" deja ver ese cuadro, no solo el nombre a secas.
  function peekHeight(sheetEl) {
    const ref = sheetEl.querySelector('#rutaInfo') || sheetEl.querySelector('h2');
    if (!ref) return 70;
    return Math.ceil(ref.getBoundingClientRect().bottom - sheetEl.getBoundingClientRect().top) + 12;
  }
  function enterPeek(sheetEl) { sheetEl.style.maxHeight = peekHeight(sheetEl) + 'px'; sheetEl.classList.add('peek'); }
  function exitPeek(sheetEl) { sheetEl.classList.remove('peek'); sheetEl.style.maxHeight = ''; }

  // Arrastrar el tirador hacia abajo minimiza la ficha a solo el nombre («peek»); arrastrar hacia
  // arriba (o tocar el tirador) la vuelve a abrir del todo. Así se puede ver el mapa sin cerrar la ficha.
  // Delegado en el propio <section>, no en .grab: card.innerHTML se reemplaza en cada select() y
  // un listener puesto directamente en .grab quedaría huérfano en cuanto se recrease el nodo.
  function setupDrag(sheetEl) {
    const THRESH = 60;
    let startY = 0, dragging = false, wasPeek = false, moved = false, pid = null;
    sheetEl.addEventListener('pointerdown', e => {
      if (!e.target.closest('.grab')) return;
      dragging = true; moved = false; wasPeek = sheetEl.classList.contains('peek');
      startY = e.clientY; pid = e.pointerId;
      sheetEl.classList.add('dragging');
      sheetEl.setPointerCapture(pid);
    });
    sheetEl.addEventListener('pointermove', e => {
      if (!dragging || e.pointerId !== pid) return;
      const dy = e.clientY - startY;
      if (Math.abs(dy) > 4) moved = true;
      // 700-899px el CSS centra la hoja con translateX(-50%): hay que conservarlo al arrastrar
      const baseT = (window.innerWidth >= 700 && window.innerWidth < 900) ? 'translateX(-50%) ' : '';
      if ((!wasPeek && dy > 0) || (wasPeek && dy < 0)) sheetEl.style.transform = baseT + 'translateY(' + dy + 'px)';
    });
    const endDrag = e => {
      if (!dragging || e.pointerId !== pid) return;
      dragging = false;
      sheetEl.classList.remove('dragging');
      sheetEl.style.transform = '';
      const dy = e.clientY - startY;
      if (!wasPeek && dy > THRESH) enterPeek(sheetEl);
      else if (wasPeek && dy < -THRESH) exitPeek(sheetEl);
      else if (!moved) { if (sheetEl.classList.contains('peek')) exitPeek(sheetEl); else enterPeek(sheetEl); }
    };
    sheetEl.addEventListener('pointerup', endDrag);
    sheetEl.addEventListener('pointercancel', endDrag);
  }
  setupDrag(card); setupDrag(listEl);
  const row = (label, html) => html ? '<div class="r"><dt>' + label + '</dt><dd>' + html + '</dd></div>' : '';
  const CAT_LABEL = { dentro: 'Dentro de 20 km', fuera: '20-24 km', otro: 'Otro espacio' };

  function cardHtml(p) {
    const q = p.located ? p.lat + ',' + p.lon : encodeURIComponent(p.name + ' ' + cleanAddr(p.addr) + ' Madrid');
    const g = 'https://www.google.com/maps/search/?api=1&query=' + q;
    const ruta = p.located ? 'https://www.google.com/maps/dir/?api=1&origin=' + ORIGEN.center[1] + ',' + ORIGEN.center[0] + '&destination=' + p.lat + ',' + p.lon : '';
    const a = 'https://maps.apple.com/?q=' + encodeURIComponent(p.name) + (p.located ? '&ll=' + p.lat + ',' + p.lon : '');
    const web = safeUrl(p.src);
    const insta = p.insta ? 'https://www.instagram.com/' + encodeURIComponent(p.insta) + '/' : '';
    const cat = '<span class="chip cat" data-c="' + esc(p.categoria) + '">' + esc(CAT_LABEL[p.categoria] || p.categoria) + '</span>';
    const distChip = p.dist != null ? '<span class="chip">' + esc(String(p.dist).replace('.', ',')) + ' km</span>' : '';

    const aforo = p.aforoDet || (p.aforoMin != null || p.aforoMax != null ? [p.aforoMin, p.aforoMax].filter(x => x != null).join('–') + ' invitados' : '');
    const precioInv = p.precioInv != null ? eur(p.precioInv) + '/invitado' : '';
    // precio por invitado: directo si lo da el local, si no se deriva del coste para 120 que trae el Excel
    const perGuest = p.precioInv != null ? p.precioInv : (p.precio120 != null ? p.precio120 / 120 : null);
    const precioEst = perGuest != null ? '≈ ' + eur(perGuest * 250) + ' – ' + eur(perGuest * 300) + ' para 250-300 invitados' : '';
    const googleR = p.googleNota != null ? p.googleNota.toString().replace('.', ',') + ' / 5' + (p.googleResenas != null ? ' (' + p.googleResenas + ' reseñas)' : '') : '';

    const detailRows = [
      row('Tipo de espacio', esc(p.tipo)), row('Dirección', esc(p.addr)), row('Tiempo en coche desde Ayala 71', p.tiempoCoche != null ? p.tiempoCoche + ' min' : ''),
      row('Aforo', esc(aforo)), row('Salones / espacios', esc(p.salones)), row('¿Uno o varios?', esc(p.unoVarios)),
      row('Precio orientativo', esc(p.precio)), row('Menú', [esc(precioInv), esc(precioEst)].filter(Boolean).join(' · ')),
      row('Catering', esc(p.catering)), row('¿Catering libre?', esc(p.cateringLibre)),
      row('Terraza / jardín', esc(p.terraza)), row('Zona de novios', esc(p.novios)), row('Ceremonia en la finca', esc(p.ceremonia)),
      row('Alojamiento', esc(p.alojamiento)), row('Parking', esc(p.parking)), row('Horario', esc(p.horario)),
      row('Valoración Google', esc(googleR)), row('bodas.net', esc(p.bodasNet)), row('Teléfono', p.tel ? '<a href="tel:' + esc(p.tel.replace(/\s+/g, '')) + '">' + esc(p.tel) + '</a>' : ''),
      row('Notas', esc(p.notas))
    ].join('');

    const estadoChip = '<span class="chip estado" style="background:' + estadoColor(p.estado) + ';color:#fff;border-color:transparent">' + esc(ESTADO_LABEL[p.estado] || ESTADO_LABEL.pendiente) + '</span>';
    // foto guardada en el propio repositorio (fotos/): si el archivo faltara, se oculta sola
    const foto = p.foto ? '<img class="photo" src="' + esc(p.foto) + '" alt="" loading="lazy" onerror="this.remove()">' : '';

    return '<div class="grab"></div><button class="close" type="button" data-close aria-label="Cerrar">✕</button>' +
      '<h2>' + esc(p.name) + '</h2><p class="zone">' + esc(p.zona) + '</p>' +
      (p.located ? '<p id="rutaInfo" class="ruta-box"></p>' : '') +
      '<div class="chips">' + cat + distChip + estadoChip + '</div>' +
      foto +
      '<div class="actions"><a class="main" href="' + g + '" target="_blank" rel="noopener">Google Maps</a>' +
      (ruta ? '<a href="' + ruta + '" target="_blank" rel="noopener">Ruta desde Ayala 71</a>' : '') +
      '<a href="' + a + '" target="_blank" rel="noopener">Apple Maps</a>' +
      (web ? '<a href="' + esc(web) + '" target="_blank" rel="noopener">Web</a>' : '') +
      (insta ? '<a href="' + esc(insta) + '" target="_blank" rel="noopener">Instagram</a>' : '') +
      '<button type="button" class="editbtn" data-edit="' + p.idx + '">Editar</button></div>' +
      '<dl>' + detailRows + '</dl>' +
      (p._edited ? '<p class="note">Datos editados en este móvil.</p>' : '') +
      (p.approx ? '<p class="note">Ubicación aproximada en el mapa. Para afinarla, escribe Latitud y Longitud en el Excel.</p>' : '') +
      (!p.located ? '<p class="note">Todavía sin ubicar en el mapa. Añade Latitud y Longitud en el Excel.</p>' : '');
  }

  function select(p, fly) {
    sel = p; setSel(p);
    card.innerHTML = cardHtml(p);
    openSheet(card);
    drawRuta(p);
    if (fly && p.located) {
      map.flyTo({
        center: [p.lon, p.lat], zoom: Math.max(map.getZoom(), 14.5), bearing: map.getBearing(),
        pitch: auto ? targetPitch(15) : map.getPitch(),
        offset: [0, -Math.round(window.innerHeight * 0.17)], speed: 0.9, curve: 1.4
      });
    }
  }

  /* ================== LISTA Y FILTROS ================== */
  function renderList() {
    const t = norm($('search').value);
    const items = places.filter(p => match(p) && (!t || norm(p.name + ' ' + p.zona + ' ' + p.tipo).includes(t)));
    $('items').innerHTML = items.length ? items.map(p =>
      '<button type="button" class="item" data-i="' + p.idx + '"><span class="dot" style="background:' + estadoColor(p.estado) + '"></span>' +
      '<span class="t"><b>' + esc(p.name) + '</b><span>' + esc(p.zona) + (p.dist != null ? ' · ' + esc(String(p.dist).replace('.', ',')) + ' km' : '') + '</span></span>' +
      (p.located ? '' : '<span class="tag">sin ubicar</span>') + '</button>').join('') : '<p class="empty">No hay fincas con ese filtro.</p>';
    document.querySelectorAll('#bar [data-f]').forEach(b => {
      const f = b.dataset.f;
      const n = places.filter(p => f === 'all' || p.categoria === f).length;
      const label = f === 'all' ? 'Todas' : f === 'dentro' ? 'Dentro 20 km' : f === 'fuera' ? '20-24 km' : 'Otros';
      b.textContent = label + ' (' + n + ')';
      b.classList.toggle('on', f === filter);
    });
    $('btnList').textContent = 'Lista (' + places.length + ')';
  }
  $('items').addEventListener('click', e => {
    const b = e.target.closest('.item'); if (!b) return;
    const p = places[+b.dataset.i]; if (p) select(p, true);
  });
  $('search').addEventListener('input', renderList);
  $('btnList').addEventListener('click', () => { renderList(); openSheet(listEl); });
  document.querySelectorAll('#bar [data-f]').forEach(b => b.addEventListener('click', () => {
    filter = b.dataset.f;
    syncPins();
    renderList();
  }));
  document.addEventListener('click', e => { if (e.target.closest('[data-close]')) closeSheets(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeSheets(); });
  map.on('click', e => {
    const r = 5;
    const cf = map.getLayer('pin-clusters') ? map.queryRenderedFeatures([[e.point.x - 12, e.point.y - 12], [e.point.x + 12, e.point.y + 12]], { layers: ['pin-clusters'] }) : [];
    if (cf.length) {
      const f0 = cf[0], src = map.getSource('fincas');
      Promise.resolve(src.getClusterExpansionZoom(f0.properties.cluster_id))
        .then(z => map.easeTo({ center: f0.geometry.coordinates, zoom: Math.min((+z || 14) + 0.4, 17.5), duration: 700 }))
        .catch(() => map.easeTo({ center: f0.geometry.coordinates, zoom: map.getZoom() + 2 }));
      return;
    }
    const f = map.getLayer('pins') ? map.queryRenderedFeatures([[e.point.x - r, e.point.y - r], [e.point.x + r, e.point.y + r]], { layers: ['pins'] }) : [];
    if (f.length) { const p = places[f[0].properties.idx]; if (p) select(p, true); return; }
    if (card.classList.contains('open')) closeSheets();
  });

  /* ================== ARRANQUE ================== */
  function setPlaces(list) {
    closeSheets();
    places = list.map((p, i) => { p.idx = i; if (window.VT.edits) window.VT.edits.apply(p); return p; });
    syncPins();
    renderList();
    fitAll();
    geocodeAll();
    precacheOffline(places);   // en segundo plano; no bloquea nada de lo anterior
  }

  function start() {
    loadData().then(r => setPlaces(r.list)).catch(err => {
      msg.style.display = 'block';
      msg.textContent = 'No se han podido cargar las fincas. ' + (err && err.message ? err.message : '');
    });
  }

  // Intro: de la Tierra vista desde lejos a Ayala 71. Solo al terminar se fijan los límites normales
  // del mapa y se cargan las fincas (así no compite con el fitBounds de setPlaces()).
  function lockToOrigen() { map.setMinZoom(9.5); map.setMaxBounds(BOUNDS); }
  function playGlobeIntro() {
    const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduceMotion) { map.jumpTo(ORIGEN); lockToOrigen(); start(); return; }
    map.once('moveend', () => { lockToOrigen(); start(); });
    map.flyTo({ center: ORIGEN.center, zoom: ORIGEN.zoom, pitch: 0, bearing: 0, duration: 2800, curve: 1.3, essential: true });
  }
  if (map.loaded()) playGlobeIntro(); else map.once('load', playGlobeIntro);

  // Cargar tu propio Excel: se guarda en este dispositivo y se usa en cada visita hasta que lo quites
  const customInfo = $('customXlsxInfo'), customName = $('customXlsxName');
  function showCustomXlsxUI(name) { if (!customInfo) return; customName.textContent = name; customInfo.style.display = ''; }
  function hideCustomXlsxUI() { if (!customInfo) return; customInfo.style.display = 'none'; }
  (function initCustomXlsxUI() { const c = getCustomXlsx(); if (c) showCustomXlsxUI(c.name); })();

  $('loadXlsx').addEventListener('click', () => $('file').click());
  $('file').addEventListener('change', async e => {
    const f = e.target.files && e.target.files[0]; if (!f) return;
    try {
      userMoved = false;
      const buf = await f.arrayBuffer();
      setPlaces(await parseWorkbook(buf));
      try { localStorage.setItem(CUSTOM_XLSX_KEY, JSON.stringify({ name: f.name, data: bufToB64(buf) })); showCustomXlsxUI(f.name); }
      catch (e2) { /* no cupo en el almacenamiento: se usa solo en esta visita */ }
      geoBox.classList.add('on'); geoText.textContent = 'Usando tu Excel: ' + f.name + ' (guardado en este dispositivo)';
      setTimeout(() => { if (!geoRunning && !places.some(p => !p.located)) geoBox.classList.remove('on'); }, 3500);
    } catch (err) {
      msg.style.display = 'block'; msg.textContent = 'No se pudo leer ese Excel.';
      setTimeout(() => { msg.style.display = 'none'; }, 3500);
    }
    e.target.value = '';
  });
  const clearBtn = $('clearXlsx');
  if (clearBtn) clearBtn.addEventListener('click', () => {
    clearCustomXlsx(); hideCustomXlsxUI();
    loadData().then(r => setPlaces(r.list)).catch(() => {});
  });

  // para el editor (editor.js)
  window.VT = window.VT || {};
  window.VT.app = { places: () => places, select: select, map: map, COL: COL, SHEETS: SHEETS, geocode: geocodeAll,
    ensureXLSX: loadSheetJS, refresh: () => { syncPins(); renderList(); } };

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => { navigator.serviceWorker.register('sw.js').catch(() => {}); });
  }
})();
