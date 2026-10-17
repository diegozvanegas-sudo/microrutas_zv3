// Genera los datos livianos del mapa rápido (mapa_rapido.html) en data/rapido/.
// - Un archivo de lotes por cuadrilla, para cargar solo la que se necesita.
// - Coordenadas con 6 decimales (~10 cm), sin altura, y sin vértices que no cambian la forma (Douglas-Peucker ~0,5 m).
// - Cada lote ya trae la dirección, descripción, frecuencia y comuna (no hace falta cargar direcciones_*.js).
// Volver a correr si cambian las capas de data/:  node tools/generar_mapa_rapido.js
var fs = require('fs'), path = require('path');
var raiz = path.join(__dirname, '..'), salida = path.join(raiz, 'data', 'rapido');
if (!fs.existsSync(salida)) fs.mkdirSync(salida);

function leerGeo(archivo) {
    var t = fs.readFileSync(path.join(raiz, 'data', archivo), 'utf8');
    return JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1));
}
function leerVar(archivo, nombre) {
    var t = fs.readFileSync(path.join(raiz, 'data', archivo), 'utf8'), ctx = {};
    new Function('ctx', t.replace('var ' + nombre, 'ctx.' + nombre))(ctx);
    return ctx[nombre];
}

// ---------- Simplificación ----------
function distSeg(p, a, b) {
    var dx = b[0] - a[0], dy = b[1] - a[1], l = dx * dx + dy * dy, t = l ? ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l : 0;
    t = Math.max(0, Math.min(1, t));
    var x = a[0] + t * dx - p[0], y = a[1] + t * dy - p[1];
    return Math.sqrt(x * x + y * y);
}
function douglas(pts, tol) {
    if (pts.length < 3) return pts;
    var marca = new Array(pts.length), pila = [[0, pts.length - 1]];
    marca[0] = marca[pts.length - 1] = true;
    while (pila.length) {
        var r = pila.pop(), max = 0, k = -1;
        for (var i = r[0] + 1; i < r[1]; i++) { var d = distSeg(pts[i], pts[r[0]], pts[r[1]]); if (d > max) { max = d; k = i; } }
        if (max > tol) { marca[k] = true; pila.push([r[0], k], [k, r[1]]); }
    }
    return pts.filter(function (p, i) { return marca[i]; });
}
function redondear(p) { return [Math.round(p[0] * 1e6) / 1e6, Math.round(p[1] * 1e6) / 1e6]; }
function anillo(r, tol) {
    var pts = r.map(redondear).filter(function (p, i, a) { return !i || p[0] !== a[i - 1][0] || p[1] !== a[i - 1][1]; });
    var s = douglas(pts, tol);
    return s.length >= 4 ? s : pts.length >= 4 ? pts : null; // un anillo cerrado necesita al menos 4 puntos
}
function linea(l, tol) { var s = douglas(l.map(redondear), tol); return s.length >= 2 ? s : null; }
function geometria(g, tol) {
    if (!g) return null;
    if (g.type === 'Polygon' || g.type === 'MultiPolygon') {
        var polys = (g.type === 'Polygon' ? [g.coordinates] : g.coordinates).map(function (p) {
            var ext = anillo(p[0], tol);
            return ext ? [ext].concat(p.slice(1).map(function (h) { return anillo(h, tol); }).filter(Boolean)) : null;
        }).filter(Boolean);
        return polys.length ? { type: 'MultiPolygon', coordinates: polys } : null;
    }
    if (g.type === 'LineString' || g.type === 'MultiLineString') {
        var ls = (g.type === 'LineString' ? [g.coordinates] : g.coordinates).map(function (l) { return linea(l, tol); }).filter(Boolean);
        return ls.length ? { type: 'MultiLineString', coordinates: ls } : null;
    }
    return g;
}

// Geometría compacta: cada anillo/línea es una lista plana de enteros (grados × 1e6), el primer
// punto absoluto y los demás como diferencia con el anterior: [lon0, lat0, dLon1, dLat1, ...].
// El mapa lo convierte de vuelta a [lat, lon] para Leaflet.
function codificar(g) {
    var plano = function (pts) {
        var out = [], px = 0, py = 0;
        pts.forEach(function (p) { var x = Math.round(p[0] * 1e6), y = Math.round(p[1] * 1e6); out.push(x - px, y - py); px = x; py = y; });
        return out;
    };
    return g.type === 'MultiPolygon' ? g.coordinates.map(function (pol) { return pol.map(plano); }) : g.coordinates.map(plano);
}

// ---------- Textos (misma lógica que la ventana del mapa completo) ----------
function arreglarTildes(s) {
    s = String(s == null ? '' : s);
    try { return /Ã|Â/.test(s) ? decodeURIComponent(escape(s)) : s; } catch (e) { return s; }
}
function limpio(v) { var s = String(v == null ? '' : v).trim(); return s === '' || s === 'null' ? undefined : s; }

var TOL = 0.000005; // ~0,5 m en Medellín
var prog = leerVar('direcciones_lotes.js', 'direccionesLotes');
var extras = leerVar('direcciones_fid.js', 'extrasLotes');
var zonas = leerGeo('zonasverdes__areas_zonas_verdes_2.js');

var grupos = {}, indice = { cuadrillas: {}, lotes: {}, mrs: {} };
zonas.features.forEach(function (f) {
    var p = f.properties || {}, g = geometria(f.geometry, TOL);
    if (!g) return;
    var mr = limpio(p.Microruta), cuad = limpio(p.CUADRILLA) || 'sin', lote = limpio(p.No_Lote);
    var pr = (mr && prog[mr]) || {}, ex = (p.fid != null && extras[p.fid]) || {};
    var dañada = /ï¿½|�/.test(String(p.Descripcio || ''));
    var desc = dañada && pr.desc ? pr.desc : arreglarTildes(p.Descripcio || pr.desc || '').replace(/ï¿½|�/g, '?');
    var props = {
        l: lote, m: mr, c: cuad === 'sin' ? undefined : cuad,
        a: limpio(p.Area_campo) || (p.Area_m2 ? String(Math.round(p.Area_m2)) : undefined),
        d: limpio(p.DIRECCION) || limpio(pr.d) || limpio(ex.d), s: limpio(desc), f: limpio(pr.frec), k: limpio(ex.c)
    };
    (grupos[cuad] = grupos[cuad] || []).push({ properties: props, geometry: g });
    if (lote && lote !== 'NA' && !indice.lotes[lote]) indice.lotes[lote] = cuad;
    if (mr && !indice.mrs[mr.toUpperCase()]) indice.mrs[mr.toUpperCase()] = cuad;
});

function limites(features) {
    var b = [90, 180, -90, -180];
    features.forEach(function (f) {
        JSON.stringify(f.geometry.coordinates).replace(/\[(-?[\d.]+),(-?[\d.]+)\]/g, function (_, x, y) {
            x = +x; y = +y;
            if (y < b[0]) b[0] = y; if (x < b[1]) b[1] = x; if (y > b[2]) b[2] = y; if (x > b[3]) b[3] = x;
            return '';
        });
    });
    return [[b[0], b[1]], [b[2], b[3]]];
}
function escribir(nombre, llamada, clave, datos) {
    var txt = llamada + '(' + JSON.stringify(clave) + ',' + JSON.stringify(datos) + ');\n';
    fs.writeFileSync(path.join(salida, nombre), txt);
    return txt.length;
}

// Dirección, descripción, frecuencia y comuna se repiten en todos los polígonos de una microrruta:
// cada texto va una vez en "t" y el lote guarda su posición (el mapa los resuelve al cargar)
function deduplicar(features) {
    var t = [], pos = {};
    features.forEach(function (f) {
        ['d', 's', 'f', 'k'].forEach(function (k) {
            var v = f.properties[k];
            if (v === undefined) return;
            if (pos[v] === undefined) { pos[v] = t.length; t.push(v); }
            f.properties[k] = pos[v];
        });
    });
    return { t: t, lotes: features.map(function (f) { var p = f.properties; p.g = codificar(f.geometry); return p; }) };
}

var total = 0;
Object.keys(grupos).forEach(function (c) {
    var bytes = escribir('lotes_' + c + '.js', 'cargarLotes', c, deduplicar(grupos[c]));
    indice.cuadrillas[c] = { n: grupos[c].length, kb: Math.round(bytes / 1024), b: limites(grupos[c]) };
    total += bytes;
    console.log('lotes_' + c + '.js', grupos[c].length + ' lotes', Math.round(bytes / 1024) + ' KB');
});

// Capas de apoyo (se cargan solo si se activan)
var capas = {
    comunas: { archivo: 'Comunas_1.js', tol: 0.00002, props: { n: 'NOMBRE' } },
    operativos: { archivo: 'Operativos_4.js', tol: TOL, props: { s: 'Descripcion', l: 'No_Lote', a: 'Area_campo' } },
    rutas_sif: { archivo: 'Rutas_SIF_3.js', tol: TOL, props: { s: 'Descripcion', a: 'area_estimada' } },
    espartillos: { archivo: 'Espartillos_6.js', tol: TOL, props: { a: 'area', s: 'Descripci', d: 'Direccion', o: 'Codigo' } }
};
indice.capas = {};
Object.keys(capas).forEach(function (k) {
    var cfg = capas[k], geo = leerGeo(cfg.archivo);
    var feats = geo.features.map(function (f) {
        var g = geometria(f.geometry, cfg.tol), props = {};
        Object.keys(cfg.props).forEach(function (c) { props[c] = limpio(arreglarTildes(f.properties[cfg.props[c]])); });
        if (!g) return null;
        props.g = codificar(g);
        return props;
    }).filter(Boolean);
    var lineas = geo.features.some(function (f) { return f.geometry && /Line/.test(f.geometry.type); });
    var bytes = escribir('capa_' + k + '.js', 'cargarCapa', k, { tipo: lineas ? 'linea' : 'poligono', items: feats });
    indice.capas[k] = { kb: Math.round(bytes / 1024) };
    console.log('capa_' + k + '.js', Math.round(bytes / 1024) + ' KB (antes ' + Math.round(fs.statSync(path.join(raiz, 'data', cfg.archivo)).size / 1024) + ' KB)');
});

fs.writeFileSync(path.join(salida, 'indice.js'),
    '// Índice del mapa rápido. Generado con tools/generar_mapa_rapido.js, no editar a mano.\nvar indiceRapido = ' + JSON.stringify(indice) + ';\n');
console.log('Todos los lotes: ' + Math.round(total / 1024) + ' KB (antes ' + Math.round(fs.statSync(path.join(raiz, 'data/zonasverdes__areas_zonas_verdes_2.js')).size / 1024) + ' KB)');
console.log('indice.js ' + Math.round(fs.statSync(path.join(salida, 'indice.js')).size / 1024) + ' KB');
