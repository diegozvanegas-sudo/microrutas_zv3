// Genera data/centros_lotes.js: punto central [lat, lon] de cada lote, para la pre-ruta de subsanaciones.
// Volver a correr si cambia la capa de zonas verdes:  node tools/generar_centros_lotes.js
var fs = require('fs'), path = require('path');
var raiz = path.join(__dirname, '..');
var texto = fs.readFileSync(path.join(raiz, 'data/zonasverdes__areas_zonas_verdes_2.js'), 'utf8');
var geo = JSON.parse(texto.slice(texto.indexOf('{'), texto.lastIndexOf('}') + 1));

// Centroide ponderado por área de los anillos exteriores (en grados basta a escala de ciudad)
function sumar(anillo, acc) {
    for (var i = 0, j = anillo.length - 1; i < anillo.length; j = i++) {
        var f = anillo[j][0] * anillo[i][1] - anillo[i][0] * anillo[j][1];
        acc.a += f; acc.x += (anillo[j][0] + anillo[i][0]) * f; acc.y += (anillo[j][1] + anillo[i][1]) * f;
    }
}
var acum = {};
geo.features.forEach(function (f) {
    var lote = String(f.properties.No_Lote == null ? '' : f.properties.No_Lote).trim();
    if (!lote || lote === 'NA' || !f.geometry) return;
    var polys = f.geometry.type === 'Polygon' ? [f.geometry.coordinates] : f.geometry.type === 'MultiPolygon' ? f.geometry.coordinates : [];
    var acc = acum[lote] = acum[lote] || { a: 0, x: 0, y: 0, n: 0, sx: 0, sy: 0 };
    polys.forEach(function (p) {
        var a = { a: 0, x: 0, y: 0 };
        sumar(p[0], a);
        if (a.a < 0) { a.a = -a.a; a.x = -a.x; a.y = -a.y; }
        acc.a += a.a; acc.x += a.x; acc.y += a.y;
        p[0].forEach(function (c) { acc.sx += c[0]; acc.sy += c[1]; acc.n++; });
    });
});
var centros = {};
Object.keys(acum).forEach(function (l) {
    var c = acum[l], lon, lat;
    if (c.a > 0) { lon = c.x / (3 * c.a); lat = c.y / (3 * c.a); } else if (c.n) { lon = c.sx / c.n; lat = c.sy / c.n; } else return;
    centros[l] = [+lat.toFixed(6), +lon.toFixed(6)];
});
fs.writeFileSync(path.join(raiz, 'data/centros_lotes.js'),
    '// Centro [lat, lon] de cada lote (No_Lote). Generado con tools/generar_centros_lotes.js, no editar a mano.\n' +
    'var centrosLotes = ' + JSON.stringify(centros) + ';\n');
console.log(Object.keys(centros).length + ' lotes');
