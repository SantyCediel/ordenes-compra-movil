// Ubica cada orden en un municipio a partir del texto libre del campo «Campo / lugar de entrega».
// Corrige errores de escritura, entiende nombres de campos y estaciones conocidos, usa el departamento
// escrito para los nombres repetidos y, si el texto no dice dónde (una dirección, una persona), toma el
// lugar más común de las demás órdenes del mismo contrato. Funciona igual en el celular y en Node.
(function (raiz) {
  "use strict";
  const TILDES = new RegExp("[" + String.fromCharCode(0x300) + "-" + String.fromCharCode(0x36f) + "]", "g");
  const norm = s => String(s ?? "").normalize("NFD").replace(TILDES, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();

  // nombres cortos con que se conoce a municipios de nombre oficial largo
  const CORTOS = {
    "BOGOTA": ["BOGOTA D C", "BOGOTA D C"], "CARTAGENA": ["CARTAGENA DE INDIAS", "BOLIVAR"], "CUCUTA": ["SAN JOSE DE CUCUTA", "NORTE DE SANTANDER"],
    "TOLU": ["SANTIAGO DE TOLU", "SUCRE"], "UBATE": ["VILLA DE SAN DIEGO DE UBATE", "CUNDINAMARCA"], "MOMPOX": ["SANTA CRUZ DE MOMPOX", "BOLIVAR"],
    "MOMPOS": ["SANTA CRUZ DE MOMPOX", "BOLIVAR"], "BUGA": ["GUADALAJARA DE BUGA", "VALLE DEL CAUCA"], "MARIQUITA": ["SAN SEBASTIAN DE MARIQUITA", "TOLIMA"],
    "TUMACO": ["SAN ANDRES DE TUMACO", "NARINO"], "BARRANCA": ["BARRANCABERMEJA", "SANTANDER"], "BCA": ["BARRANCABERMEJA", "SANTANDER"],
  };
  // campos, estaciones y sitios que no son municipios (verificados en fuentes públicas)
  const SITIOS = {
    "SAN CAYETANO": ["BOGOTA D C", "BOGOTA D C"],          // Complejo Logístico San Cayetano, Engativá (Bogotá)
    "CELTA": ["FUNZA", "CUNDINAMARCA"],                     // Parque Industrial Celta Trade Park
    "RUBIALES": ["PUERTO GAITAN", "META"], "CEPSA": ["PUERTO GAITAN", "META"], "CPE6": ["PUERTO GAITAN", "META"], "CPE 6": ["PUERTO GAITAN", "META"],
    "CUSIANA": ["TAURAMENA", "CASANARE"],                   // estación Cusiana
    "GALAN": ["BARRANCABERMEJA", "SANTANDER"], "LA CIRA": ["BARRANCABERMEJA", "SANTANDER"], "EL CENTRO": ["BARRANCABERMEJA", "SANTANDER"],
    "CHIMITA": ["GIRON", "SANTANDER"],                      // poliducto Galán–Chimitá
    "SEBASTOPOL": ["CIMITARRA", "SANTANDER"],               // planta Sebastopol, Puerto Olaya
    "CHIQUILLO": ["REMEDIOS", "ANTIOQUIA"],                 // estación Chiquillo de Ocensa
    "SANTA ROSA": ["BOLIVAR", "SANTANDER"],                 // planta Santa Rosa de Cenit, vía Landázuri
    "BELENCITO": ["NOBSA", "BOYACA"], "BONGA MAMEY": ["SAN PEDRO", "SUCRE"],
  };
  // palabras que no dicen dónde: se quitan antes de buscar
  const RUIDO = new Set(("ENTREGAR A EST ESTACION BODEGA BODEGAS ALMACEN ALMACENES OFICINA OFICINAS COMPLEJO LOGISTICO SALA SEGURIDAD PARQUE " +
    "INDUSTRIAL CORREG CORREGIMIENTO VEREDA CAMPO NUCLEO AREA ZONA X IV VII Y CONFIPETROL ENEL OCENSA CENIT ECOPETROL PAREX HOCOL " +
    "FRONTERA EMERALD LATAM GCT PLANTA BASE SEDE PTO").split(" "));
  const DEPTO_ALIAS = { "VALLE": "VALLE DEL CAUCA", "GUAJIRA": "LA GUAJIRA", "PTYO": "PUTUMAYO", "BOGOTA": "BOGOTA D C", "NORTE SANTANDER": "NORTE DE SANTANDER" };
  const DIRECCION = /\b(CL|CLL|CALLE|CRA|KR|CARRERA|CR|AV|AVENIDA|MZ|MANZANA|CASA|CS|APTO|TRANSVERSAL|DIAGONAL|KM)\b/;
  const QUITAR = /\b(CAMINOS?|CONJ\w*|CONJUNTO|BARRIO|URB\w*|EDIFICIO|AUTOPISTA)\b( (DE|DEL|LA|LOS))?( \w+){0,2}/g;

  // distancia de edición con transposiciones (Damerau, versión OSA), con corte temprano
  function distancia(a, b, tope) {
    if (Math.abs(a.length - b.length) > tope) return tope + 1;
    let p2 = null, p = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) {
      const c = [i]; let min = i;
      for (let j = 1; j <= b.length; j++) {
        let v = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        if (p2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) v = Math.min(v, p2[j - 2] + 1);
        c.push(v); if (v < min) min = v;
      }
      if (min > tope) return tope + 1;
      p2 = p; p = c;
    }
    return p[b.length];
  }
  const tolerancia = n => (n <= 4 ? 0 : n <= 7 ? 1 : n <= 11 ? 2 : 3);

  function crear(municipios) {
    // municipios: [codigo, nombre, departamento, x, y]
    const porCodigo = new Map(), indice = new Map(), deptos = new Map();
    const poner = (clave, cod, peso) => { if (!clave) return; const l = indice.get(clave) || []; if (!l.some(e => e.cod === cod)) l.push({ cod, peso }); indice.set(clave, l); };
    for (const m of municipios) {
      const [cod, nombre, dpto] = m, n = norm(nombre), d = norm(dpto);
      porCodigo.set(cod, { cod, nombre, dpto, x: m[3], y: m[4], d });
      poner(n, cod, 0); poner(n.replace(/ /g, ""), cod, 0);
      deptos.set(d, d);
    }
    for (const [k, v] of Object.entries(DEPTO_ALIAS)) deptos.set(k, v);
    const buscar = (nombre, dpto) => [...porCodigo.values()].find(m => norm(m.nombre) === nombre && m.d === dpto);
    const sitios = new Map();
    for (const [k, [n, d]] of Object.entries({ ...CORTOS, ...SITIOS })) { const m = buscar(n, d); if (m) sitios.set(k, m.cod); }
    const claves = [...indice.keys()];

    // ¿qué departamento menciona el texto? (con tolerancia a errores en nombres largos)
    function depto(tokens) {
      for (let w = 3; w >= 1; w--) for (let i = 0; i + w <= tokens.length; i++) {
        const t = tokens.slice(i, i + w).join(" ");
        if (deptos.has(t)) return { d: deptos.get(t), i, w };
        if (t.length >= 6) for (const k of deptos.keys()) if (k.length >= 6 && distancia(t, k, 1) <= 1) return { d: deptos.get(k), i, w };
      }
      return null;
    }

    // resultado: { cods: [candidatos], como: exacto|corregido|sitio, fuerza: 2 fuerte | 1 débil }
    function leer(texto) {
      let t = norm(texto).replace(/\bENTREGAR A\b/g, " ").replace(/\bPTO\b/g, "PUERTO");
      const esDireccion = DIRECCION.test(t);
      t = t.replace(QUITAR, " ").replace(/\s+/g, " ").trim();
      const tokens = t.split(" ").filter(Boolean);
      if (!tokens.length || /\bKM ?\d|\bKM\b/.test(norm(texto))) return null; // un kilómetro de vía no dice el municipio
      const hint = depto(tokens);
      const util = tokens.filter((x, i) => !RUIDO.has(x) && !/\d/.test(x) && !(hint && i >= hint.i && i < hint.i + hint.w));
      const hallazgos = [];
      for (let w = Math.min(4, tokens.length); w >= 1; w--) for (let i = 0; i + w <= tokens.length; i++) {
        const k = tokens.slice(i, i + w).join(" "), pegado = k.replace(/ /g, "");
        if (sitios.has(k)) { hallazgos.push({ i, w, cods: [sitios.get(k)], como: "sitio" }); continue; }
        for (const clave of w > 1 ? [k, pegado] : [k]) if (indice.has(clave) && !(w === 1 && RUIDO.has(clave))) { // VILLA GARZON = VILLAGARZON
          const l = indice.get(clave), min = Math.min(...l.map(e => e.peso));
          hallazgos.push({ i, w, cods: l.filter(e => e.peso === min).map(e => e.cod), como: "exacto" });
          break;
        }
      }
      // primero lo que cuadra con el departamento escrito; luego lo que aparece antes y lo más largo
      const cuadra = x => !hint || x.cods.some(c => porCodigo.get(c).d === hint.d);
      let h = hallazgos.sort((a, b) => cuadra(b) - cuadra(a) || a.i - b.i || b.w - a.w)[0];
      if (!h && !esDireccion) {
        // error de escritura: el municipio más parecido, si es claramente el más cercano
        let mejor = null;
        for (let w = 1; w <= Math.min(3, tokens.length); w++) for (let i = 0; i + w <= tokens.length; i++) {
          const k = tokens.slice(i, i + w).join(" ");
          if (k.length < 5 || RUIDO.has(k) || /\d/.test(k)) continue;
          const tol = tolerancia(k.length);
          for (const c of claves) {
            if (c[0] !== k[0] && c[1] !== k[1]) continue;
            const d = distancia(k, c, tol + 1);
            const vale = d <= tol || (d === tol + 1 && c.slice(0, 3) === k.slice(0, 3) && k.length >= 7);
            if (vale && (!mejor || d < mejor.d || (d === mejor.d && w > mejor.w))) mejor = { d, w, i, c };
          }
          for (const [s, cod] of sitios) if (s.length >= 5 && distancia(k, s, tolerancia(k.length)) <= tolerancia(k.length) && (!mejor || mejor.d > 0)) mejor = { d: 0, w, i, c: s, cod };
        }
        if (mejor) {
          const l = mejor.cod ? [{ cod: mejor.cod, peso: 0 }] : indice.get(mejor.c), min = Math.min(...l.map(e => e.peso));
          h = { i: mejor.i, w: mejor.w, cods: l.filter(e => e.peso === min).map(e => e.cod), como: "corregido" };
        }
      }
      if (!h) return null;
      let cods = h.cods;
      if (hint) { const f = cods.filter(c => porCodigo.get(c).d === hint.d); if (f.length) cods = f; }
      // ¿sobran palabras que no son lugares? (p. ej. el nombre de una persona): el hallazgo es débil
      const resto = util.filter((x, j) => { const pos = tokens.indexOf(x); return pos < h.i || pos >= h.i + h.w; });
      const fuerza = resto.length === 0 ? 2 : 1;
      // parecidos cercanos, por si el cliente trabaja en otro sitio (CIÉNEGA de Boyacá vs. CIÉNAGA de Magdalena)
      const tramo = tokens.slice(h.i, h.i + h.w).join(" "), tolA = Math.max(1, tolerancia(tramo.length)), alternativas = [];
      if (tramo.length >= 5 && h.como !== "sitio") for (const c of claves) {
        if (c[0] !== tramo[0] || c === tramo) continue;
        const dist = distancia(tramo, c, tolA + 1);
        if (dist <= tolA + 1) for (const e of indice.get(c)) if (!cods.includes(e.cod) && !alternativas.some(a => a.cod === e.cod)) alternativas.push({ cod: e.cod, d: dist });
      }
      // un nombre bien escrito solo se cambia por otro a UNA letra (Ciénega/Ciénaga); uno mal escrito, por los cercanos
      const cambiables = alternativas.filter(a => h.como === "exacto" ? a.d <= 1 : true).map(a => a.cod);
      return { cods, como: h.como, fuerza, hint: hint?.d, alternativas: cambiables };
    }

    // ubica todas las órdenes; o.ca = texto, o.ct = contrato, o.c = cliente
      const ctDe = o => (o.ct ? o.c + "|" + o.ct : ""); // el contrato, siempre dentro del mismo cliente
    function ubicar(ordenes) {
      const cache = new Map(), lectura = o => { const k = o.ca || ""; if (!cache.has(k)) cache.set(k, leer(k)); return cache.get(k); };
      // 1) lo seguro: un solo candidato y sin palabras de más
      const porGrupo = { ct: new Map(), c: new Map() }, propio = new Map();
      const sumar = (mapa, k, cod) => { if (!k) return; const m = mapa.get(k) || new Map(); m.set(cod, (m.get(cod) || 0) + 1); mapa.set(k, m); };
      for (const o of ordenes) {
        const r = lectura(o);
        if (r && r.cods.length === 1 && r.fuerza === 2 && r.como !== "corregido") { // solo lo escrito sin errores
          sumar(porGrupo.ct, ctDe(o), r.cods[0]); sumar(porGrupo.c, o.c, r.cods[0]);
          sumar(propio, "ct|" + ctDe(o) + "|" + (o.ca || ""), r.cods[0]); sumar(propio, "c|" + o.c + "|" + (o.ca || ""), r.cods[0]);
        }
      }
      // dónde trabaja el contrato (o el cliente) según sus OTRAS órdenes: sin contar las del mismo texto
      const otras = (g, k, o) => {
        const base = porGrupo[g].get(k); if (!base) return null;
        const m = new Map(base), menos = propio.get(g + "|" + k + "|" + (o.ca || ""));
        if (menos) for (const [cod, n] of menos) { const v = (m.get(cod) || 0) - n; if (v > 0) m.set(cod, v); else m.delete(cod); }
        return m.size ? m : null;
      };
      const cuenta = m => (m ? [...m.values()].reduce((s, n) => s + n, 0) : 0);
      // el contrato si tiene suficientes órdenes; si no, el cliente
      const contexto = o => { const a = otras("ct", ctDe(o), o); return cuenta(a) >= 3 ? a : otras("c", o.c, o) || a; };
      const masComun = m => m && [...m].sort((a, b) => b[1] - a[1])[0]?.[0];
      const deptosDe = m => { const d = new Map(); if (m) for (const [cod, n] of m) { const x = porCodigo.get(cod).d; d.set(x, (d.get(x) || 0) + n); } return d; };
      const resultado = [];
      for (const o of ordenes) {
        const r = lectura(o), ctx = contexto(o);
        let cod = null, como = null;
        if (r) {
          let cods = r.cods;
          if (cods.length > 1) { // nombre repetido: el departamento donde trabaja ese contrato
            const d = deptosDe(ctx); const f = cods.filter(c => d.has(porCodigo.get(c).d)).sort((a, b) => d.get(porCodigo.get(b).d) - d.get(porCodigo.get(a).d));
            cods = f.length ? f : cods;
          }
          const d = deptosDe(ctx), total = [...d.values()].reduce((s, n) => s + n, 0);
          const peso = c => (d.get(porCodigo.get(c).d) || 0) / (total || 1), enCtx = c => peso(c) >= 0.25;
          const coherente = !total || enCtx(cods[0]) || !!r.hint;
          // no cuadra con donde trabaja el cliente y un municipio casi igual SÍ sale en sus otras órdenes:
          // es un error de escritura que dio con otro municipio (CIENEGA de Boyacá por CIÉNAGA de Magdalena)
          if (!coherente && total >= 3 && r.como !== "sitio") {
            const alt = r.alternativas.filter(enCtx).sort((a, b) => peso(b) - peso(a))[0];
            if (alt) { cod = alt; como = "corregido"; }
          }
          // una corrección con palabras de más (p. ej. el nombre de una persona) no se acepta
          const conocido = !!(ctx && ctx.has(cods[0])); // el cliente ya tiene otras órdenes en ese municipio
          if (!cod && (r.fuerza === 2 || (r.como !== "corregido" && (coherente || conocido)))) { cod = cods[0]; como = r.como; }
        }
        if (!cod) {
          const c1 = masComun(porGrupo.ct.get(ctDe(o))), c2 = masComun(porGrupo.c.get(o.c));
          if (c1) { cod = c1; como = "contrato"; } else if (c2) { cod = c2; como = "cliente"; }
        }
        resultado.push(cod ? { cod, como } : null);
      }
      // el mismo texto, el mismo lugar: las estimaciones de un texto repetido se unifican
      const votos = new Map();
      ordenes.forEach((o, i) => { const r = resultado[i]; if (!r) return; const k = norm(o.ca); const m = votos.get(k) || new Map(); m.set(r.cod, (m.get(r.cod) || 0) + 1); votos.set(k, m); });
      ordenes.forEach((o, i) => {
        const r = resultado[i];
        if (r && r.como !== "contrato" && r.como !== "cliente") return;
        const g = masComun(votos.get(norm(o.ca)));
        if (g && (!r || g !== r.cod)) resultado[i] = { cod: g, como: "cliente" };
      });
      return resultado;
    }

    return { leer, ubicar, municipio: cod => porCodigo.get(cod), norm };
  }

  const api = { crear, norm, distancia };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else raiz.Lugares = api;
})(typeof window !== "undefined" ? window : globalThis);
