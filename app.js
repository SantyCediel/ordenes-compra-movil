// Órdenes de Compra para el celular. Lee _AUTOMATIZACION/movil/datos.json de tu OneDrive
// con tu cuenta (Microsoft Graph, solo lectura). Sin servidor propio: nada sale de Microsoft 365.
"use strict";

const C = window.CONFIG_OC || {};
const ALCANCES = ["User.Read", "Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
// dentro de la aplicación del computador: los datos vienen del disco y los PDF los abre Windows
const LOCAL = new URLSearchParams(location.search).has("local");
// solo el mapa, dentro del Inicio de la aplicación del computador
const INCRUSTADO = new URLSearchParams(location.search).has("incrustado");
if (INCRUSTADO) document.documentElement.classList.add("incrustado");
const escritorio = m => { try { window.chrome?.webview?.postMessage(m); return !!window.chrome?.webview; } catch { return false; } };
const $ = s => document.querySelector(s);
const vista = $("#vista");

let D = null;          // datos.json
let msalApp = null, cuenta = null;
let actual = "inicio"; // vista de fondo
let instalar = null;   // aviso de instalación de Android
let dibujado = false;
let coloresCliente = null; // cliente -> casilla de color (fija, por lo comprado en total)

// ---------------------------------------------------------------- utilidades
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const TILDES = new RegExp("[" + String.fromCharCode(0x300) + "-" + String.fromCharCode(0x36f) + "]", "g");
const normal = s => String(s ?? "").normalize("NFD").replace(TILDES, "").toLowerCase();
const fmtCop = new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
const fmtNum = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 2 });
const fmtFecha = new Intl.DateTimeFormat("es-CO", { day: "numeric", month: "short", year: "numeric" });
const fmtDia = new Intl.DateTimeFormat("es-CO", { weekday: "long", day: "numeric", month: "long" });
const fmtMes = new Intl.DateTimeFormat("es-CO", { month: "short" });
const fmtMesLargo = new Intl.DateTimeFormat("es-CO", { month: "long", year: "numeric" });
const NL = String.fromCharCode(10);

function dinero(v, m) {
  if (v == null) return "—";
  if (!m || m === "COP") return fmtCop.format(v);
  return m + " " + new Intl.NumberFormat("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
}
function corto(v) {
  if (v == null) return "—";
  if (Math.abs(v) >= 1e9) return "$" + fmtNum.format(Math.round(v / 1e8) / 10) + " mil M";
  if (Math.abs(v) >= 1e6) return "$" + fmtNum.format(Math.round(v / 1e5) / 10) + " M";
  return fmtCop.format(v);
}
const aFecha = iso => iso ? new Date(iso.slice(0, 10) + "T12:00:00") : null;
const fecha = iso => iso ? fmtFecha.format(aFecha(iso)).split(".").join("") : "—";
const fmtDiaMes = new Intl.DateTimeFormat("es-CO", { day: "numeric", month: "short" });
const fechaCorta = iso => !iso ? "—" : iso.slice(0, 4) === hoyIso().slice(0, 4) ? fmtDiaMes.format(aFecha(iso)).split(".").join("") : fecha(iso);
const hoyIso = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
function hace(iso) {
  const min = Math.round((Date.now() - new Date(iso)) / 60000);
  if (min < 1) return "hace un momento";
  if (min < 60) return `hace ${min} min`;
  const h = Math.round(min / 60);
  if (h < 24) return `hace ${h} h`;
  const d = Math.round(h / 24);
  return d === 1 ? "ayer" : `hace ${d} días`;
}
const diasEntre = (a, b) => Math.round((aFecha(b) - aFecha(a)) / 864e5);
const isoDe = d => new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
function mesInicio(atras) { const d = aFecha(hoyIso()); d.setDate(1); d.setMonth(d.getMonth() - atras); return isoDe(d); }
const masDias = (iso, n) => { const d = aFecha(iso); d.setDate(d.getDate() + n); return isoDe(d); };
const plural = (n, uno, varios) => `${fmtNum.format(n)} ${n === 1 ? uno : varios}`;

const RELLENO = new Set(["SAS", "S.A.S", "S.A.S.", "SA", "S.A", "S.A.", "LTDA", "LTDA.", "DE", "DEL", "LA", "EL", "LOS", "Y", "&", "E", "S", "A", "CIA", "Y/O"]);
function iniciales(n) {
  const p = String(n || "?").toUpperCase().split(" ").filter(x => x && !RELLENO.has(x));
  return ((p[0] || "?")[0] + (p[1] ? p[1][0] : (p[0] || "")[1] || "")).replace(/[^A-Z0-9ÁÉÍÓÚÑ]/g, "") || "?";
}
const DEGRADADOS = [["#2563eb", "#7c3aed"], ["#0891b2", "#2563eb"], ["#059669", "#0891b2"], ["#d97706", "#dc2626"],
  ["#db2777", "#7c3aed"], ["#7c3aed", "#2563eb"], ["#0d9488", "#65a30d"], ["#ea580c", "#db2777"]];
function avatar(n, clase = "avatar") {
  let h = 0; for (const c of String(n)) h = (h * 31 + c.charCodeAt(0)) | 0;
  const [a, b] = DEGRADADOS[Math.abs(h) % DEGRADADOS.length];
  return `<div class="${clase}" style="background:linear-gradient(135deg,${a},${b})">${esc(iniciales(n))}</div>`;
}

const ESTADO = {
  0: ["Pendiente", "t-gris"], 1: ["Enviada", "t-azul"], 2: ["Confirmada", "t-verde"], 3: ["Antigua", "t-gris"], 4: ["Por enviar", "t-ambar"],
};
function chipEstado(o) {
  const [t, c] = ESTADO[o.e] || ["—", "t-gris"];
  return `<span class="chip ${c}">${t}</span>`;
}
function chips(o) {
  let h = chipEstado(o);
  if (o.venc) h += ` <span class="chip t-rojo">Entrega vencida</span>`;
  else if (o.r) h += ` <span class="chip t-verde">Recibida</span>`;
  return h;
}

// color de la franja de cada fila: lo que pide atención salta a la vista
function tono(o) {
  if (o.venc) return "rojo";
  if (o.e === 4 || o.sc) return "ambar";
  if (o.r || o.e === 2) return "verde";
  if (o.e === 1) return "azul";
  return "gris";
}
// lo siguiente que hay que saber de la orden, en palabras
function proximo(o) {
  const hoy = hoyIso();
  if (o.e === 4) return o.venc ? ["Sin enviar y con entrega vencida", "c-rojo"] : ["Falta darle Enviar", "c-ambar"];
  if (o.venc) { const n = diasEntre(o.en, hoy); return [`Entrega vencida hace ${plural(n, "día", "días")}`, "c-rojo"]; }
  if (o.sc) return [`Sin confirmar · ${plural(o.d, "día", "días")}`, "c-ambar"];
  if (!o.r && o.e !== 3 && o.en) {
    const n = diasEntre(hoy, o.en);
    return [n === 0 ? "Llega hoy" : n === 1 ? "Llega mañana" : n <= 30 ? `Llega en ${n} días` : "Entrega " + fecha(o.en), n <= 1 ? "c-azul" : ""];
  }
  if (o.r) return ["Recibida", "c-verde"];
  return [fecha(o.f), ""];
}

function aviso(t) {
  const a = $("#aviso"); a.textContent = t; a.hidden = false;
  clearTimeout(aviso.t); aviso.t = setTimeout(() => (a.hidden = true), 3200);
}

function contar(el) {
  const fin = Number(el.dataset.n); if (!fin) { el.textContent = "0"; return; }
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) { el.textContent = fmtNum.format(fin); return; }
  const t0 = performance.now();
  const paso = t => { const p = Math.min(1, (t - t0) / 900), k = 1 - Math.pow(1 - p, 4); el.textContent = fmtNum.format(Math.round(fin * k)); if (p < 1) requestAnimationFrame(paso); };
  requestAnimationFrame(paso);
}

// ---------------------------------------------------------------- guardado local (IndexedDB)
const almacen = {
  abrir() { return this.db ||= new Promise((ok, mal) => { const r = indexedDB.open("ordenes-compra", 1); r.onupgradeneeded = () => r.result.createObjectStore("kv"); r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error); }); },
  async leer(k) { try { const db = await this.abrir(); return await new Promise(ok => { const r = db.transaction("kv").objectStore("kv").get(k); r.onsuccess = () => ok(r.result); r.onerror = () => ok(null); }); } catch { return null; } },
  async poner(k, v) { try { const db = await this.abrir(); await new Promise(ok => { const t = db.transaction("kv", "readwrite"); t.objectStore("kv").put(v, k); t.oncomplete = ok; t.onerror = ok; }); } catch { } },
  async borrar() { try { const db = await this.abrir(); await new Promise(ok => { const t = db.transaction("kv", "readwrite"); t.objectStore("kv").clear(); t.oncomplete = ok; }); } catch { } },
};
const pref = (k, v) => { try { if (v === undefined) return localStorage.getItem("oc." + k); localStorage.setItem("oc." + k, v); } catch { return null; } };
const ruta = () => pref("ruta") || C.ruta || "COMPRAS/ORDENES DE COMPRA";

// ---------------------------------------------------------------- cuenta y Graph
async function token() {
  try { return (await msalApp.acquireTokenSilent({ scopes: ALCANCES, account: cuenta })).accessToken; }
  catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) { await msalApp.acquireTokenRedirect({ scopes: ALCANCES, account: cuenta }); }
    throw e;
  }
}
async function graph(camino) {
  const r = await fetch(GRAPH + camino, { headers: { Authorization: "Bearer " + await token() } });
  if (!r.ok) { const e = new Error("Graph " + r.status); e.estado = r.status; throw e; }
  return r.json();
}
const enDrive = rel => "/me/drive/root:/" + rel.split("/").filter(Boolean).map(encodeURIComponent).join("/") + ":";

function preparar(d) {
  const hoy = hoyIso();
  for (const o of d.ordenes) {
    o.venc = !o.r && o.e !== 3 && !!o.en && o.en < hoy;
    o.sc = o.e === 1 && (o.d ?? 0) >= (d.diasAlerta || 8);
    o.busca = normal([o.n, o.p, o.c, o.ct, o.ca, o.rq, o.at, ...(o.it || []).map(a => a[1] + " " + a[2] + " " + a[0])].join(" "));
  }
  d.porNumero = new Map();
  for (const o of d.ordenes) if (!d.porNumero.has(o.n)) d.porNumero.set(o.n, o);
  return d;
}

function poner(datos, etag) {
  D = preparar(datos);
  coloresCliente = null;
  setTimeout(() => asegurarLugares().catch(() => { }), 1200); // el mapa y los municipios, sin estorbar el arranque
  $("#actualizado").textContent = (DEMO ? "Datos de ejemplo · " : "") + "Actualizado " + hace(D.generado) + (D.equipo && !DEMO ? " · " + D.equipo : "");
  const n = D.ordenes.filter(o => o.sc).length;
  $("#insignia").hidden = !n; $("#insignia").textContent = n > 99 ? "99+" : n;
  if (etag !== undefined) almacen.poner("datos", { etag, datos, cuando: Date.now() });
}

async function sincronizar(manual) {
  const b = $("#btnRecargar"); b.classList.add("girando");
  try {
    if (DEMO) { poner(await (await fetch("demo.json")).json()); dibujar(); return; }
    if (LOCAL) {
      const datos = await (await fetch("datos.json", { cache: "no-store" })).json();
      if (D && D.generado === datos.generado) { if (manual) aviso("Ya tienes lo último"); return; }
      poner(datos); dibujar(); if (manual) aviso("Actualizado"); return;
    }
    if (!cuenta) return;
    const meta = await graph(enDrive(ruta() + "/_AUTOMATIZACION/movil/datos.json"));
    const guardado = await almacen.leer("datos");
    if (guardado && guardado.etag === meta.eTag && D) { if (manual) aviso("Ya tienes lo último"); return; }
    const r = await fetch(meta["@microsoft.graph.downloadUrl"]);
    poner(await r.json(), meta.eTag);
    revisarNovedades();
    dibujar();
    if (manual) aviso("Actualizado");
  } catch (e) {
    if (e.estado === 404) mostrarError("No encontré los datos en tu OneDrive", `Busqué en «${esc(ruta())}/_AUTOMATIZACION/movil». Abre la aplicación del PC una vez (genera el archivo) o corrige la carpeta en Ajustes.`);
    else if (e.estado === 401 || e.estado === 403) mostrarError("Sin permiso para leer tu OneDrive", "Cierra sesión y vuelve a entrar. Si te pide aprobación de un administrador, la empresa debe autorizar la aplicación.");
    else if (D) aviso(navigator.onLine ? "No se pudo actualizar; muestro lo último guardado" : "Sin conexión: muestro lo último guardado");
    else mostrarError("No se pudieron traer las órdenes", esc(e.message || e));
  } finally { b.classList.remove("girando"); }
}

function mostrarError(titulo, texto) {
  if (D) { aviso(titulo); return; }
  vista.innerHTML = `<div class="vacio entra"><div class="ilus"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg></div><b>${esc(titulo)}</b><p>${texto}</p>
    <p><a class="boton" href="#/ajustes">Ajustes</a></p></div>`;
}

// ---------------------------------------------------------------- vistas
const ICO = {
  borrador: '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
  reloj: '<svg viewBox="0 0 24 24"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5"/></svg>',
  alerta: '<svg viewBox="0 0 24 24"><path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17h.01"/></svg>',
  mes: '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
  ok: '<svg viewBox="0 0 24 24"><path d="m5 12 5 5 9-10"/></svg>',
  enviar: '<svg viewBox="0 0 24 24"><path d="M4 12 20 4l-6 16-3-7z"/></svg>',
  caja: '<svg viewBox="0 0 24 24"><path d="M3 7l9-4 9 4v10l-9 4-9-4z"/><path d="M3 7l9 4 9-4M12 11v10"/></svg>',
  pdf: '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></svg>',
  correo: '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>',
  buscar: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  pago: '<svg viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20M6 15h4"/></svg>',
  compartir: '<svg viewBox="0 0 24 24"><path d="M12 15V3M7 8l5-5 5 5M5 13v6a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-6"/></svg>',
  flecha: '<svg class="flecha" viewBox="0 0 24 24"><path d="m9 6 6 6-6 6"/></svg>',
  mapaFit: '<svg viewBox="0 0 24 24"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>',
  play: '<svg viewBox="0 0 24 24"><path d="M8 5.5v13l11-6.5z" fill="currentColor"/></svg>',
  pausa: '<svg viewBox="0 0 24 24"><path d="M8 5v14M16 5v14" stroke-width="3"/></svg>',
  pin: '<svg viewBox="0 0 24 24"><path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/></svg>',
};

function fila(o) {
  const [txt, c] = proximo(o);
  return `<a class="fila tarjeta tono-${tono(o)}" href="#/orden/${esc(o.n)}">${avatar(o.p)}
    <div class="txt"><div class="l1"><span>Orden ${esc(o.n)}</span><span class="monto">${esc(corto(o.cop ?? o.v))}</span></div>
    <div class="l2">${esc(o.p)}</div><div class="l3">${chipEstado(o)}<span class="${c}">${esc(txt)}</span></div></div></a>`;
}

// lo que pide atención hoy, en orden de importancia
function atencion() {
  const hoy = hoyIso(), manana = masDias(hoy, 1), semana = masDias(hoy, -7);
  const pend = D.ordenes.filter(o => !o.r && o.e !== 3 && o.en);
  const items = [];
  const borr = D.ordenes.filter(o => o.e === 4).length;
  if (borr) items.push(["ambar", ICO.borrador, plural(borr, "borrador por enviar", "borradores por enviar"), "Están en Outlook esperando que les des Enviar", "#/ordenes/borrador"]);
  const nuevasV = D.ordenes.filter(o => o.venc && o.en >= semana).length;
  if (nuevasV) items.push(["rojo", ICO.alerta, plural(nuevasV, "entrega se venció esta semana", "entregas se vencieron esta semana"), "Todavía sin entrada de mercancía", "#/ordenes/vencida"]);
  const hoyN = pend.filter(o => o.en === hoy).length, manN = pend.filter(o => o.en === manana).length;
  if (hoyN || manN) items.push(["azul", ICO.caja, [hoyN && `${hoyN} llega${hoyN > 1 ? "n" : ""} hoy`, manN && `${manN} llega${manN > 1 ? "n" : ""} mañana`].filter(Boolean).join(" · "), "Entregas pactadas con el proveedor", "#/entregas"]);
  const viejas = D.ordenes.filter(o => o.sc && o.d >= 15).length;
  if (viejas) items.push(["ambar", ICO.reloj, plural(viejas, "orden lleva", "órdenes llevan") + " +15 días sin confirmar", "Conviene escribirle al proveedor", "#/seguimiento"]);
  if (!items.length) return `<div class="tarjeta atencion"><div class="aviso-fila"><div class="ico t-verde">${ICO.ok}</div><div><b>Todo al día</b><small>No hay nada urgente hoy</small></div></div></div>`;
  return `<div class="tarjeta atencion">${items.map(([t, i, b, s, h]) => `<a class="aviso-fila" href="${h}"><div class="ico t-${t}">${i}</div><div><b>${esc(b)}</b><small>${esc(s)}</small></div>${ICO.flecha}</a>`).join("")}</div>`;
}

// tendencia de 6 meses dentro de la tarjeta principal (una sola serie: línea blanca sobre el degradado)
function lineaPortada(sumas, meses) {
  const max = Math.max(1, ...sumas), n = sumas.length;
  const pts = sumas.map((s, i) => [i / (n - 1) * 100, 37 - s / max * 31]);
  const linea = pts.map((p, i) => (i ? "L" : "M") + p[0].toFixed(2) + " " + p[1].toFixed(2)).join(" ");
  const [ux, uy] = pts[n - 1];
  return `<div class="p-linea"><div class="p-graf"><svg viewBox="0 0 100 40" preserveAspectRatio="none" aria-hidden="true">
      <defs><linearGradient id="relleno-portada" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff" stop-opacity=".32"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient></defs>
      <path d="${linea} L100 40 L0 40 Z" fill="url(#relleno-portada)" stroke="none"/>
      <path class="trazo" d="${linea}" fill="none" stroke="#fff" stroke-width="2.2" vector-effect="non-scaling-stroke"/></svg>
      <i class="p-punto" style="left:${ux}%;top:${uy / 40 * 100}%"></i></div>
    <div class="p-meses">${meses.map(m => `<span>${esc(fmtMes.format(aFecha(m + "-01")).replace(".", ""))}</span>`).join("")}</div></div>`;
}

function inicio() {
  const hoy = hoyIso(), mes = hoy.slice(0, 7);
  const borr = D.ordenes.filter(o => o.e === 4), sc = D.ordenes.filter(o => o.sc), venc = D.ordenes.filter(o => o.venc);
  const delMes = D.ordenes.filter(o => (o.f || "").startsWith(mes));
  const hora = new Date().getHours();
  const saludo = hora < 12 ? "Buenos días" : hora < 19 ? "Buenas tardes" : "Buenas noches";
  const nombre = (cuenta?.name || "").split(" ")[0];
  // valor por mes, los últimos 6
  const meses = [5, 4, 3, 2, 1, 0].map(i => mesInicio(i).slice(0, 7));
  const sumas = meses.map(m => D.ordenes.filter(o => (o.f || "").startsWith(m)).reduce((s, o) => s + (o.cop || 0), 0));
  const resumen = borr.length ? `Tienes ${borr.length} borrador${borr.length > 1 ? "es" : ""} por enviar en Outlook.` :
    sc.length ? `${sc.length} órdenes esperan confirmación del proveedor.` : "Todo al día.";
  // el mes hasta hoy contra el mismo corte del mes pasado
  const finAnt = aFecha(hoy); finAnt.setDate(0);
  const corteAnt = isoDe(new Date(finAnt.getFullYear(), finAnt.getMonth(), Math.min(aFecha(hoy).getDate(), finAnt.getDate()), 12));
  const totMes = delMes.reduce((s, o) => s + (o.cop || 0), 0);
  const totAnt = D.ordenes.filter(o => o.f && o.f >= mesInicio(1) && o.f <= corteAnt).reduce((s, o) => s + (o.cop || 0), 0);
  const delta = totAnt ? Math.round((totMes - totAnt) / totAnt * 100) : null;
  const provMes = new Set(delMes.map(o => o.p)).size;
  vista.innerHTML = `<div class="entra">
    <h1>${saludo}${nombre ? ", <span class='deg'>" + esc(nombre) + "</span>" : ""}</h1><p class="sub">${esc(resumen)}</p>
    <a class="portada" href="#/tablero">
      <div class="p-tope"><span>Comprado en ${esc(fmtMesLargo.format(aFecha(hoy)).split(" ")[0])}</span><span class="p-ir">Tablero ${ICO.flecha}</span></div>
      <div class="p-valor">${esc(corto(totMes))}</div>
      <div class="p-delta">${delta == null ? "Sin compras el mes pasado a esta altura" : `<b>${delta > 0 ? "▲" : delta < 0 ? "▼" : "="} ${Math.abs(delta)}%</b> vs. el mismo corte de ${esc(fmtMesLargo.format(finAnt).split(" ")[0])}`}</div>
      ${lineaPortada(sumas, meses)}
      <div class="p-pies"><div><b>${fmtNum.format(delMes.length)}</b><small>órdenes</small></div><div><b>${fmtNum.format(provMes)}</b><small>proveedores</small></div><div><b>${esc(corto(delMes.length ? totMes / delMes.length : 0))}</b><small>promedio</small></div></div>
    </a>
    ${tarjetaNovedades()}
    <h2>Atención</h2>${atencion()}
    <h2>Resumen</h2>
    <div class="cifras">
      <a class="cifra tarjeta" href="#/ordenes/borrador"><div class="ico t-ambar">${ICO.borrador}</div><b data-n="${borr.length}">0</b><span>Por enviar</span><small>borradores en Outlook</small></a>
      <a class="cifra tarjeta" href="#/seguimiento"><div class="ico t-azul">${ICO.reloj}</div><b data-n="${sc.length}">0</b><span>Sin confirmar</span><small>enviadas hace ${D.diasAlerta || 8}+ días</small></a>
      <a class="cifra tarjeta" href="#/entregas"><div class="ico t-rojo">${ICO.alerta}</div><b data-n="${venc.length}">0</b><span>Entregas vencidas</span><small>sin llegada · ${esc(corto(venc.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
      <a class="cifra tarjeta" href="#/ordenes/mes"><div class="ico t-verde">${ICO.mes}</div><b data-n="${delMes.length}">0</b><span>Órdenes del mes</span><small>${esc(corto(delMes.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
    </div>
    <h2>¿Dónde se entrega? <small class="h2-nota">últimos 12 meses</small></h2>${tarjetaMapa()}
    ${instalable()}
    <h2>Recientes</h2><div class="lista">${D.ordenes.slice(0, 8).map(fila).join("")}</div></div>`;
  vista.querySelectorAll("[data-n]").forEach(contar);
  { const R = rango("12m"); pintarMapa(D.ordenes.filter(o => o.f && o.f >= R.desde && o.f <= R.hasta), R); }
  if ($("#novCerrar")) $("#novCerrar").onclick = () => { NOVEDADES = null; $(".novedades")?.remove(); };
}

function instalable() {
  if (LOCAL) return "";
  if (matchMedia("(display-mode: standalone)").matches || navigator.standalone || pref("sinInstalar")) return "";
  const ios = /iphone|ipad/i.test(navigator.userAgent);
  return `<div class="tarjeta" style="margin-top:10px;padding:14px 16px;display:flex;gap:12px;align-items:center">
    <img src="iconos/icono-192.png" width="42" height="42" style="border-radius:12px" alt="">
    <div style="flex:1"><b>Tenla como aplicación</b><br><small class="sub">${ios ? "En Safari: Compartir → «Agregar a inicio»." : "Instálala y ábrela desde tu pantalla de inicio."}</small></div>
    ${!ios && instalar ? '<button class="boton primario" id="btnInstalar">Instalar</button>' : ""}</div>`;
}

let filtro = "todas", texto = "";
const FILTROS = [["todas", "Todas"], ["borrador", "Por enviar"], ["enviada", "Enviadas"], ["sinconf", "Sin confirmar"],
  ["confirmada", "Confirmadas"], ["vencida", "Vencidas"], ["recibida", "Recibidas"], ["mes", "Este mes"]];
const PRUEBA = {
  todas: () => true, borrador: o => o.e === 4, enviada: o => o.e === 1, sinconf: o => o.sc, confirmada: o => o.e === 2,
  vencida: o => o.venc, recibida: o => o.r, mes: o => (o.f || "").startsWith(hoyIso().slice(0, 7)),
};

// faceta: viene del tablero o del detalle (todas las órdenes de un proveedor, cliente o tipo)
let faceta = null, ordenLista = "recientes";
const FACETAS = { p: "Proveedor", c: "Cliente", tc: "Tipo de compra", mu: "Municipio", f: "Fecha" };
const ORDENES_POR = [["recientes", "Más recientes"], ["valor", "Mayor valor"], ["entrega", "Entrega más próxima"]];

function ordenes(arg) {
  if (arg) {
    const m = /^(p|c|tc|mu|f|q):(.*)$/.exec(arg);
    if (m) {
      const v = decodeURIComponent(m[2]);
      if (m[1] === "q") { texto = v; faceta = null; } else { faceta = { k: m[1], v }; texto = ""; }
      filtro = "todas";
    } else if (PRUEBA[arg]) { filtro = arg; faceta = null; }
  }
  if (faceta?.k === "mu" && !D.ubicado) { // primero hay que ubicar las órdenes en el mapa
    vista.innerHTML = `<div class="entra"><h1>Órdenes</h1><div class="esq" style="height:300px;margin-top:16px"></div></div>`;
    asegurarLugares().then(() => { if (actual === "ordenes") ordenes(); });
    return;
  }
  const valorFaceta = faceta && (faceta.k === "mu" ? nombreMuni(faceta.v) + " · " + deptoMuni(faceta.v) : faceta.k === "f" ? fecha(faceta.v) : faceta.v);
  vista.innerHTML = `<div class="entra"><h1>Órdenes</h1>
    <div class="buscar"><div class="caja tarjeta">${ICO.buscar}<input id="q" type="search" placeholder="Número, proveedor, municipio o artículo" value="${esc(texto)}" autocomplete="off" enterkeyhint="search"></div>
    <div class="filtros">${FILTROS.map(([k, t]) => `<button data-f="${k}" class="${k === filtro ? "si" : ""}">${t}</button>`).join("")}</div>
    ${faceta ? `<div class="faceta">${faceta.k === "mu" ? ICO.pin : ICO.buscar}<span>${FACETAS[faceta.k]}: <b>${esc(valorFaceta)}</b></span><button id="quitarFaceta" aria-label="Quitar este filtro">✕</button></div>` : ""}</div>
    <div class="contador"><span id="cuenta"></span><select id="ordenPor" aria-label="Ordenar">${ORDENES_POR.map(([k, t]) => `<option value="${k}"${k === ordenLista ? " selected" : ""}>${t}</option>`).join("")}</select></div>
    <div class="lista" id="res"></div><div id="mas"></div></div>`;
  $("#ordenPor").onchange = e => { ordenLista = e.target.value; listar(); };
  if ($("#quitarFaceta")) $("#quitarFaceta").onclick = () => { faceta = null; history.replaceState(null, "", "#/ordenes"); ordenes(); };
  const q = $("#q");
  let t;
  q.addEventListener("input", () => { clearTimeout(t); t = setTimeout(() => { texto = q.value; listar(); }, 160); });
  vista.querySelectorAll("[data-f]").forEach(b => b.onclick = () => {
    filtro = b.dataset.f; vista.querySelectorAll("[data-f]").forEach(x => x.classList.toggle("si", x === b)); listar();
    b.scrollIntoView({ inline: "center", behavior: "smooth", block: "nearest" });
  });
  listar();
}

let observador;
function listar() {
  const palabras = normal(texto).split(" ").filter(Boolean);
  const r = D.ordenes.filter(o => PRUEBA[filtro](o) && (!faceta || o[faceta.k] === faceta.v) && palabras.every(p => o.busca.includes(p)));
  const pendiente = o => !!o.en && !o.r && o.e !== 3;
  if (ordenLista === "valor") r.sort((a, b) => (b.cop || 0) - (a.cop || 0));
  else if (ordenLista === "entrega") r.sort((a, b) => (pendiente(b) - pendiente(a)) || (a.en || "9").localeCompare(b.en || "9"));
  else r.sort((a, b) => (b.f || "").localeCompare(a.f || "") || String(b.n).localeCompare(String(a.n)));
  const total = r.reduce((s, o) => s + (o.cop || 0), 0);
  $("#cuenta").textContent = `${plural(r.length, "orden", "órdenes")} · ${corto(total)}`;
  const res = $("#res"); res.innerHTML = "";
  if (!r.length) { res.innerHTML = `<div class="vacio"><div class="ilus">${ICO.buscar}</div><b>Ninguna orden coincide</b>Prueba con otro número, parte del proveedor o un artículo.</div>`; return; }
  // con «Más recientes» la lista va por meses, cada uno con su total
  const porMes = new Map();
  if (ordenLista === "recientes") for (const o of r) { const k = (o.f || "").slice(0, 7); const g = porMes.get(k) || { n: 0, v: 0 }; g.n++; g.v += o.cop || 0; porMes.set(k, g); }
  let ultimo = null;
  const pintar = o => {
    let h = "";
    if (ordenLista === "recientes") {
      const k = (o.f || "").slice(0, 7);
      if (k !== ultimo) { ultimo = k; const g = porMes.get(k); h = `<div class="mes-cab"><b>${esc(k ? fmtMesLargo.format(aFecha(k + "-01")) : "Sin fecha")}</b><span>${plural(g.n, "orden", "órdenes")} · ${esc(corto(g.v))}</span></div>`; }
    }
    return h + fila(o);
  };
  let n = 0;
  const mas = () => { res.insertAdjacentHTML("beforeend", r.slice(n, n + 40).map(pintar).join("")); n += 40; };
  mas();
  observador?.disconnect();
  observador = new IntersectionObserver(e => { if (e[0].isIntersecting && n < r.length) mas(); }, { rootMargin: "600px" });
  observador.observe($("#mas"));
}

function seguimiento() {
  const sc = D.ordenes.filter(o => o.sc);
  const grupos = new Map();
  for (const o of sc) { if (!grupos.has(o.p)) grupos.set(o.p, []); grupos.get(o.p).push(o); }
  const lista = [...grupos].map(([p, os]) => ({ p, os: os.sort((a, b) => (b.d || 0) - (a.d || 0)), max: Math.max(...os.map(o => o.d || 0)) }))
    .sort((a, b) => b.max - a.max);
  const tope = Math.max(1, ...lista.map(g => g.max));
  vista.innerHTML = `<div class="entra"><h1>Seguimiento</h1>
    <p class="sub">${sc.length} órdenes de ${lista.length} proveedores, enviadas hace ${D.diasAlerta || 8} días o más, sin confirmar.</p>
    ${lista.length ? "" : `<div class="vacio"><div class="ilus">${ICO.ok}</div><b>Todo confirmado</b>No hay órdenes esperando respuesta.</div>`}
    ${lista.map(g => {
      const color = g.max <= 15 ? "var(--verde)" : g.max <= 30 ? "var(--ambar)" : "var(--rojo)";
      const correo = g.os.find(o => o.em)?.em;
      const asunto = encodeURIComponent("Seguimiento órdenes de compra " + g.os.map(o => o.n).join(", "));
      return `<div class="grupo tarjeta"><header><div style="display:flex;gap:10px;align-items:center;min-width:0">${avatar(g.p)}<div style="min-width:0"><b>${esc(g.p)}</b><br><small>${g.os.length} ${g.os.length > 1 ? "órdenes" : "orden"}</small></div></div>
        ${correo ? `<a class="redondo" href="mailto:${esc(correo)}?subject=${asunto}" aria-label="Escribir al proveedor">${ICO.correo}</a>` : ""}</header>
        <div class="medidor"><i style="width:${g.max / tope * 100}%;background:${color}"></i></div>
        ${g.os.map(o => `<a class="fila" href="#/orden/${esc(o.n)}"><div class="txt"><div class="l1"><span>Orden ${esc(o.n)}</span><span class="dias" style="color:${color}">${o.d} días</span></div>
          <div class="l3"><span>enviada ${esc(fecha(o.fe))}</span><span class="monto">${esc(corto(o.cop ?? o.v))}</span>${o.ce ? '<span class="chip t-azul">Contestó</span>' : ""}</div></div></a>`).join("")}</div>`;
    }).join("")}</div>`;
}

function entregas() {
  const hoy = hoyIso();
  const venc = D.ordenes.filter(o => o.venc).sort((a, b) => a.en.localeCompare(b.en));
  const limite = new Date(Date.now() + 60 * 864e5).toISOString().slice(0, 10);
  const prox = D.ordenes.filter(o => !o.r && o.e !== 3 && o.en && o.en >= hoy && o.en <= limite).sort((a, b) => a.en.localeCompare(b.en));
  const porDia = new Map();
  for (const o of prox) { if (!porDia.has(o.en)) porDia.set(o.en, []); porDia.get(o.en).push(o); }
  const dia = iso => { const d = aFecha(iso); return `<div class="dia ${iso === hoy ? "hoy" : ""}"><small>${esc(fmtMes.format(d).replace(".", ""))}</small><b>${d.getDate()}</b></div>`; };
  vista.innerHTML = `<div class="entra"><h1>Entregas</h1><p class="sub">${prox.length} por llegar en los próximos 60 días · ${venc.length} vencidas</p>
    ${venc.length ? `<h2 style="color:var(--rojo)">Vencidas · sin entrada de mercancía</h2><div class="lista">${venc.slice(-10).reverse().map(fila).join("")}</div>
      ${venc.length > 10 ? `<p><a class="boton" href="#/ordenes/vencida">Ver las ${venc.length} vencidas</a></p>` : ""}` : ""}
    <h2>Próximas</h2>
    ${porDia.size ? [...porDia].map(([iso, os]) => `<div class="fecha-grupo">${dia(iso)}<div><b>${esc(fmtDia.format(aFecha(iso)))}</b><br><small class="sub">${os.length} entrega${os.length > 1 ? "s" : ""} · ${esc(corto(os.reduce((s, o) => s + (o.cop || 0), 0)))}</small></div></div>
      <div class="lista">${os.map(fila).join("")}</div>`).join("") : `<div class="vacio"><div class="ilus">${ICO.mes}</div><b>Nada pactado</b>No hay entregas en los próximos 60 días.</div>`}</div>`;
}

// ---------------------------------------------------------------- tablero
let periodo = pref("periodo") || "12m";
const PERIODOS = [["mes", "Este mes"], ["3m", "3 meses"], ["12m", "12 meses"], ["ano", "Este año"], ["todo", "Todo"]];
const anoAntes = iso => (Number(iso.slice(0, 4)) - 1) + iso.slice(4);

// el periodo elegido, cómo se agrupa el gráfico y contra qué se compara (el mismo periodo del año pasado)
function rango(p) {
  const hoy = hoyIso();
  if (p === "todo") {
    const desde = D.ordenes.reduce((m, o) => o.f && o.f < m ? o.f : m, hoy);
    return { desde, hasta: hoy, cubeta: "ano" };
  }
  const desde = p === "mes" ? mesInicio(0) : p === "3m" ? mesInicio(2) : p === "12m" ? mesInicio(11) : hoy.slice(0, 4) + "-01-01";
  return { desde, hasta: hoy, antes: [anoAntes(desde), anoAntes(hoy)], cubeta: p === "mes" ? "dia" : p === "3m" ? "semana" : "mes" };
}
const lunes = d => { d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return d; };
function cubetas(desde, hasta, tipo) {
  const out = [], fin = aFecha(hasta);
  let d = aFecha(desde);
  if (tipo === "semana") lunes(d);
  if (tipo === "mes" || tipo === "ano") d.setDate(1);
  if (tipo === "ano") d.setMonth(0);
  while (d <= fin) {
    out.push(isoDe(d));
    if (tipo === "dia") d.setDate(d.getDate() + 1);
    else if (tipo === "semana") d.setDate(d.getDate() + 7);
    else if (tipo === "mes") d.setMonth(d.getMonth() + 1);
    else d.setFullYear(d.getFullYear() + 1);
  }
  return out;
}
function claveDe(f, tipo) {
  if (tipo === "dia") return f;
  if (tipo === "mes") return f.slice(0, 7) + "-01";
  if (tipo === "ano") return f.slice(0, 4) + "-01-01";
  return isoDe(lunes(aFecha(f)));
}
function etiqueta(k, tipo, larga) {
  const d = aFecha(k), mes = fmtMes.format(d).replace(".", "");
  if (tipo === "dia") return larga ? fecha(k) : String(d.getDate());
  if (tipo === "semana") return larga ? "Semana del " + fecha(k) : d.getDate() + " " + mes;
  if (tipo === "mes") return larga ? fmtMesLargo.format(d) : mes;
  return k.slice(0, 4);
}
function techo(max) { // un tope redondo para la escala
  const p = Math.pow(10, Math.floor(Math.log10(max)));
  return [1, 2, 2.5, 5, 10].map(m => m * p).find(v => v >= max);
}

function agrupar(lista, clave) {
  const m = new Map();
  for (const o of lista) { const k = clave(o) || "(sin dato)"; const g = m.get(k) || { k, v: 0, n: 0 }; g.v += o.cop || 0; g.n++; m.set(k, g); }
  return [...m.values()].sort((a, b) => b.v - a.v);
}
function articulos(lista) {
  const m = new Map();
  for (const o of lista) {
    const tasa = o.v ? (o.cop || 0) / o.v : 1;
    for (const a of o.it || []) {
      const clave = claveArt(a); if (!clave) continue;
      const g = m.get(clave) || { k: String(a[1] || "").trim() + (a[2] ? " " + a[2] : ""), clave, v: 0, n: 0, q: 0, u: a[3] };
      g.v += (a[6] || 0) * tasa; g.n++; g.q += a[4] || 0; m.set(clave, g);
    }
  }
  return [...m.values()].sort((a, b) => b.v - a.v);
}

// el color de cada cliente es fijo (por lo comprado en total), no cambia con el periodo
function casillaCliente(c) {
  coloresCliente ||= new Map(agrupar(D.ordenes, o => o.c).slice(0, 5).map((g, i) => [g.k, i + 1]));
  return coloresCliente.get(c) || 0;
}

function kpi(titulo, valor, antes, fmt) {
  const p = antes ? Math.round((valor - antes) / antes * 100) : null;
  return `<div class="kpi tarjeta"><span>${titulo}</span><b>${esc(fmt(valor))}</b>
    <small>${p == null ? "&nbsp;" : `${p > 0 ? "▲" : p < 0 ? "▼" : "="} ${Math.abs(p)}% vs. año pasado`}</small></div>`;
}

function grafico(cub, tipo, sumas, cuentas) {
  const tope = techo(Math.max(1, ...sumas)), salto = Math.ceil(cub.length / 7);
  const total = sumas.reduce((s, v) => s + v, 0), n = cuentas.reduce((s, v) => s + v, 0);
  return `<div class="tarjeta grafico" id="tendencia">
    <div class="lectura"><small id="tdEtq">Todo el periodo</small><b id="tdVal">${esc(corto(total))}</b><small id="tdCnt">${plural(n, "orden", "órdenes")}</small></div>
    <div class="plano">
      <div class="rejilla"><span data-v="${esc(corto(tope))}"></span><span data-v="${esc(corto(tope / 2))}"></span><span data-v="0"></span></div>
      <div class="columnas">${cub.map((k, i) => `<button data-i="${i}" aria-label="${esc(etiqueta(k, tipo, true))}: ${esc(dinero(sumas[i]))}, ${plural(cuentas[i], "orden", "órdenes")}">
        <span class="col"><i style="height:${sumas[i] ? Math.max(1.5, sumas[i] / tope * 100) : 0}%;animation-delay:${Math.min(i * 25, 500)}ms"></i></span>
        <small><span>${(cub.length - 1 - i) % salto === 0 ? esc(etiqueta(k, tipo)) : ""}</span></small></button>`).join("")}</div>
    </div>
    <p class="pista">Toca o desliza sobre las barras para ver cada ${tipo === "dia" ? "día" : tipo === "semana" ? "semana" : tipo === "mes" ? "mes" : "año"}.</p>
    <details class="como-tabla"><summary>Ver como tabla</summary><table><thead><tr><th></th><th>Comprado</th><th>Órdenes</th></tr></thead><tbody>
      ${cub.map((k, i) => `<tr><td>${esc(etiqueta(k, tipo, true))}</td><td>${esc(corto(sumas[i]))}</td><td>${cuentas[i]}</td></tr>`).join("")}</tbody></table></details></div>`;
}

function activarGrafico(cub, tipo, sumas, cuentas) {
  const cont = $("#tendencia .columnas"); if (!cont) return;
  const bs = [...cont.children];
  const elegir = i => {
    cont.classList.add("elegido"); bs.forEach((b, j) => b.classList.toggle("si", j === i));
    $("#tdEtq").textContent = etiqueta(cub[i], tipo, true);
    $("#tdVal").textContent = corto(sumas[i]);
    $("#tdCnt").textContent = plural(cuentas[i], "orden", "órdenes");
  };
  const porX = x => { const r = cont.getBoundingClientRect(); return Math.max(0, Math.min(bs.length - 1, Math.floor((x - r.left) / r.width * bs.length))); };
  cont.addEventListener("pointerdown", e => { elegir(porX(e.clientX)); try { cont.setPointerCapture(e.pointerId); } catch { } });
  cont.addEventListener("pointermove", e => { if (e.buttons || e.pointerType === "mouse") elegir(porX(e.clientX)); });
  bs.forEach((b, i) => b.addEventListener("focus", () => elegir(i)));
}

function ranking(grupos, campo, cuantos, sub) {
  if (!grupos.length) return `<div class="vacio" style="padding:24px">Sin órdenes en este periodo.</div>`;
  const max = grupos[0].v || 1, top = grupos.slice(0, cuantos), resto = grupos.slice(cuantos);
  return top.map(g => `<a class="rank" href="${campo === "art" ? "#/articulo/" + encodeURIComponent(g.clave) : `#/ordenes/${campo}:${encodeURIComponent(g.k)}`}"><div class="r1"><span>${esc(g.k)}</span><b>${esc(corto(g.v))}</b></div>
      <div class="r2"><i style="width:${Math.max(0.8, g.v / max * 100)}%"></i></div><small>${esc(sub(g))}</small></a>`).join("") +
    (resto.length ? `<div class="rank resto"><div class="r1"><span>Otros ${fmtNum.format(resto.length)}</span><b>${esc(corto(resto.reduce((s, g) => s + g.v, 0)))}</b></div></div>` : "");
}

function partes(grupos) {
  const total = grupos.reduce((s, g) => s + g.v, 0) || 1;
  const seg = new Map();
  for (const g of grupos) {
    const s = casillaCliente(g.k), k = s ? g.k : "Otros clientes";
    const x = seg.get(k) || { k, v: 0, n: 0, s }; x.v += g.v; x.n += g.n; seg.set(k, x);
  }
  const lista = [...seg.values()].sort((a, b) => (a.s || 9) - (b.s || 9));
  if (!lista.length) return `<div class="vacio" style="padding:24px">Sin órdenes en este periodo.</div>`;
  return `<div class="pila">${lista.filter(x => x.v > 0).map(x => `<i style="flex:${x.v};background:var(--s${x.s || "o"})"></i>`).join("")}</div>
    <div class="leyenda">${lista.map(x => `<${x.s ? `a href="#/ordenes/c:${encodeURIComponent(x.k)}"` : "div"} class="ley"><i style="background:var(--s${x.s || "o"})"></i><span>${esc(x.k)}<small>${plural(x.n, "orden", "órdenes")}</small></span><b>${esc(corto(x.v))}</b><em>${x.v / total < 0.005 ? "&lt;1" : Math.round(x.v / total * 100)}%</em></${x.s ? "a" : "div"}>`).join("")}</div>`;
}

function tablero(arg, sinAnimar) {
  if (arg && PERIODOS.some(p => p[0] === arg)) periodo = arg;
  const R = rango(periodo);
  const entre = (o, a, b) => o.f && o.f >= a && o.f <= b;
  const sel = D.ordenes.filter(o => entre(o, R.desde, R.hasta));
  const ant = R.antes ? D.ordenes.filter(o => entre(o, R.antes[0], R.antes[1])) : null;
  const suma = l => l.reduce((s, o) => s + (o.cop || 0), 0);
  const provs = l => new Set(l.map(o => o.p)).size;
  const cub = cubetas(R.desde, R.hasta, R.cubeta), pos = new Map(cub.map((k, i) => [k, i]));
  const sumas = cub.map(() => 0), cuentas = cub.map(() => 0);
  for (const o of sel) { const i = pos.get(claveDe(o.f, R.cubeta)); if (i != null) { sumas[i] += o.cop || 0; cuentas[i]++; } }
  const tot = suma(sel), n = sel.length;
  const fmtN = v => fmtNum.format(Math.round(v));
  vista.innerHTML = `<div class="${sinAnimar ? "" : "entra"}"><div class="h1-fila"><h1>Tablero</h1><a class="boton" href="#/presentar">${ICO.play}Presentar</a></div>
    <div class="filtros periodos">${PERIODOS.map(([k, t]) => `<button data-per="${k}" class="${k === periodo ? "si" : ""}">${t}</button>`).join("")}</div>
    <p class="sub rango">${esc(fecha(R.desde))} – ${esc(fecha(R.hasta))}${R.antes ? " · comparado con el mismo periodo del año pasado" : ""}</p>
    <div class="kpis">
      ${kpi("Comprado", tot, ant && suma(ant), corto)}
      ${kpi("Órdenes", n, ant && ant.length, fmtN)}
      ${kpi("Promedio por orden", n ? tot / n : 0, ant && ant.length ? suma(ant) / ant.length : null, corto)}
      ${kpi("Proveedores", provs(sel), ant && provs(ant), fmtN)}
    </div>
    <h2>Compras en el tiempo</h2>${grafico(cub, R.cubeta, sumas, cuentas)}
    <h2>Calendario de compras</h2>${calendario(sel, R)}
    <h2>¿Dónde se entrega?</h2>${tarjetaMapa()}
    <h2>Por cliente</h2><div class="tarjeta bloque">${partes(agrupar(sel, o => o.c))}</div>
    <h2>Proveedores principales</h2><div class="tarjeta bloque">${ranking(agrupar(sel, o => o.p), "p", 8, g => plural(g.n, "orden", "órdenes"))}</div>
    <h2>Tipo de compra</h2><div class="tarjeta bloque">${ranking(agrupar(sel, o => o.tc), "tc", 6, g => plural(g.n, "orden", "órdenes"))}</div>
    <h2>Precios que subieron</h2><div class="tarjeta bloque">${preciosQueSubieron(sel)}</div>
    <h2>Artículos más comprados</h2><div class="tarjeta bloque">${ranking(articulos(sel), "art", 8, g => `${fmtNum.format(g.q)} ${g.u || "und"} · en ${plural(g.n, "orden", "órdenes")}`)}</div>
    <p class="nota">Valores en pesos, con impuestos (los artículos, antes de impuestos). Las órdenes en dólares o euros van convertidas. Toca un cliente, proveedor o artículo para ver sus órdenes.</p></div>`;
  vista.querySelectorAll("[data-per]").forEach(b => b.onclick = () => {
    periodo = b.dataset.per; pref("periodo", periodo); tablero(null, true);
  });
  activarGrafico(cub, R.cubeta, sumas, cuentas);
  activarCalendario(sel);
  pintarMapa(sel, R);
}

// ---------------------------------------------------------------- calendario: cuánto se compró cada día
// Una celda por día (semanas en columnas); el tono sube con lo comprado (5 niveles de un solo azul, por quintiles).
function calendario(sel, R) {
  const fin = aFecha(R.hasta), tope = new Date(fin); tope.setDate(tope.getDate() - 7 * 52 + 1);
  let ini = aFecha(R.desde); const recortado = ini < tope; if (recortado) ini = tope;
  lunes(ini);
  const desde = isoDe(ini), porDia = new Map();
  for (const o of sel) { if (!o.f || o.f < desde) continue; const g = porDia.get(o.f) || { v: 0, n: 0 }; g.v += o.cop || 0; g.n++; porDia.set(o.f, g); }
  const vals = [...porDia.values()].map(g => g.v).sort((a, b) => a - b);
  const q = p => vals[Math.min(vals.length - 1, Math.floor(p * vals.length))] ?? 0, cortes = [q(0.2), q(0.4), q(0.6), q(0.8)];
  const nivel = v => (!v ? 0 : 1 + cortes.filter(c => v > c).length);
  const celdas = [], meses = [], hoy = hoyIso();
  let semana = 0, mesAnt = "";
  for (const d = new Date(ini); d <= fin || d.getDay() !== 1; d.setDate(d.getDate() + 1)) {
    const iso = isoDe(d), dia = (d.getDay() + 6) % 7;
    if (dia === 0 && iso !== desde) semana++;
    if (dia === 0 && iso.slice(0, 7) !== mesAnt) { mesAnt = iso.slice(0, 7); meses.push([semana, fmtMes.format(d).replace(".", "")]); }
    if (d > fin) { celdas.push(`<i class="d fuera"></i>`); continue; }
    const g = porDia.get(iso);
    celdas.push(`<button class="d n${nivel(g?.v)}${iso === hoy ? " hoy" : ""}" data-f="${iso}" aria-label="${esc(fecha(iso))}: ${g ? esc(corto(g.v)) + ", " + plural(g.n, "orden", "órdenes") : "sin compras"}"></button>`);
  }
  const dias = porDia.size, activo = [...porDia].sort((a, b) => b[1].v - a[1].v)[0];
  return `<div class="tarjeta calendario">
    <div class="lectura"><small id="calEtq">${recortado ? "Últimas 52 semanas" : "Todo el periodo"}</small><b id="calVal">${plural(dias, "día con compras", "días con compras")}</b>
      <small id="calSub">${activo ? "el más fuerte: " + esc(fechaCorta(activo[0])) + " · " + esc(corto(activo[1].v)) : ""}</small></div>
    <div class="cal-marco" id="calMarco"><div class="cal-meses" style="--sem:${semana + 1}">${meses.map(([s, m]) => `<span style="grid-column:${s + 1}">${esc(m)}</span>`).join("")}</div>
      <div class="cal-cuerpo"><div class="cal-dias"><span>L</span><span></span><span>M</span><span></span><span>V</span><span></span><span>D</span></div>
      <div class="cal" style="--sem:${semana + 1}">${celdas.join("")}</div></div></div>
    <div class="cal-pie"><a id="calVer" class="enlace" hidden></a><span class="cal-ley">Menos <i class="d n0"></i><i class="d n1"></i><i class="d n2"></i><i class="d n3"></i><i class="d n4"></i><i class="d n5"></i> Más</span></div></div>`;
}
function activarCalendario(sel) {
  const marco = $("#calMarco"); if (!marco) return;
  marco.scrollLeft = marco.scrollWidth; // lo más reciente a la vista
  const porDia = new Map(); for (const o of sel) { if (!o.f) continue; const g = porDia.get(o.f) || { v: 0, n: 0 }; g.v += o.cop || 0; g.n++; porDia.set(o.f, g); }
  marco.addEventListener("click", e => {
    const b = e.target.closest("button.d"); if (!b) return;
    marco.querySelectorAll(".d.si").forEach(x => x.classList.remove("si")); b.classList.add("si");
    const g = porDia.get(b.dataset.f), ver = $("#calVer");
    $("#calEtq").textContent = fmtDia.format(aFecha(b.dataset.f));
    $("#calVal").textContent = g ? corto(g.v) : "Sin compras";
    $("#calSub").textContent = g ? plural(g.n, "orden", "órdenes") : "";
    ver.hidden = !g; ver.href = "#/ordenes/f:" + b.dataset.f; ver.textContent = "Ver las órdenes de ese día";
  });
}

// ---------------------------------------------------------------- precios: historial por artículo
// El código del PDF no identifica el artículo (es una categoría + un consecutivo): se usa la descripción y su
// complemento (p. ej. la medida). Los servicios y los saltos de más de 5 veces (cambio de unidad o de empaque)
// no se comparan.
let PRECIOS = null;
const claveArt = a => normal(String(a[1] || "") + " " + String(a[2] || "")).replace(/[^a-z0-9]+/g, " ").trim();
const esServicio = d => /^\s*servicio/i.test(String(d || ""));
const comparable = (a, b) => { const r = a / b; return r > 0.2 && r < 5; };
function precios() {
  if (PRECIOS?.d === D) return PRECIOS.m;
  const m = new Map();
  for (const o of D.ordenes) {
    const tasa = o.v ? (o.cop || 0) / o.v : 1;
    for (const a of o.it || []) {
      if (!(a[4] > 0) || !(a[5] > 0)) continue;
      const k = claveArt(a); if (!k) continue;
      const l = m.get(k) || []; l.push({ o, a, pu: a[5] * tasa, q: a[4] }); m.set(k, l);
    }
  }
  for (const l of m.values()) l.sort((x, y) => (x.o.f || "").localeCompare(y.o.f || "") || String(x.o.n).localeCompare(String(y.o.n)));
  PRECIOS = { d: D, m }; return m;
}
// la compra anterior del mismo artículo (en otra orden) y cuánto cambió el precio unitario
function cambioPrecio(o, a) {
  if (esServicio(a[1])) return null;
  const l = precios().get(claveArt(a)); if (!l || l.length < 2) return null;
  const i = l.findIndex(x => x.a === a); if (i <= 0) return null;
  const ant = l.slice(0, i).reverse().find(x => x.o !== o); if (!ant) return null;
  if (!comparable(l[i].pu, ant.pu)) return null;
  return { pct: Math.round((l[i].pu / ant.pu - 1) * 100), ant, l };
}
function etiquetaPrecio(o, a) {
  const l = precios().get(claveArt(a)); if (!l || l.length < 2) return "";
  const enlace = "#/articulo/" + encodeURIComponent(claveArt(a)), cp = cambioPrecio(o, a);
  if (cp && Math.abs(cp.pct) >= 1) return `<a class="precio ${cp.pct > 0 ? "sube" : "baja"}" href="${enlace}">${cp.pct > 0 ? "▲" : "▼"} ${Math.abs(cp.pct)}% vs. ${esc(fechaCorta(cp.ant.o.f))}</a>`;
  return `<a class="precio" href="${enlace}">${cp ? "= mismo precio · " : ""}Historial (${l.length})</a>`;
}

// en el tablero: lo que más subió de precio en el periodo, por el dinero de más que costó
function preciosQueSubieron(sel) {
  const enSel = new Set(sel), cambios = [];
  for (const [k, l] of precios()) {
    if (l.length < 2 || esServicio(l[0].a[1])) continue;
    let i = -1; for (let j = l.length - 1; j >= 0; j--) if (enSel.has(l[j].o)) { i = j; break; }
    if (i <= 0) continue;
    const ult = l[i], ant = l.slice(0, i).reverse().find(x => x.o !== ult.o);
    if (!ant || !comparable(ult.pu, ant.pu)) continue;
    const pct = (ult.pu / ant.pu - 1) * 100;
    cambios.push({ k, ult, ant, pct, impacto: (ult.pu - ant.pu) * ult.q });
  }
  const suben = cambios.filter(c => c.pct >= 5).sort((a, b) => b.impacto - a.impacto), bajan = cambios.filter(c => c.pct <= -5);
  const extra = suben.reduce((s, c) => s + c.impacto, 0);
  if (!cambios.length) return `<div class="vacio" style="padding:24px">No hay artículos con compras anteriores para comparar en este periodo.</div>`;
  return `<div class="precios-resumen"><div><b class="c-rojo">▲ ${fmtNum.format(suben.length)}</b><small>subieron 5 % o más</small></div><div><b class="c-verde">▼ ${fmtNum.format(bajan.length)}</b><small>bajaron 5 % o más</small></div><div><b>${esc(corto(extra))}</b><small>de más por las subidas</small></div></div>` +
    suben.slice(0, 6).map(c => `<a class="rank" href="#/articulo/${encodeURIComponent(c.k)}"><div class="r1"><span>${esc(c.ult.a[1])}</span><b class="c-rojo">▲ ${Math.round(c.pct)}%</b></div>
      <small>${esc(dinero(c.ant.pu))} → ${esc(dinero(c.ult.pu))} · ${esc(c.ult.o.p)}</small></a>`).join("");
}

// pantalla de un artículo: su precio en el tiempo, por proveedor y cada compra
function articulo(arg) {
  const k = decodeURIComponent(arg || ""), l = precios().get(k);
  if (!l) { vista.innerHTML = `<div class="vacio entra"><div class="ilus">${ICO.buscar}</div><b>No encontré ese artículo</b><p><a class="boton" href="#/tablero">Ir al tablero</a></p></div>`; return; }
  const a0 = l[l.length - 1].a, pus = l.map(x => x.pu), ult = l[l.length - 1], pri = l[0];
  const servicio = esServicio(a0[1]), varia = !servicio && comparable(ult.pu, pri.pu) ? Math.round((ult.pu / pri.pu - 1) * 100) : null;
  const provs = new Map();
  for (const x of l) { const g = provs.get(x.o.p) || { p: x.o.p, s: 0, q: 0, n: 0, ult: x }; g.s += x.pu * x.q; g.q += x.q; g.n++; g.ult = x; provs.set(x.o.p, g); }
  const lp = [...provs.values()].map(g => ({ ...g, prom: g.s / g.q })).sort((a, b) => a.prom - b.prom);
  vista.innerHTML = `<div class="entra">
    <button class="volver" id="btnVolver">${ICO.flecha}Volver</button>
    <h1 class="h1-art">${esc(a0[1])}</h1>
    <p class="sub">${a0[2] ? esc(a0[2]) + " · " : ""}${plural(l.length, "compra", "compras")} entre ${esc(fecha(pri.o.f))} y ${esc(fecha(ult.o.f))}${servicio ? " · servicio: el precio depende de cada trabajo" : ""}</p>
    <div class="kpis">
      <div class="kpi tarjeta"><span>Último precio</span><b>${esc(dinero(ult.pu))}</b><small>${esc(fechaCorta(ult.o.f))} · ${esc(ult.a[3] || "unidad")}</small></div>
      <div class="kpi tarjeta"><span>Variación</span><b class="${varia > 0 ? "c-rojo" : varia < 0 ? "c-verde" : ""}">${varia == null ? "—" : (varia > 0 ? "▲ " : varia < 0 ? "▼ " : "") + Math.abs(varia) + "%"}</b><small>desde la primera compra</small></div>
      <div class="kpi tarjeta"><span>Más bajo</span><b>${esc(dinero(Math.min(...pus)))}</b><small>unitario</small></div>
      <div class="kpi tarjeta"><span>Más alto</span><b>${esc(dinero(Math.max(...pus)))}</b><small>unitario</small></div>
    </div>
    <h2>Precio unitario en el tiempo</h2>
    <div class="tarjeta grafico" id="precioGraf"><div class="lectura"><small id="pgEtq">Toca un punto</small><b id="pgVal">${esc(dinero(ult.pu))}</b><small id="pgSub">${esc(ult.o.p)}</small></div><div id="pgLienzo"></div></div>
    <h2>Por proveedor</h2>
    <div class="tarjeta bloque">${lp.map((g, i) => `<a class="rank" href="#/ordenes/p:${encodeURIComponent(g.p)}"><div class="r1"><span>${esc(g.p)}</span><b>${esc(dinero(g.prom))}</b></div>
      <small>${i === 0 && lp.length > 1 && g.prom < lp[1].prom * 0.995 ? '<span class="chip t-verde">Mejor precio</span> ' : ""}${plural(g.n, "compra", "compras")} · última ${esc(fechaCorta(g.ult.o.f))} · promedio por unidad</small></a>`).join("")}</div>
    <h2>Compras</h2>
    <div class="tarjeta bloque">${[...l].reverse().map(x => `<a class="rank" href="#/orden/${esc(x.o.n)}"><div class="r1"><span>Orden ${esc(x.o.n)} · ${esc(x.o.p)}</span><b>${esc(dinero(x.pu))}</b></div>
      <small>${esc(fecha(x.o.f))} · ${esc(fmtNum.format(x.q))} ${esc(x.a[3] || "")}${x.o.m && x.o.m !== "COP" ? " · en " + esc(x.o.m) + ", convertido" : ""}</small></a>`).join("")}</div></div>`;
  $("#btnVolver").onclick = () => (history.length > 1 ? history.back() : (location.hash = "#/tablero"));
  lineaPrecio(l);
}

// una sola serie: el precio unitario de cada compra, en su fecha
function lineaPrecio(l) {
  const cont = $("#pgLienzo"); if (!cont) return;
  const W = cont.clientWidth || 320, H = 180, pi = 10, pd = 62, pa = 14, pb = 26;
  const ts = l.map(x => aFecha(x.o.f).getTime()), t0 = Math.min(...ts), t1 = Math.max(...ts);
  // eje ajustado al rango de precios (en una línea de precios importa la variación, no el cero)
  const vmin = Math.min(...l.map(x => x.pu)), vmax = Math.max(...l.map(x => x.pu)), holgura = Math.max((vmax - vmin) * 0.25, vmax * 0.03);
  const lo = Math.max(0, vmin - holgura), hi = vmax + holgura;
  const X = t => pi + (t1 === t0 ? (W - pi - pd) / 2 : (t - t0) / (t1 - t0) * (W - pi - pd));
  const Y = v => pa + (1 - (v - lo) / (hi - lo || 1)) * (H - pa - pb);
  const pts = l.map((x, i) => ({ x: X(ts[i]), y: Y(x.pu), d: x }));
  const linea = pts.map((p, i) => (i ? "L" : "M") + p.x.toFixed(1) + " " + p.y.toFixed(1)).join(" ");
  cont.innerHTML = `<svg class="linea-precio" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Precio unitario en cada compra">
    ${[0, 0.5, 1].map(f => { const v = lo + (hi - lo) * f; return `<line x1="${pi}" x2="${W - pd}" y1="${Y(v)}" y2="${Y(v)}" class="rej"/><text x="${W - pd + 8}" y="${Y(v) + 4}" class="eje">${esc(corto(v))}</text>`; }).join("")}
    <path d="${linea}" class="trazo"/>
    ${pts.map((p, i) => `<circle cx="${p.x}" cy="${p.y}" r="4.5" data-i="${i}"/>`).join("")}
    <text x="${pi}" y="${H - 6}" class="eje">${esc(fechaCorta(l[0].o.f))}</text><text x="${W - pd}" y="${H - 6}" class="eje" text-anchor="end">${esc(fechaCorta(l[l.length - 1].o.f))}</text></svg>`;
  const svg = cont.querySelector("svg"), cs = [...svg.querySelectorAll("circle")];
  const elegir = i => {
    cs.forEach((c, j) => c.classList.toggle("si", j === i));
    const x = pts[i].d; $("#pgEtq").textContent = fecha(x.o.f) + " · orden " + x.o.n; $("#pgVal").textContent = dinero(x.pu); $("#pgSub").textContent = x.o.p;
  };
  const cercano = e => { const r = svg.getBoundingClientRect(), x = (e.clientX - r.left) / r.width * W; let m = 0; pts.forEach((p, i) => { if (Math.abs(p.x - x) < Math.abs(pts[m].x - x)) m = i; }); return m; };
  svg.addEventListener("pointerdown", e => elegir(cercano(e)));
  svg.addEventListener("pointermove", e => { if (e.buttons || e.pointerType === "mouse") elegir(cercano(e)); });
  elegir(pts.length - 1);
}

// ---------------------------------------------------------------- mapa: dónde se entregan las órdenes
// mapa.json: municipios del DANE (DIVIPOLA) y contorno de departamentos (geoBoundaries), ya proyectados.
// lugares.js lee el texto libre del lugar de entrega, corrige errores y estima lo que no dice dónde.
let MAPA = null, UBIC = null, cargandoLugares = null;
const COMO = { exacto: ["Tal cual", "t-verde"], sitio: ["Sitio conocido", "t-azul"], corregido: ["Corregido", "t-ambar"], cliente: ["Estimado", "t-gris"], contrato: ["Estimado", "t-gris"] };

function asegurarLugares() {
  if (D && D.ubicado) return Promise.resolve(true);
  return cargandoLugares ||= (async () => {
    try {
      MAPA ||= await (await fetch("mapa.json")).json();
      UBIC ||= Lugares.crear(MAPA.municipios);
      if (D && !D.ubicado) {
        const r = UBIC.ubicar(D.ordenes);
        D.ordenes.forEach((o, i) => {
          o.mu = r[i]?.cod || null; o.muc = r[i]?.como || null;
          if (o.mu) { const m = UBIC.municipio(o.mu); o.busca += " " + normal(m.nombre + " " + m.dpto); } // buscar por municipio
        });
        D.ubicado = true;
      }
      return true;
    } finally { cargandoLugares = null; }
  })();
}
// «PUERTO GAITÁN» -> «Puerto Gaitán»: más fácil de leer
const MENORES = new Set(["de", "del", "la", "las", "los", "y", "el"]);
const titulo = s => s.toLowerCase().split(" ").map((p, i) => (i && MENORES.has(p) ? p : p.charAt(0).toUpperCase() + p.slice(1))).join(" ");
const nombreMuni = cod => { const m = cod && UBIC?.municipio(cod); return m ? titulo(m.nombre.replace(", D.C.", "")) : "—"; };
const deptoMuni = cod => { const m = cod && UBIC?.municipio(cod); return m ? titulo(m.dpto.replace(", D.C.", "")).replace(/^Bogotá$/, "Bogotá D.C.") : ""; };
// ancho real de un texto (para que los nombres del mapa no se encimen)
const medidor = document.createElement("canvas").getContext("2d");
const anchoTexto = (t, fs) => { medidor.font = "650 100px Inter, system-ui, sans-serif"; return medidor.measureText(t).width / 100 * fs; };

function tarjetaMapa() {
  return `<div class="tarjeta mapa-caja" id="mapaCaja"><div class="esq" style="height:380px"></div></div>
    <div class="tarjeta linea-tiempo" id="mapaTiempo" hidden></div>
    <div id="mapaSel"></div><div class="tarjeta bloque" id="mapaLista" style="margin-top:10px"></div>
    <a class="ver-mas tarjeta enlace-lugares" id="enlaceLugares" href="#/lugares" hidden></a>`;
}

function porMunicipio(lista) {
  const grupos = new Map();
  for (const o of lista) {
    if (!o.mu) continue;
    const g = grupos.get(o.mu) || { cod: o.mu, v: 0, n: 0, est: 0, cl: new Map() };
    g.v += o.cop || 0; g.n++; if (o.muc === "cliente" || o.muc === "contrato") g.est++;
    g.cl.set(o.c, (g.cl.get(o.c) || 0) + (o.cop || 0)); grupos.set(o.mu, g);
  }
  return [...grupos.values()].sort((a, b) => b.v - a.v).map(g => ({ ...g, ...UBIC.municipio(g.cod) }));
}

// R: el rango del tablero (para la línea de tiempo por meses)
async function pintarMapa(sel, R) {
  const caja = $("#mapaCaja"); if (!caja) return;
  try { await asegurarLugares(); } catch { caja.innerHTML = `<div class="vacio" style="padding:24px">No se pudo cargar el mapa. Revisa la conexión.</div>`; return; }
  if ($("#mapaCaja") !== caja) return;
  const todos = porMunicipio(sel);
  const sinUbicar = sel.filter(o => !o.mu).length;
  if (!todos.length) { caja.innerHTML = `<div class="vacio" style="padding:24px">Sin órdenes en este periodo.</div>`; $("#mapaLista").innerHTML = ""; return; }

  // meses del periodo (para animar); la escala de los círculos es la misma en todos los meses
  const meses = R && R.cubeta !== "dia" ? cubetas(R.desde, R.hasta, "mes").map(k => k.slice(0, 7)) : [];
  const porMes = new Map(meses.map(m => [m, porMunicipio(sel.filter(o => (o.f || "").startsWith(m)))]));
  const maxMes = Math.max(1, ...[...porMes.values()].flatMap(l => l.map(g => g.v)));
  let mes = null, elegido = null;

  // encuadre: todos los puntos del periodo con margen y la proporción de la caja
  const ancho = caja.clientWidth || 360, alto = INCRUSTADO ? Math.max(300, Math.min(420, window.innerHeight - 96)) : Math.round(Math.min(480, Math.max(320, ancho * 1.08)));
  let x0 = Math.min(...todos.map(p => p.x)), x1 = Math.max(...todos.map(p => p.x)), y0 = Math.min(...todos.map(p => p.y)), y1 = Math.max(...todos.map(p => p.y));
  const margen = 80; x0 -= margen; x1 += margen; y0 -= margen; y1 += margen;
  let w = Math.max(260, x1 - x0), h = Math.max(260, y1 - y0);
  if (w / h > ancho / alto) h = w * alto / ancho; else w = h * ancho / alto;
  const inicial = { x: (x0 + x1) / 2 - w / 2, y: (y0 + y1) / 2 - h / 2, w, h }, vb = { ...inicial };

  const nd = s => normal(s).replace(/[^a-z0-9]+/g, " ").trim();
  const deptoDe = new Map(MAPA.deptos.map((d, i) => [nd(d.n), i]));
  caja.innerHTML = `<svg class="mapa" viewBox="${vb.x} ${vb.y} ${vb.w} ${vb.h}" style="height:${alto}px" role="img" aria-label="Mapa de Colombia con lo comprado por municipio">
      <defs><radialGradient id="burbuja" cx="38%" cy="32%" r="70%">
        <stop offset="0" style="stop-color:color-mix(in srgb, var(--s1) 45%, #fff);stop-opacity:.95"/>
        <stop offset="1" style="stop-color:var(--s1);stop-opacity:.62"/></radialGradient></defs>
      <g class="deptos">${MAPA.deptos.map((d, i) => `<path data-d="${i}" d="${d.d}"><title>${esc(d.n)}</title></path>`).join("")}</g>
      <g class="burbujas"></g><g class="nombres"></g></svg>
    <div class="mapa-mes" id="mapaMes" hidden></div>
    <div class="mapa-botones"><button class="redondo" data-z="in" aria-label="Acercar">+</button><button class="redondo" data-z="out" aria-label="Alejar">−</button><button class="redondo" data-z="fit" aria-label="Encuadrar">${ICO.mapaFit}</button></div>
    <div class="mapa-leyenda"><i></i><span>Tamaño = valor comprado</span></div>`;
  const svg = caja.querySelector("svg"), gB = svg.querySelector(".burbujas"), gN = svg.querySelector(".nombres");
  const caminos = [...svg.querySelectorAll(".deptos path")];

  const actuales = () => (mes ? porMes.get(mes) : todos);
  const escala = () => (mes ? maxMes : todos[0].v || 1);
  const radio = p => vb.w * (0.011 + 0.05 * Math.sqrt(p.v / escala()));

  // nombres sin encimarse: se prueba arriba, abajo, a la derecha y a la izquierda; si no cabe, no se pone
  function nombres(pts) {
    const fs = vb.w * 0.029, alto = fs * 1.15, ocupado = [];
    const choca = r => ocupado.some(o => r.x < o.x + o.w && r.x + r.w > o.x && r.y < o.y + o.h && r.y + r.h > o.y);
    const dentro = r => r.x >= vb.x && r.y >= vb.y && r.x + r.w <= vb.x + vb.w && r.y + r.h <= vb.y + vb.h;
    for (const p of pts) { const r = radio(p) * 0.55; ocupado.push({ x: p.x - r, y: p.y - r, w: 2 * r, h: 2 * r, propio: p.cod }); } // el centro de cada círculo
    const salida = [];
    for (const p of pts) {
      if (salida.length >= 14) break;
      const t = nombreMuni(p.cod), an = anchoTexto(t, fs) + fs * 0.3, r = radio(p), sep = fs * 0.25;
      const opciones = [
        { x: p.x - an / 2, y: p.y - r - sep - alto, tx: p.x, ty: p.y - r - sep - fs * 0.22, a: "middle" },
        { x: p.x - an / 2, y: p.y + r + sep, tx: p.x, ty: p.y + r + sep + fs * 0.9, a: "middle" },
        { x: p.x + r + sep, y: p.y - alto / 2, tx: p.x + r + sep, ty: p.y + fs * 0.35, a: "start" },
        { x: p.x - r - sep - an, y: p.y - alto / 2, tx: p.x - r - sep, ty: p.y + fs * 0.35, a: "end" },
      ];
      for (const o of opciones) {
        const caja = { x: o.x, y: o.y, w: an, h: alto };
        const otros = ocupado.filter(x => x.propio !== p.cod);
        if (!dentro(caja) || otros.some(x => caja.x < x.x + x.w && caja.x + caja.w > x.x && caja.y < x.y + x.h && caja.y + caja.h > x.y)) continue;
        ocupado.push(caja); salida.push(`<text x="${o.tx}" y="${o.ty}" font-size="${fs}" text-anchor="${o.a}">${esc(t)}</text>`); break;
      }
    }
    return salida.join("");
  }

  const dibujar = () => {
    const pts = actuales(), total = new Map();
    for (const p of pts) { const k = deptoDe.get(nd(p.dpto)); if (k != null) total.set(k, (total.get(k) || 0) + p.v); }
    const maxD = Math.max(1, ...total.values());
    caminos.forEach((c, i) => { const v = total.get(i); c.style.fill = v ? `color-mix(in srgb, var(--s1) ${Math.round(5 + 15 * Math.sqrt(v / maxD))}%, var(--sup2))` : ""; });
    // primero los grandes, para que los pequeños queden encima y se puedan tocar
    gB.innerHTML = [...pts].sort((a, b) => b.v - a.v).map(p => `<g data-cod="${p.cod}" class="${p.cod === elegido ? "si" : ""}"><circle cx="${p.x}" cy="${p.y}" r="${radio(p)}"/><circle class="centro" cx="${p.x}" cy="${p.y}" r="${vb.w * 0.0045}"/><title>${esc(nombreMuni(p.cod))}: ${esc(corto(p.v))}</title></g>`).join("");
    gN.innerHTML = nombres(pts);
  };
  const aplicar = () => { svg.setAttribute("viewBox", `${vb.x} ${vb.y} ${vb.w} ${vb.h}`); dibujar(); };

  const elegir = cod => {
    elegido = cod; dibujar();
    const lista = actuales(), p = lista.find(x => x.cod === cod);
    if (!p) { $("#mapaSel").innerHTML = cod ? `<div class="tarjeta mapa-sel"><p class="nota" style="margin:0">${esc(nombreMuni(cod))}: sin órdenes en ${esc(fmtMesLargo.format(aFecha(mes + "-01")))}.</p></div>` : ""; return; }
    const top = [...p.cl].sort((a, b) => b[1] - a[1]).slice(0, 3), suma = lista.reduce((s, g) => s + g.v, 0);
    $("#mapaSel").innerHTML = `<div class="tarjeta mapa-sel"><div class="ms-tope"><div><b>${esc(nombreMuni(p.cod))}</b><small>${esc(deptoMuni(p.cod))}${mes ? " · " + esc(fmtMesLargo.format(aFecha(mes + "-01"))) : ""}</small></div><button class="redondo" id="msCerrar" aria-label="Cerrar">✕</button></div>
      <div class="ms-cifras"><div><b>${esc(corto(p.v))}</b><small>comprado</small></div><div><b>${fmtNum.format(p.n)}</b><small>${p.n === 1 ? "orden" : "órdenes"}</small></div><div><b>${Math.round(p.v / (suma || 1) * 100)}%</b><small>del ${mes ? "mes" : "periodo"}</small></div></div>
      ${top.length ? `<div class="ms-clientes">${top.map(([c, v]) => `<span><i></i>${esc(c)} · ${esc(corto(v))}</span>`).join("")}</div>` : ""}
      ${p.est ? `<p class="nota" style="margin:8px 0 0">${plural(p.est, "orden ubicada", "órdenes ubicadas")} por estimación (el texto no decía el lugar).</p>` : ""}
      <a class="boton primario ancho" href="#/ordenes/mu:${p.cod}">Ver sus órdenes</a></div>`;
    $("#msCerrar").onclick = () => elegir(null);
  };

  // tocar, arrastrar, pellizcar y la rueda del ratón
  const punteros = new Map(); let movido = false, tocado = null, recorrido = 0;
  const enUnidades = (px, py) => { const r = svg.getBoundingClientRect(); return [vb.x + (px - r.left) / r.width * vb.w, vb.y + (py - r.top) / r.height * vb.h]; };
  const zoom = (f, cx, cy) => {
    const nw = Math.min(MAPA.ancho * 1.3, Math.max(30, vb.w * f)), k = nw / vb.w;
    vb.x = cx - (cx - vb.x) * k; vb.y = cy - (cy - vb.y) * k; vb.w = nw; vb.h *= k; aplicar();
  };
  svg.addEventListener("pointerdown", e => {
    if (!punteros.size) { tocado = e.target.closest("[data-cod]")?.dataset.cod ?? null; recorrido = 0; movido = false; }
    punteros.set(e.pointerId, [e.clientX, e.clientY]); if (punteros.size > 1) movido = true;
    try { svg.setPointerCapture(e.pointerId); } catch { }
  });
  svg.addEventListener("pointermove", e => {
    if (!punteros.has(e.pointerId)) return;
    const antes = punteros.get(e.pointerId); punteros.set(e.pointerId, [e.clientX, e.clientY]);
    const r = svg.getBoundingClientRect();
    if (punteros.size === 1) {
      const dx = e.clientX - antes[0], dy = e.clientY - antes[1];
      recorrido += Math.abs(dx) + Math.abs(dy); if (recorrido > 8) movido = true;
      vb.x -= dx / r.width * vb.w; vb.y -= dy / r.height * vb.h; aplicar();
    } else if (punteros.size === 2) {
      const otro = [...punteros].find(([id]) => id !== e.pointerId)[1];
      const dAntes = Math.hypot(antes[0] - otro[0], antes[1] - otro[1]), dAhora = Math.hypot(e.clientX - otro[0], e.clientY - otro[1]);
      if (dAhora > 0) zoom(dAntes / dAhora, ...enUnidades((e.clientX + otro[0]) / 2, (e.clientY + otro[1]) / 2));
    }
  });
  svg.addEventListener("pointerup", e => {
    if (punteros.size === 1 && !movido) { elegir(tocado); if (tocado) vibrar(); } // un toque, no un arrastre
    punteros.delete(e.pointerId);
  });
  svg.addEventListener("pointercancel", e => punteros.delete(e.pointerId));
  svg.addEventListener("wheel", e => { e.preventDefault(); zoom(e.deltaY > 0 ? 1.18 : 1 / 1.18, ...enUnidades(e.clientX, e.clientY)); }, { passive: false });
  svg.addEventListener("dblclick", e => zoom(0.6, ...enUnidades(e.clientX, e.clientY)));
  caja.querySelectorAll("[data-z]").forEach(b => b.onclick = () => {
    if (b.dataset.z === "fit") Object.assign(vb, inicial), aplicar();
    else zoom(b.dataset.z === "in" ? 0.7 : 1 / 0.7, vb.x + vb.w / 2, vb.y + vb.h / 2);
  });

  // la lista: el mismo contenido del mapa, para leerlo sin tocar círculos
  const pintarLista = () => {
    const lista = actuales(), maxL = lista[0]?.v || 1;
    $("#mapaLista").innerHTML = lista.length ? lista.slice(0, 8).map(g => `<a class="rank" href="#/ordenes/mu:${g.cod}"><div class="r1"><span>${esc(nombreMuni(g.cod))} <small class="dep">${esc(deptoMuni(g.cod))}</small></span><b>${esc(corto(g.v))}</b></div>
        <div class="r2"><i style="width:${Math.max(0.8, g.v / maxL * 100)}%"></i></div><small>${plural(g.n, "orden", "órdenes")}${g.est ? ` · ${g.est} estimada${g.est > 1 ? "s" : ""}` : ""}</small></a>`).join("") +
      (lista.length > 8 ? `<div class="rank resto"><div class="r1"><span>Otros ${fmtNum.format(lista.length - 8)} municipios</span><b>${esc(corto(lista.slice(8).reduce((s, g) => s + g.v, 0)))}</b></div></div>` : "")
      : `<div class="vacio" style="padding:20px">Sin órdenes en este mes.</div>`;
  };

  // línea de tiempo: ver el mapa mes a mes, o en reproducción
  const tiempo = $("#mapaTiempo");
  if (meses.length > 1) {
    tiempo.hidden = false;
    tiempo.innerHTML = `<button class="redondo" id="tPlay" aria-label="Reproducir mes a mes">${ICO.play}</button>
      <div class="t-centro"><div class="t-etq"><b id="tEtq">Todo el periodo</b><small id="tVal"></small></div>
      <input type="range" id="tRango" min="0" max="${meses.length}" value="0" aria-label="Mes"></div>`;
    const rango = $("#tRango"), etq = $("#tEtq"), val = $("#tVal"), btn = $("#tPlay"), cartel = $("#mapaMes");
    let reloj = null;
    const ir = i => {
      mes = i ? meses[i - 1] : null; rango.value = i;
      const l = actuales(), suma = l.reduce((s, g) => s + g.v, 0);
      etq.textContent = mes ? fmtMesLargo.format(aFecha(mes + "-01")) : "Todo el periodo";
      val.textContent = corto(suma);
      cartel.hidden = !mes; cartel.textContent = mes ? fmtMesLargo.format(aFecha(mes + "-01")) : "";
      rango.style.setProperty("--p", (i / meses.length * 100) + "%");
      dibujar(); pintarLista(); if (elegido) elegir(elegido);
    };
    const parar = () => { clearInterval(reloj); reloj = null; btn.innerHTML = ICO.play; btn.classList.remove("activo"); };
    btn.onclick = () => {
      if (reloj) return parar();
      btn.innerHTML = ICO.pausa; btn.classList.add("activo");
      let i = Number(rango.value) >= meses.length ? 1 : Number(rango.value) + 1; ir(i);
      reloj = setInterval(() => { if (!document.body.contains(btn)) return clearInterval(reloj); i++; if (i > meses.length) { parar(); ir(0); return; } ir(i); }, reducido() ? 1600 : 1100);
    };
    rango.oninput = () => { parar(); ir(Number(rango.value)); };
    ir(0);
  } else { tiempo.hidden = true; aplicar(); pintarLista(); }

  const corr = sel.filter(o => o.muc === "corregido").length, est = sel.filter(o => o.muc === "cliente" || o.muc === "contrato").length;
  const enl = $("#enlaceLugares");
  if (corr || est || sinUbicar) { enl.hidden = false; enl.innerHTML = `Revisar cómo se leyeron los lugares · ${[corr && plural(corr, "corregido", "corregidos"), est && plural(est, "estimado", "estimados"), sinUbicar && `${sinUbicar} sin ubicar`].filter(Boolean).join(" · ")}`; }
}

// revisión: qué se leyó en cada texto escrito
let filtroLugares = "revisar";
function lugares(arg) {
  if (arg) filtroLugares = arg;
  vista.innerHTML = `<div class="entra"><h1>Lugares</h1><p class="sub">Cómo se leyó el lugar de entrega que está escrito en cada orden. Los errores de escritura se corrigen solos; si el texto no dice dónde (una dirección, una persona), se estima con las demás órdenes del mismo cliente.</p>
    <div id="lugCuerpo"><div class="esq" style="height:300px"></div></div></div>`;
  asegurarLugares().then(() => {
    const cuerpo = $("#lugCuerpo"); if (!cuerpo) return;
    const grupos = new Map();
    for (const o of D.ordenes) {
      const t = (o.ca || "").trim() || "(vacío)", id = t + "|" + o.mu + "|" + o.muc;
      const g = grupos.get(id) || { t, mu: o.mu, como: o.muc, n: 0 }; g.n++; grupos.set(id, g);
    }
    const todos = [...grupos.values()].sort((a, b) => b.n - a.n);
    const cuenta = f => D.ordenes.filter(f).length;
    const FL = [["revisar", "Para revisar", g => g.como !== "exacto" && g.como !== "sitio"], ["corregido", "Corregidos", g => g.como === "corregido"],
      ["estimado", "Estimados", g => g.como === "cliente" || g.como === "contrato" || !g.mu], ["todos", "Todos", () => true]];
    const pintar = () => {
      const f = (FL.find(x => x[0] === filtroLugares) || FL[0])[2], l = todos.filter(f);
      $("#lugLista").innerHTML = l.length ? l.map(g => { const [t, c] = g.mu ? COMO[g.como] || ["—", "t-gris"] : ["Sin ubicar", "t-rojo"];
        return `<a class="lug tarjeta" href="#/ordenes/q:${encodeURIComponent(g.t)}"><div class="lug-t">«${esc(g.t)}»</div>
          <div class="lug-m">${ICO.flecha}<b>${esc(nombreMuni(g.mu))}</b><small>${esc(deptoMuni(g.mu))}</small></div>
          <div class="lug-p"><span class="chip ${c}">${t}</span><small>${plural(g.n, "orden", "órdenes")}</small></div></a>`; }).join("")
        : `<div class="vacio"><div class="ilus">${ICO.ok}</div><b>Nada que revisar</b>Todos los lugares se leyeron tal cual.</div>`;
    };
    cuerpo.innerHTML = `<div class="kpis lug-kpis">
        <div class="kpi tarjeta"><span>Tal cual</span><b>${fmtNum.format(cuenta(o => o.muc === "exacto" || o.muc === "sitio"))}</b><small>órdenes</small></div>
        <div class="kpi tarjeta"><span>Corregidos</span><b>${fmtNum.format(cuenta(o => o.muc === "corregido"))}</b><small>errores de escritura</small></div>
        <div class="kpi tarjeta"><span>Estimados</span><b>${fmtNum.format(cuenta(o => o.muc === "cliente" || o.muc === "contrato"))}</b><small>por el cliente</small></div>
        <div class="kpi tarjeta"><span>Sin ubicar</span><b>${fmtNum.format(cuenta(o => !o.mu))}</b><small>órdenes</small></div></div>
      <div class="filtros" style="margin-top:12px">${FL.map(([k, t]) => `<button data-fl="${k}" class="${k === filtroLugares ? "si" : ""}">${t}</button>`).join("")}</div>
      <div class="lista" id="lugLista" style="margin-top:10px"></div>
      <p class="nota">Municipios: DIVIPOLA del DANE (datos.gov.co). Mapa: geoBoundaries / OpenStreetMap.</p>`;
    cuerpo.querySelectorAll("[data-fl]").forEach(b => b.onclick = () => {
      filtroLugares = b.dataset.fl; cuerpo.querySelectorAll("[data-fl]").forEach(x => x.classList.toggle("si", x === b)); pintar();
    });
    pintar();
  });
}

// ---------------------------------------------------------------- presentación: el tablero en pantalla grande
// Diapositivas que pasan solas (para una reunión o una pantalla). Flechas o toques para moverse; Esc para salir.
let presentacion = null;
function terminarPresentacion() {
  if (!presentacion) return;
  clearInterval(presentacion.reloj); document.removeEventListener("keydown", presentacion.teclas);
  document.documentElement.classList.remove("presentando");
  if (document.fullscreenElement) document.exitFullscreen?.().catch(() => { });
  presentacion = null;
}
function presentar() {
  terminarPresentacion();
  const R = rango("12m"), entre = o => o.f && o.f >= R.desde && o.f <= R.hasta;
  const sel = D.ordenes.filter(entre), ant = D.ordenes.filter(o => o.f && o.f >= R.antes[0] && o.f <= R.antes[1]);
  const suma = l => l.reduce((s, o) => s + (o.cop || 0), 0);
  const cub = cubetas(R.desde, R.hasta, "mes"), pos = new Map(cub.map((k, i) => [k, i]));
  const sumas = cub.map(() => 0), cuentas = cub.map(() => 0);
  for (const o of sel) { const i = pos.get(claveDe(o.f, "mes")); if (i != null) { sumas[i] += o.cop || 0; cuentas[i]++; } }
  const hoy = hoyIso(), delMes = D.ordenes.filter(o => (o.f || "").startsWith(hoy.slice(0, 7)));
  const fmtN = v => fmtNum.format(Math.round(v)), provs = l => new Set(l.map(o => o.p)).size;
  const diapositivas = [
    { t: "Resumen del mes", html: () => `<div class="pr-portada"><small>Comprado en ${esc(fmtMesLargo.format(aFecha(hoy)))}</small><b>${esc(corto(suma(delMes)))}</b>
        <div class="pr-trio"><div><b>${fmtN(delMes.length)}</b><small>órdenes</small></div><div><b>${fmtN(provs(delMes))}</b><small>proveedores</small></div><div><b>${fmtN(D.ordenes.filter(o => o.venc).length)}</b><small>entregas vencidas</small></div></div></div>` },
    { t: "Últimos 12 meses", html: () => `<div class="kpis pr-kpis">${kpi("Comprado", suma(sel), suma(ant), corto)}${kpi("Órdenes", sel.length, ant.length, fmtN)}${kpi("Promedio por orden", sel.length ? suma(sel) / sel.length : 0, ant.length ? suma(ant) / ant.length : null, corto)}${kpi("Proveedores", provs(sel), provs(ant), fmtN)}</div>${grafico(cub, "mes", sumas, cuentas)}`, despues: () => activarGrafico(cub, "mes", sumas, cuentas) },
    { t: "¿Dónde se entrega?", html: () => `<div class="pr-dos"><div>${tarjetaMapa()}</div></div>`, despues: () => pintarMapa(sel, R).then(() => setTimeout(() => $("#tPlay")?.click(), 900)) },
    { t: "Proveedores principales", html: () => `<div class="tarjeta bloque">${ranking(agrupar(sel, o => o.p), "p", 8, g => plural(g.n, "orden", "órdenes"))}</div>` },
    { t: "Por cliente", html: () => `<div class="tarjeta bloque">${partes(agrupar(sel, o => o.c))}</div>` },
    { t: "Precios que subieron", html: () => `<div class="tarjeta bloque">${preciosQueSubieron(sel)}</div>` },
    { t: "Pendientes", html: () => atencion() },
  ];
  document.documentElement.classList.add("presentando");
  vista.innerHTML = `<div class="presenta"><div class="pr-barra"><i id="prProg"></i></div>
    <header class="pr-cabeza"><div><small>Órdenes de Compra · ${esc(fecha(hoy))}</small><h1 id="prTitulo"></h1></div>
      <div class="pr-controles"><button class="redondo" id="prAnt" aria-label="Anterior">‹</button><span id="prNum"></span><button class="redondo" id="prPausa" aria-label="Pausar">${ICO.pausa}</button><button class="redondo" id="prSig" aria-label="Siguiente">›</button><button class="boton" id="prSalir">Salir</button></div></header>
    <div class="pr-cuerpo" id="prCuerpo"></div></div>`;
  const P = presentacion = { i: 0, reloj: null, pausa: false, DUR: 12000 };
  const ir = i => {
    P.i = (i + diapositivas.length) % diapositivas.length; const d = diapositivas[P.i];
    $("#prTitulo").textContent = d.t; $("#prNum").textContent = `${P.i + 1} / ${diapositivas.length}`;
    $("#prCuerpo").innerHTML = `<div class="entra">${d.html()}</div>`; d.despues?.();
    const barra = $("#prProg"); barra.style.animation = "none"; void barra.offsetWidth; barra.style.animation = P.pausa ? "none" : `avance ${P.DUR}ms linear forwards`;
  };
  const arrancar = () => { clearInterval(P.reloj); if (!P.pausa) P.reloj = setInterval(() => ir(P.i + 1), P.DUR); };
  $("#prAnt").onclick = () => { ir(P.i - 1); arrancar(); };
  $("#prSig").onclick = () => { ir(P.i + 1); arrancar(); };
  $("#prPausa").onclick = () => { P.pausa = !P.pausa; $("#prPausa").innerHTML = P.pausa ? ICO.play : ICO.pausa; ir(P.i); arrancar(); };
  $("#prSalir").onclick = () => { terminarPresentacion(); location.hash = "#/tablero"; };
  P.teclas = e => {
    if (e.key === "ArrowRight" || e.key === " ") { e.preventDefault(); $("#prSig").click(); }
    else if (e.key === "ArrowLeft") $("#prAnt").click();
    else if (e.key === "Escape") $("#prSalir").click();
  };
  document.addEventListener("keydown", P.teclas);
  document.documentElement.requestFullscreen?.().catch(() => { });
  ir(0); arrancar();
}

// ---------------------------------------------------------------- avisos: lo nuevo desde la última vez
// Al traer datos nuevos se compara con lo que ya habías visto. Si diste permiso, sale también un aviso del sistema.
// (Un aviso con la aplicación cerrada necesitaría un servidor propio que la despierte; esta app no tiene servidor.)
let NOVEDADES = null;
function revisarNovedades() {
  let antes = null; try { antes = JSON.parse(pref("vistos") || "null"); } catch { }
  const ahora = { venc: D.ordenes.filter(o => o.venc).map(o => o.n), borr: D.ordenes.filter(o => o.e === 4).map(o => o.n), sc: D.ordenes.filter(o => o.sc).map(o => o.n), conf: D.ordenes.filter(o => o.e === 2).map(o => o.n) };
  pref("vistos", JSON.stringify(ahora));
  if (!antes) return;
  const nuevos = k => { const s = new Set(antes[k] || []); return ahora[k].filter(n => !s.has(n)); };
  const nv = { venc: nuevos("venc"), borr: nuevos("borr"), sc: nuevos("sc"), conf: nuevos("conf") };
  if (!nv.venc.length && !nv.borr.length && !nv.sc.length && !nv.conf.length) return;
  NOVEDADES = nv;
  const lineas = [nv.venc.length && plural(nv.venc.length, "entrega se venció", "entregas se vencieron"), nv.borr.length && plural(nv.borr.length, "borrador nuevo por enviar", "borradores nuevos por enviar"),
    nv.conf.length && plural(nv.conf.length, "orden confirmada", "órdenes confirmadas"), nv.sc.length && plural(nv.sc.length, "orden pasó a «sin confirmar»", "órdenes pasaron a «sin confirmar»")].filter(Boolean);
  if (pref("avisos") === "si" && "Notification" in window && Notification.permission === "granted")
    navigator.serviceWorker?.getRegistration().then(r => r?.showNotification("Órdenes de Compra", { body: lineas.join(NL), icon: "iconos/icono-192.png", badge: "iconos/icono-192.png", tag: "novedades", data: { url: "#/inicio" } })).catch(() => { });
}
function tarjetaNovedades() {
  const nv = NOVEDADES; if (!nv) return "";
  const filas = [[nv.venc, "rojo", ICO.alerta, "entrega se venció", "entregas se vencieron", "#/ordenes/vencida"], [nv.borr, "ambar", ICO.borrador, "borrador nuevo por enviar", "borradores nuevos por enviar", "#/ordenes/borrador"],
    [nv.conf, "verde", ICO.ok, "orden confirmada por el proveedor", "órdenes confirmadas por el proveedor", "#/ordenes/confirmada"], [nv.sc, "ambar", ICO.reloj, "orden pasó a «sin confirmar»", "órdenes pasaron a «sin confirmar»", "#/seguimiento"]].filter(f => f[0].length);
  return `<div class="tarjeta novedades"><div class="nov-tope"><b>Desde tu última visita</b><button class="redondo" id="novCerrar" aria-label="Ocultar">✕</button></div>
    ${filas.map(([l, t, i, uno, varios, h]) => `<a class="aviso-fila" href="${h}"><div class="ico t-${t}">${i}</div><div><b>${esc(plural(l.length, uno, varios))}</b><small>${esc(l.slice(0, 4).join(", "))}${l.length > 4 ? "…" : ""}</small></div>${ICO.flecha}</a>`).join("")}</div>`;
}
async function activarAvisos() {
  if (!("Notification" in window)) { aviso("Este celular no permite avisos desde la web"); return; }
  const p = await Notification.requestPermission();
  if (p === "granted") { pref("avisos", "si"); aviso("Listo: te avisaré de lo nuevo"); } else { pref("avisos", "no"); aviso("Sin permiso para avisos"); }
  dibujar();
}

function ajustes() {
  const tema = pref("tema") || "sistema";
  vista.innerHTML = `<div class="entra"><h1>Ajustes</h1>
    <div class="tarjeta">
      <div class="campo"><label>Cuenta</label><small>${cuenta ? esc(cuenta.name || "") + " · " + esc(cuenta.username) : DEMO ? "Modo de ejemplo" : LOCAL ? "Datos de este computador (sin iniciar sesión)" : "Sin iniciar sesión"}</small></div>
      <div class="campo"><label for="ruta">Carpeta de órdenes en tu OneDrive</label><input id="ruta" value="${esc(ruta())}" autocapitalize="characters"><small>La que contiene _AUTOMATIZACION. Normalmente «COMPRAS/ORDENES DE COMPRA».</small></div>
      <div class="campo"><label>Avisos</label><small>${pref("avisos") === "si" && "Notification" in window && Notification.permission === "granted" ? "Activados: cuando abras la app y haya algo nuevo (entregas vencidas, borradores, confirmaciones) te sale un aviso." : "Te avisa de lo nuevo cada vez que abres la app: entregas vencidas, borradores por enviar y órdenes confirmadas."}</small>${pref("avisos") === "si" ? "" : '<button class="boton" id="btnAvisos">Activar avisos</button>'}</div>
      <div class="campo"><label>Modo presentación</label><small>El tablero en pantalla completa, con diapositivas que pasan solas. Ideal para reuniones.</small><a class="boton" href="#/presentar">Presentar</a></div>
      <div class="campo"><label for="tema">Tema</label><select id="tema"><option value="sistema">Como el celular</option><option value="claro">Claro</option><option value="oscuro">Oscuro</option></select></div>
    </div>
    <div class="acciones">
      ${cuenta ? '<button class="boton" id="btnSalir">Cerrar sesión</button>' : ""}
      <button class="boton" id="btnBorrar">Borrar lo guardado en este celular</button>
    </div>
    ${instalable()}
    <p class="nota">Los datos se leen de tu OneDrive con tu cuenta y solo con permiso de lectura. Se actualizan cada vez que el PC (o el portátil) corre la herramienta.${D ? " Última generación: " + esc(new Date(D.generado).toLocaleString("es-CO")) + "." : ""}</p></div>`;
  $("#tema").value = tema;
  if ($("#btnAvisos")) $("#btnAvisos").onclick = activarAvisos;
  $("#tema").onchange = e => { pref("tema", e.target.value); aplicarTema(); };
  $("#ruta").onchange = e => { pref("ruta", e.target.value.trim().replace(/^[/]+|[/]+$/g, "")); almacen.borrar(); D = null; sincronizar(true); };
  $("#btnBorrar").onclick = async () => { await almacen.borrar(); aviso("Listo: se borró lo guardado"); };
  if ($("#btnSalir")) $("#btnSalir").onclick = async () => { await almacen.borrar(); msalApp.logoutRedirect({ account: cuenta }); };
}

function aplicarTema() {
  const t = pref("tema");
  if (t === "claro" || t === "oscuro") document.documentElement.dataset.tema = t; else delete document.documentElement.dataset.tema;
}

// ---------------------------------------------------------------- detalle de una orden
function paso(clase, ico, titulo, sub) {
  return `<div class="paso ${clase}"><div class="punto">${ico}</div><div><b>${titulo}</b><small>${sub}</small></div></div>`;
}

function abrirOrden(n) {
  const o = D?.porNumero.get(n);
  if (!o) { aviso("No encontré la orden " + n); return; }
  const enviada = o.e === 1 || o.e === 2 || !!o.fe;
  const pasos = [
    paso("hecho", ICO.ok, "Emitida", esc(fecha(o.f))),
    o.e === 4 ? paso("ahora", ICO.borrador, "En borrador", "falta darle Enviar en Outlook")
      : o.e === 3 ? paso("", ICO.enviar, "Enviada", "anterior al buzón: no se sabe")
      : enviada ? paso("hecho", ICO.ok, "Enviada al proveedor", esc(fecha(o.fe)) + (o.d != null ? ` · hace ${o.d} días` : ""))
      : paso("ahora", ICO.enviar, "Por enviar", "todavía no sale por correo"),
    o.e === 2 ? paso("hecho", ICO.ok, "Confirmada", "el proveedor la devolvió")
      : paso(o.sc ? "alerta" : o.e === 1 ? "ahora" : "", ICO.reloj, "Confirmada", o.e === 1 ? (o.ce ? "contestó, pero no la ha devuelto firmada" : "todavía no") : "—"),
    o.r ? paso("hecho", ICO.ok, "Entregada", "ya hay entrada de mercancía")
      : paso(o.venc ? "alerta" : "", ICO.caja, o.venc ? "Entrega vencida" : "Entregada", o.en ? "pactada para " + esc(fecha(o.en)) : "sin fecha pactada"),
  ];
  const hoy = hoyIso();
  const it = o.it || [];
  const imp = o.v != null && o.st != null && Math.abs(o.v - o.st) > 1 ? o.v - o.st : 0;
  const asunto = encodeURIComponent("Orden de compra N° " + o.n);

  // lo clave de un vistazo
  const hechos = [];
  if (o.en) {
    let s = fechaCorta(o.en);
    if (!o.r && o.e !== 3) { const d = diasEntre(hoy, o.en); s += d < 0 ? ` · hace ${-d} d` : d === 0 ? " · hoy" : ` · en ${d} d`; }
    hechos.push([ICO.caja, o.venc ? "Entrega vencida" : o.r ? "Entregada" : "Entrega", s, o.venc ? "c-rojo" : ""]);
  }
  if (o.fe) hechos.push([ICO.enviar, "Enviada", fechaCorta(o.fe) + (o.d != null ? ` · hace ${o.d} d` : ""), o.sc ? "c-ambar" : ""]);
  hechos.push([ICO.mes, "Emitida", fechaCorta(o.f), ""]);
  if (o.fp) hechos.push([ICO.pago, "Pago", o.fp, ""]);

  // artículos: cada uno con su peso dentro de la orden
  const base = o.st || it.reduce((s, a) => s + (a[6] || 0), 0) || 1;
  const articulo = a => { const p = (a[6] || 0) / base * 100; return `<div class="articulo"><div class="d">${esc(a[1] || "(sin descripción)")}</div>
      ${a[2] ? `<div class="c">${esc(a[2])}</div>` : ""}<div class="q">${esc(fmtNum.format(a[4] ?? 0))} ${esc(a[3] || "")} × ${esc(dinero(a[5], o.m))}${a[0] ? " · " + esc(a[0]) : ""}</div>
      <div class="t">${esc(dinero(a[6], o.m))}<small>${p >= 1 ? Math.round(p) : "&lt;1"}%</small></div>
      ${etiquetaPrecio(o, a) ? `<div class="historial">${etiquetaPrecio(o, a)}</div>` : ""}
      <div class="peso"><i style="width:${Math.min(100, p)}%"></i></div></div>`; };
  const MUESTRA = 6;

  // historial con el mismo proveedor
  const delProv = D.ordenes.filter(x => x.p === o.p).sort((a, b) => (b.f || "").localeCompare(a.f || ""));
  const otras = delProv.filter(x => x !== o).slice(0, 4);
  const desde = delProv.reduce((m, x) => x.f && (!m || x.f < m) ? x.f : m, "");
  const enlaceProv = "#/ordenes/p:" + encodeURIComponent(o.p);

  const grupo = (titulo, filas) => {
    const h = filas.filter(([, v]) => v).map(([k, v, crudo]) => `<div class="dato"><span>${k}</span><b>${crudo ? v : esc(v)}</b></div>`).join("");
    return h ? `<h2>${titulo}</h2><div class="tarjeta datos">${h}</div>` : "";
  };

  $("#hojaCuerpo").innerHTML = `
    <div class="cabeza">${avatar(o.p)}<div style="min-width:0"><h3>Orden ${esc(o.n)}</h3><p><a class="enlace" href="${enlaceProv}">${esc(o.p)}</a></p></div></div>
    <div class="tarjeta heroe tono-${tono(o)}">
      <div>${chips(o)}</div>
      <div class="total">${esc(dinero(o.v, o.m))}</div>
      <small class="sub">${o.m && o.m !== "COP" && o.cop ? "≈ " + esc(dinero(o.cop)) + (imp ? " · " : "") : ""}${imp ? "Subtotal " + esc(dinero(o.st, o.m)) + " + impuestos " + esc(dinero(imp, o.m)) : ""}</small>
      <div class="hechos">${hechos.map(([i, t, v, c]) => `<div class="hecho"><div class="ico">${i}</div><div><small>${t}</small><b class="${c}">${esc(v)}</b></div></div>`).join("")}</div>
    </div>
    <div class="acciones">
      <button class="boton primario" id="btnPdf">${ICO.pdf}Ver el PDF</button>
      ${o.em ? `<a class="boton" href="mailto:${esc(o.em)}?subject=${asunto}">${ICO.correo}Escribir</a>` : ""}
      <button class="boton" id="btnCompartir">${ICO.compartir}Compartir</button>
    </div>
    <h2>Recorrido</h2><div class="tarjeta recorrido">${pasos.join("")}</div>
    <h2>Artículos${it.length ? " · " + it.length : ""}</h2>
    <div class="tarjeta">${it.length ? it.slice(0, MUESTRA).map(articulo).join("") +
        (it.length > MUESTRA ? `<div id="masArt" hidden>${it.slice(MUESTRA).map(articulo).join("")}</div><button class="ver-mas" id="btnMasArt">Ver los ${it.length} artículos</button>` : "") +
        `<div class="sumas"><span>Subtotal</span><span>${esc(dinero(o.st, o.m))}</span></div>` +
        (imp ? `<div class="sumas fina"><span>Impuestos</span><span>${esc(dinero(imp, o.m))}</span></div><div class="sumas fina total-final"><span>Total</span><span>${esc(dinero(o.v, o.m))}</span></div>` : "")
      : `<div class="vacio" style="padding:24px">No se pudieron leer los artículos de esta orden.</div>`}</div>
    <h2>La orden</h2><div class="tarjeta vista-previa" id="previa"><div class="esq" style="width:100%;height:220px"></div></div>
    <h2>Con este proveedor</h2>
    <div class="tarjeta prov">
      <div class="mini3"><div><b>${fmtNum.format(delProv.length)}</b><small>${delProv.length === 1 ? "orden" : "órdenes"}</small></div>
        <div><b>${esc(corto(delProv.reduce((s, x) => s + (x.cop || 0), 0)))}</b><small>comprado</small></div>
        <div><b>${desde ? desde.slice(0, 4) : "—"}</b><small>primera orden</small></div></div>
      ${otras.map(x => `<a class="fila mini" href="#/orden/${esc(x.n)}"><div class="txt"><div class="l1"><span>Orden ${esc(x.n)}</span><span class="monto">${esc(corto(x.cop ?? x.v))}</span></div>
        <div class="l3">${chipEstado(x)}<span>${esc(fecha(x.f))}</span></div></div>${ICO.flecha}</a>`).join("")}
      ${delProv.length > 1 ? `<a class="ver-mas" href="${enlaceProv}">Ver sus ${fmtNum.format(delProv.length)} órdenes</a>` : ""}
    </div>
    ${grupo("Proyecto", [["Cliente", o.c], ["Contrato", o.ct], ["Lugar escrito", o.ca], ["Municipio", `<span id="datoMuni">…</span>`, true], ["Solicitud (RQ)", o.rq]])}
    ${grupo("Proveedor", [["NIT", o.nit], ["Contacto", o.at], ["Correo", o.em ? `<a href="mailto:${esc(o.em)}?subject=${asunto}">${esc(o.em)}</a>` : "", true]])}
    ${grupo("Compra", [["Forma de pago", o.fp], ["Tipo de compra", o.tc], ["Responsable", o.rs], ["Moneda", o.m && o.m !== "COP" ? o.m : ""]])}`;
  const h = $("#hoja"); h.hidden = false; h.classList.remove("saliendo");
  $("#hoja .panel").scrollTop = 0;
  $("#btnPdf").onclick = () => abrirPdf(o);
  asegurarLugares().then(() => {
    const el = $("#datoMuni"); if (!el) return;
    if (!o.mu) { el.textContent = "Sin ubicar"; return; }
    const nota = o.muc === "corregido" ? "corregido de lo escrito" : o.muc === "cliente" || o.muc === "contrato" ? "estimado: el texto no dice el lugar" : "";
    el.innerHTML = `<a href="#/ordenes/mu:${o.mu}">${esc(nombreMuni(o.mu))}</a> · ${esc(deptoMuni(o.mu))}${nota ? `<small class="nota-dato">${nota}</small>` : ""}`;
  }).catch(() => { const el = $("#datoMuni"); if (el) el.textContent = "—"; });
  if ($("#btnMasArt")) $("#btnMasArt").onclick = e => { $("#masArt").hidden = false; e.currentTarget.remove(); };
  $("#btnCompartir").onclick = async () => {
    const t = [`Orden de compra N° ${o.n}`, o.p, dinero(o.v, o.m), (ESTADO[o.e] || [""])[0], o.en ? "Entrega: " + fecha(o.en) : ""].filter(Boolean).join(NL);
    try {
      if (navigator.share) await navigator.share({ title: "Orden " + o.n, text: t });
      else { await navigator.clipboard.writeText(t); aviso("Copiado: ya lo puedes pegar"); }
    } catch { }
  };
  vistaPrevia(o);
}

async function vistaPrevia(o) {
  const el = $("#previa");
  if (LOCAL) { el.innerHTML = "<p>Usa «Ver el PDF» para abrir la orden.</p>"; return; }
  if (DEMO || !cuenta || !D.raizOneDrive && D.raizOneDrive !== "") { el.innerHTML = "<p>La imagen de la orden aparece cuando entras con tu cuenta.</p>"; return; }
  try {
    const camino = enDrive((D.raizOneDrive ? D.raizOneDrive + "/" : "") + o.a);
    let r = await graph(camino + "/thumbnails?select=c1200x1700");
    let url = r.value?.[0]?.c1200x1700?.url;
    if (!url) { r = await graph(camino + "/thumbnails"); url = r.value?.[0]?.large?.url; }
    if (!url) throw new Error("sin imagen");
    if ($("#previa") !== el) return;
    el.innerHTML = `<img src="${esc(url)}" alt="Primera página de la orden ${esc(o.n)}">`;
    el.onclick = () => { const v = $("#visor"); v.querySelector("img").src = url; v.hidden = false; v.classList.remove("zoom"); };
  } catch { el.innerHTML = "<p>No se pudo mostrar la imagen. Usa «Ver el PDF».</p>"; }
}

async function abrirPdf(o) {
  if (LOCAL) { if (!escritorio({ tipo: "pdf", a: o.a, n: o.n })) aviso("No se pudo abrir el PDF"); return; }
  if (DEMO || !cuenta) { aviso("Entra con tu cuenta para abrir el PDF"); return; }
  try {
    const it = await graph(enDrive((D.raizOneDrive ? D.raizOneDrive + "/" : "") + o.a) + "?$select=webUrl");
    window.open(it.webUrl, "_blank", "noopener");
  } catch { aviso("No encontré el PDF en tu OneDrive"); }
}

function cerrarHoja(sinAnimar) {
  const h = $("#hoja"), panel = $("#hoja .panel");
  if (h.hidden) return;
  if (sinAnimar) { h.hidden = true; panel.style.transform = ""; return; }
  h.classList.add("saliendo");
  setTimeout(() => { h.hidden = true; h.classList.remove("saliendo"); panel.style.transform = ""; }, 260);
}

// ---------------------------------------------------------------- navegación
// el mapa solo (12 meses), para incrustarlo: a la izquierda el mapa y su línea de tiempo; a la derecha el detalle y la lista
function mapa() {
  const R = rango("12m"), sel = D.ordenes.filter(o => o.f && o.f >= R.desde && o.f <= R.hasta);
  vista.innerHTML = `<div class="mapa-solo"><div class="ms-izq"><div class="tarjeta mapa-caja" id="mapaCaja"><div class="esq" style="height:340px"></div></div><div class="tarjeta linea-tiempo" id="mapaTiempo" hidden></div></div>
    <div class="ms-der"><div id="mapaSel"></div><div class="tarjeta bloque" id="mapaLista"></div><a class="ver-mas tarjeta enlace-lugares" id="enlaceLugares" href="#/lugares" hidden></a></div></div>`;
  pintarMapa(sel, R);
}
const VISTAS = { inicio, ordenes, tablero, seguimiento, entregas, ajustes, lugares, articulo, presentar, mapa };
const PESTANAS = ["inicio", "ordenes", "tablero", "seguimiento", "entregas"];
const reducido = () => matchMedia("(prefers-reduced-motion: reduce)").matches;
const vibrar = (ms = 8) => { try { navigator.vibrate?.(ms); } catch { } };

function dibujar(arg) {
  if (actual !== "presentar") terminarPresentacion();
  if (!D && actual !== "ajustes") return;
  VISTAS[actual](arg);
  dibujado = true;
  document.querySelectorAll("#pestanas a").forEach(a => a.classList.toggle("si", a.dataset.p === actual));
  const i = PESTANAS.indexOf(actual), ind = $("#indicador");
  ind.style.setProperty("--i", Math.max(0, i)); ind.classList.toggle("oculto", i < 0);
  observarTitulo();
  const bi = $("#btnInstalar"); if (bi) bi.onclick = async () => { instalar.prompt(); await instalar.userChoice; instalar = null; dibujar(); };
}

// al bajar, el título grande pasa a la barra de arriba
function observarTitulo() {
  const h1 = vista.querySelector("h1");
  $("#tituloBarra").textContent = h1 ? h1.textContent : "";
  revisarTitulo();
}
function revisarTitulo() {
  const h1 = vista.querySelector("h1"), barra = $("#barra");
  barra.classList.toggle("compacta", !!h1 && h1.getBoundingClientRect().bottom < barra.offsetHeight);
}
let pendienteTitulo = false;
window.addEventListener("scroll", () => {
  if (pendienteTitulo) return;
  pendienteTitulo = true;
  requestAnimationFrame(() => { pendienteTitulo = false; revisarTitulo(); });
}, { passive: true });

function alCambiar() {
  const h = location.hash.replace(/^#/, "").replace(/^[/]/, "");
  if (/^(code|error|state)=/.test(h)) return;
  const [p, arg] = h.split("/");
  // incrustado: cualquier otro destino se abre en la vista Análisis de la aplicación del computador
  if (INCRUSTADO && dibujado && p !== "mapa") { escritorio({ tipo: "ir", hash: "#/" + h }); history.replaceState(null, "", location.pathname + location.search + "#/mapa"); return; }
  if (p === "orden") {
    if (!dibujado) { actual = "ordenes"; dibujar(); }
    if (D) abrirOrden(decodeURIComponent(arg || ""));
    return;
  }
  const antes = !$("#hoja").hidden;
  cerrarHoja();
  const nueva = VISTAS[p] ? p : "inicio";
  if (antes && nueva === actual && !arg) return; // volver de una orden: la lista queda igual
  const cambiar = () => { actual = nueva; dibujar(arg); window.scrollTo({ top: 0 }); };
  // entre pestañas: la vista se desliza hacia el lado de la pestaña y el indicador viaja con ella
  if (document.startViewTransition && dibujado && !antes && nueva !== actual && !reducido()) {
    const de = PESTANAS.indexOf(actual), a = PESTANAS.indexOf(nueva), raiz = document.documentElement;
    raiz.dataset.dir = a >= 0 && de >= 0 && a < de ? "atras" : "adelante";
    raiz.classList.add("en-transicion");
    document.startViewTransition(cambiar).finished.finally(() => raiz.classList.remove("en-transicion"));
  } else cambiar();
}

// gestos de celular: jalar hacia abajo para actualizar y deslizar la hoja para cerrarla
function gestos() {
  const jalar = $("#jalar"), hoja = $("#hoja"), panel = $("#hoja .panel"), UMBRAL = 80;
  let y0 = null, d = 0;
  window.addEventListener("touchstart", e => {
    y0 = window.scrollY <= 0 && hoja.hidden && D && !jalar.classList.contains("cargando") ? e.touches[0].clientY : null; d = 0;
  }, { passive: true });
  window.addEventListener("touchmove", e => {
    if (y0 == null) return;
    d = Math.max(0, e.touches[0].clientY - y0);
    if (window.scrollY > 0) { y0 = null; d = 0; }
    const k = Math.min(1, d / UMBRAL);
    jalar.style.setProperty("--k", k); jalar.classList.toggle("visible", d > 8);
    if (d > UMBRAL && !jalar.classList.contains("listo")) vibrar(6);
    jalar.classList.toggle("listo", d > UMBRAL);
  }, { passive: true });
  window.addEventListener("touchend", async () => {
    if (y0 == null) return;
    y0 = null;
    if (d > UMBRAL) {
      jalar.classList.add("cargando");
      try { await sincronizar(true); } finally { jalar.classList.remove("cargando", "listo", "visible"); }
    } else jalar.classList.remove("visible", "listo");
    d = 0;
  });

  let h0 = null, dh = 0;
  panel.addEventListener("touchstart", e => {
    h0 = panel.scrollTop <= 0 ? e.touches[0].clientY : null; dh = 0; panel.style.transition = "none";
  }, { passive: true });
  panel.addEventListener("touchmove", e => {
    if (h0 == null) return;
    dh = Math.max(0, e.touches[0].clientY - h0);
    if (panel.scrollTop > 0) dh = 0;
    panel.style.transform = dh ? `translateY(${dh}px)` : "";
  }, { passive: true });
  panel.addEventListener("touchend", () => {
    if (h0 == null) return;
    h0 = null; panel.style.transition = "";
    if (dh > 110) { vibrar(); history.length > 1 ? history.back() : (location.hash = "#/" + actual); }
    else panel.style.transform = "";
  });
}

document.addEventListener("click", e => {
  if (e.target.closest("[data-cerrar]")) history.length > 1 ? history.back() : (location.hash = "#/" + actual);
  if (e.target.closest("[data-cerrar-visor]")) { $("#visor").hidden = true; return; }
  if (e.target.closest("#visor img")) $("#visor").classList.toggle("zoom");
});
$("#pestanas").addEventListener("click", e => { if (e.target.closest("a")) vibrar(); });
$("#btnRecargar").onclick = () => sincronizar(true);
gestos();
$("#btnAjustes").onclick = () => (location.hash = "#/ajustes");
window.addEventListener("hashchange", alCambiar);
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); instalar = e; if (actual === "inicio" && D) dibujar(); });
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && cuenta && D && Date.now() - (sincronizar.ultima || 0) > 120000) { sincronizar.ultima = Date.now(); sincronizar(); } });

// ---------------------------------------------------------------- arranque
function pantallaIngreso() {
  vista.innerHTML = `<div class="ingreso entra"><img src="iconos/icono-192.png" alt="">
    <h1>Órdenes de <span class="deg">Compra</span></h1>
    <p>Tus órdenes, su estado y sus artículos, desde el celular. Entra con tu correo de la empresa.</p>
    <button class="boton primario" id="btnEntrar">${ICO.ok}Entrar con Microsoft</button>
    <p class="nota">Solo pide permiso para leer tus archivos de OneDrive. No envía correos ni cambia nada.</p></div>`;
  $("#btnEntrar").onclick = () => msalApp.loginRedirect({ scopes: ALCANCES, prompt: "select_account" });
}

async function arrancar() {
  aplicarTema();
  if ("serviceWorker" in navigator && !DEMO && !LOCAL) navigator.serviceWorker.register("sw.js").catch(() => { });
  vista.innerHTML = `<div class="entra"><div class="esq" style="height:38px;width:60%;margin:8px 0 18px"></div>
    <div class="cifras">${'<div class="esq" style="height:128px"></div>'.repeat(4)}</div>
    <div class="esq" style="height:170px;margin-top:10px"></div></div>`;

  if (DEMO || LOCAL) { await sincronizar(); alCambiar(); return; }
  if (!C.clientId) {
    vista.innerHTML = `<div class="vacio entra"><div class="ilus">${ICO.alerta}</div><b>Falta configurar la aplicación</b>Hay que poner el «Id. de aplicación» de Microsoft Entra en config.js.</div>`;
    return;
  }
  msalApp = new msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: C.autoridad, redirectUri: location.origin + location.pathname, navigateToLoginRequestUrl: false },
    cache: { cacheLocation: "localStorage" },
  });
  await msalApp.initialize();
  try { const r = await msalApp.handleRedirectPromise(); if (r) { cuenta = r.account; history.replaceState(null, "", location.pathname + "#/inicio"); } }
  catch (e) { aviso("No se pudo iniciar sesión: " + (e.errorMessage || e.message || e)); }
  cuenta ||= msalApp.getAllAccounts()[0] || null;
  if (cuenta) msalApp.setActiveAccount(cuenta);

  const guardado = await almacen.leer("datos");
  if (guardado?.datos) { poner(guardado.datos); }
  if (!cuenta) { pantallaIngreso(); return; }
  alCambiar();
  sincronizar.ultima = Date.now();
  await sincronizar();
}

arrancar();
