// Órdenes de Compra para el celular. Lee _AUTOMATIZACION/movil/datos.json de tu OneDrive
// con tu cuenta (Microsoft Graph, solo lectura). Sin servidor propio: nada sale de Microsoft 365.
"use strict";

const C = window.CONFIG_OC || {};
const ALCANCES = ["User.Read", "Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
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
    if (!cuenta) return;
    const meta = await graph(enDrive(ruta() + "/_AUTOMATIZACION/movil/datos.json"));
    const guardado = await almacen.leer("datos");
    if (guardado && guardado.etag === meta.eTag && D) { if (manual) aviso("Ya tienes lo último"); return; }
    const r = await fetch(meta["@microsoft.graph.downloadUrl"]);
    poner(await r.json(), meta.eTag);
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
    <h2>Atención</h2>${atencion()}
    <h2>Resumen</h2>
    <div class="cifras">
      <a class="cifra tarjeta" href="#/ordenes/borrador"><div class="ico t-ambar">${ICO.borrador}</div><b data-n="${borr.length}">0</b><span>Por enviar</span><small>borradores en Outlook</small></a>
      <a class="cifra tarjeta" href="#/seguimiento"><div class="ico t-azul">${ICO.reloj}</div><b data-n="${sc.length}">0</b><span>Sin confirmar</span><small>enviadas hace ${D.diasAlerta || 8}+ días</small></a>
      <a class="cifra tarjeta" href="#/entregas"><div class="ico t-rojo">${ICO.alerta}</div><b data-n="${venc.length}">0</b><span>Entregas vencidas</span><small>sin llegada · ${esc(corto(venc.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
      <a class="cifra tarjeta" href="#/ordenes/mes"><div class="ico t-verde">${ICO.mes}</div><b data-n="${delMes.length}">0</b><span>Órdenes del mes</span><small>${esc(corto(delMes.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
    </div>
    ${instalable()}
    <h2>Recientes</h2><div class="lista">${D.ordenes.slice(0, 8).map(fila).join("")}</div></div>`;
  vista.querySelectorAll("[data-n]").forEach(contar);
}

function instalable() {
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
const FACETAS = { p: "Proveedor", c: "Cliente", tc: "Tipo de compra", mu: "Municipio" };
const ORDENES_POR = [["recientes", "Más recientes"], ["valor", "Mayor valor"], ["entrega", "Entrega más próxima"]];

function ordenes(arg) {
  if (arg) {
    const m = /^(p|c|tc|mu|q):(.*)$/.exec(arg);
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
  const valorFaceta = faceta && (faceta.k === "mu" ? nombreMuni(faceta.v) + " · " + deptoMuni(faceta.v) : faceta.v);
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
      const k = String(a[1] || "").trim(); if (!k) continue;
      const g = m.get(k) || { k, v: 0, n: 0, q: 0, u: a[3] };
      g.v += (a[6] || 0) * tasa; g.n++; g.q += a[4] || 0; m.set(k, g);
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
  return top.map(g => `<a class="rank" href="#/ordenes/${campo}:${encodeURIComponent(g.k)}"><div class="r1"><span>${esc(g.k)}</span><b>${esc(corto(g.v))}</b></div>
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
  vista.innerHTML = `<div class="${sinAnimar ? "" : "entra"}"><h1>Tablero</h1>
    <div class="filtros periodos">${PERIODOS.map(([k, t]) => `<button data-per="${k}" class="${k === periodo ? "si" : ""}">${t}</button>`).join("")}</div>
    <p class="sub rango">${esc(fecha(R.desde))} – ${esc(fecha(R.hasta))}${R.antes ? " · comparado con el mismo periodo del año pasado" : ""}</p>
    <div class="kpis">
      ${kpi("Comprado", tot, ant && suma(ant), corto)}
      ${kpi("Órdenes", n, ant && ant.length, fmtN)}
      ${kpi("Promedio por orden", n ? tot / n : 0, ant && ant.length ? suma(ant) / ant.length : null, corto)}
      ${kpi("Proveedores", provs(sel), ant && provs(ant), fmtN)}
    </div>
    <h2>Compras en el tiempo</h2>${grafico(cub, R.cubeta, sumas, cuentas)}
    <h2>¿Dónde se entrega?</h2>${tarjetaMapa()}
    <h2>Por cliente</h2><div class="tarjeta bloque">${partes(agrupar(sel, o => o.c))}</div>
    <h2>Proveedores principales</h2><div class="tarjeta bloque">${ranking(agrupar(sel, o => o.p), "p", 8, g => plural(g.n, "orden", "órdenes"))}</div>
    <h2>Tipo de compra</h2><div class="tarjeta bloque">${ranking(agrupar(sel, o => o.tc), "tc", 6, g => plural(g.n, "orden", "órdenes"))}</div>
    <h2>Artículos más comprados</h2><div class="tarjeta bloque">${ranking(articulos(sel), "q", 8, g => `${fmtNum.format(g.q)} ${g.u || "und"} · en ${plural(g.n, "orden", "órdenes")}`)}</div>
    <p class="nota">Valores en pesos, con impuestos (los artículos, antes de impuestos). Las órdenes en dólares o euros van convertidas. Toca un cliente, proveedor o artículo para ver sus órdenes.</p></div>`;
  vista.querySelectorAll("[data-per]").forEach(b => b.onclick = () => {
    periodo = b.dataset.per; pref("periodo", periodo); tablero(null, true);
  });
  activarGrafico(cub, R.cubeta, sumas, cuentas);
  pintarMapa(sel);
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
const nombreMuni = cod => { const m = cod && UBIC?.municipio(cod); return m ? m.nombre.replace(", D.C.", "") : "—"; };
const deptoMuni = cod => { const m = cod && UBIC?.municipio(cod); return m ? m.dpto.replace(", D.C.", " D.C.") : ""; };

function tarjetaMapa() {
  return `<div class="tarjeta mapa-caja" id="mapaCaja"><div class="esq" style="height:380px"></div></div>
    <div id="mapaSel"></div><div class="tarjeta bloque" id="mapaLista" style="margin-top:10px"></div>
    <a class="ver-mas tarjeta enlace-lugares" id="enlaceLugares" href="#/lugares" hidden></a>`;
}

async function pintarMapa(sel) {
  const caja = $("#mapaCaja"); if (!caja) return;
  try { await asegurarLugares(); } catch { caja.innerHTML = `<div class="vacio" style="padding:24px">No se pudo cargar el mapa. Revisa la conexión.</div>`; return; }
  if ($("#mapaCaja") !== caja) return;
  const grupos = new Map();
  for (const o of sel) {
    if (!o.mu) continue;
    const g = grupos.get(o.mu) || { cod: o.mu, v: 0, n: 0, est: 0, cl: new Map() };
    g.v += o.cop || 0; g.n++; if (o.muc === "cliente" || o.muc === "contrato") g.est++;
    g.cl.set(o.c, (g.cl.get(o.c) || 0) + (o.cop || 0)); grupos.set(o.mu, g);
  }
  const lista = [...grupos.values()].sort((a, b) => b.v - a.v);
  const sinUbicar = sel.filter(o => !o.mu).length;
  if (!lista.length) { caja.innerHTML = `<div class="vacio" style="padding:24px">Sin órdenes en este periodo.</div>`; $("#mapaLista").innerHTML = ""; return; }
  const pts = lista.map(g => ({ ...g, ...UBIC.municipio(g.cod) }));
  const max = pts[0].v || 1;

  // encuadre inicial: los puntos con un margen, con la proporción de la caja
  const ancho = caja.clientWidth || 360, alto = Math.round(Math.min(460, Math.max(300, ancho * 1.05)));
  let x0 = Math.min(...pts.map(p => p.x)), x1 = Math.max(...pts.map(p => p.x)), y0 = Math.min(...pts.map(p => p.y)), y1 = Math.max(...pts.map(p => p.y));
  const margen = 70; x0 -= margen; x1 += margen; y0 -= margen; y1 += margen;
  let w = Math.max(240, x1 - x0), h = Math.max(240, y1 - y0);
  const prop = ancho / alto;
  if (w / h > prop) h = w / prop; else w = h * prop;
  const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
  const inicial = { x: cx - w / 2, y: cy - h / 2, w, h };
  const vb = { ...inicial };

  caja.innerHTML = `<svg class="mapa" viewBox="${vb.x} ${vb.y} ${vb.w} ${vb.h}" style="height:${alto}px" role="img" aria-label="Mapa de Colombia con las órdenes por municipio">
      <g class="deptos">${MAPA.deptos.map(d => `<path d="${d.d}"><title>${esc(d.n)}</title></path>`).join("")}</g>
      <g class="burbujas"></g><g class="nombres"></g></svg>
    <div class="mapa-botones"><button class="redondo" data-z="in" aria-label="Acercar">+</button><button class="redondo" data-z="out" aria-label="Alejar">−</button><button class="redondo" data-z="fit" aria-label="Encuadrar">${ICO.mapaFit}</button></div>
    <div class="mapa-leyenda"><i></i><span>Tamaño = valor comprado</span></div>`;
  const svg = caja.querySelector("svg"), gB = svg.querySelector(".burbujas"), gN = svg.querySelector(".nombres");
  let elegido = null;
  const radio = p => vb.w * (0.012 + 0.055 * Math.sqrt(p.v / max));
  const dibujarPuntos = () => {
    // primero los grandes, para que los pequeños queden encima y se puedan tocar
    gB.innerHTML = [...pts].sort((a, b) => b.v - a.v).map(p => `<circle data-cod="${p.cod}" cx="${p.x}" cy="${p.y}" r="${radio(p)}" class="${p.cod === elegido ? "si" : ""}"><title>${esc(nombreMuni(p.cod))}: ${esc(corto(p.v))}</title></circle>`).join("");
    const fs = vb.w * 0.032;
    gN.innerHTML = pts.slice(0, 6).map(p => `<text x="${p.x}" y="${p.y - radio(p) - fs * 0.35}" font-size="${fs}">${esc(nombreMuni(p.cod))}</text>`).join("");
  };
  const aplicar = () => { svg.setAttribute("viewBox", `${vb.x} ${vb.y} ${vb.w} ${vb.h}`); dibujarPuntos(); };
  const elegir = cod => {
    elegido = cod; dibujarPuntos();
    const p = pts.find(x => x.cod === cod); if (!p) { $("#mapaSel").innerHTML = ""; return; }
    const top = [...p.cl].sort((a, b) => b[1] - a[1]).slice(0, 3);
    $("#mapaSel").innerHTML = `<div class="tarjeta mapa-sel entra"><div class="ms-tope"><div><b>${esc(nombreMuni(p.cod))}</b><small>${esc(deptoMuni(p.cod))}</small></div><button class="redondo" id="msCerrar" aria-label="Cerrar">✕</button></div>
      <div class="ms-cifras"><div><b>${esc(corto(p.v))}</b><small>comprado</small></div><div><b>${fmtNum.format(p.n)}</b><small>${p.n === 1 ? "orden" : "órdenes"}</small></div><div><b>${Math.round(p.v / lista.reduce((s, g) => s + g.v, 0) * 100)}%</b><small>del periodo</small></div></div>
      ${top.length ? `<div class="ms-clientes">${top.map(([c, v]) => `<span><i></i>${esc(c)} · ${esc(corto(v))}</span>`).join("")}</div>` : ""}
      ${p.est ? `<p class="nota" style="margin:8px 0 0">${plural(p.est, "orden ubicada", "órdenes ubicadas")} por estimación (el texto no decía el lugar).</p>` : ""}
      <a class="boton primario ancho" href="#/ordenes/mu:${p.cod}">Ver sus órdenes</a></div>`;
    $("#msCerrar").onclick = () => elegir(null);
  };

  // tocar, arrastrar, pellizcar y la rueda del ratón
  const punteros = new Map(); let movido = false;
  const enUnidades = (px, py) => { const r = svg.getBoundingClientRect(); return [vb.x + (px - r.left) / r.width * vb.w, vb.y + (py - r.top) / r.height * vb.h]; };
  const zoom = (f, cx, cy) => {
    const nw = Math.min(MAPA.ancho * 1.3, Math.max(30, vb.w * f)), k = nw / vb.w;
    vb.x = cx - (cx - vb.x) * k; vb.y = cy - (cy - vb.y) * k; vb.w = nw; vb.h *= k; aplicar();
  };
  // el círculo tocado se recuerda al apoyar el dedo (después el mapa captura el puntero)
  let tocado = null, recorrido = 0;
  svg.addEventListener("pointerdown", e => {
    if (!punteros.size) { tocado = e.target.closest("circle")?.dataset.cod ?? null; recorrido = 0; movido = false; }
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
      movido = true;
      const otro = [...punteros].find(([id]) => id !== e.pointerId)[1];
      const dAntes = Math.hypot(antes[0] - otro[0], antes[1] - otro[1]), dAhora = Math.hypot(e.clientX - otro[0], e.clientY - otro[1]);
      if (dAhora > 0) zoom(dAntes / dAhora, ...enUnidades((e.clientX + otro[0]) / 2, (e.clientY + otro[1]) / 2));
    }
  });
  svg.addEventListener("pointerup", e => {
    if (punteros.size === 1 && !movido && e.type === "pointerup") { elegir(tocado); if (tocado) vibrar(); } // un toque, no un arrastre
    punteros.delete(e.pointerId);
  });
  svg.addEventListener("pointercancel", e => punteros.delete(e.pointerId));
  svg.addEventListener("wheel", e => { e.preventDefault(); zoom(e.deltaY > 0 ? 1.18 : 1 / 1.18, ...enUnidades(e.clientX, e.clientY)); }, { passive: false });
  svg.addEventListener("dblclick", e => zoom(0.6, ...enUnidades(e.clientX, e.clientY)));
  caja.querySelectorAll("[data-z]").forEach(b => b.onclick = () => {
    if (b.dataset.z === "fit") Object.assign(vb, inicial), aplicar();
    else zoom(b.dataset.z === "in" ? 0.7 : 1 / 0.7, vb.x + vb.w / 2, vb.y + vb.h / 2);
  });
  aplicar();

  // la lista: el mismo contenido del mapa, para leerlo sin tocar círculos
  const maxL = lista[0].v || 1;
  $("#mapaLista").innerHTML = lista.slice(0, 8).map(g => `<a class="rank" href="#/ordenes/mu:${g.cod}"><div class="r1"><span>${esc(nombreMuni(g.cod))} <small class="dep">${esc(deptoMuni(g.cod))}</small></span><b>${esc(corto(g.v))}</b></div>
      <div class="r2"><i style="width:${Math.max(0.8, g.v / maxL * 100)}%"></i></div><small>${plural(g.n, "orden", "órdenes")}${g.est ? ` · ${g.est} estimada${g.est > 1 ? "s" : ""}` : ""}</small></a>`).join("") +
    (lista.length > 8 ? `<div class="rank resto"><div class="r1"><span>Otros ${fmtNum.format(lista.length - 8)} municipios</span><b>${esc(corto(lista.slice(8).reduce((s, g) => s + g.v, 0)))}</b></div></div>` : "");
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

function ajustes() {
  const tema = pref("tema") || "sistema";
  vista.innerHTML = `<div class="entra"><h1>Ajustes</h1>
    <div class="tarjeta">
      <div class="campo"><label>Cuenta</label><small>${cuenta ? esc(cuenta.name || "") + " · " + esc(cuenta.username) : DEMO ? "Modo de ejemplo" : "Sin iniciar sesión"}</small></div>
      <div class="campo"><label for="ruta">Carpeta de órdenes en tu OneDrive</label><input id="ruta" value="${esc(ruta())}" autocapitalize="characters"><small>La que contiene _AUTOMATIZACION. Normalmente «COMPRAS/ORDENES DE COMPRA».</small></div>
      <div class="campo"><label for="tema">Tema</label><select id="tema"><option value="sistema">Como el celular</option><option value="claro">Claro</option><option value="oscuro">Oscuro</option></select></div>
    </div>
    <div class="acciones">
      ${cuenta ? '<button class="boton" id="btnSalir">Cerrar sesión</button>' : ""}
      <button class="boton" id="btnBorrar">Borrar lo guardado en este celular</button>
    </div>
    ${instalable()}
    <p class="nota">Los datos se leen de tu OneDrive con tu cuenta y solo con permiso de lectura. Se actualizan cada vez que el PC (o el portátil) corre la herramienta.${D ? " Última generación: " + esc(new Date(D.generado).toLocaleString("es-CO")) + "." : ""}</p></div>`;
  $("#tema").value = tema;
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
const VISTAS = { inicio, ordenes, tablero, seguimiento, entregas, ajustes, lugares };
const PESTANAS = ["inicio", "ordenes", "tablero", "seguimiento", "entregas"];
const reducido = () => matchMedia("(prefers-reduced-motion: reduce)").matches;
const vibrar = (ms = 8) => { try { navigator.vibrate?.(ms); } catch { } };

function dibujar(arg) {
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
  if ("serviceWorker" in navigator && !DEMO) navigator.serviceWorker.register("sw.js").catch(() => { });
  vista.innerHTML = `<div class="entra"><div class="esq" style="height:38px;width:60%;margin:8px 0 18px"></div>
    <div class="cifras">${'<div class="esq" style="height:128px"></div>'.repeat(4)}</div>
    <div class="esq" style="height:170px;margin-top:10px"></div></div>`;

  if (DEMO) { await sincronizar(); alCambiar(); return; }
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
