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

// ---------------------------------------------------------------- utilidades
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const TILDES = new RegExp("[" + String.fromCharCode(0x300) + "-" + String.fromCharCode(0x36f) + "]", "g");
const normal = s => String(s ?? "").normalize("NFD").replace(TILDES, "").toLowerCase();
const fmtCop = new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
const fmtNum = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 2 });
const fmtFecha = new Intl.DateTimeFormat("es-CO", { day: "numeric", month: "short", year: "numeric" });
const fmtDia = new Intl.DateTimeFormat("es-CO", { weekday: "long", day: "numeric", month: "long" });
const fmtMes = new Intl.DateTimeFormat("es-CO", { month: "short" });

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
function chips(o) {
  const [t, c] = ESTADO[o.e] || ["—", "t-gris"];
  let h = `<span class="chip ${c}">${t}</span>`;
  if (o.venc) h += ` <span class="chip t-rojo">Entrega vencida</span>`;
  else if (o.r) h += ` <span class="chip t-verde">Recibida</span>`;
  return h;
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
};

function fila(o) {
  return `<a class="fila tarjeta" href="#/orden/${esc(o.n)}">${avatar(o.p)}
    <div class="txt"><div class="l1"><span>Orden ${esc(o.n)}</span><span class="monto">${esc(corto(o.cop ?? o.v))}</span></div>
    <div class="l2">${esc(o.p)}</div><div class="l3">${chips(o)}<span>${esc(fecha(o.f))}</span></div></div></a>`;
}

function inicio() {
  const hoy = hoyIso(), mes = hoy.slice(0, 7);
  const borr = D.ordenes.filter(o => o.e === 4), sc = D.ordenes.filter(o => o.sc), venc = D.ordenes.filter(o => o.venc);
  const delMes = D.ordenes.filter(o => (o.f || "").startsWith(mes));
  const hora = new Date().getHours();
  const saludo = hora < 12 ? "Buenos días" : hora < 19 ? "Buenas tardes" : "Buenas noches";
  const nombre = (cuenta?.name || "").split(" ")[0];
  // valor por mes, los últimos 6
  const meses = [];
  for (let i = 5; i >= 0; i--) { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - i); meses.push(d.toISOString().slice(0, 7)); }
  const sumas = meses.map(m => D.ordenes.filter(o => (o.f || "").startsWith(m)).reduce((s, o) => s + (o.cop || 0), 0));
  const max = Math.max(1, ...sumas);
  const resumen = borr.length ? `Tienes ${borr.length} borrador${borr.length > 1 ? "es" : ""} por enviar en Outlook.` :
    sc.length ? `${sc.length} órdenes esperan confirmación del proveedor.` : "Todo al día.";
  vista.innerHTML = `<div class="entra">
    <h1>${saludo}${nombre ? ", <span class='deg'>" + esc(nombre) + "</span>" : ""}</h1><p class="sub">${esc(resumen)}</p>
    <div class="cifras">
      <a class="cifra tarjeta" href="#/ordenes/borrador"><div class="ico t-ambar">${ICO.borrador}</div><b data-n="${borr.length}">0</b><span>Por enviar</span><small>borradores en Outlook</small></a>
      <a class="cifra tarjeta" href="#/seguimiento"><div class="ico t-azul">${ICO.reloj}</div><b data-n="${sc.length}">0</b><span>Sin confirmar</span><small>enviadas hace ${D.diasAlerta || 8}+ días</small></a>
      <a class="cifra tarjeta" href="#/entregas"><div class="ico t-rojo">${ICO.alerta}</div><b data-n="${venc.length}">0</b><span>Entregas vencidas</span><small>sin llegada · ${esc(corto(venc.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
      <a class="cifra tarjeta" href="#/ordenes/mes"><div class="ico t-verde">${ICO.mes}</div><b data-n="${delMes.length}">0</b><span>Órdenes del mes</span><small>${esc(corto(delMes.reduce((s, o) => s + (o.cop || 0), 0)))}</small></a>
    </div>
    <div class="tarjeta grafico" style="margin-top:10px"><div class="tope"><div><small class="sub">Comprado en ${esc(fmtMes.format(aFecha(hoy)).replace(".", ""))}</small><br><b>${esc(corto(sumas[5]))}</b></div><small class="sub">últimos 6 meses</small></div>
      <div class="barras">${sumas.map((s, i) => `<div><i style="height:${Math.max(3, s / max * 100)}%;animation-delay:${i * 60}ms"></i><small>${esc(fmtMes.format(aFecha(meses[i] + "-01")).replace(".", ""))}</small></div>`).join("")}</div></div>
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

function ordenes(arg) {
  if (arg && PRUEBA[arg]) filtro = arg;
  vista.innerHTML = `<div class="entra"><h1>Órdenes</h1>
    <div class="buscar"><div class="caja tarjeta">${ICO.buscar}<input id="q" type="search" placeholder="Número, proveedor, cliente o artículo" value="${esc(texto)}" autocomplete="off" enterkeyhint="search"></div>
    <div class="filtros">${FILTROS.map(([k, t]) => `<button data-f="${k}" class="${k === filtro ? "si" : ""}">${t}</button>`).join("")}</div></div>
    <div class="contador" id="cuenta"></div><div class="lista" id="res"></div><div id="mas"></div></div>`;
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
  const r = D.ordenes.filter(o => PRUEBA[filtro](o) && palabras.every(p => o.busca.includes(p)));
  const total = r.reduce((s, o) => s + (o.cop || 0), 0);
  $("#cuenta").textContent = `${fmtNum.format(r.length)} ${r.length === 1 ? "orden" : "órdenes"} · ${corto(total)}`;
  const res = $("#res"); res.innerHTML = "";
  if (!r.length) { res.innerHTML = `<div class="vacio"><div class="ilus">${ICO.buscar}</div><b>Ninguna orden coincide</b>Prueba con otro número, parte del proveedor o un artículo.</div>`; return; }
  let n = 0;
  const mas = () => { res.insertAdjacentHTML("beforeend", r.slice(n, n + 40).map(fila).join("")); n += 40; };
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
  const it = o.it || [];
  const datos = [["Cliente", o.c], ["Contrato", o.ct], ["Campo", o.ca], ["Solicitud (RQ)", o.rq], ["NIT", o.nit], ["Contacto", o.at],
    ["Correo", o.em ? `<a href="mailto:${esc(o.em)}?subject=${encodeURIComponent("Orden de compra N° " + o.n)}">${esc(o.em)}</a>` : ""],
    ["Forma de pago", o.fp], ["Tipo de compra", o.tc], ["Responsable", o.rs]]
    .filter(([, v]) => v).map(([k, v]) => `<div class="dato"><span>${k}</span><b>${k === "Correo" ? v : esc(v)}</b></div>`).join("");
  $("#hojaCuerpo").innerHTML = `
    <div class="cabeza">${avatar(o.p)}<div style="min-width:0"><h3>Orden ${esc(o.n)}</h3><p>${esc(o.p)}</p></div></div>
    <div style="margin-top:12px">${chips(o)}</div>
    <div class="total">${esc(dinero(o.v, o.m))}</div>
    ${o.m && o.m !== "COP" && o.cop ? `<small class="sub">≈ ${esc(dinero(o.cop))}</small>` : ""}
    <div class="acciones">
      <button class="boton primario" id="btnPdf">${ICO.pdf}Ver el PDF</button>
      ${o.em ? `<a class="boton" href="mailto:${esc(o.em)}?subject=${encodeURIComponent("Orden de compra N° " + o.n)}">${ICO.correo}Escribir</a>` : ""}
    </div>
    <h2>Recorrido</h2><div class="tarjeta recorrido">${pasos.join("")}</div>
    <h2>Artículos${it.length ? " · " + it.length : ""}</h2>
    <div class="tarjeta">${it.length ? it.map(a => `<div class="articulo"><div class="d">${esc(a[1] || "(sin descripción)")}</div>
        ${a[2] ? `<div class="c">${esc(a[2])}</div>` : ""}<div class="q">${esc(fmtNum.format(a[4] ?? 0))} ${esc(a[3] || "")} × ${esc(dinero(a[5], o.m))}${a[0] ? " · " + esc(a[0]) : ""}</div>
        <div class="t">${esc(dinero(a[6], o.m))}</div></div>`).join("") +
        `<div class="sumas"><span>Subtotal</span><span>${esc(dinero(o.st, o.m))}</span></div>` +
        (o.v && o.st && Math.abs(o.v - o.st) > 1 ? `<div class="sumas" style="border-top:0;padding-top:0"><span>Total con impuestos</span><span>${esc(dinero(o.v, o.m))}</span></div>` : "")
      : `<div class="vacio" style="padding:24px">No se pudieron leer los artículos de esta orden.</div>`}</div>
    <h2>La orden</h2><div class="tarjeta vista-previa" id="previa"><div class="esq" style="width:100%;height:220px"></div></div>
    <h2>Datos</h2><div class="tarjeta datos">${datos}</div>`;
  const h = $("#hoja"); h.hidden = false; h.classList.remove("saliendo");
  $("#hoja .panel").scrollTop = 0;
  $("#btnPdf").onclick = () => abrirPdf(o);
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
  const h = $("#hoja");
  if (h.hidden) return;
  if (sinAnimar) { h.hidden = true; return; }
  h.classList.add("saliendo");
  setTimeout(() => { h.hidden = true; h.classList.remove("saliendo"); }, 260);
}

// ---------------------------------------------------------------- navegación
const VISTAS = { inicio, ordenes, seguimiento, entregas, ajustes };

function dibujar(arg) {
  if (!D && actual !== "ajustes") return;
  VISTAS[actual](arg);
  dibujado = true;
  document.querySelectorAll("#pestanas a").forEach(a => a.classList.toggle("si", a.dataset.p === actual));
  const bi = $("#btnInstalar"); if (bi) bi.onclick = async () => { instalar.prompt(); await instalar.userChoice; instalar = null; dibujar(); };
}

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
  actual = nueva;
  dibujar(arg);
  window.scrollTo({ top: 0 });
}

document.addEventListener("click", e => {
  if (e.target.closest("[data-cerrar]")) history.length > 1 ? history.back() : (location.hash = "#/" + actual);
  if (e.target.closest("[data-cerrar-visor]")) { $("#visor").hidden = true; return; }
  if (e.target.closest("#visor img")) $("#visor").classList.toggle("zoom");
});
$("#btnRecargar").onclick = () => sincronizar(true);
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
