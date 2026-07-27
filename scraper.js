const cheerio = require("cheerio");
const ExcelJS = require("exceljs");
const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const pLimit = require("p-limit").default;
const Database = require("better-sqlite3");

// ─── CARGAR CONFIGURACIÓN DESDE config.json ───────────────────────
const configPath = path.join(__dirname, "config.json");
if (!fs.existsSync(configPath)) {
  console.error(" No se encontró config.json. Crea el archivo antes de ejecutar.");
  process.exit(1);
}
const CONFIG = JSON.parse(fs.readFileSync(configPath, "utf-8"));
// ─────────────────────────────────────────────────────────────────

// ─── SQLITE ────────────────────────────────────────────────────────
const DB_PATH = path.join(__dirname, "contactos.db");
const db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.exec(`
  CREATE TABLE IF NOT EXISTS negocios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT DEFAULT '',
    categoria TEXT DEFAULT '',
    valoracion TEXT DEFAULT '',
    telefono_maps TEXT DEFAULT '',
    telefono_web TEXT DEFAULT '',
    correo TEXT DEFAULT '',
    whatsapp TEXT DEFAULT '',
    instagram TEXT DEFAULT '',
    facebook TEXT DEFAULT '',
    tiktok TEXT DEFAULT '',
    direccion TEXT DEFAULT '',
    web TEXT DEFAULT '',
    url_maps TEXT DEFAULT '',
    busqueda TEXT DEFAULT '',
    metodo TEXT DEFAULT '',
    via TEXT DEFAULT '',
    estado TEXT DEFAULT '',
    creado_en DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

const insertNegocio = db.prepare(`
  INSERT INTO negocios (
    nombre, categoria, valoracion, telefono_maps, telefono_web,
    correo, whatsapp, instagram, facebook, tiktok,
    direccion, web, url_maps, busqueda, metodo, via, estado
  ) VALUES (
    @nombre, @categoria, @valoracion, @telefono_maps, @telefono_web,
    @correo, @whatsapp, @instagram, @facebook, @tiktok,
    @direccion, @web, @url_maps, @busqueda, @metodo, @via, @estado
  )
`);

function guardarNegocio(datos) {
  insertNegocio.run({
    nombre: datos.Nombre || '',
    categoria: datos.Categoría || '',
    valoracion: datos.Valoración || '',
    telefono_maps: datos['Teléfono Maps'] || '',
    telefono_web: datos['Teléfono Web'] || '',
    correo: datos.Correo || '',
    whatsapp: datos.WhatsApp || '',
    instagram: datos.Instagram || '',
    facebook: datos.Facebook || '',
    tiktok: datos.TikTok || '',
    direccion: datos.Dirección || '',
    web: datos.Web || '',
    url_maps: datos.URLMaps || '',
    busqueda: datos.Búsqueda || '',
    metodo: datos.Método || '',
    via: datos.Vía || '',
    estado: datos.Estado || '',
  });
}

function contarNegocios() {
  const row = db.prepare('SELECT COUNT(*) as count FROM negocios').get();
  return row.count;
}
// ─────────────────────────────────────────────────────────────────

// Dominios de herramientas de desarrollo / tracking que no son correos reales
const DOMINIOS_BASURA =
  /sentry\.io|example\.com|amazonaws\.com|cloudfront\.net|w3\.org|schema\.org|hotjar\.com|klaviyo\.com|googleapis\.com|gstatic\.com|jquery\.com|bootstrapcdn\.com/i;

// Extensiones de archivo para filtrar falsos correos
const EXTENSIONES_NO_CORREO =
  /\.(webp|png|jpg|jpeg|gif|svg|mp4|mp3|pdf|zip|ico|woff|woff2|ttf)$/i;

// Regex de emojis para limpiar texto antes de guardar en Excel
const REGEX_EMOJI =
  /[\u{1F000}-\u{1FFFF}]|[\u{2600}-\u{27FF}]|[\u{2B00}-\u{2BFF}]|[\u{FE00}-\u{FEFF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA00}-\u{1FA9F}]|[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]/gu;

// Regex para teléfonos peruanos y correos
const REGEX = {
  telefono: /(\+?51[\s\-]?)?(9\d{8}|\d{7,8})\b/g,
  correo: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g,
};

// ─── GENERADOR DE BÚSQUEDAS ───────────────────────────────────────

function generarTerminosDeBusqueda() {
  const terminos = [];
  for (const categoria of CONFIG.CATEGORIAS) {
    for (const distrito of CONFIG.DISTRITOS) {
      terminos.push({ termino: `${categoria} ${distrito}`, categoria });
    }
  }
  return terminos;
}

// ─── UTILIDADES ───────────────────────────────────────────────────

function esperarMs(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function esperarAleatorio(minMs, maxMs) {
  const ms = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
  return new Promise((r) => setTimeout(r, ms));
}

function limpiarTexto(texto) {
  if (!texto || texto === "—") return texto;
  return (
    texto
      .replace(REGEX_EMOJI, "")
      .replace(/[\u00AD\u200B\u200C\u200D\uFEFF]/g, "")
      .replace(/\s+/g, " ")
      .trim() || "—"
  );
}

function normalizarTelefonoPe(digitos) {
  if (/^51[9][0-9]{8}$/.test(digitos)) return digitos.slice(2);
  if (/^51[0-9]{7,8}$/.test(digitos)) return digitos.slice(2);
  if (/^0[1-9][0-9]{6,8}$/.test(digitos)) return digitos.slice(1);
  return digitos;
}

function esTelefonoValido(t) {
  const digitos = t.replace(/\D/g, "");
  if (digitos.length < 7 || digitos.length > 15) return false;
  if (digitos.length === 11 && /^(10|20)/.test(digitos)) return false;
  if (/^20[0-9]{6}$/.test(digitos)) return false;
  return true;
}

function limpiarTelefonos(texto) {
  const matches = texto.match(REGEX.telefono) || [];
  const validos = matches
    .map((t) => normalizarTelefonoPe(t.replace(/\D/g, "")))
    .filter((d) => esTelefonoValido(d));
  return [...new Set(validos)].slice(0, 3).join(" | ");
}

function extraerNumeroDeWhatsApp(href) {
  try {
    const parsed = new URL(href);
    const porPath = parsed.pathname.match(/\/([\d]+)/);
    const porParam = parsed.searchParams.get("phone");
    return (porPath?.[1] || porParam || "").replace(/\D/g, "");
  } catch (_) {
    return "";
  }
}

function esUrlWhatsApp(href) {
  return /^https?:\/\/(wa\.me|(api\.|chat\.)?whatsapp\.com|wa\.link)\b|^\/\/(wa\.me|(api\.|chat\.)?whatsapp\.com|wa\.link)\b|^whatsapp:\/\//i.test(href);
}

function buscarUrlContacto($, urlBase) {
  try {
    const base = new URL(urlBase);
    const origen = base.origin;
    let encontrada = null;

    $("a[href]").each((_, el) => {
      if (encontrada) return;

      const href = ($(el).attr("href") || "").trim();
      if (
        !href ||
        href.toLowerCase().startsWith("javascript") ||
        href.toLowerCase().startsWith("mailto") ||
        href.toLowerCase().startsWith("tel")
      )
        return;

      const contieneContacto = CONFIG.palabrasContacto.some((p) =>
        href.toLowerCase().includes(p)
      );
      if (!contieneContacto) return;

      try {
        const urlObj = new URL(href, urlBase);
        if (urlObj.origin === origen) encontrada = urlObj.href;
      } catch (_) {}
    });

    return encontrada;
  } catch (_) {
    return null;
  }
}

function combinarDatosContacto(datosPrincipal, datosContacto) {
  const combinar = (a, b) => {
    if (!a || a === "—") return b || "—";
    if (!b || b === "—") return a;
    const set = new Set([...a.split(" | "), ...b.split(" | ")]);
    return [...set].slice(0, 5).join(" | ");
  };

  return {
    telefonoWeb: combinar(datosPrincipal.telefonoWeb, datosContacto.telefonoWeb),
    correo: combinar(datosPrincipal.correo, datosContacto.correo),
    whatsapp: combinar(datosPrincipal.whatsapp, datosContacto.whatsapp),
    instagram: combinar(datosPrincipal.instagram, datosContacto.instagram),
    facebook: combinar(datosPrincipal.facebook, datosContacto.facebook),
    tiktok: combinar(datosPrincipal.tiktok, datosContacto.tiktok),
  };
}

function extraerTelefonosWeb($, textoVisible) {
  const telefonosEncontrados = new Set();

  $("a[href^='tel:']").each((_, el) => {
    const href = $(el).attr("href") || "";
    const tel = href.replace(/^tel:/i, "").trim();
    const digitos = tel.replace(/\D/g, "");
    if (digitos.length >= 7) telefonosEncontrados.add(normalizarTelefonoPe(digitos));
  });

  $("a[href]").each((_, el) => {
    const href = ($(el).attr("href") || "").trim();
    if (!esUrlWhatsApp(href)) return;
    const numero = extraerNumeroDeWhatsApp(href);
    if (numero.length >= 7) telefonosEncontrados.add(normalizarTelefonoPe(numero));
  });

  const matches = textoVisible.match(REGEX.telefono) || [];
  matches.forEach((t) => {
    const digitos = normalizarTelefonoPe(t.replace(/\D/g, ""));
    telefonosEncontrados.add(digitos);
  });

  const validos = [...telefonosEncontrados].filter(esTelefonoValido);
  return [...new Set(validos)].slice(0, 3).join(" | ");
}

function extraerCorreos($, textoVisible) {
  const correos = new Set();
  const textoSanitizado = textoVisible
    .replace(/[\u00AD\u200B\u200C\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ");

  $("a[href^='mailto:']").each((_, el) => {
    const href = $(el).attr("href") || "";
    const correo = href.replace(/^mailto:/i, "").split("?")[0].trim().toLowerCase();
    if (correo && !EXTENSIONES_NO_CORREO.test(correo) && !DOMINIOS_BASURA.test(correo)) {
      correos.add(correo);
    }
  });

  $("a").each((_, el) => {
    const texto = $(el).text().trim().toLowerCase();
    const matches = texto.match(REGEX.correo) || [];
    for (const m of matches) {
      if (!EXTENSIONES_NO_CORREO.test(m) && !DOMINIOS_BASURA.test(m)) correos.add(m);
    }
  });

  const matches = textoSanitizado.match(REGEX.correo) || [];
  for (const m of matches) {
    const correo = m.toLowerCase();
    if (!EXTENSIONES_NO_CORREO.test(correo) && !DOMINIOS_BASURA.test(correo)) {
      correos.add(correo);
    }
  }

  return [...correos].slice(0, 5).join(" | ");
}

function extraerRedesSociales($) {
  const redes = {
    WhatsApp: new Set(),
    Instagram: new Set(),
    Facebook: new Set(),
    TikTok: new Set(),
  };

  $("a[href]").each((_, el) => {
    const href = ($(el).attr("href") || "").trim();
    if (!href || href.toLowerCase().startsWith("javascript")) return;

    if (esUrlWhatsApp(href)) {
      try {
        const parsed = new URL(href);
        if (/wa\.link/i.test(href)) {
          redes.WhatsApp.add(href.split("?")[0]);
          return;
        }
        if (parsed.protocol === 'whatsapp:' && !parsed.searchParams.get('phone') && !parsed.pathname.match(/\/([\d]+)/)) return;
        const porPath = parsed.pathname.match(/\/([\d]+)/);
        const porParam = parsed.searchParams.get("phone");
        const numero = (porPath?.[1] || porParam || "").replace(/\D/g, "");
        if (numero.length >= 7) {
          redes.WhatsApp.add(`https://wa.me/${numero}`);
        } else {
          parsed.searchParams.delete("text");
          redes.WhatsApp.add(parsed.toString());
        }
      } catch (_) {
        redes.WhatsApp.add(href);
      }
    } else if (/instagram\.com\//i.test(href)) {
      const m = href.match(/instagram\.com\/([a-zA-Z0-9._]{2,40})/i);
      if (m) redes.Instagram.add(`instagram.com/${m[1]}`);
        } else if (/facebook\.com\//i.test(href) && !/sharer|login|recover|\/photo(?:\/|\.php)|profile\.php|\/dialog\/|l\.php|\/followers?\/|\/following\/|\/messages(?:\/|$)|\/events(?:\/|$)|\/about/i.test(href) && !/facebook\.com\/facebook(?:\/|$)/i.test(href)) {
      const m = href.match(/facebook\.com\/([^\s"'<>?#]+)/i);
      if (m) redes.Facebook.add(`facebook.com/${m[1]}`);
    } else if (/tiktok\.com\/@/i.test(href)) {
      const m = href.match(/tiktok\.com\/@([^\s"'<>?/]+)/i);
      if (m) redes.TikTok.add(`tiktok.com/@${m[1]}`);
    }
  });

  return {
    WhatsApp: [...redes.WhatsApp].slice(0, 3).join(" | "),
    Instagram: [...redes.Instagram].slice(0, 3).join(" | "),
    Facebook: [...redes.Facebook].slice(0, 3).join(" | "),
    TikTok: [...redes.TikTok].slice(0, 3).join(" | "),
  };
}

function textoLimpio($) {
  const bloques =
    "p,div,li,td,th,h1,h2,h3,h4,h5,h6,br,tr,section,article,header,footer,nav,aside";
  $(bloques).each((_, el) => $(el).append(" "));
  return $("body").text().replace(/\s+/g, " ").trim();
}

function extraerDatosDeHtml(html) {
  const $ = cheerio.load(html);
  $("script,style,noscript").remove();
  const textoVisible = textoLimpio($);
  const redes = extraerRedesSociales($);

  return {
    $,
    telefonoWeb: extraerTelefonosWeb($, textoVisible),
    correo: extraerCorreos($, textoVisible),
    whatsapp: redes.WhatsApp,
    instagram: redes.Instagram,
    facebook: redes.Facebook,
    tiktok: redes.TikTok,
  };
}

// ─── VISITAR WEB DEL NEGOCIO ──────────────────────────────────────

async function visitarWebConPuppeteer(url, browser) {
  const page = await browser.newPage();
  await page.setUserAgent(
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
  );
  await page.setRequestInterception(true);
  page.on("request", (req) => {
    ["image", "stylesheet", "font", "media"].includes(req.resourceType())
      ? req.abort()
      : req.continue();
  });

  let html = "";
  try {
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 20000 });
    await esperarAleatorio(800, 1500);
  } catch (e) {
    console.warn(`     Timeout o error cargando ${url}, procesando lo que haya...`);
  } finally {
    html = await page.content();
    await page.close();
  }
  return html;
}

async function visitarWebConFetch(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 12000);

  const res = await fetch(url, {
    signal: controller.signal,
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    },
  });

  clearTimeout(timer);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const html = await res.text();
  const $chk = cheerio.load(html);
  $chk("script,style").remove();

  if ($chk("body").text().trim().length < CONFIG.umbralTextoUtil) {
    throw new Error("página vacía o renderizada por JS");
  }

  return html;
}

async function visitarUrl(url, browser) {
  try {
    const html = await visitarWebConFetch(url);
    return { html, via: "fetch" };
  } catch (errFetch) {
    console.warn(`   Fetch falló (${errFetch.message}), usando Puppeteer...`);
  }

  try {
    const html = await visitarWebConPuppeteer(url, browser);
    return { html, via: "puppeteer" };
  } catch (errPuppeteer) {
    throw new Error(`puppeteer: ${errPuppeteer.message}`);
  }
}

// ─── POOL DE PÁGINAS ─────────────────────────────────────────────

async function crearPoolPaginas(browser, n) {
  const paginas = [];
  for (let i = 0; i < n; i++) {
    const p = await browser.newPage();
    await p.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
    );
    await p.setRequestInterception(true);
    p.on("request", (req) => {
      ["image", "stylesheet", "font", "media"].includes(req.resourceType())
        ? req.abort()
        : req.continue();
    });
    paginas.push(p);
  }
  return paginas;
}

function crearGestorPool(paginas) {
  const libres = [...paginas];
  const esperando = [];

  function obtener() {
    if (libres.length > 0) return Promise.resolve(libres.pop());
    return new Promise((resolve) => esperando.push(resolve));
  }

  function liberar(pagina) {
    const siguiente = esperando.shift();
    if (siguiente) siguiente(pagina);
    else libres.push(pagina);
  }

  return { obtener, liberar };
}

async function extraerFichaNegocio(pagina, enlace) {
  await pagina.goto(enlace, { waitUntil: "domcontentloaded", timeout: 15000 });
  await pagina.waitForSelector("h1", { timeout: 6000 }).catch(() => {});

  const datos = await pagina.evaluate(() => {
    const txt = (sel) => document.querySelector(sel)?.textContent?.trim() || "";

    const nombre = txt("h1");

    let telefono = "";
    document.querySelectorAll('button[data-item-id^="phone"]').forEach((btn) => {
      telefono =
        btn.getAttribute("aria-label")?.replace(/^Teléfono:\s*/i, "") || "";
    });
    if (!telefono) {
      document.querySelectorAll("button[aria-label]").forEach((btn) => {
        const label = btn.getAttribute("aria-label") || "";
        if (/^\+?[\d\s\-().]{7,}$/.test(label.trim())) telefono = label.trim();
      });
    }

    let direccion = "";
    document.querySelectorAll('button[data-item-id="address"]').forEach((btn) => {
      direccion =
        btn.getAttribute("aria-label")?.replace(/^Dirección:\s*/i, "") || "";
    });

    let web = "";
    document.querySelectorAll('a[data-item-id="authority"]').forEach((a) => {
      web = a.href || "";
    });
    if (!web) {
      document.querySelectorAll("a[aria-label]").forEach((a) => {
        if (/sitio web/i.test(a.getAttribute("aria-label") || "")) web = a.href;
      });
    }

    const categoria =
      txt('button[jsaction*="category"]') ||
      document.querySelector(".DkEaL")?.textContent?.trim() ||
      "";

    const valoracion =
      txt(".F7nice span") || txt('[aria-label*="estrellas"]') || "";

    return { nombre, telefono, direccion, web, categoria, valoracion };
  });

  if (datos.nombre) return { ...datos, urlMaps: enlace };
  return null;
}

// ─── GOOGLE MAPS: SCRAPING PRINCIPAL ─────────────────────────────

async function detectarBloqueo(page) {
  const bloqueado = await page.evaluate(() =>
    document.body.innerText.includes("unusual traffic") ||
    document.body.innerText.includes("tráfico inusual") ||
    document.body.innerText.includes("tráfico anormal") ||
    !!document.querySelector('form[action*="sorry"]')
  );
  return bloqueado;
}

async function buscarEnMaps(termino, browser) {
  const page = await browser.newPage();
  await page.setUserAgent(
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
  );
  await page.setRequestInterception(true);
  page.on("request", (req) => {
    ["image", "stylesheet", "font", "media"].includes(req.resourceType())
      ? req.abort()
      : req.continue();
  });

  try {
    const url = `https://www.google.com/maps/search/${encodeURIComponent(termino)}`;
    console.log(`   Abriendo Maps: ${url}`);
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });

    if (await detectarBloqueo(page)) {
      console.warn(`   Google bloqueó la búsqueda (tráfico inusual). Esperando 60s...`);
      await esperarMs(60000);
      await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
      if (await detectarBloqueo(page)) {
        throw new Error("Google bloqueó la búsqueda después de reintentar");
      }
    }

    await page.waitForSelector('[role="feed"]', { timeout: 15000 });

    let prevCount = 0;
    let sameCount = 0;
    const maxIteraciones = 20;

    for (let s = 0; s < maxIteraciones; s++) {
      const enlacesActuales = await page.$$eval('a[href*="/maps/place/"]', (els) =>
        [...new Set(els.map((a) => a.href).filter((h) => h.includes("/maps/place/")))]
      );

      if (enlacesActuales.length === prevCount) sameCount++;
      else sameCount = 0;
      prevCount = enlacesActuales.length;

      if (enlacesActuales.length >= CONFIG.maxResultadosPorBusqueda) break;
      if (sameCount >= 3) break;

      await page.evaluate(() => {
        const panel = document.querySelector('[role="feed"]');
        if (panel) panel.scrollBy(0, 2000);
      });
      await esperarAleatorio(1000, 1800);
    }

    const enlaces = await page.$$eval('a[href*="/maps/place/"]', (els) =>
      [...new Set(els.map((a) => a.href).filter((h) => h.includes("/maps/place/")))]
    );

    console.log(`   Encontrados ${enlaces.length} negocios en "${termino}"`);
    return enlaces;
  } finally {
    await page.close();
  }
}

// ─── PROCESAR FICHAS EN PARALELO ────────────────────────────────

class BlockError extends Error {
  constructor(message) {
    super(message);
    this.name = 'BlockError';
  }
}

async function procesarFichasEnParalelo(enlaces, gestor, concurrenciaFichas) {
  const limite = Math.min(enlaces.length, CONFIG.maxResultadosPorBusqueda);
  if (limite === 0) return [];

  const limitFichas = pLimit(concurrenciaFichas);
  let totalAttempted = 0;
  let nullCount = 0;

  console.log(`   Extrayendo ${limite} fichas (${concurrenciaFichas} en paralelo)...`);

  const resultados = await Promise.all(
    enlaces.slice(0, limite).map((enlace) =>
      limitFichas(async () => {
        const pagina = await gestor.obtener();
        totalAttempted++;
        try {
          const datos = await extraerFichaNegocio(pagina, enlace);
          await esperarAleatorio(300, 700);
          if (!datos) nullCount++;
          return datos;
        } catch (err) {
          console.warn(`   Error en ficha: ${err.message}`);
          nullCount++;
          return null;
        } finally {
          gestor.liberar(pagina);
        }
      })
    )
  );

  const validos = resultados.filter(Boolean);
  const nullRatio = totalAttempted > 0 ? nullCount / totalAttempted : 0;

  if (nullRatio > 0.6) {
    console.warn(`   Alta tasa de fichas vacías (${(nullRatio * 100).toFixed(0)}%). Posible bloqueo de Google.`);
    throw new BlockError(`Posible bloqueo: ${(nullRatio * 100).toFixed(0)}% de fichas vacías`);
  }

  return validos;
}

// ─── PROCESAR UN NEGOCIO ──────────────────────────────────────────

async function procesarNegocio(negocio, browser, terminoBusqueda, categoriaUsuario) {
  let datos = {
    telefonoWeb: "",
    correo: "",
    whatsapp: "",
    instagram: "",
    facebook: "",
    tiktok: "",
  };
  let metodo = "maps";
  let via = "—";

  // Extraer redes desde la URL de Maps si la web es una red social
  if (negocio.web) {
    const web = negocio.web;
    if (/facebook\.com\//i.test(web)) {
      const m = web.match(/facebook\.com\/([^\s"'<>?#]+)/i);
      if (m) datos.facebook = `facebook.com/${m[1]}`;
    } else if (/instagram\.com\//i.test(web)) {
      const m = web.match(/instagram\.com\/([a-zA-Z0-9._]{2,40})/i);
      if (m) datos.instagram = `instagram.com/${m[1]}`;
    } else if (/tiktok\.com\/@/i.test(web)) {
      const m = web.match(/tiktok\.com\/@([^\s"'<>?/]+)/i);
      if (m) datos.tiktok = `tiktok.com/@${m[1]}`;
    } else if (esUrlWhatsApp(web)) {
      datos.whatsapp = web;
    }
  }

  // Saltar webs que son redes sociales (no se pueden scrapear)
  const esRedSocial = /facebook\.com|instagram\.com|tiktok\.com|wa\.me|wa\.link/i.test(negocio.web);

  if (CONFIG.visitarWebDelNegocio && negocio.web && !esRedSocial) {
    try {
      const resultadoPrincipal = await visitarUrl(negocio.web, browser);
      const datosPrincipal = extraerDatosDeHtml(resultadoPrincipal.html);
      via = resultadoPrincipal.via;
      metodo = "maps+web";

      const urlContacto = buscarUrlContacto(datosPrincipal.$, negocio.web);

      if (urlContacto && urlContacto !== negocio.web) {
        try {
          console.log(`   Visitando subpagina contacto: ${urlContacto}`);
          const resultadoContacto = await visitarUrl(urlContacto, browser);
          const datosContacto = extraerDatosDeHtml(resultadoContacto.html);
          datos = combinarDatosContacto(datosPrincipal, datosContacto);
        } catch (_) {
          datos = datosPrincipal;
        }
      } else {
        datos = datosPrincipal;
      }
    } catch (err) {
      console.warn(`   No se pudo visitar web (${negocio.web}): ${err.message}`);
    }
  }

  const telefonoMaps = negocio.telefono
    ? limpiarTelefonos(negocio.telefono) || negocio.telefono
    : "";

  return {
    Nombre: limpiarTexto(negocio.nombre) || "—",
    Categoría: categoriaUsuario || limpiarTexto(negocio.categoria) || "—",
    Valoración: negocio.valoracion || "—",
    "Teléfono Maps": telefonoMaps || "—",
    "Teléfono Web": datos.telefonoWeb || "—",
    Correo: datos.correo || "—",
    WhatsApp: datos.whatsapp || "—",
    Instagram: datos.instagram || "—",
    Facebook: datos.facebook || "—",
    TikTok: datos.tiktok || "—",
    Dirección: limpiarTexto(negocio.direccion) || "—",
    Web: negocio.web || "—",
    URLMaps: negocio.urlMaps || "—",
    Búsqueda: terminoBusqueda,
    Método: metodo,
    Vía: via,
    Estado: "OK",
  };
}

// ─── WRAPPER SEGURO PARA CONCURRENCIA ─────────────────────────────

async function procesarNegocioConError(negocio, browser, termino, categoria) {
  try {
    return await procesarNegocio(negocio, browser, termino, categoria);
  } catch (err) {
    console.error(`   Error inesperado: ${err.message}`);
    return {
      Nombre: negocio.nombre || "ERROR",
      Categoría: negocio.categoria || "",
      Valoración: "",
      "Teléfono Maps": negocio.telefono ? limpiarTelefonos(negocio.telefono) || negocio.telefono : "",
      "Teléfono Web": "",
      Correo: "",
      WhatsApp: "",
      Instagram: "",
      Facebook: "",
      TikTok: "",
      Dirección: negocio.direccion || "",
      Web: negocio.web || "",
      URLMaps: negocio.urlMaps || "",
      Búsqueda: termino,
      Método: "maps",
      Vía: "—",
      Estado: err.message.slice(0, 100),
    };
  }
}

// ─── GUARDAR EXCEL ────────────────────────────────────────────────

async function guardarExcel(datos, ruta) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Contactos");

  ws.columns = [
    { header: "Nombre", key: "Nombre", width: 35 },
    { header: "Categoría", key: "Categoría", width: 25 },
    { header: "Valoración", key: "Valoración", width: 12 },
    { header: "Teléfono Maps", key: "Teléfono Maps", width: 22 },
    { header: "Teléfono Web", key: "Teléfono Web", width: 22 },
    { header: "Correo", key: "Correo", width: 35 },
    { header: "WhatsApp", key: "WhatsApp", width: 35 },
    { header: "Instagram", key: "Instagram", width: 30 },
    { header: "Facebook", key: "Facebook", width: 35 },
    { header: "TikTok", key: "TikTok", width: 28 },
    { header: "Dirección", key: "Dirección", width: 40 },
    { header: "Web", key: "Web", width: 40 },
    { header: "URL Maps", key: "URLMaps", width: 45 },
    { header: "Búsqueda", key: "Búsqueda", width: 35 },
    { header: "Estado", key: "Estado", width: 14 },
  ];

  ws.getRow(1).eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
  ws.getRow(1).height = 22;

  datos.forEach((r, i) => {
    const row = ws.addRow(r);

    if (i % 2 === 0) {
      row.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFE8F0FE" },
        };
      });
    }

    if (r.Estado !== "OK") {
      row.getCell("Estado").font = { color: { argb: "FFCC0000" }, bold: true };
    }

    row.alignment = { vertical: "middle" };
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 15 } };
  ws.views = [{ state: "frozen", ySplit: 1 }];

  await wb.xlsx.writeFile(ruta);
}

// ─── REANUDACIÓN ─────────────────────────────────────────────────

function terminoYaProcesado(termino) {
  const row = db.prepare(
    `SELECT COUNT(*) as count FROM negocios WHERE busqueda = ? AND nombre != 'ERROR-BUSQUEDA'`
  ).get(termino);
  return row.count > 0;
}

// ─── MAIN ─────────────────────────────────────────────────────────

async function main() {
  const terminos = generarTerminosDeBusqueda();
  const totalCombinaciones = CONFIG.CATEGORIAS.length * CONFIG.DISTRITOS.length;
  const concurrencia = CONFIG.concurrencia || 5;
  const concurrenciaFichas = CONFIG.concurrenciaFichas || 4;

  console.log(`\n Maps Extractor v2`);
  console.log(`   Categorias    : ${CONFIG.CATEGORIAS.length}`);
  console.log(`   Distritos     : ${CONFIG.DISTRITOS.length}`);
  console.log(`   Combinaciones : ${totalCombinaciones}`);
  console.log(`   Max. negocios por busqueda : ${CONFIG.maxResultadosPorBusqueda}`);
  console.log(`   Concurrencia fichas: ${concurrenciaFichas}`);
  console.log(`   Concurrencia webs: ${concurrencia}`);
  console.log(`${"─".repeat(50)}\n`);

  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--lang=es-PE"],
  });

  const pool = await crearPoolPaginas(browser, concurrenciaFichas);
  const gestor = crearGestorPool(pool);
  const limit = pLimit(concurrencia);

  try {
    for (let i = 0; i < terminos.length; i++) {
      const { termino, categoria: catUsuario } = terminos[i];
      console.log(`\n[Busqueda ${i + 1}/${terminos.length}] "${termino}"`);

      // Reanudación: saltar si ya fue procesado
      if (terminoYaProcesado(termino)) {
        console.log(`   Ya procesado anteriormente, saltando...`);
        continue;
      }

      let enlaces = [];
      try {
        enlaces = await buscarEnMaps(termino, browser);
      } catch (err) {
        console.error(` Error al buscar en Maps: ${err.message}`);
        guardarNegocio({
          Nombre: "ERROR-BUSQUEDA",
          Categoría: catUsuario,
          Búsqueda: termino,
          Estado: err.message.slice(0, 100),
        });
        continue;
      }

      let negocios = [];
      let fichasOk = false;
      for (let intento = 0; intento < 3; intento++) {
        try {
          negocios = await procesarFichasEnParalelo(enlaces, gestor, concurrenciaFichas);
          fichasOk = true;
          break;
        } catch (err) {
          if (err.name !== 'BlockError') throw err;
          if (intento < 2) {
            console.warn(`   Reintentando tras posible bloqueo (intento ${intento + 1}/3) en 60s...`);
            await esperarMs(60000);
          }
        }
      }
      if (!fichasOk) {
        console.error(`   Bloqueo persistente en fichas. Saltando término.`);
        guardarNegocio({
          Nombre: "ERROR-BUSQUEDA",
          Categoría: catUsuario,
          Búsqueda: termino,
          Estado: "Bloqueo persistente en extracción de fichas",
        });
        continue;
      }

      console.log(`   Extraidas ${negocios.length} fichas. Enriquiciendo (concurrencia: ${concurrencia})...`);

      const promises = negocios.map((negocio) =>
        limit(() => procesarNegocioConError(negocio, browser, termino, catUsuario))
      );

      const resultados = await Promise.all(promises);

      for (const datos of resultados) {
        guardarNegocio(datos);
      }

      const total = contarNegocios();
      console.log(`   Almacenados: ${resultados.length} registros  Total en BD: ${total}`);

      if (i < terminos.length - 1) {
        await esperarAleatorio(
          CONFIG.esperaMsEntreBusquedas,
          CONFIG.esperaMsEntreBusquedas + 2000
        );
      }
    }
  } finally {
    for (const p of pool) await p.close().catch(() => {});
    await browser.close();
  }

  // ─── EXPORTAR A EXCEL DESDE SQLITE ─────────────────────────────

  const todos = db.prepare(
    `SELECT nombre, categoria, valoracion, telefono_maps, telefono_web,
            correo, whatsapp, instagram, facebook, tiktok,
            direccion, web, url_maps, busqueda, estado
     FROM negocios ORDER BY id`
  ).all();

  const exportData = todos.map((r) => ({
    Nombre: r.nombre || "—",
    Categoría: r.categoria || "—",
    Valoración: r.valoracion || "—",
    "Teléfono Maps": r.telefono_maps || "—",
    "Teléfono Web": r.telefono_web || "—",
    Correo: r.correo || "—",
    WhatsApp: r.whatsapp || "—",
    Instagram: r.instagram || "—",
    Facebook: r.facebook || "—",
    TikTok: r.tiktok || "—",
    Dirección: r.direccion || "—",
    Web: r.web || "—",
    URLMaps: r.url_maps || "—",
    Búsqueda: r.busqueda || "—",
    Estado: r.estado || "—",
  }));

  await guardarExcel(exportData, CONFIG.archivoExcel);

  const ok = todos.filter((r) => r.estado === "OK").length;
  const err = todos.filter((r) => r.estado !== "OK").length;

  console.log(`\n${"═".repeat(50)}`);
  console.log(` Proceso completo`);
  console.log(`   Busquedas     : ${terminos.length}`);
  console.log(`   Negocios      : ${todos.length}`);
  console.log(`   Exitosos      : ${ok}`);
  console.log(`   Con error     : ${err}`);
  console.log(`   Archivo       : ${CONFIG.archivoExcel}`);
  console.log(`   BD SQLite     : ${DB_PATH}`);
  console.log(`${"═".repeat(50)}\n`);

  db.close();
}

main().catch(console.error);
