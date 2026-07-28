const cheerio = require("cheerio");
const ExcelJS = require("exceljs");
const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");

// ─── CARGAR CONFIGURACIÓN DESDE config.json ───────────────────────
const configPath = path.join(__dirname, "config.json");
if (!fs.existsSync(configPath)) {
  console.error("❌  No se encontró config.json. Crea el archivo antes de ejecutar.");
  process.exit(1);
}
const CONFIG = JSON.parse(fs.readFileSync(configPath, "utf-8"));
// ─────────────────────────────────────────────────────────────────

// Dominios de herramientas de desarrollo / tracking que no son correos reales
const DOMINIOS_BASURA =
  /sentry\.io|example\.com|amazonaws\.com|cloudfront\.net|w3\.org|schema\.org|hotjar\.com|klaviyo\.com|googleapis\.com|gstatic\.com|jquery\.com|bootstrapcdn\.com/i;

// Extensiones de archivo para filtrar falsos correos
const EXTENSIONES_NO_CORREO =
  /\.(webp|png|jpg|jpeg|gif|svg|mp4|mp3|pdf|zip|ico|woff|woff2|ttf)$/i;

// Regex de emojis para limpiar texto antes de guardar en Excel
const REGEX_EMOJI =
  /[\u{1F000}-\u{1FFFF}]|[\u{2600}-\u{27FF}]|[\u{2B00}-\u{2BFF}]|[\u{FE00}-\u{FEFF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA00}-\u{1FA9F}]/gu;

// Regex para teléfonos peruanos y correos
const REGEX = {
  telefono: /(\+?51[\s\-]?)?9\d{8}\b/g,
  correo: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g,
};

// ─── GENERADOR DE BÚSQUEDAS ───────────────────────────────────────

function generarTerminosDeBusqueda() {
  const terminos = [];
  for (const categoria of CONFIG.CATEGORIAS) {
    for (const distrito of CONFIG.DISTRITOS) {
      terminos.push(`${categoria} ${distrito}`);
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
  return /wa\.me|whatsapp\.com|api\.whatsapp\.com|whatsapp:\/\/|wa\.link/i.test(href);
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
        if (/wa\.link/i.test(href)) {
          redes.WhatsApp.add(href.split("?")[0]);
          return;
        }
        const parsed = new URL(href);
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
    } else if (/facebook\.com\//i.test(href) && !/sharer/i.test(href)) {
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
    await esperarMs(1000);
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
    const html = await visitarWebConPuppeteer(url, browser);
    return { html, via: "puppeteer" };
  } catch (errPuppeteer) {
    console.warn(`   Puppeteer falló (${errPuppeteer.message}), intentando fetch...`);
  }

  try {
    const html = await visitarWebConFetch(url);
    return { html, via: "fetch" };
  } catch (errFetch) {
    throw new Error(`fetch: ${errFetch.message}`);
  }
}

// ─── GOOGLE MAPS: SCRAPING PRINCIPAL ─────────────────────────────

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

  const negocios = [];

  try {
    const url = `https://www.google.com/maps/search/${encodeURIComponent(termino)}`;
    console.log(`   Abriendo Maps: ${url}`);
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
    await page.waitForSelector('[role="feed"]', { timeout: 15000 });

    // Scroll robusto: se detiene cuando la cantidad de enlaces deja de crecer
    let prevCount = 0;
    let sameCount = 0;
    const maxIteraciones = 20;

    for (let s = 0; s < maxIteraciones; s++) {
      const enlacesActuales = await page.$$eval('a[href*="/maps/place/"]', (els) =>
        [...new Set(els.map((a) => a.href).filter((h) => h.includes("/maps/place/")))]
      );

      if (enlacesActuales.length === prevCount) {
        sameCount++;
      } else {
        sameCount = 0;
      }

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
    const limite = Math.min(enlaces.length, CONFIG.maxResultadosPorBusqueda);

    for (let i = 0; i < limite; i++) {
      const enlaceNegocio = enlaces[i];
      console.log(`   [${i + 1}/${limite}] Extrayendo ficha...`);

      try {
        await page.goto(enlaceNegocio, { waitUntil: "domcontentloaded", timeout: 20000 });
        await page.waitForSelector("h1", { timeout: 8000 }).catch(() => {});

        const datos = await page.evaluate(() => {
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

        if (datos.nombre) {
          negocios.push({ ...datos, urlMaps: enlaceNegocio });
        }
      } catch (err) {
        console.warn(`   Error en ficha: ${err.message}`);
      }

      await esperarAleatorio(
        CONFIG.esperaMsEntreNegocios,
        CONFIG.esperaMsEntreNegocios + 1000
      );
    }
  } finally {
    await page.close();
  }

  return negocios;
}

// ─── PROCESAR UN NEGOCIO ──────────────────────────────────────────

async function procesarNegocio(negocio, browser, terminoBusqueda) {
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

  if (CONFIG.visitarWebDelNegocio && negocio.web) {
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
    Categoría: limpiarTexto(negocio.categoria) || "—",
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
    { header: "Método", key: "Método", width: 12 },
    { header: "Vía", key: "Vía", width: 12 },
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
    if (r.Vía === "puppeteer") {
      row.getCell("Vía").font = { color: { argb: "FF1A7F37" }, bold: true };
    } else if (r.Vía === "fetch") {
      row.getCell("Vía").font = { color: { argb: "FFB45309" }, bold: true };
    }

    row.alignment = { vertical: "middle" };
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 17 } };
  ws.views = [{ state: "frozen", ySplit: 1 }];

  await wb.xlsx.writeFile(ruta);
}

// ─── MAIN ─────────────────────────────────────────────────────────

async function main() {
  const terminos = generarTerminosDeBusqueda();
  const totalCombinaciones = CONFIG.CATEGORIAS.length * CONFIG.DISTRITOS.length;

  console.log(`\n Maps Extractor iniciado`);
  console.log(`   Categorias    : ${CONFIG.CATEGORIAS.length}`);
  console.log(`   Distritos     : ${CONFIG.DISTRITOS.length}`);
  console.log(`   Combinaciones : ${totalCombinaciones}`);
  console.log(`   Max. negocios por busqueda : ${CONFIG.maxResultadosPorBusqueda}`);
  console.log(
    `   Max. negocios estimados    : ~${totalCombinaciones * CONFIG.maxResultadosPorBusqueda}`
  );
  console.log(`${"─".repeat(50)}\n`);

  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--lang=es-PE"],
  });

  const resultados = [];

  const filaErrorBase = {
    Nombre: "",
    Categoría: "",
    Valoración: "",
    "Teléfono Maps": "",
    "Teléfono Web": "",
    Correo: "",
    WhatsApp: "",
    Instagram: "",
    Facebook: "",
    TikTok: "",
    Dirección: "",
    Web: "",
    URLMaps: "",
    Búsqueda: "",
    Método: "—",
    Vía: "—",
    Estado: "",
  };

  try {
    for (let i = 0; i < terminos.length; i++) {
      const termino = terminos[i];
      console.log(`\n[Busqueda ${i + 1}/${terminos.length}] "${termino}"`);

      let negocios = [];
      try {
        negocios = await buscarEnMaps(termino, browser);
      } catch (err) {
        console.error(` Error al buscar en Maps: ${err.message}`);
        resultados.push({
          ...filaErrorBase,
          Nombre: "ERROR-BUSQUEDA",
          Categoría: termino,
          Búsqueda: termino,
          Estado: err.message.slice(0, 100),
        });
        continue;
      }

      console.log(`   Procesando ${negocios.length} negocios...`);

      for (let j = 0; j < negocios.length; j++) {
        const negocio = negocios[j];
        console.log(`   [${j + 1}/${negocios.length}] ${negocio.nombre}`);

        try {
          const datos = await procesarNegocio(negocio, browser, termino);
          resultados.push(datos);
          console.log(
            `     Maps: ${datos["Teléfono Maps"] || "—"}  Web: ${
              datos["Teléfono Web"] || "—"
            }  IG: ${datos.Instagram || "—"}  Via: ${datos.Vía}`
          );
        } catch (err) {
          console.error(`   Error: ${err.message}`);
          resultados.push({
            ...filaErrorBase,
            Nombre: negocio.nombre || "ERROR",
            Categoría: negocio.categoria || "",
            "Teléfono Maps": negocio.telefono || "",
            Dirección: negocio.direccion || "",
            Web: negocio.web || "",
            URLMaps: negocio.urlMaps || "",
            Búsqueda: termino,
            Método: "maps",
            Estado: err.message.slice(0, 100),
          });
        }
      }

      const tempRuta = CONFIG.archivoExcel.replace(".xlsx", "_temp.xlsx");
      await guardarExcel(resultados, tempRuta);
      console.log(`\n Guardado parcial: ${resultados.length} registros → ${tempRuta}`);

      if (i < terminos.length - 1) {
        await esperarAleatorio(
          CONFIG.esperaMsEntreBusquedas,
          CONFIG.esperaMsEntreBusquedas + 2000
        );
      }
    }
  } finally {
    await browser.close();
  }

  await guardarExcel(resultados, CONFIG.archivoExcel);

  const ok = resultados.filter((r) => r.Estado === "OK").length;
  const err = resultados.filter((r) => r.Estado !== "OK").length;
  const viaPuppeteer = resultados.filter((r) => r.Vía === "puppeteer").length;
  const viaFetch = resultados.filter((r) => r.Vía === "fetch").length;

  console.log(`\n${"═".repeat(50)}`);
  console.log(` Proceso completo`);
  console.log(`   Busquedas     : ${terminos.length}`);
  console.log(`   Negocios      : ${resultados.length}`);
  console.log(`   Exitosos      : ${ok}`);
  console.log(`   Con error     : ${err}`);
  console.log(`   Via puppeteer : ${viaPuppeteer}`);
  console.log(`   Via fetch     : ${viaFetch}`);
  console.log(`   Archivo       : ${CONFIG.archivoExcel}`);
  console.log(`${"═".repeat(50)}\n`);
}

main().catch(console.error);
