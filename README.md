# Google Maps Contact Scraper - Perú

![Node.js](https://img.shields.io/badge/Node.js-18%2B-339933?logo=node.js&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-blue)
![Status](https://img.shields.io/badge/status-activo-brightgreen)

Herramienta para recolectar datos de contacto de negocios desde Google Maps a escala nacional. Diseñada para proyectos de logística que requieren identificar proveedores y colaboradores en múltiples categorías y distritos.

---

## ¿Qué extrae?

| Campo | Fuente |
|---|---|
| Nombre del negocio | Google Maps |
| Categoría | Configurada por el usuario (exacta) |
| Valoración | Google Maps |
| Dirección | Google Maps |
| Teléfono | Google Maps + Web del negocio |
| Correo electrónico | Web del negocio |
| WhatsApp | Web del negocio |
| Instagram | Web del negocio |
| Facebook | Web del negocio |
| TikTok | Web del negocio |
| URL de Maps | Google Maps |

---

## ¿Cómo funciona?

```
config.json (categorías × distritos)
        ↓
Google Maps → scroll infinito → lista de enlaces
        ↓
Pool persistente (4 páginas paralelas) → extrae fichas
        ↓
Fetch-first / Puppeteer fallback → web de cada negocio
        ↓
Detecta subpágina de contacto y la visita también
        ↓
Almacena en SQLite (contactos.db) → exporta a contactos.xlsx
        ↓
Reanudación automática: si se interrumpe, retoma donde quedó
```

---

## Instalación

**Requisitos:** Node.js 18 o superior

```bash
# 1. Clona el repositorio
git clone https://github.com/C1FR4/google-maps-contact-scraper.git

# 2. Instala las dependencias
cd google-maps-contact-scraper
npm install

# 3. Edita config.json con tus categorías y distritos

# 4. Ejecuta
npm start
```

---

## Configuración (`config.json`)

```json
{
  "archivoExcel": "contactos.xlsx",
  "maxResultadosPorBusqueda": 50,
  "visitarWebDelNegocio": true,
  "esperaMsEntreBusquedas": 3000,
  "umbralTextoUtil": 400,
  "concurrencia": 5,
  "concurrenciaFichas": 4,

  "palabrasContacto": [
    "contacto", "contactenos", "contact", "contactus",
    "ubicacion", "ubicaciones", "locales", "tiendas", "sucursales"
  ],

  "CATEGORIAS": [
    "Cafetería",
    "Restaurante peruano",
    "Tienda de ropa"
  ],

  "DISTRITOS": [
    "Miraflores Lima",
    "Barranco Lima",
    "San Isidro Lima",
    "Lince Lima",
    "Jesús María Lima",
    "Magdalena del Mar Lima",
    "Pueblo Libre Lima",
    "San Miguel Lima",
    "Surquillo Lima",
    "Santiago de Surco Lima"
  ]
}
```

| Parámetro | Descripción | Default |
|---|---|---|
| `archivoExcel` | Nombre del archivo de salida | `contactos.xlsx` |
| `maxResultadosPorBusqueda` | Negocios máximos por término de búsqueda | `50` |
| `visitarWebDelNegocio` | Enriquecer con datos de la web | `true` |
| `esperaMsEntreBusquedas` | Pausa entre búsquedas (ms) | `3000` |
| `umbralTextoUtil` | Mínimo de texto para considerar página válida | `400` |
| `concurrencia` | Cuántas webs visita en paralelo | `5` |
| `concurrenciaFichas` | Cuántas fichas de Maps extrae en paralelo | `4` |
| `palabrasContacto` | Palabras clave para detectar subpáginas de contacto | — |
| `CATEGORIAS` | Tipos de negocio a buscar | — |
| `DISTRITOS` | Distritos / zonas donde buscar | — |

---

## Output

Genera `contactos.xlsx` con:

- 15 columnas limpias (sin columnas técnicas)
- Encabezado azul oscuro con texto blanco, fila congelada
- Filas con colores alternos (blanco / azul claro)
- Filtros automáticos en todas las columnas
- Categoría exacta como la configuró el usuario
- Teléfonos normalizados (sin prefijo +51)
- Enlaces directos funcionales (WhatsApp, redes sociales)

---

## Características técnicas

- **Fetch-first**: intenta obtener la web con `fetch` (rápido); si la página requiere JS, cae automáticamente a Puppeteer.
- **Extracción paralela de fichas**: pool de 4 páginas de Puppeteer extrayendo fichas de Maps simultáneamente.
- **Pool persistente**: las páginas se crean una vez y se reúsan entre términos de búsqueda.
- **Reanudación automática**: si el proceso se interrumpe, al retomar salta los términos ya procesados (basado en SQLite).
- **Detección de bloqueo**: si Google detecta tráfico automatizado, espera 60s y reintenta una vez antes de fallar controladamente.
- **Filtro de redes sociales**: descarta enlaces falsos de Facebook (login, share, recover, photo, l.php, messages, events, etc.) y WhatsApp sin número.
- **Extracción desde URL social**: si un negocio tiene Instagram/Facebook como su "sitio web" en Maps, el handle se extrae directamente.

---

## Tiempo estimado

Para 30 términos (ej: 3 categorías × 10 distritos):

| Fase | Por término |
|---|---|
| Maps scroll + enlaces | ~15s |
| Fichas en paralelo | ~20s |
| Enriquecimiento web | ~60–90s |
| **Total estimado** | **30–45 minutos** |

~200–500 negocios en el Excel final.

---

## Notas

- Usa **fetch** como primera opción (rápido) y **Puppeteer** como fallback automático para páginas renderizadas con JS.
- Detecta y visita subpáginas de contacto dentro del mismo dominio.
- Normaliza teléfonos peruanos (elimina prefijo +51, 51, 0).
- Filtra correos falsos de dominios de tracking (Sentry, Hotjar, Klaviyo, etc.) y extensiones de archivo.
- Agrega pausas aleatorias para simular comportamiento humano.
- No requiere API keys ni proxies pagados.

---

## Disclaimer

Este proyecto se proporciona exclusivamente con **fines educativos y de investigación**. El scraping automatizado de Google Maps puede violar sus [Términos de Servicio](https://policies.google.com/terms). 

El mantenedor de este repositorio **no se responsabiliza** por el uso que terceros hagan de esta herramienta. Úsala bajo tu propio criterio y responsabilidad.

---


