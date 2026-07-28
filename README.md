# Google Maps Contact Scraper - Perú

![Node.js](https://img.shields.io/badge/Node.js-18%2B-339933?logo=node.js&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-blue)
![Status](https://img.shields.io/badge/status-activo-brightgreen)

Herramienta de automatización para recolectar datos de contacto de negocios desde Google Maps. Desarrollada para un proyecto de logística para identificar potenciales proveedores y colaboradores.
---

## ¿Qué extrae?

| Campo | Fuente |
|---|---|
| Nombre del negocio | Google Maps |
| Categoría | Google Maps |
| Valoración (rating) | Google Maps |
| Dirección | Google Maps |
| Teléfono | Google Maps + Web |
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
Google Maps → lista de negocios
        ↓
Visita web de cada negocio (Puppeteer / Fetch)
        ↓  
Detecta subpágina de contacto y la visita también
        ↓
contactos.xlsx  
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
**Importante:** Antes de ejecutar el script, asegúrate de modificar el archivo de configuración con los distritos y categorías que necesites. Los valores actuales son solo de prueba.
```json
{
  "archivoExcel": "contactos.xlsx",
  "maxResultadosPorBusqueda": 20,
  "visitarWebDelNegocio": true,

  "CATEGORIAS": [
    "Cafetería",
    "Restaurante peruano"
  ],

  "DISTRITOS": [
    "Miraflores Lima",
    "Barranco Lima"
  ]
}
```

| Parámetro | Descripción | Default |
|---|---|---|
| `archivoExcel` | Nombre del archivo de salida | `contactos.xlsx` |
| `maxResultadosPorBusqueda` | Negocios máximos por búsqueda | `20` |
| `visitarWebDelNegocio` | Enriquecer con datos de la web | `true` |
| `esperaMsEntreBusquedas` | Pausa entre búsquedas (ms) | `3000` |
| `CATEGORIAS` | Lista de tipos de negocio a buscar | ver archivo |
| `DISTRITOS` | Lista de distritos / ciudades | ver archivo |

---

## Output

Genera `contactos.xlsx` con encabezados formateados, filas con colores alternos, filtros automáticos y primera fila fija. También guarda un archivo `_temp.xlsx` progresivo para no perder datos si se interrumpe la ejecución.

---

## Notas

- Usa **Puppeteer** como primera opción para visitar webs y **fetch** como fallback automático.
- Detecta y visita subpáginas de contacto dentro del mismo dominio.
- Normaliza teléfonos peruanos (elimina prefijo `+51`, `51`).
- Filtra correos falsos provenientes de dominios de tracking o herramientas (Sentry, Hotjar, etc.).
- Agrega pausas aleatorias para simular comportamiento humano.

---

## Disclaimer

Este proyecto es para fines educativos y de investigación. Úsalo respetando los [Términos de Servicio de Google](https://policies.google.com/terms) y la normativa de privacidad aplicable en tu país.
