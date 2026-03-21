<div align="center">

<!-- Animated SVG Logo -->
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 520 260" width="520" height="260">
  <defs>
    <!-- Gradients -->
    <linearGradient id="iconGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FF6B4A"/>
      <stop offset="50%" stop-color="#D35230"/>
      <stop offset="100%" stop-color="#A92B1A"/>
    </linearGradient>
    <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FF6B4A" stop-opacity="0.6"/>
      <stop offset="100%" stop-color="#A92B1A" stop-opacity="0.1"/>
    </linearGradient>
    <linearGradient id="textGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#FFFFFF"/>
      <stop offset="100%" stop-color="#E0E0E0"/>
    </linearGradient>
    <linearGradient id="painGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#FF6B4A"/>
      <stop offset="100%" stop-color="#D35230"/>
    </linearGradient>
    <radialGradient id="glowBg" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="#D35230" stop-opacity="0.15"/>
      <stop offset="100%" stop-color="#D35230" stop-opacity="0"/>
    </radialGradient>
    <!-- Filters -->
    <filter id="glow">
      <feGaussianBlur stdDeviation="4" result="blur"/>
      <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
    <filter id="softGlow">
      <feGaussianBlur stdDeviation="8" result="blur"/>
      <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
    <filter id="iconShadow">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#A92B1A" flood-opacity="0.5"/>
    </filter>
  </defs>

  <!-- Background glow pulse -->
  <circle cx="100" cy="110" r="90" fill="url(#glowBg)">
    <animate attributeName="r" values="80;100;80" dur="5s" repeatCount="indefinite"/>
    <animate attributeName="opacity" values="0.6;1;0.6" dur="5s" repeatCount="indefinite"/>
  </circle>

  <!-- Outer orbit ring 1 — slow rotation -->
  <ellipse cx="100" cy="110" rx="72" ry="72" fill="none" stroke="url(#ringGrad)" stroke-width="0.8" stroke-dasharray="8 12" opacity="0.3">
    <animateTransform attributeName="transform" type="rotate" from="0 100 110" to="360 100 110" dur="30s" repeatCount="indefinite"/>
  </ellipse>

  <!-- Outer orbit ring 2 — tilted, reverse -->
  <ellipse cx="100" cy="110" rx="80" ry="55" fill="none" stroke="#C43E1C" stroke-width="0.5" stroke-dasharray="4 8" opacity="0.2" transform="rotate(-20 100 110)">
    <animateTransform attributeName="transform" type="rotate" from="360 100 110" to="0 100 110" dur="25s" repeatCount="indefinite"/>
  </ellipse>

  <!-- Inner orbit ring — fast -->
  <circle cx="100" cy="110" r="58" fill="none" stroke="#FF6B4A" stroke-width="0.4" stroke-dasharray="3 6" opacity="0.25">
    <animateTransform attributeName="transform" type="rotate" from="0 100 110" to="-360 100 110" dur="15s" repeatCount="indefinite"/>
  </circle>

  <!-- Orbiting particle 1 — large, glowing -->
  <circle r="3" fill="#FF6B4A" opacity="0.9" filter="url(#glow)">
    <animateMotion dur="6s" repeatCount="indefinite" path="M100,38 A72,72 0 1,1 99.99,38"/>
    <animate attributeName="opacity" values="0.4;1;0.4" dur="6s" repeatCount="indefinite"/>
  </circle>

  <!-- Orbiting particle 2 — medium, counter-clockwise -->
  <circle r="2" fill="#D35230" opacity="0.7" filter="url(#glow)">
    <animateMotion dur="8s" repeatCount="indefinite" path="M100,38 A72,72 0 1,0 99.99,38"/>
    <animate attributeName="opacity" values="0.3;0.8;0.3" dur="8s" repeatCount="indefinite"/>
  </circle>

  <!-- Orbiting particle 3 — small, inner orbit -->
  <circle r="1.5" fill="#FF8A6A" opacity="0.6">
    <animateMotion dur="4.5s" repeatCount="indefinite" path="M100,52 A58,58 0 1,1 99.99,52"/>
    <animate attributeName="opacity" values="0.2;0.7;0.2" dur="4.5s" repeatCount="indefinite"/>
  </circle>

  <!-- Orbiting particle 4 — tilted orbit -->
  <circle r="1.8" fill="#FF6B4A" opacity="0.5" filter="url(#glow)">
    <animateMotion dur="10s" repeatCount="indefinite" path="M100,30 A80,55 0 1,1 99.99,30"/>
  </circle>

  <!-- Floating sparkles -->
  <circle cx="45" cy="55" r="1" fill="#FF6B4A" opacity="0">
    <animate attributeName="opacity" values="0;0.8;0" dur="3s" begin="0s" repeatCount="indefinite"/>
    <animate attributeName="r" values="0.5;1.5;0.5" dur="3s" begin="0s" repeatCount="indefinite"/>
  </circle>
  <circle cx="155" cy="65" r="1" fill="#D35230" opacity="0">
    <animate attributeName="opacity" values="0;0.6;0" dur="4s" begin="1.5s" repeatCount="indefinite"/>
    <animate attributeName="r" values="0.5;1.2;0.5" dur="4s" begin="1.5s" repeatCount="indefinite"/>
  </circle>
  <circle cx="60" cy="165" r="1" fill="#FF8A6A" opacity="0">
    <animate attributeName="opacity" values="0;0.5;0" dur="3.5s" begin="0.8s" repeatCount="indefinite"/>
    <animate attributeName="r" values="0.3;1;0.3" dur="3.5s" begin="0.8s" repeatCount="indefinite"/>
  </circle>
  <circle cx="148" cy="155" r="1" fill="#C43E1C" opacity="0">
    <animate attributeName="opacity" values="0;0.7;0" dur="2.8s" begin="2s" repeatCount="indefinite"/>
    <animate attributeName="r" values="0.5;1.3;0.5" dur="2.8s" begin="2s" repeatCount="indefinite"/>
  </circle>

  <!-- Icon background — rounded rect with shadow -->
  <rect x="64" y="74" width="72" height="72" rx="18" fill="url(#iconGrad)" filter="url(#iconShadow)">
    <animate attributeName="rx" values="18;20;18" dur="6s" repeatCount="indefinite"/>
  </rect>

  <!-- Icon "P" letter -->
  <text x="100" y="125" text-anchor="middle" font-family="'Segoe UI','Helvetica Neue',Arial,sans-serif" font-size="48" font-weight="800" fill="white" filter="url(#glow)">P</text>

  <!-- "Power" text -->
  <text x="200" y="108" font-family="'Segoe UI','Helvetica Neue',Arial,sans-serif" font-size="44" font-weight="800" fill="url(#textGrad)" letter-spacing="-1">Power</text>

  <!-- "Pain" text — accent color -->
  <text x="370" y="108" font-family="'Segoe UI','Helvetica Neue',Arial,sans-serif" font-size="44" font-weight="800" fill="url(#painGrad)" letter-spacing="-1">Pain</text>

  <!-- Tagline -->
  <text x="200" y="138" font-family="'Segoe UI','Helvetica Neue',Arial,sans-serif" font-size="14" fill="#999" letter-spacing="3" font-weight="400">FIX BROKEN FONTS IN POWERPOINT</text>

  <!-- Decorative line under tagline -->
  <line x1="200" y1="148" x2="480" y2="148" stroke="url(#ringGrad)" stroke-width="1" opacity="0.4">
    <animate attributeName="x2" values="200;480;200" dur="8s" repeatCount="indefinite"/>
  </line>

  <!-- Version badge -->
  <rect x="200" y="158" width="44" height="18" rx="9" fill="#D35230" opacity="0.15"/>
  <text x="222" y="171" text-anchor="middle" font-family="'SF Mono','Fira Code',monospace" font-size="9" fill="#FF6B4A" font-weight="600">v1.0</text>
</svg>

<br/>

**Fix broken fonts in PowerPoint — in one click.**

[![MIT License](https://img.shields.io/badge/License-MIT-C43E1C?style=flat-square)](LICENSE)
[![Built with Bun](https://img.shields.io/badge/Bun-Runtime-C43E1C?style=flat-square&logo=bun&logoColor=white)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-Strict-C43E1C?style=flat-square&logo=typescript&logoColor=white)](https://typescriptlang.org)
[![Hono](https://img.shields.io/badge/Hono-Backend-C43E1C?style=flat-square)](https://hono.dev)

---

<p>
  <b>EN</b> · Font won't change in PowerPoint? We fix it.<br/>
  <b>RU</b> · Шрифт не меняется в PowerPoint? Мы починим.<br/>
  <b>UK</b> · Шрифт не змінюється в PowerPoint? Ми виправимо.<br/>
  <b>ZH</b> · PowerPoint 中字体无法更改？我们来修复。<br/>
  <b>DE</b> · Schriftart ändert sich nicht in PowerPoint? Wir reparieren es.<br/>
  <b>ES</b> · ¿La fuente no cambia en PowerPoint? Lo arreglamos.<br/>
  <b>FR</b> · La police ne change pas dans PowerPoint ? On répare ça.
</p>

</div>

---

## The Problem

PowerPoint silently ignores your font changes in three common scenarios:

| Problem | Root Cause | What PowerPain Does |
|---------|-----------|---------------------|
| Font doesn't change via UI | `lang="zh-CN"` → PowerPoint reads `<a:ea>`, ignores `<a:latin>` | Sets `lang="en-US"`, writes explicit `<a:latin>` and `<a:ea>` |
| Theme font overrides everything | `+mj-lt` / `+mn-lt` references in run properties | Replaces all theme references with target font |
| CJK font in theme | `等线`, `等线 Light` in `majorFont`/`minorFont` | Normalizes `theme1.xml` directly |
| Layout/master inheritance | Fonts defined in layouts override slide-level settings | Processes all layers: slides, layouts, masters, themes |

## Features

- **One-click fix** — upload `.pptx`, get fixed file back
- **Deep normalization** — processes `a:rPr`, `a:defRPr`, `a:endParaRPr` across all XML files
- **Theme repair** — fixes `majorFont`/`minorFont` in theme files
- **11 font choices** — Arial, Calibri, Times New Roman, Helvetica, Verdana, Tahoma, Georgia, Segoe UI, Roboto, Open Sans, Inter
- **Zero storage** — files processed in memory, never saved to disk
- **7 languages** — EN, RU, UK, ZH, DE, ES, FR
- **Privacy-first** — no accounts, no tracking, no analytics
- **Modern UI** — dark theme, PowerPoint color palette, responsive design

## Tech Stack

| Layer | Technology | Why |
|-------|-----------|-----|
| Runtime | [Bun](https://bun.sh) | Fastest JS runtime, native TypeScript |
| Backend | [Hono](https://hono.dev) | 14KB, fastest Bun-native HTTP framework |
| PPTX Processing | [JSZip](https://stuk.github.io/jszip/) + [@xmldom/xmldom](https://github.com/xmldom/xmldom) | Lightweight ZIP + XML manipulation |
| Frontend | Vanilla HTML/CSS/JS | No framework overhead for a single page |
| QR Code | [qrcode](https://github.com/soldair/node-qrcode) | MIT-licensed, SVG generation |

**All dependencies are MIT-licensed.**

## Quick Start

```bash
# Clone
git clone https://github.com/yatskovskyi/PowerPain.git
cd PowerPain

# Install
bun install

# Run
bun run start
```

Open **http://localhost:3000**

### Development (auto-reload)

```bash
bun run dev
```

## Project Structure

```
PowerPain/
├── src/
│   ├── server.ts              # Hono server, API routes, static files
│   ├── pptx-normalizer.ts     # Core PPTX font normalization engine
│   ├── cleanup.ts             # TTL cleaner for orphan temp files
│   └── public/
│       ├── index.html          # Single-page UI with full SEO
│       ├── style.css           # Dark theme, PowerPoint palette
│       ├── app.js              # i18n, drag&drop, upload logic
│       ├── robots.txt          # Search engine + AI bot rules
│       ├── sitemap.xml         # XML sitemap
│       └── .well-known/
│           └── ai-plugin.json  # AI plugin manifest
├── package.json
├── tsconfig.json
├── LICENSE                     # MIT
└── README.md
```

## API

### `POST /api/process`

Upload a `.pptx` file, receive the fixed version.

```bash
curl -X POST http://localhost:3000/api/process \
  -F "file=@presentation.pptx" \
  -F "targetFont=Arial" \
  -o fixed.pptx
```

**Parameters:**
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | File | Yes | `.pptx` file (max 100 MB) |
| `targetFont` | String | No | Target font name (default: `Arial`) |

**Response:** Fixed `.pptx` file as download, with `X-Stats` header containing processing statistics.

### `GET /api/fonts`

Returns the list of available target fonts.

### `GET /api/qr`

Returns the donation QR code as SVG.

## Environment

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `3000` | Server port |

## How to Verify

1. Open a PowerPoint file with CJK font issues
2. Upload it to PowerPain
3. Download the fixed file
4. Open in PowerPoint — fonts should now be changeable
5. Check: `lang` attributes should be `en-US`, no `+mj-lt` references remain

## SEO & Discoverability

- Full meta tags (OG, Twitter Card, JSON-LD)
- `FAQPage` structured data for rich snippets
- `WebApplication` schema markup
- Hreflang tags for 6 languages
- XML sitemap
- `robots.txt` with AI crawler instructions
- `.well-known/ai-plugin.json` manifest

## Contributing

Issues and PRs welcome at [github.com/yatskovskyi/PowerPain](https://github.com/yatskovskyi/PowerPain).

## License

[MIT](LICENSE) — Dmytro Yatskovskyi

---

<div align="center">

**If PowerPain saved your day — consider [supporting the project](https://powerpain.yatskovskyi.top/#donate) ❤️**

`USDT (TRC-20): TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ`

</div>
