<div align="center">

<!-- Animated SVG Logo -->
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 160" width="400" height="160">
  <defs>
    <linearGradient id="rg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#D35230"/>
      <stop offset="100%" stop-color="#A92B1A"/>
    </linearGradient>
    <filter id="g">
      <feGaussianBlur stdDeviation="6" result="b"/>
      <feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
  </defs>
  <circle cx="60" cy="80" r="50" fill="#D35230" opacity="0.08">
    <animate attributeName="r" values="45;55;45" dur="4s" repeatCount="indefinite"/>
  </circle>
  <circle cx="60" cy="80" r="56" fill="none" stroke="#C43E1C" stroke-width="0.5" opacity="0.2" stroke-dasharray="6 4">
    <animateTransform attributeName="transform" type="rotate" from="0 60 80" to="360 60 80" dur="20s" repeatCount="indefinite"/>
  </circle>
  <circle r="2" fill="#D35230" opacity="0.8" filter="url(#g)">
    <animateMotion dur="6s" repeatCount="indefinite" path="M60,24 A56,56 0 1,1 59.99,24"/>
  </circle>
  <circle r="1.5" fill="#FF6B4A" opacity="0.5">
    <animateMotion dur="9s" repeatCount="indefinite" path="M60,24 A56,56 0 1,0 59.99,24"/>
  </circle>
  <rect x="28" y="48" width="64" height="64" rx="14" fill="url(#rg)" filter="url(#g)"/>
  <text x="60" y="94" text-anchor="middle" font-family="Arial,sans-serif" font-size="42" font-weight="bold" fill="white">P</text>
  <text x="140" y="92" font-family="Arial,sans-serif" font-size="38" font-weight="800" fill="white">PowerPain</text>
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
- **6 languages** — EN, UK, ZH, DE, ES, FR
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

**If PowerPain saved your day — consider [supporting the project](https://powerpain.dev/#donate) ❤️**

`USDT (TRC-20): TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ`

</div>
