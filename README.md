<div align="center">

<!-- Animated SVG Logo — GitHub renders SMIL animations via <img> -->
<a href="https://powerpain.yatskovskyi.top">
  <img src="src/public/img/logo-animated.svg" alt="PowerPain — Fix Broken Fonts in PowerPoint" width="680"/>
</a>

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

## Backstory

I'm currently writing academic papers, and part of the process involves redrawing figures in PowerPoint so they can later be inserted into Word. Don't ask me why — that's just how the workflow goes.

At some point I ran into a maddening issue: when you import shapes or figures into `.pptx`, the text inside them **refuses to change font**. You select everything, pick Arial — nothing happens. The font dropdown shows one thing, the slide shows another. It's completely broken.

This isn't a one-off thing. When you have dozens of figures to process, doing it manually is not an option. I've always solved problems like this with code — so I dug into the `.pptx` XML, found the root causes (CJK language attributes, theme font references, broken inheritance chains), and built a tool to fix it automatically.

That tool is **PowerPain**.

---

## The Problem

PowerPoint silently ignores your font changes in these common scenarios:

| Problem | Root Cause | What PowerPain Does |
|---------|-----------|---------------------|
| Font doesn't change via UI | `lang="zh-CN"` → PowerPoint reads `<a:ea>`, ignores `<a:latin>` | Sets `lang="en-US"`, writes explicit `<a:latin>` and `<a:ea>` |
| Theme font overrides everything | `+mj-lt` / `+mn-lt` references in run properties | Replaces all theme references with target font |
| CJK font in theme | `等线`, `等线 Light` in `majorFont`/`minorFont` | Normalizes `theme1.xml` directly |
| Layout/master inheritance | Fonts defined in layouts override slide-level settings | Processes all layers: slides, layouts, masters, themes |

## Features

- **One-click fix** — upload `.pptx`, get fixed file back
- **100% client-side** — all processing runs in your browser via JSZip + DOMParser, files never leave your device
- **Deep normalization** — processes `a:rPr`, `a:defRPr`, `a:endParaRPr` across all XML files
- **Theme repair** — fixes `majorFont`/`minorFont` in theme files
- **11 font choices** — Arial, Calibri, Times New Roman, Helvetica, Verdana, Tahoma, Georgia, Segoe UI, Roboto, Open Sans, Inter
- **Zero upload** — nothing is sent to any server, ever
- **7 languages** — EN, RU, UK, ZH, DE, ES, FR
- **Privacy-first** — no accounts, no tracking, no analytics
- **Modern UI** — dark theme, PowerPoint color palette, responsive design

## Tech Stack

| Layer | Technology | Why |
|-------|-----------|-----|
| Runtime | [Bun](https://bun.sh) | Fastest JS runtime, native TypeScript |
| Static Server | [Hono](https://hono.dev) | 14KB, fastest Bun-native HTTP framework |
| PPTX Processing | [JSZip](https://stuk.github.io/jszip/) + DOMParser (browser) | Client-side ZIP + XML manipulation, no server needed |
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
│   ├── server.ts              # Hono server — static files, QR API, security headers
│   └── public/
│       ├── index.html          # Single-page UI with full SEO
│       ├── style.css           # Dark theme, PowerPoint palette
│       ├── app.js              # Client-side PPTX processing, i18n, drag&drop
│       ├── jszip.min.js        # JSZip library for in-browser ZIP handling
│       ├── 404.html            # Custom 404 page
│       ├── robots.txt          # Search engine + AI bot rules
│       ├── sitemap.xml         # XML sitemap
│       └── img/                # Favicon, OG image, donate photo
├── package.json
├── tsconfig.json
├── LICENSE                     # MIT
└── README.md
```

## Architecture

All PPTX processing happens **entirely in the browser**. The server only serves static files and the donation QR code.

### How client-side processing works

1. User drops a `.pptx` file
2. `JSZip` unpacks the archive in the browser
3. `DOMParser` parses each XML file (slides, layouts, masters, themes)
4. JavaScript normalizes fonts: fixes `lang` attributes, replaces theme references (`+mj-lt`, `+mn-lt`), sets explicit `<a:latin>` and `<a:ea>` typefaces
5. `XMLSerializer` converts the DOM back to XML strings
6. `JSZip` repacks everything into a new `.pptx` blob
7. User downloads the fixed file — nothing was ever uploaded

### Server API

The server is minimal — it only has one API endpoint:

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

## Security

- **CSP** — strict Content-Security-Policy (self + Google Fonts only)
- **HSTS** — enforced HTTPS with 1-year max-age
- **X-Frame-Options: DENY** — clickjacking protection
- **No server processing** — zero attack surface for file uploads
- **Rate limiting** — dotfile / sensitive path blocking

## SEO & Discoverability

- Full meta tags (OG, Twitter Card, JSON-LD)
- `FAQPage` structured data for rich snippets
- `WebApplication` schema markup
- Hreflang tags for 7 languages
- XML sitemap
- `robots.txt` with AI crawler instructions

## Contributing

Issues and PRs welcome at [github.com/yatskovskyi/PowerPain](https://github.com/yatskovskyi/PowerPain).

## License

[MIT](LICENSE) — Dmytro Yatskovskyi

---

<div align="center">

**If PowerPain saved your day — consider [supporting the project](https://powerpain.yatskovskyi.top/#donate)**

`USDT (TRC-20): TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ`

</div>
