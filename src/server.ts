import { Hono } from "hono";
import { serveStatic } from "hono/bun";
import { join } from "node:path";
import QRCode from "qrcode";

const WALLET_ADDRESS = "TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ";

const app = new Hono();

const PORT = parseInt(process.env.PORT || "3000", 10);

// ─── Security headers ────────────────────────────────────────────────────────

app.use("*", async (c, next) => {
  await next();

  // Prevent MIME-type sniffing
  c.res.headers.set("X-Content-Type-Options", "nosniff");

  // Clickjacking protection
  c.res.headers.set("X-Frame-Options", "DENY");

  // XSS filter (legacy browsers)
  c.res.headers.set("X-XSS-Protection", "1; mode=block");

  // Referrer policy — don't leak full URLs
  c.res.headers.set("Referrer-Policy", "strict-origin-when-cross-origin");

  // Restrict browser features
  c.res.headers.set(
    "Permissions-Policy",
    "camera=(), microphone=(), geolocation=(), interest-cohort=()"
  );

  // Content Security Policy
  c.res.headers.set(
    "Content-Security-Policy",
    [
      "default-src 'self'",
      "script-src 'self'",
      "style-src 'self' 'unsafe-inline'",
      "img-src 'self' data: blob:",
      "font-src 'self'",
      "connect-src 'self'",
      "object-src 'none'",
      "frame-ancestors 'none'",
      "base-uri 'self'",
      "form-action 'self'",
    ].join("; ")
  );

  // HSTS (browsers will enforce HTTPS for 1 year)
  c.res.headers.set("Strict-Transport-Security", "max-age=31536000; includeSubDomains");

  // Hide server identity
  c.res.headers.delete("X-Powered-By");
  c.res.headers.delete("Server");
});

// ─── Block dotfile / sensitive path access ───────────────────────────────────

app.use("*", async (c, next) => {
  const path = new URL(c.req.url).pathname;

  // Block access to hidden files, source maps, backups, configs
  if (
    /\/\./.test(path) ||
    /\.(map|ts|env|log|bak|sql|sh|git)$/i.test(path) ||
    /\/(node_modules|\.git|tmp)\//i.test(path)
  ) {
    return c.text("Not Found", 404);
  }

  await next();
});

// ─── API ────────────────────────────────────────────────────────────────────

// QR code for donation wallet
let qrCache: string | null = null;

app.get("/api/qr", async (c) => {
  if (!qrCache) {
    qrCache = await QRCode.toString(WALLET_ADDRESS, {
      type: "svg",
      margin: 1,
      color: { dark: "#C43E1C", light: "#FFFFFF" },
      width: 256,
    });
  }
  return new Response(qrCache, {
    headers: {
      "Content-Type": "image/svg+xml",
      "Cache-Control": "public, max-age=86400",
    },
  });
});

// ─── Cache headers for static assets ─────────────────────────────────────────

app.use("/*", async (c, next) => {
  await next();

  const path = new URL(c.req.url).pathname;
  if (/\.(svg|png|jpg|jpeg|webp|ico|woff2?|ttf|eot)$/i.test(path)) {
    // Images and fonts — cache aggressively
    c.res.headers.set("Cache-Control", "public, max-age=86400, immutable");
  } else if (/\.(js|css)$/i.test(path)) {
    // JS/CSS — short cache, revalidate on deploy
    c.res.headers.set("Cache-Control", "public, max-age=300, must-revalidate");
  } else if (path === "/" || path.endsWith(".html")) {
    c.res.headers.set("Cache-Control", "no-cache");
  }
});

// ─── Static files ────────────────────────────────────────────────────────────

app.use("/*", serveStatic({ root: "./src/public" }));

// ─── 404 page ────────────────────────────────────────────────────────────────

let notFoundHtml: string | null = null;

app.all("*", async (c) => {
  if (!notFoundHtml) {
    const file = Bun.file(join(import.meta.dir, "public", "404.html"));
    notFoundHtml = await file.text();
  }
  return c.html(notFoundHtml, 404);
});

// ─── Start ──────────────────────────────────────────────────────────────────

console.log(`PowerPain running at http://localhost:${PORT}`);

export default {
  port: PORT,
  fetch: app.fetch,
};
