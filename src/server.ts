import { Hono } from "hono";
import { serveStatic } from "hono/bun";
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import QRCode from "qrcode";
import { normalizePptxBuffer } from "./pptx-normalizer";
import { removeTempFile, startCleanupTimer } from "./cleanup";

const WALLET_ADDRESS = "TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ";

const app = new Hono();

const PORT = parseInt(process.env.PORT || "3000", 10);
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200 MB
const TEMP_DIR = join(import.meta.dir, "..", "tmp");

const ALLOWED_FONTS = [
  "Arial",
  "Calibri",
  "Times New Roman",
  "Helvetica",
  "Verdana",
  "Tahoma",
  "Georgia",
  "Segoe UI",
  "Roboto",
  "Open Sans",
  "Inter",
];

// ─── Rate limiter (per IP) ───────────────────────────────────────────────────

const RATE_WINDOW_MS = 60 * 1000; // 1 minute
const RATE_MAX_REQUESTS = 10; // max 10 uploads per minute per IP
const rateBuckets = new Map<string, { count: number; resetAt: number }>();

function isRateLimited(ip: string): boolean {
  const now = Date.now();
  const bucket = rateBuckets.get(ip);
  if (!bucket || now > bucket.resetAt) {
    rateBuckets.set(ip, { count: 1, resetAt: now + RATE_WINDOW_MS });
    return false;
  }
  bucket.count++;
  return bucket.count > RATE_MAX_REQUESTS;
}

// Clean up stale rate-limit buckets every 5 minutes
setInterval(() => {
  const now = Date.now();
  for (const [ip, bucket] of rateBuckets) {
    if (now > bucket.resetAt) rateBuckets.delete(ip);
  }
}, 5 * 60 * 1000).unref();

// Ensure temp directory exists
await mkdir(TEMP_DIR, { recursive: true });
startCleanupTimer(TEMP_DIR);

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

app.post("/api/process", async (c) => {
  const requestId = crypto.randomUUID().slice(0, 8);
  let tempInputPath: string | null = null;

  try {
    // Rate limiting
    const clientIp =
      c.req.header("x-forwarded-for")?.split(",")[0]?.trim() ||
      c.req.header("x-real-ip") ||
      "unknown";

    if (isRateLimited(clientIp)) {
      console.log(`[${requestId}] Rate limited: ${clientIp}`);
      return c.json({ error: "Too many requests. Please try again later." }, 429);
    }

    const contentType = c.req.header("content-type") || "";
    if (!contentType.includes("multipart/form-data")) {
      return c.json({ error: "Expected multipart/form-data" }, 400);
    }

    const formData = await c.req.formData();
    const file = formData.get("file");
    const targetFont = (formData.get("targetFont") as string) || "Arial";

    if (!file || !(file instanceof File)) {
      return c.json({ error: "No file uploaded" }, 400);
    }

    // Validate font
    if (!ALLOWED_FONTS.includes(targetFont)) {
      return c.json({ error: `Font not allowed. Choose from: ${ALLOWED_FONTS.join(", ")}` }, 400);
    }

    // Validate extension
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      return c.json({ error: "Only .pptx files are accepted" }, 400);
    }

    // Validate MIME type
    const validMimes = [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "application/octet-stream",
      "application/zip",
    ];
    if (!validMimes.includes(file.type)) {
      return c.json({ error: "Invalid file type" }, 400);
    }

    // Validate size
    if (file.size > MAX_FILE_SIZE) {
      return c.json({ error: `File too large. Maximum size: ${MAX_FILE_SIZE / 1024 / 1024} MB` }, 400);
    }

    console.log(`[${requestId}] Upload started: ${file.name} (${(file.size / 1024 / 1024).toFixed(1)} MB)`);

    // Read file into buffer (no temp file needed for input — process in memory)
    const arrayBuffer = await file.arrayBuffer();
    const inputBuffer = Buffer.from(arrayBuffer);

    // Validate it's actually a ZIP (PPTX is a ZIP archive)
    if (inputBuffer[0] !== 0x50 || inputBuffer[1] !== 0x4b) {
      return c.json({ error: "File is not a valid PPTX archive" }, 400);
    }

    console.log(`[${requestId}] Processing started, font: ${targetFont}`);

    const { output, stats } = await normalizePptxBuffer(inputBuffer, { targetFont });

    console.log(
      `[${requestId}] Processing finished: ${stats.themes} theme(s), ${stats.masters} master(s), ${stats.layouts} layout(s), ${stats.slides} slide(s)`
    );

    // Build output filename
    const baseName = file.name.replace(/\.pptx$/i, "");
    const outputName = `${baseName}-fixed.pptx`;

    return new Response(output, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(outputName)}"`,
        "Content-Length": output.length.toString(),
        "X-Stats": JSON.stringify(stats),
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error(`[${requestId}] Error:`, message);
    return c.json({ error: "Processing failed. The file may be corrupted or not a valid PPTX." }, 500);
  } finally {
    if (tempInputPath) {
      await removeTempFile(tempInputPath);
      console.log(`[${requestId}] Cleanup done`);
    }
  }
});

app.get("/api/fonts", (c) => {
  return c.json({ fonts: ALLOWED_FONTS });
});

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
  if (/\.(js|css|svg|png|jpg|jpeg|webp|ico|woff2?|ttf|eot)$/i.test(path)) {
    c.res.headers.set("Cache-Control", "public, max-age=86400, immutable");
  } else if (path === "/" || path.endsWith(".html")) {
    c.res.headers.set("Cache-Control", "no-cache");
  }
});

// ─── Static files ────────────────────────────────────────────────────────────

app.use("/*", serveStatic({ root: "./src/public" }));

// ─── Catch-all: block unmatched routes ───────────────────────────────────────

app.all("*", (c) => c.text("Not Found", 404));

// ─── Start ──────────────────────────────────────────────────────────────────

console.log(`PowerPain running at http://localhost:${PORT}`);

export default {
  port: PORT,
  fetch: app.fetch,
};
