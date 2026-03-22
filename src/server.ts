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

// Store completed results for download (auto-expire after 10 min)
const resultStore = new Map<string, { path: string; name: string; stats: import("./pptx-normalizer").NormalizeStats; createdAt: number }>();

setInterval(() => {
  const now = Date.now();
  for (const [id, entry] of resultStore) {
    if (now - entry.createdAt > 10 * 60 * 1000) {
      removeTempFile(entry.path);
      resultStore.delete(id);
    }
  }
}, 60 * 1000).unref();

app.post("/api/process", async (c) => {
  const requestId = crypto.randomUUID().slice(0, 8);

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

    if (!ALLOWED_FONTS.includes(targetFont)) {
      return c.json({ error: `Font not allowed. Choose from: ${ALLOWED_FONTS.join(", ")}` }, 400);
    }

    if (!file.name.toLowerCase().endsWith(".pptx")) {
      return c.json({ error: "Only .pptx files are accepted" }, 400);
    }

    const validMimes = [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "application/octet-stream",
      "application/zip",
    ];
    if (!validMimes.includes(file.type)) {
      return c.json({ error: "Invalid file type" }, 400);
    }

    if (file.size > MAX_FILE_SIZE) {
      return c.json({ error: `File too large. Maximum size: ${MAX_FILE_SIZE / 1024 / 1024} MB` }, 400);
    }

    console.log(`[${requestId}] Upload started: ${file.name} (${(file.size / 1024 / 1024).toFixed(1)} MB)`);

    const arrayBuffer = await file.arrayBuffer();
    const inputBuffer = Buffer.from(arrayBuffer);

    if (inputBuffer[0] !== 0x50 || inputBuffer[1] !== 0x4b) {
      return c.json({ error: "File is not a valid PPTX archive" }, 400);
    }

    const baseName = file.name.replace(/\.pptx$/i, "");
    const outputName = `${baseName}-fixed.pptx`;

    console.log(`[${requestId}] Processing started, font: ${targetFont}`);

    // SSE stream for live progress
    const stream = new ReadableStream({
      async start(controller) {
        const encoder = new TextEncoder();
        const send = (event: string, data: object) => {
          controller.enqueue(encoder.encode(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`));
        };

        try {
          send("progress", { stage: "upload", message: "File received, unpacking archive...", percent: 5 });

          const { output, stats } = await normalizePptxBuffer(inputBuffer, { targetFont }, (stage, current, total) => {
            const stageMap: Record<string, { msg: string; base: number; weight: number }> = {
              unzip:   { msg: "Unpacking PPTX archive...", base: 5, weight: 10 },
              themes:  { msg: `Fixing themes (${current + 1}/${total})...`, base: 15, weight: 5 },
              slides:  { msg: `Fixing slides (${current + 1}/${total})...`, base: 20, weight: 40 },
              layouts: { msg: `Fixing layouts (${current + 1}/${total})...`, base: 60, weight: 10 },
              masters: { msg: `Fixing masters (${current + 1}/${total})...`, base: 70, weight: 5 },
              zip:     { msg: "Compressing fixed file...", base: 75, weight: 20 },
            };

            const s = stageMap[stage];
            if (s) {
              const stagePercent = total > 0 ? current / total : 0;
              const percent = Math.round(s.base + s.weight * stagePercent);
              send("progress", { stage, message: s.msg, percent });
            }
          });

          send("progress", { stage: "saving", message: "Saving result...", percent: 95 });

          // Save result to temp file for download
          const tempPath = join(TEMP_DIR, `${requestId}.pptx`);
          await Bun.write(tempPath, output);
          resultStore.set(requestId, { path: tempPath, name: outputName, stats, createdAt: Date.now() });

          console.log(
            `[${requestId}] Processing finished: ${stats.themes} theme(s), ${stats.masters} master(s), ${stats.layouts} layout(s), ${stats.slides} slide(s)`
          );

          send("done", { downloadId: requestId, fileName: outputName, stats });

        } catch (err) {
          const message = err instanceof Error ? err.message : "Unknown error";
          console.error(`[${requestId}] Error:`, message);
          send("error", { error: "Processing failed. The file may be corrupted or not a valid PPTX." });
        }

        controller.close();
      },
    });

    return new Response(stream, {
      headers: {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "X-Accel-Buffering": "no",
      },
    });

  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error(`[${requestId}] Error:`, message);
    return c.json({ error: "Processing failed. The file may be corrupted or not a valid PPTX." }, 500);
  }
});

// Download endpoint for completed results
app.get("/api/download/:id", async (c) => {
  const id = c.req.param("id");
  const entry = resultStore.get(id);

  if (!entry) {
    return c.json({ error: "File not found or expired. Please process again." }, 404);
  }

  const file = Bun.file(entry.path);
  if (!(await file.exists())) {
    resultStore.delete(id);
    return c.json({ error: "File not found. Please process again." }, 404);
  }

  const buffer = await file.arrayBuffer();

  // Clean up after download
  removeTempFile(entry.path);
  resultStore.delete(id);

  return new Response(buffer, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": `attachment; filename="${encodeURIComponent(entry.name)}"`,
      "Content-Length": buffer.byteLength.toString(),
      "X-Stats": JSON.stringify(entry.stats),
    },
  });
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
