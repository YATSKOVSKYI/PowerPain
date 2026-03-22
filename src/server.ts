import { Hono } from "hono";
import { serveStatic } from "hono/bun";
import { mkdir, writeFile, readFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { existsSync } from "node:fs";
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

// ─── Job store for async processing with polling ────────────────────────────

interface Job {
  status: "uploading" | "processing" | "done" | "error";
  stage: string;
  message: string;
  percent: number;
  fileName?: string;
  filePath?: string;
  stats?: import("./pptx-normalizer").NormalizeStats;
  error?: string;
  createdAt: number;
  // Chunked upload state
  totalChunks: number;
  receivedChunks: number;
  totalSize: number;
  uploadDir?: string;
  targetFont: string;
  originalName: string;
}

const jobStore = new Map<string, Job>();

// Clean up old jobs every minute
setInterval(() => {
  const now = Date.now();
  for (const [id, job] of jobStore) {
    if (now - job.createdAt > 15 * 60 * 1000) {
      if (job.filePath) removeTempFile(job.filePath);
      if (job.uploadDir) rm(job.uploadDir, { recursive: true, force: true }).catch(() => {});
      jobStore.delete(id);
    }
  }
}, 60 * 1000).unref();

// Step 1: Init upload — returns jobId and tells client how many chunks to send
app.post("/api/upload/init", async (c) => {
  const requestId = crypto.randomUUID().slice(0, 8);

  try {
    const clientIp =
      c.req.header("x-forwarded-for")?.split(",")[0]?.trim() ||
      c.req.header("x-real-ip") ||
      "unknown";

    if (isRateLimited(clientIp)) {
      return c.json({ error: "Too many requests. Please try again later." }, 429);
    }

    const body = await c.req.json();
    const { fileName, fileSize, targetFont } = body as {
      fileName: string;
      fileSize: number;
      targetFont: string;
    };

    if (!fileName || !fileSize) {
      return c.json({ error: "Missing fileName or fileSize" }, 400);
    }

    if (!fileName.toLowerCase().endsWith(".pptx")) {
      return c.json({ error: "Only .pptx files are accepted" }, 400);
    }

    if (fileSize > MAX_FILE_SIZE) {
      return c.json({ error: `File too large. Maximum size: ${MAX_FILE_SIZE / 1024 / 1024} MB` }, 400);
    }

    const font = targetFont || "Arial";
    if (!ALLOWED_FONTS.includes(font)) {
      return c.json({ error: `Font not allowed. Choose from: ${ALLOWED_FONTS.join(", ")}` }, 400);
    }

    const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MB chunks
    const totalChunks = Math.ceil(fileSize / CHUNK_SIZE);

    // Create upload directory for this job
    const uploadDir = join(TEMP_DIR, `upload-${requestId}`);
    await mkdir(uploadDir, { recursive: true });

    const job: Job = {
      status: "uploading",
      stage: "upload",
      message: "Waiting for file chunks...",
      percent: 0,
      totalChunks,
      receivedChunks: 0,
      totalSize: fileSize,
      uploadDir,
      targetFont: font,
      originalName: fileName,
      createdAt: Date.now(),
    };
    jobStore.set(requestId, job);

    console.log(`[${requestId}] Upload init: ${fileName} (${(fileSize / 1024 / 1024).toFixed(1)} MB, ${totalChunks} chunks)`);

    return c.json({ jobId: requestId, chunkSize: CHUNK_SIZE, totalChunks });

  } catch (err) {
    console.error(`[${requestId}] Init error:`, err);
    return c.json({ error: "Failed to initialize upload" }, 500);
  }
});

// Step 2: Upload a single chunk
app.post("/api/upload/:id/chunk/:index", async (c) => {
  const id = c.req.param("id");
  const index = parseInt(c.req.param("index"), 10);
  const job = jobStore.get(id);

  if (!job || !job.uploadDir) {
    return c.json({ error: "Job not found or expired." }, 404);
  }

  if (job.status !== "uploading") {
    return c.json({ error: "Upload already completed." }, 400);
  }

  if (isNaN(index) || index < 0 || index >= job.totalChunks) {
    return c.json({ error: "Invalid chunk index." }, 400);
  }

  try {
    const chunkData = Buffer.from(await c.req.arrayBuffer());
    const chunkPath = join(job.uploadDir, `chunk-${String(index).padStart(6, "0")}`);
    await writeFile(chunkPath, chunkData);

    job.receivedChunks++;
    const uploadPercent = Math.round((job.receivedChunks / job.totalChunks) * 40); // Upload = 0-40%
    job.percent = uploadPercent;
    job.message = `Uploading... (${job.receivedChunks}/${job.totalChunks} chunks)`;

    // All chunks received — start processing
    if (job.receivedChunks >= job.totalChunks) {
      job.message = "All chunks received, assembling file...";
      job.percent = 40;
      job.status = "processing";

      // Assemble and process in background
      assembleAndProcess(id, job);
    }

    return c.json({ received: job.receivedChunks, total: job.totalChunks });

  } catch (err) {
    console.error(`[${id}] Chunk ${index} error:`, err);
    return c.json({ error: "Failed to save chunk" }, 500);
  }
});

async function assembleAndProcess(id: string, job: Job) {
  try {
    // Assemble chunks into a single buffer
    const chunkFiles: string[] = [];
    for (let i = 0; i < job.totalChunks; i++) {
      chunkFiles.push(join(job.uploadDir!, `chunk-${String(i).padStart(6, "0")}`));
    }

    job.message = "Assembling file...";
    job.percent = 42;

    const chunks: Buffer[] = [];
    for (const chunkPath of chunkFiles) {
      chunks.push(Buffer.from(await readFile(chunkPath)));
    }
    const inputBuffer = Buffer.concat(chunks);

    // Clean up chunk files
    rm(job.uploadDir!, { recursive: true, force: true }).catch(() => {});
    job.uploadDir = undefined;

    // Validate ZIP magic bytes
    if (inputBuffer[0] !== 0x50 || inputBuffer[1] !== 0x4b) {
      job.status = "error";
      job.error = "File is not a valid PPTX archive";
      job.message = job.error;
      return;
    }

    console.log(`[${id}] Processing started, font: ${job.targetFont}`);

    job.message = "Unpacking PPTX archive...";
    job.percent = 45;

    const { output, stats } = await normalizePptxBuffer(inputBuffer, { targetFont: job.targetFont }, (stage, current, total) => {
      // Processing = 45-90%
      const stageMap: Record<string, { msg: string; base: number; weight: number }> = {
        unzip:   { msg: "Unpacking PPTX archive...", base: 45, weight: 5 },
        themes:  { msg: `Fixing themes (${current + 1}/${total})...`, base: 50, weight: 3 },
        slides:  { msg: `Fixing slides (${current + 1}/${total})...`, base: 53, weight: 25 },
        layouts: { msg: `Fixing layouts (${current + 1}/${total})...`, base: 78, weight: 4 },
        masters: { msg: `Fixing masters (${current + 1}/${total})...`, base: 82, weight: 3 },
        zip:     { msg: "Compressing fixed file...", base: 85, weight: 10 },
      };
      const s = stageMap[stage];
      if (s) {
        const stagePercent = total > 0 ? current / total : 0;
        job.stage = stage;
        job.message = s.msg;
        job.percent = Math.round(s.base + s.weight * stagePercent);
      }
    });

    job.message = "Saving result...";
    job.percent = 95;

    const baseName = job.originalName.replace(/\.pptx$/i, "");
    const outputName = `${baseName}-fixed.pptx`;
    const tempPath = join(TEMP_DIR, `${id}.pptx`);
    await Bun.write(tempPath, output);

    job.status = "done";
    job.percent = 100;
    job.message = "Done!";
    job.fileName = outputName;
    job.filePath = tempPath;
    job.stats = stats;

    console.log(
      `[${id}] Processing finished: ${stats.themes} theme(s), ${stats.masters} master(s), ${stats.layouts} layout(s), ${stats.slides} slide(s)`
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error(`[${id}] Error:`, message);
    job.status = "error";
    job.message = "Processing failed. The file may be corrupted or not a valid PPTX.";
    job.error = job.message;
  }
}

// Poll job status
app.get("/api/status/:id", (c) => {
  const id = c.req.param("id");
  const job = jobStore.get(id);

  if (!job) {
    return c.json({ error: "Job not found or expired." }, 404);
  }

  return c.json({
    status: job.status,
    stage: job.stage,
    message: job.message,
    percent: job.percent,
    ...(job.status === "done" && { fileName: job.fileName, stats: job.stats }),
    ...(job.status === "error" && { error: job.error }),
  });
});

// Download completed result
app.get("/api/download/:id", async (c) => {
  const id = c.req.param("id");
  const job = jobStore.get(id);

  if (!job || job.status !== "done" || !job.filePath) {
    return c.json({ error: "File not found or expired. Please process again." }, 404);
  }

  const file = Bun.file(job.filePath);
  if (!(await file.exists())) {
    jobStore.delete(id);
    return c.json({ error: "File not found. Please process again." }, 404);
  }

  const buffer = await file.arrayBuffer();

  // Clean up after download
  removeTempFile(job.filePath);
  jobStore.delete(id);

  return new Response(buffer, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": `attachment; filename="${encodeURIComponent(job.fileName!)}"`,
      "Content-Length": buffer.byteLength.toString(),
      "X-Stats": JSON.stringify(job.stats),
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
