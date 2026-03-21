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
const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB
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

// Ensure temp directory exists
await mkdir(TEMP_DIR, { recursive: true });
startCleanupTimer(TEMP_DIR);

// ─── API ────────────────────────────────────────────────────────────────────

app.post("/api/process", async (c) => {
  const requestId = crypto.randomUUID().slice(0, 8);
  let tempInputPath: string | null = null;

  try {
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

// ─── Static files ───────────────────────────────────────────────────────────

app.use("/*", serveStatic({ root: "./src/public" }));

// ─── Start ──────────────────────────────────────────────────────────────────

console.log(`PowerPain running at http://localhost:${PORT}`);

export default {
  port: PORT,
  fetch: app.fetch,
};
