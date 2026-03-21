import { readdir, stat, rm } from "node:fs/promises";
import { join } from "node:path";

const TEMP_MAX_AGE_MS = 10 * 60 * 1000; // 10 minutes
const CLEANUP_INTERVAL_MS = 5 * 60 * 1000; // every 5 minutes

let timer: ReturnType<typeof setInterval> | null = null;

export function startCleanupTimer(tempDir: string): void {
  if (timer) return;

  timer = setInterval(() => cleanOrphanFiles(tempDir), CLEANUP_INTERVAL_MS);
  // Don't prevent process exit
  if (timer && typeof timer === "object" && "unref" in timer) {
    timer.unref();
  }
}

export function stopCleanupTimer(): void {
  if (timer) {
    clearInterval(timer);
    timer = null;
  }
}

async function cleanOrphanFiles(tempDir: string): Promise<void> {
  try {
    const entries = await readdir(tempDir);
    const now = Date.now();
    let cleaned = 0;

    for (const entry of entries) {
      const fullPath = join(tempDir, entry);
      try {
        const info = await stat(fullPath);
        if (now - info.mtimeMs > TEMP_MAX_AGE_MS) {
          await rm(fullPath, { recursive: true, force: true });
          cleaned++;
        }
      } catch {
        // File may have been deleted already
      }
    }

    if (cleaned > 0) {
      console.log(`[cleanup] Removed ${cleaned} orphan temp file(s)`);
    }
  } catch {
    // Temp dir may not exist yet
  }
}

export async function removeTempFile(filePath: string): Promise<void> {
  try {
    await rm(filePath, { force: true });
  } catch {
    console.warn(`[cleanup] Failed to remove: ${filePath}`);
  }
}
