import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { spawn, spawnSync } from "node:child_process";
import { performance } from "node:perf_hooks";
import { getConversionLimitsFromEnv } from "../limits";

export type BackgroundRenderTimings = {
  libreofficeMs?: number;
  pdftoppmMs?: number;
  totalMs?: number;
};

export async function renderBackgrounds(
  inputPptx: string,
  outDir: string,
  timings: BackgroundRenderTimings = {},
): Promise<void> {
  const libreOfficeBinary = resolveLibreOfficeBinary();
  if (!libreOfficeBinary) {
    const candidates = getLibreOfficeCandidates();
    const hintText = [
      "Install LibreOffice and ensure it is on PATH.",
      `Tried: ${candidates.join(", ")}`,
    ];
    throw new Error(
      `Missing dependency: LibreOffice.\n${hintText
        .map((hint) => `- ${hint}`)
        .join("\n")}`,
    );
  }
  await ensureBinary("pdftoppm", [
    "Install poppler-utils (pdftoppm) and ensure it is on PATH.",
  ]);

  const backgroundsDir = path.join(outDir, "backgrounds");
  await fs.mkdir(backgroundsDir, { recursive: true });

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-"));
  try {
    const totalStart = performance.now();
    const pptxName = path.basename(inputPptx);
    const pdfName = pptxName.replace(/\.pptx$/i, ".pdf");
    const pdfPath = path.join(tmpDir, pdfName);
    const limits = getConversionLimitsFromEnv();

    const libreOfficeStart = performance.now();
    await runCommandWithTimeout(
      libreOfficeBinary,
      [
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        tmpDir,
        inputPptx,
      ],
      limits.libreOfficeTimeoutMs,
    );
    timings.libreofficeMs = performance.now() - libreOfficeStart;

    const pdftoppmStart = performance.now();
    await runCommandWithTimeout(
      "pdftoppm",
      ["-png", "-r", "144", pdfPath, path.join(backgroundsDir, "slide")],
      limits.pdftoppmTimeoutMs,
    );
    timings.pdftoppmMs = performance.now() - pdftoppmStart;

    await normalizeBackgroundNames(backgroundsDir);
    timings.totalMs = performance.now() - totalStart;
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true });
  }
}

export function resolveLibreOfficeBinary(): string | null {
  const candidates = getLibreOfficeCandidates();
  for (const candidate of candidates) {
    if (isBinaryAvailable(candidate)) {
      return candidate;
    }
  }
  return null;
}

export function getLibreOfficeCandidates(): string[] {
  if (os.platform() !== "win32") {
    return ["libreoffice"];
  }
  return [
    "soffice",
    "soffice.com",
    "soffice.exe",
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files\\LibreOffice\\program\\soffice.com",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
  ];
}

async function ensureBinary(command: string, hints: string[]): Promise<void> {
  if (isBinaryAvailable(command)) {
    return;
  }
  const hintText = hints.map((hint) => `- ${hint}`).join("\n");
  throw new Error(`Missing dependency: ${command}.\n${hintText}`);
}

export function isBinaryAvailable(command: string): boolean {
  const result = spawnSync(command, ["--version"], { stdio: "ignore" });
  return result.error == null;
}

async function normalizeBackgroundNames(backgroundsDir: string): Promise<void> {
  const files = await fs.readdir(backgroundsDir);
  const slideFiles = files
    .filter((file: string) => file.startsWith("slide-") && file.endsWith(".png"))
    .sort((a: string, b: string) => extractNumber(a) - extractNumber(b));

  await Promise.all(
    slideFiles.map(async (file: string, index: number) => {
      const target = `slide-${index + 1}.png`;
      if (file === target) {
        return;
      }
      await fs.rename(
        path.join(backgroundsDir, file),
        path.join(backgroundsDir, target),
      );
    }),
  );
}

function extractNumber(fileName: string): number {
  const match = fileName.match(/slide-(\d+)/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER;
  }
  return Number(match[1]);
}

function runCommandWithTimeout(
  command: string,
  args: string[],
  timeoutMs: number,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: "ignore" });
    let settled = false;
    const timeoutId = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      try {
        child.kill("SIGKILL");
      } catch {
        child.kill();
      }
      reject(new Error(`${command} timed out after ${timeoutMs}ms.`));
    }, timeoutMs);

    child.on("error", (error: Error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutId);
      reject(error);
    });

    child.on("exit", (code: number | null, signal: NodeJS.Signals | null) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timeoutId);
      if (code === 0) {
        resolve();
        return;
      }
      const details = signal ? `signal ${signal}` : `code ${code ?? "unknown"}`;
      reject(new Error(`${command} failed with ${details}.`));
    });
  });
}
