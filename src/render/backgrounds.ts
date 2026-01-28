import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFile, spawnSync } from "node:child_process";
import { getLimits } from "../limits";

export async function renderBackgrounds(
  inputPptx: string,
  outDir: string,
): Promise<void> {
  const limits = getLimits();
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
  const pptxName = path.basename(inputPptx);
  const pdfName = pptxName.replace(/\.pptx$/i, ".pdf");
  const pdfPath = path.join(tmpDir, pdfName);

  try {
    await execFileWithTimeout(
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
      "LibreOffice",
    );

    await execFileWithTimeout(
      "pdftoppm",
      ["-png", "-r", "144", pdfPath, path.join(backgroundsDir, "slide")],
      limits.pdftoppmTimeoutMs,
      "pdftoppm",
    );

    await normalizeBackgroundNames(backgroundsDir);
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
    .filter((file) => file.startsWith("slide-") && file.endsWith(".png"))
    .sort((a, b) => extractNumber(a) - extractNumber(b));

  await Promise.all(
    slideFiles.map(async (file, index) => {
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

async function execFileWithTimeout(
  command: string,
  args: string[],
  timeoutMs: number,
  label: string,
): Promise<void> {
  try {
    await new Promise<void>((resolve, reject) => {
      const child = execFile(
        command,
        args,
        { timeout: timeoutMs, killSignal: "SIGKILL" },
        (error) => {
          if (error) {
            reject(error);
            return;
          }
          resolve();
        },
      );

      child.on("error", (error) => {
        reject(error);
      });
    });
  } catch (error) {
    if (error instanceof Error && "killed" in error) {
      const errorDetails = error as {
        killed?: boolean;
        signal?: string;
        code?: string;
      };
      if (
        errorDetails.killed ||
        errorDetails.signal === "SIGKILL" ||
        errorDetails.code === "ETIMEDOUT"
      ) {
        throw new Error(`${label} timed out after ${timeoutMs}ms.`);
      }
    }
    throw error;
  }
}
