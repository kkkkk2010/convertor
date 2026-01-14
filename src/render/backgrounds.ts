import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

export async function renderBackgrounds(
  inputPptx: string,
  outDir: string,
): Promise<void> {
  const libreOfficeBinary = getLibreOfficeBinary();
  await ensureBinary(libreOfficeBinary, [
    "Install LibreOffice and ensure it is on PATH.",
  ]);
  await ensureBinary("pdftoppm", [
    "Install poppler-utils (pdftoppm) and ensure it is on PATH.",
  ]);

  const backgroundsDir = path.join(outDir, "backgrounds");
  await fs.mkdir(backgroundsDir, { recursive: true });

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-"));
  const pptxName = path.basename(inputPptx);
  const pdfName = pptxName.replace(/\.pptx$/i, ".pdf");
  const pdfPath = path.join(tmpDir, pdfName);

  await execFileAsync(libreOfficeBinary, [
    "--headless",
    "--convert-to",
    "pdf",
    "--outdir",
    tmpDir,
    inputPptx,
  ]);

  await execFileAsync("pdftoppm", [
    "-png",
    "-r",
    "144",
    pdfPath,
    path.join(backgroundsDir, "slide"),
  ]);

  await normalizeBackgroundNames(backgroundsDir);
}

function getLibreOfficeBinary(): string {
  return os.platform() === "win32" ? "soffice" : "libreoffice";
}

async function ensureBinary(command: string, hints: string[]): Promise<void> {
  try {
    await execFileAsync("which", [command]);
  } catch {
    const hintText = hints.map((hint) => `- ${hint}`).join("\n");
    throw new Error(
      `Missing dependency: ${command}.\n${hintText}`,
    );
  }
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
