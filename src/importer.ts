import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import { convertPptxToOutZip } from "./convert";

async function main() {
  try {
    const { input, outRaw } = parseArgs(process.argv.slice(2));
    await ensureInput(input);
    const outIsZip = outRaw.toLowerCase().endsWith(".zip");
    const outZipPath = outIsZip
      ? path.resolve(outRaw)
      : path.join(path.resolve(outRaw), "out.zip");
    const outDir = outIsZip
      ? path.dirname(outZipPath)
      : path.resolve(outRaw);

    console.log(`outDir: ${outDir}`);
    console.log(`outZipPath: ${outZipPath}`);
    console.log(`outIsZip: ${outIsZip}`);

    await fs.mkdir(outDir, { recursive: true });
    if (fsSync.existsSync(outZipPath)) {
      const stat = fsSync.statSync(outZipPath);
      if (stat.isDirectory()) {
        throw new Error(
          `Expected a .zip FILE but path is a DIRECTORY: ${outZipPath}`,
        );
      }
      if (stat.isFile()) {
        fsSync.unlinkSync(outZipPath);
      }
    }

    const inputBuffer = await fs.readFile(input);
    const zipBuffer = await convertPptxToOutZip(inputBuffer);
    await writeZipSafely(outZipPath, zipBuffer);
    const stats = zipBuffer.conversionStats;
    console.log(
      `✔ Exported presentation to ${path.basename(outZipPath)} (${stats?.slideCount ?? 0} slides, ${stats?.imageCount ?? 0} images)`,
    );
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unknown error occurred.";
    console.error(message);
    process.exit(1);
  }
}

function parseArgs(argv: string[]): { input: string; outRaw: string } {
  const args = new Map<string, string>();
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg.startsWith("--")) {
      args.set(arg, argv[i + 1]);
      i += 1;
    }
  }

  const input = args.get("--input");
  const outRaw = args.get("--out");
  if (!input || !outRaw) {
    throw new Error(
      "Usage: node dist/importer.js --input <path/to/input.pptx> --out <path/to/out.zip>",
    );
  }
  return { input, outRaw };
}

async function ensureInput(inputPath: string): Promise<void> {
  try {
    const stat = await fs.stat(inputPath);
    if (!stat.isFile()) {
      throw new Error("Input path is not a file.");
    }
  } catch {
    throw new Error(`Input file not found: ${inputPath}`);
  }
}

async function writeZipSafely(
  outZipPath: string,
  zipBuffer: Buffer,
): Promise<void> {
  const tempDir = path.dirname(outZipPath);
  const tempFile = path.join(
    tempDir,
    `.out-${process.pid}-${Date.now()}.zip`,
  );
  await fs.writeFile(tempFile, zipBuffer);
  const stat = fsSync.statSync(tempFile);
  if (!stat.isFile() || stat.size <= 1024) {
    await fs.rm(tempFile, { force: true });
    throw new Error(
      `Expected zip file larger than 1KB, got ${stat.size} bytes at ${tempFile}`,
    );
  }
  await fs.rename(tempFile, outZipPath);
}

void main();
