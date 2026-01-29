import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import { convertPptxToOutZipWithDependencies } from "./convert";
import { renderBackgrounds } from "./render/backgrounds";
import { getConversionLimitsFromEnv } from "./limits";

async function main() {
  try {
    const { input, outRaw } = parseArgs(process.argv.slice(2));
    const limits = getConversionLimitsFromEnv();
    await ensureInput(input, limits.maxInputBytes);
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
    const { zipBuffer, slideCount, totalImages } =
      await convertPptxToOutZipWithDependencies(
      inputBuffer,
      { renderBackgrounds },
      limits,
    );
    await writeZipAtomically(outZipPath, zipBuffer);
    const zipStat = fsSync.statSync(outZipPath);
    if (!zipStat.isFile()) {
      throw new Error(`Expected zip file at ${outZipPath}`);
    }
    console.log(`wrote zip file size=${zipStat.size}`);
    console.log(
      `✔ Exported presentation to ${path.basename(outZipPath)} (${slideCount} slides, ${totalImages} images)`,
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

async function ensureInput(inputPath: string, maxInputBytes: number): Promise<void> {
  let stat: fsSync.Stats | null = null;
  try {
    stat = await fs.stat(inputPath);
  } catch {
    stat = null;
  }
  if (!stat || !stat.isFile()) {
    throw new Error(`Input file not found: ${inputPath}`);
  }
  if (stat.size > maxInputBytes) {
    throw new Error(
      `PPTX exceeds max input size (${stat.size} bytes > ${maxInputBytes} bytes).`,
    );
  }
}

async function writeZipAtomically(
  outZipPath: string,
  data: Buffer,
): Promise<void> {
  const outDir = path.dirname(outZipPath);
  await fs.mkdir(outDir, { recursive: true });
  const tempPath = path.join(
    outDir,
    `.tmp-${path.basename(outZipPath)}-${Date.now()}`,
  );
  await fs.writeFile(tempPath, data);
  await fs.rename(tempPath, outZipPath);
}

void main();
