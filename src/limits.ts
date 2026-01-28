export type Limits = {
  maxPptxSizeBytes: number;
  maxZipEntries: number;
  maxTotalUncompressedBytes: number;
  maxEntryBytes: number;
  libreOfficeTimeoutMs: number;
  pdftoppmTimeoutMs: number;
};

const DEFAULT_LIMITS: Limits = {
  maxPptxSizeBytes: 50 * 1024 * 1024,
  maxZipEntries: 5000,
  maxTotalUncompressedBytes: 500 * 1024 * 1024,
  maxEntryBytes: 50 * 1024 * 1024,
  libreOfficeTimeoutMs: 120_000,
  pdftoppmTimeoutMs: 120_000,
};

export function getLimits(env = process.env): Limits {
  return {
    maxPptxSizeBytes: parseEnvNumber(
      env,
      "PPTX_IMPORTER_MAX_PPTX_SIZE_BYTES",
      DEFAULT_LIMITS.maxPptxSizeBytes,
    ),
    maxZipEntries: parseEnvNumber(
      env,
      "PPTX_IMPORTER_MAX_ZIP_ENTRIES",
      DEFAULT_LIMITS.maxZipEntries,
    ),
    maxTotalUncompressedBytes: parseEnvNumber(
      env,
      "PPTX_IMPORTER_MAX_TOTAL_UNCOMPRESSED_BYTES",
      DEFAULT_LIMITS.maxTotalUncompressedBytes,
    ),
    maxEntryBytes: parseEnvNumber(
      env,
      "PPTX_IMPORTER_MAX_ENTRY_BYTES",
      DEFAULT_LIMITS.maxEntryBytes,
    ),
    libreOfficeTimeoutMs: parseEnvNumber(
      env,
      "PPTX_IMPORTER_LIBREOFFICE_TIMEOUT_MS",
      DEFAULT_LIMITS.libreOfficeTimeoutMs,
    ),
    pdftoppmTimeoutMs: parseEnvNumber(
      env,
      "PPTX_IMPORTER_PDFTOPPM_TIMEOUT_MS",
      DEFAULT_LIMITS.pdftoppmTimeoutMs,
    ),
  };
}

function parseEnvNumber(
  env: NodeJS.ProcessEnv,
  key: string,
  fallback: number,
): number {
  const raw = env[key];
  if (!raw) {
    return fallback;
  }
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Invalid ${key} value: ${raw}`);
  }
  return parsed;
}
