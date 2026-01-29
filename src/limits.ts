export type ConversionLimits = {
  maxInputBytes: number;
  maxZipEntries: number;
  maxTotalUncompressedBytes: number;
  maxFileUncompressedBytes: number;
  libreOfficeTimeoutMs: number;
  pdftoppmTimeoutMs: number;
};

const DEFAULT_LIMITS: ConversionLimits = {
  maxInputBytes: 50 * 1024 * 1024,
  maxZipEntries: 5000,
  maxTotalUncompressedBytes: 200 * 1024 * 1024,
  maxFileUncompressedBytes: 50 * 1024 * 1024,
  libreOfficeTimeoutMs: 120_000,
  pdftoppmTimeoutMs: 120_000,
};

function readEnvNumber(name: string, defaultValue: number): number {
  const raw = process.env[name];
  if (!raw) {
    return defaultValue;
  }
  const value = Number(raw);
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`Invalid ${name} value: ${raw}`);
  }
  return value;
}

export function getConversionLimitsFromEnv(): ConversionLimits {
  return {
    maxInputBytes: readEnvNumber("PPTX_MAX_INPUT_BYTES", DEFAULT_LIMITS.maxInputBytes),
    maxZipEntries: readEnvNumber("PPTX_MAX_ZIP_ENTRIES", DEFAULT_LIMITS.maxZipEntries),
    maxTotalUncompressedBytes: readEnvNumber(
      "PPTX_MAX_UNCOMPRESSED_BYTES",
      DEFAULT_LIMITS.maxTotalUncompressedBytes,
    ),
    maxFileUncompressedBytes: readEnvNumber(
      "PPTX_MAX_FILE_UNCOMPRESSED_BYTES",
      DEFAULT_LIMITS.maxFileUncompressedBytes,
    ),
    libreOfficeTimeoutMs: readEnvNumber(
      "PPTX_LIBREOFFICE_TIMEOUT_MS",
      DEFAULT_LIMITS.libreOfficeTimeoutMs,
    ),
    pdftoppmTimeoutMs: readEnvNumber(
      "PPTX_PDFTOPPM_TIMEOUT_MS",
      DEFAULT_LIMITS.pdftoppmTimeoutMs,
    ),
  };
}
