export type ErrorCode =
  | "LIMIT_EXCEEDED"
  | "TIMEOUT_LIBREOFFICE"
  | "TIMEOUT_PDFTOPPM"
  | "INVALID_PPTX"
  | "UNSUPPORTED_FEATURE"
  | "QUEUE_FULL"
  | "QUEUE_TIMEOUT"
  | "INTERNAL";

export class AppError extends Error {
  readonly code: ErrorCode;

  constructor(code: ErrorCode, message: string, options?: { cause?: unknown }) {
    super(message);
    this.code = code;
    if (options?.cause !== undefined) {
      (this as { cause?: unknown }).cause = options.cause;
    }
  }
}

export function isAppError(error: unknown): error is AppError {
  return error instanceof AppError;
}

function messageFromError(error: unknown): string {
  if (error instanceof Error) {
    return error.message || "Unknown error occurred.";
  }
  return String(error);
}

function matchInvalidPptx(message: string): boolean {
  return [
    "missing file in pptx",
    "missing media file",
    "missing doc.json",
    "no slides found",
    "slide count mismatch",
    "central directory",
    "end of data",
    "corrupted",
  ].some((fragment) => message.includes(fragment));
}

function matchLimitExceeded(message: string): boolean {
  return message.includes("exceeds max") || message.includes("entry too large");
}

function matchTimeout(message: string, command: string): boolean {
  return message.includes(command) && message.includes("timed out");
}

export function toAppError(error: unknown): AppError {
  if (isAppError(error)) {
    return error;
  }
  const message = messageFromError(error);
  const normalized = message.toLowerCase();

  if (matchLimitExceeded(normalized)) {
    return new AppError("LIMIT_EXCEEDED", message, { cause: error });
  }
  if (matchTimeout(normalized, "libreoffice") || matchTimeout(normalized, "soffice")) {
    return new AppError("TIMEOUT_LIBREOFFICE", message, { cause: error });
  }
  if (matchTimeout(normalized, "pdftoppm")) {
    return new AppError("TIMEOUT_PDFTOPPM", message, { cause: error });
  }
  if (normalized.includes("unsupported")) {
    return new AppError("UNSUPPORTED_FEATURE", message, { cause: error });
  }
  if (matchInvalidPptx(normalized)) {
    return new AppError("INVALID_PPTX", message, { cause: error });
  }
  return new AppError("INTERNAL", message, { cause: error });
}
