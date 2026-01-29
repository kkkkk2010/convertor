import http from "node:http";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { randomUUID } from "node:crypto";
import { convertPptxToOutZipWithDependencies, ConversionResult, ConversionTimings } from "./convert";
import { renderBackgrounds } from "./render/backgrounds";
import { getConversionLimitsFromEnv } from "./limits";
import { AppError, ErrorCode, toAppError } from "./errors";
import { QueueOptions, RequestQueue } from "./queue";

type LogEntry = {
  level: "info" | "warn" | "error";
  requestId: string;
  event: string;
  timingsMs?: ConversionTimings;
  inputBytes?: number;
  outputBytes?: number;
  slidesCount?: number;
  imagesCount?: number;
  errorCode?: ErrorCode;
  errorMessage?: string;
  method?: string;
  url?: string;
  contentType?: string;
  contentLength?: number;
};

type Logger = (entry: LogEntry) => void;

export type ConvertHandler = (
  inputBuffer: Buffer,
  timings: ConversionTimings,
) => Promise<ConversionResult>;

export type ServerOptions = {
  queue?: QueueOptions;
  convert?: ConvertHandler;
  logger?: Logger;
};

const DEFAULT_PORT = 3001;

function readEnvNumber(name: string, defaultValue: number, allowZero = false): number {
  const raw = process.env[name];
  if (!raw) {
    return defaultValue;
  }
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0 || (!allowZero && value === 0)) {
    throw new Error(`Invalid ${name} value: ${raw}`);
  }
  return value;
}

function buildQueueOptionsFromEnv(): QueueOptions {
  return {
    maxConcurrent: readEnvNumber("PPTX_IMPORTER_MAX_CONCURRENT", 2),
    maxQueue: readEnvNumber("PPTX_IMPORTER_MAX_QUEUE", 10, true),
    queueWaitTimeoutMs: readEnvNumber(
      "PPTX_IMPORTER_QUEUE_WAIT_TIMEOUT_MS",
      120_000,
    ),
  };
}

function defaultLogger(entry: LogEntry): void {
  console.log(JSON.stringify(entry));
}

function statusForError(code: ErrorCode): number {
  switch (code) {
    case "QUEUE_FULL":
      return 429;
    case "QUEUE_TIMEOUT":
      return 503;
    case "LIMIT_EXCEEDED":
      return 413;
    case "INVALID_PPTX":
      return 400;
    case "UNSUPPORTED_FEATURE":
      return 422;
    case "TIMEOUT_LIBREOFFICE":
    case "TIMEOUT_PDFTOPPM":
      return 504;
    case "INTERNAL":
    default:
      return 500;
  }
}

function safeMessageForCode(code: ErrorCode): string {
  switch (code) {
    case "QUEUE_FULL":
      return "Queue is full.";
    case "QUEUE_TIMEOUT":
      return "Queue wait timeout exceeded.";
    case "LIMIT_EXCEEDED":
      return "Input exceeds configured limits.";
    case "INVALID_PPTX":
      return "Invalid or unsupported PPTX.";
    case "UNSUPPORTED_FEATURE":
      return "Unsupported PPTX feature.";
    case "TIMEOUT_LIBREOFFICE":
      return "LibreOffice conversion timed out.";
    case "TIMEOUT_PDFTOPPM":
      return "pdftoppm conversion timed out.";
    case "INTERNAL":
    default:
      return "Unexpected error occurred.";
  }
}

async function readRawBody(req: any, maxBytes: number): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let total = 0;
    req.on("data", (chunk: Buffer) => {
      total += chunk.length;
      if (total > maxBytes) {
        reject(
          new AppError(
            "LIMIT_EXCEEDED",
            `PPTX exceeds max input size (${total} bytes > ${maxBytes} bytes).`,
          ),
        );
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", (error: unknown) => reject(error));
  });
}

function defaultConvertHandler(): ConvertHandler {
  return async (inputBuffer: Buffer, timings: ConversionTimings) => {
    return convertPptxToOutZipWithDependencies(
      inputBuffer,
      {
        renderBackgrounds,
      },
      getConversionLimitsFromEnv(),
      timings,
    );
  };
}

export function createConverterServer(options: ServerOptions = {}): any {
  const queueOptions = options.queue ?? buildQueueOptionsFromEnv();
  const queue = new RequestQueue(queueOptions);
  const logger = options.logger ?? defaultLogger;
  const convert = options.convert ?? defaultConvertHandler();
  const limits = getConversionLimitsFromEnv();
  const debugHttp = process.env.PPTX_IMPORTER_DEBUG_HTTP === "1";

  return http.createServer(async (req: any, res: any) => {
    const requestId = randomUUID();
    res.setHeader("X-Request-Id", requestId);

    if (!req.url || req.url.split("?")[0] !== "/convert") {
      res.statusCode = 404;
      res.end("Not Found");
      return;
    }
    if (req.method !== "POST") {
      res.statusCode = 405;
      res.end("Method Not Allowed");
      return;
    }

    let closed = false;
    let released = false;
    const releaseOnce = () => {
      if (!release || released) {
        return;
      }
      released = true;
      release();
    };
    req.on("aborted", () => {
      closed = true;
      queue.cancel(requestId);
    });
    req.on("close", () => {
      closed = true;
      queue.cancel(requestId);
    });
    res.on("finish", () => {
      releaseOnce();
      if (!debugHttp) {
        return;
      }
      logger({
        level: "info",
        requestId,
        event: "http_res_finish",
      });
    });
    res.on("close", () => {
      releaseOnce();
      if (!debugHttp) {
        return;
      }
      logger({
        level: "info",
        requestId,
        event: "http_res_close",
      });
    });

    logger({
      level: "info",
      requestId,
      event: "convert_start",
    });

    let release: (() => void) | null = null;
    const timings: ConversionTimings = {};
    let inputBytes = 0;

    try {
      release = await queue.acquire(requestId);
      if (closed) {
        release();
        return;
      }
      const contentType = req.headers["content-type"]?.split(";")[0]?.trim();
      if (debugHttp) {
        logger({
          level: "info",
          requestId,
          event: "http_req",
          method: req.method,
          url: req.url,
          contentType,
          contentLength: Number(req.headers["content-length"]) || undefined,
        });
      }
      if (
        contentType &&
        contentType !==
          "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      ) {
        throw new AppError("INVALID_PPTX", "Unsupported Content-Type.");
      }
      const inputBuffer = await readRawBody(req, limits.maxInputBytes);
      inputBytes = inputBuffer.length;
      if (inputBuffer.length === 0) {
        throw new AppError("INVALID_PPTX", "Empty request body.");
      }
      if (debugHttp) {
        logger({
          level: "info",
          requestId,
          event: "body_read_done",
          inputBytes,
        });
      }
      if (debugHttp) {
        logger({
          level: "info",
          requestId,
          event: "before_convert",
        });
      }
      const result = await convert(inputBuffer, timings);
      if (closed) {
        return;
      }
      if (debugHttp) {
        logger({
          level: "info",
          requestId,
          event: "after_convert",
          outputBytes: result.zipBuffer.length,
        });
      }
      if (debugHttp) {
        logger({
          level: "info",
          requestId,
          event: "before_headers",
        });
      }
      logger({
        level: "info",
        requestId,
        event: "convert_done",
        timingsMs: result.timings ?? timings,
        inputBytes,
        outputBytes: result.zipBuffer.length,
        slidesCount: result.slideCount,
        imagesCount: result.totalImages,
      });
      res.statusCode = 200;
      res.setHeader("Content-Type", "application/zip");
      res.setHeader("Content-Length", String(result.zipBuffer.length));
      if (!closed) {
        if (debugHttp) {
          logger({
            level: "info",
            requestId,
            event: "before_pipeline",
            outputBytes: result.zipBuffer.length,
          });
        }
        try {
          await pipeline(Readable.from(result.zipBuffer), res);
          if (debugHttp) {
            logger({
              level: "info",
              requestId,
              event: "after_pipeline",
              outputBytes: result.zipBuffer.length,
            });
          }
          logger({
            level: "info",
            requestId,
            event: "http_response_end",
          });
          return;
        } catch (pipelineError) {
          if (debugHttp) {
            const message =
              pipelineError instanceof Error
                ? pipelineError.message
                : String(pipelineError);
            logger({
              level: "error",
              requestId,
              event: "pipeline_error",
              errorCode: toAppError(pipelineError).code,
              errorMessage: message,
            });
          }
          throw pipelineError;
        }
      }
      return;
    } catch (error) {
      const appError = toAppError(error);
      if (!closed && !res.headersSent) {
        res.statusCode = statusForError(appError.code);
        res.setHeader("Content-Type", "application/json");
        res.end(
          JSON.stringify({
            code: appError.code,
            message: safeMessageForCode(appError.code),
            requestId,
          }),
        );
        logger({
          level: "info",
          requestId,
          event: "http_response_end",
          errorCode: appError.code,
        });
      }
      logger({
        level: appError.code === "QUEUE_FULL" ? "warn" : "error",
        requestId,
        event: "convert_fail",
        timingsMs: timings,
        inputBytes: inputBytes || undefined,
        errorCode: appError.code,
      });
    } finally {
      releaseOnce();
    }
  });
}

export function startServer(): void {
  const port = readEnvNumber("PPTX_IMPORTER_PORT", DEFAULT_PORT, true);
  const server = createConverterServer();
  server.listen(port, () => {
    console.log(`converter server listening on ${port}`);
  });
}

if (require.main === module) {
  startServer();
}
