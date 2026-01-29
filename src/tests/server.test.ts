import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { convertPptxToOutZipWithDependencies } from "../convert";
import { createConverterServer, ConvertHandler } from "../server";

async function buildMinimalPptx(): Promise<Buffer> {
  const zip = new JSZip();
  const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`;
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:pic>
        <p:blipFill>
          <a:blip r:embed="rId1"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="914400" cy="914400"/>
          </a:xfrm>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image-1.png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
</Relationships>`;

  zip.file("ppt/presentation.xml", presentationXml);
  zip.file("ppt/slides/slide1.xml", slideXml);
  zip.file("ppt/slides/_rels/slide1.xml.rels", relsXml);
  zip.file("ppt/media/image-1.png", Buffer.from("png-content"));

  return zip.generateAsync({ type: "nodebuffer" });
}

async function stubRenderBackgrounds(_inputPptx: string, outDir: string): Promise<void> {
  const backgroundsDir = path.join(outDir, "backgrounds");
  await fs.mkdir(backgroundsDir, { recursive: true });
  const payload = "background".repeat(200);
  await fs.writeFile(path.join(backgroundsDir, "slide-1.png"), payload);
}

function createConvertHandler(delayMs = 0): ConvertHandler {
  return async (buffer, timings) => {
    if (delayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
    return convertPptxToOutZipWithDependencies(
      buffer,
      {
        renderBackgrounds: stubRenderBackgrounds,
      },
      undefined,
      timings,
    );
  };
}

async function sendRequest(
  port: number,
  body: Buffer,
  timeoutMs = 2000,
): Promise<{
  status: number;
  headers: Record<string, string | string[] | undefined>;
  body: Buffer;
}> {
  return new Promise((resolve, reject) => {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    const req = http.request(
      {
        hostname: "127.0.0.1",
        port,
        path: "/convert",
        method: "POST",
        signal: controller.signal,
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "Content-Length": body.length,
        },
      },
      (res: any) => {
        const chunks: Buffer[] = [];
        res.on("data", (chunk: Buffer) => chunks.push(chunk));
        res.on("end", () => {
          clearTimeout(timeoutId);
          resolve({
            status: res.statusCode ?? 0,
            headers: res.headers,
            body: Buffer.concat(chunks),
          });
        });
      },
    );
    req.on("error", (error: unknown) => {
      clearTimeout(timeoutId);
      reject(error);
    });
    req.write(body);
    req.end();
  });
}

async function listen(server: any): Promise<number> {
  return new Promise((resolve, reject) => {
    server.listen(0, () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        reject(new Error("Server did not bind to a port."));
        return;
      }
      resolve(address.port);
    });
  });
}

async function runSuccessTest(): Promise<void> {
  const convertHandler = createConvertHandler();
  const server = createConverterServer({
    queue: { maxConcurrent: 1, maxQueue: 1, queueWaitTimeoutMs: 1000 },
    convert: convertHandler,
  });
  const port = await listen(server);
  try {
    const inputBuffer = await buildMinimalPptx();
    const response = await sendRequest(port, inputBuffer, 4000);
    if (response.status !== 200) {
      throw new Error(`Expected 200, got ${response.status}`);
    }
    const contentType = response.headers["content-type"];
    if (
      Array.isArray(contentType)
        ? !contentType.includes("application/zip")
        : contentType !== "application/zip"
    ) {
      throw new Error(`Expected application/zip, got ${contentType}`);
    }
    const signature = Buffer.from(response.body.slice(0, 2)).toString("utf-8");
    if (signature !== "PK") {
      throw new Error("Expected ZIP signature in response body.");
    }
    const outZip = await JSZip.loadAsync(response.body);
    const files = Object.keys(outZip.files);
    if (!files.includes("doc.json")) {
      throw new Error("Expected doc.json in out.zip.");
    }
    const backgroundFiles = files.filter((file) =>
      file.startsWith("backgrounds/"),
    );
    if (backgroundFiles.length === 0) {
      throw new Error("Expected backgrounds/ entries in out.zip.");
    }
    const assetFiles = files.filter((file) => file.startsWith("assets/"));
    if (assetFiles.length === 0) {
      throw new Error("Expected assets/ entries in out.zip.");
    }
  } finally {
    server.close();
  }
}

async function runQueueFullTest(): Promise<void> {
  const inputBuffer = await buildMinimalPptx();
  let gate: ((value?: void | PromiseLike<void>) => void) | undefined;
  const convertHandler: ConvertHandler = async (buffer, timings) => {
    if (!gate) {
      await new Promise<void>((resolve) => {
        gate = resolve;
      });
    }
    return convertPptxToOutZipWithDependencies(
      buffer,
      { renderBackgrounds: stubRenderBackgrounds },
      undefined,
      timings,
    );
  };
  const server = createConverterServer({
    queue: { maxConcurrent: 1, maxQueue: 0, queueWaitTimeoutMs: 1000 },
    convert: convertHandler,
  });
  const port = await listen(server);
  try {
    const first = sendRequest(port, inputBuffer, 4000);
    await new Promise((resolve) => setTimeout(resolve, 50));
    const second = await sendRequest(port, inputBuffer, 2000);
    if (second.status !== 429) {
      throw new Error(`Expected 429, got ${second.status}`);
    }
    const errorBody = JSON.parse(second.body.toString("utf-8")) as { code?: string };
    if (errorBody.code !== "QUEUE_FULL") {
      throw new Error(`Expected QUEUE_FULL, got ${errorBody.code}`);
    }
    gate?.();
    await first;
  } finally {
    server.close();
  }
}

async function runQueueTimeoutTest(): Promise<void> {
  const inputBuffer = await buildMinimalPptx();
  let gate: ((value?: void | PromiseLike<void>) => void) | undefined;
  const convertHandler: ConvertHandler = async (buffer, timings) => {
    if (!gate) {
      await new Promise<void>((resolve) => {
        gate = resolve;
      });
    }
    return convertPptxToOutZipWithDependencies(
      buffer,
      { renderBackgrounds: stubRenderBackgrounds },
      undefined,
      timings,
    );
  };
  const server = createConverterServer({
    queue: { maxConcurrent: 1, maxQueue: 1, queueWaitTimeoutMs: 20 },
    convert: convertHandler,
  });
  const port = await listen(server);
  try {
    const first = sendRequest(port, inputBuffer, 4000);
    await new Promise((resolve) => setTimeout(resolve, 10));
    const second = await sendRequest(port, inputBuffer, 2000);
    if (second.status !== 503) {
      throw new Error(`Expected 503, got ${second.status}`);
    }
    const errorBody = JSON.parse(second.body.toString("utf-8")) as { code?: string };
    if (errorBody.code !== "QUEUE_TIMEOUT") {
      throw new Error(`Expected QUEUE_TIMEOUT, got ${errorBody.code}`);
    }
    gate?.();
    await first;
  } finally {
    server.close();
  }
}

async function main(): Promise<void> {
  await runSuccessTest();
  await runQueueFullTest();
  await runQueueTimeoutTest();
  console.log("server tests passed");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
