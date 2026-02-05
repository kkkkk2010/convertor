# PPTX to doc_json importer (MVP)

CLI utility to convert Gamma-generated PPTX into a minimal JSON format for our editor. Only text and image elements are editable; everything else is rendered as a slide background PNG.

## Requirements

Node.js 18+ and the following system tools:

- **libreoffice** (for PPTX → PDF)
- **pdftoppm** from poppler-utils (for PDF → PNG)

Example install commands:

```bash
# macOS (Homebrew)
brew install libreoffice poppler

# Ubuntu/Debian
sudo apt-get install libreoffice poppler-utils
```

If the binaries are not on your PATH, add them manually (example for Windows PowerShell):

```powershell
$env:Path += ";C:\\Program Files\\LibreOffice\\program"
```

## Install

```bash
npm install
```

## Build

```bash
npm run build
```

## CLI

```bash
npm run build
node dist/importer.js --input <path/to/input.pptx> --out <path/to/outDir>
```

## Run

```bash
node dist/importer.js --input ./Tehnologii-budushego.pptx --out ./out
```

## HTTP service (internal)

Start the service:

```bash
npm run build
npm run start:server
```

Example request (raw PPTX body):

```bash
curl -X POST http://localhost:3001/convert \
  -H "Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation" \
  --data-binary @./input.pptx \
  --output out.zip
```

On success, the response body is the `out.zip` bytes and includes `X-Request-Id`.

## Library API

Use the programmatic API to get an in-memory out.zip buffer:

```js
const fs = require("node:fs/promises");
const { convertPptxToOutZip } = require("./dist/convert.js");

const inputBuffer = await fs.readFile("./input.pptx");
const outZipBuffer = await convertPptxToOutZip(inputBuffer);
await fs.writeFile("./out.zip", outZipBuffer);
```

## Limits & timeouts

The conversion enforces configurable limits (defaults shown):

- `PPTX_MAX_INPUT_BYTES` (default: 52428800)
- `PPTX_MAX_ZIP_ENTRIES` (default: 5000)
- `PPTX_MAX_UNCOMPRESSED_BYTES` (default: 209715200)
- `PPTX_MAX_FILE_UNCOMPRESSED_BYTES` (default: 52428800)
- `PPTX_LIBREOFFICE_TIMEOUT_MS` (default: 120000)
- `PPTX_PDFTOPPM_TIMEOUT_MS` (default: 120000)

## Service concurrency & queue

The HTTP service accepts the following env vars (defaults shown):

- `PPTX_IMPORTER_PORT` (default: 3001)
- `PPTX_IMPORTER_MAX_CONCURRENT` (default: 2)
- `PPTX_IMPORTER_MAX_QUEUE` (default: 10)
- `PPTX_IMPORTER_QUEUE_WAIT_TIMEOUT_MS` (default: 120000)

If the queue is full, the service returns HTTP 429 with `QUEUE_FULL`. If a queued request
waits longer than the timeout, the service returns HTTP 503 with `QUEUE_TIMEOUT`.

## Error codes

HTTP errors are JSON:

```json
{ "code": "INVALID_PPTX", "message": "Invalid or unsupported PPTX.", "requestId": "..." }
```

| Code | Meaning |
| --- | --- |
| `LIMIT_EXCEEDED` | Input or zip limits exceeded. |
| `TIMEOUT_LIBREOFFICE` | LibreOffice conversion timed out. |
| `TIMEOUT_PDFTOPPM` | pdftoppm conversion timed out. |
| `INVALID_PPTX` | Invalid ZIP/PPTX or missing required parts. |
| `UNSUPPORTED_FEATURE` | Unsupported PPTX feature. |
| `QUEUE_FULL` | Queue is full; retry later. |
| `QUEUE_TIMEOUT` | Queue wait timeout exceeded. |
| `INTERNAL` | Unexpected server error. |

## Output structure

```
out/
  doc.json
  backgrounds/
    slide-1.png
    slide-2.png
  assets/
    images/
      slide-1-img-1.png
      slide-1-img-2.png
```

## Notes

- Coordinates are in pixels at 96 DPI (EMU → px).
- Slide size is read from `ppt/presentation.xml` if present; otherwise defaults to 13.333 x 7.5 inches (16:9).
- Text styling is simplified to a basic default for MVP.

## Docker

By default, the image installs `libreoffice` and `poppler-utils` versions currently available in Debian bookworm repositories.

Build:

```bash
docker build -t pptx-importer .
```

Optional version pinning via build args (the build tries pinned versions first and automatically falls back to the latest repository versions if those pins are unavailable):

```bash
docker build -t pptx-importer \
  --build-arg LIBREOFFICE_VERSION=1:7.4.7-1+deb12u6 \
  --build-arg POPPLER_UTILS_VERSION=22.12.0-2+deb12u1 \
  .
```
Run:

```bash
docker run --rm -v "$PWD:/work" -w /work pptx-importer \
  node dist/importer.js --input ./input.pptx --out ./out
```

### Docker Compose (service)

`docker-compose.yml` starts the converter in HTTP server mode (`node dist/server.js`), not in CLI importer mode.

```bash
docker compose up -d --build
```

Inside the Compose network, set editor `CONVERTER_URL` to `http://converter:3001` and call `POST /convert`.
The converter is reachable at `http://converter:3001/convert`.
