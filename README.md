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

## Docker

Build the image (LibreOffice + poppler-utils are pinned in the Dockerfile):

```bash
docker build -t pptx-importer .
```

Run the CLI inside Docker (mount your working directory):

```bash
docker run --rm -v "$PWD":/work -w /work pptx-importer --input ./input.pptx --out ./out.zip
```

## Library API

Programmatic conversion returns the exact same `out.zip` bytes as the CLI:

```ts
import fs from "node:fs/promises";
import { convertPptxToOutZip } from "./dist/convert";

const pptx = await fs.readFile("./input.pptx");
const outZip = await convertPptxToOutZip(pptx);
await fs.writeFile("./out.zip", outZip);
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

## Limits & timeouts

The converter enforces configurable limits and process timeouts via environment variables:

| Variable | Default | Description |
| --- | --- | --- |
| `PPTX_IMPORTER_MAX_PPTX_SIZE_BYTES` | 52428800 | Max input PPTX size in bytes. |
| `PPTX_IMPORTER_MAX_ZIP_ENTRIES` | 5000 | Max number of zip entries in the PPTX. |
| `PPTX_IMPORTER_MAX_TOTAL_UNCOMPRESSED_BYTES` | 524288000 | Max total uncompressed size of the PPTX entries. |
| `PPTX_IMPORTER_MAX_ENTRY_BYTES` | 52428800 | Max size of any single extracted entry. |
| `PPTX_IMPORTER_LIBREOFFICE_TIMEOUT_MS` | 120000 | Timeout for LibreOffice conversion. |
| `PPTX_IMPORTER_PDFTOPPM_TIMEOUT_MS` | 120000 | Timeout for pdftoppm conversion. |

## Tests

```bash
npm test
```

Tests generate a minimal PPTX in-memory using JSZip and stub the background render step.

## Manual PPTX creation (no binaries committed)

If you need a local PPTX for manual validation, create one on your machine (do not commit it):

1. Open LibreOffice or PowerPoint.
2. Create a single-slide deck with one image and a text box.
3. Save it as `example.pptx`.
4. Run: `node dist/importer.js --input ./example.pptx --out ./out.zip`

## Notes

- Coordinates are in pixels at 96 DPI (EMU → px).
- Slide size is read from `ppt/presentation.xml` if present; otherwise defaults to 13.333 x 7.5 inches (16:9).
- Text styling is simplified to a basic default for MVP.
