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

Build:

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
