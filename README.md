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
