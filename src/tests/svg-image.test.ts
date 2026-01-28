import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { convertPptxToOutZipInternal } from "../convert";

const SVG_PAYLOAD =
  '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>';
const ONE_BY_ONE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0nDkQAAAAASUVORK5CYII=";

test("SVG media is preserved with .svg extension", async () => {
  const pptxBuffer = await buildSvgPptx();
  const zipBuffer = await convertPptxToOutZipInternal(pptxBuffer, {
    renderBackgrounds: async (_inputPptx, outDir) => {
      const backgroundsDir = path.join(outDir, "backgrounds");
      await fs.mkdir(backgroundsDir, { recursive: true });
      await fs.writeFile(path.join(backgroundsDir, "slide-1.png"), "stub");
    },
  });

  const outZip = await JSZip.loadAsync(zipBuffer);
  const pngFile = outZip.file("assets/images/slide-1-img-1.png");
  assert.ok(pngFile, "png asset should exist");
  const pngBytes = await pngFile.async("nodebuffer");
  assert.ok(isPngSignature(pngBytes), "png bytes should have signature");

  const svgFile = outZip.file("assets/images/slide-1-img-2.svg");
  assert.ok(svgFile, "svg asset should exist");
  assert.ok(
    !outZip.file("assets/images/slide-1-img-2.png"),
    "png asset should not exist for svg content",
  );
  const svgBytes = await svgFile.async("string");
  assert.ok(svgBytes.startsWith("<svg"), "svg bytes should be preserved");
});

async function buildSvgPptx(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
</Types>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`,
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
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
      <p:pic>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="0"/>
            <a:ext cx="914400" cy="914400"/>
          </a:xfrm>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>`,
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image2.png"/>
</Relationships>`,
  );
  zip.file("ppt/media/image1.png", Buffer.from(ONE_BY_ONE_PNG_BASE64, "base64"));
  zip.file("ppt/media/image2.png", Buffer.from(ONE_BY_ONE_PNG_BASE64, "base64"));
  zip.file("ppt/media/image2.svg", SVG_PAYLOAD);

  return zip.generateAsync({ type: "nodebuffer" });
}

function isPngSignature(data: Buffer): boolean {
  if (data.length < 8) {
    return false;
  }
  return (
    data[0] === 0x89 &&
    data[1] === 0x50 &&
    data[2] === 0x4e &&
    data[3] === 0x47 &&
    data[4] === 0x0d &&
    data[5] === 0x0a &&
    data[6] === 0x1a &&
    data[7] === 0x0a
  );
}
