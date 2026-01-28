import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { convertPptxToOutZipInternal } from "../convert";

const ONE_BY_ONE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0nDkQAAAAASUVORK5CYII=";

test("convertPptxToOutZip creates required output structure", async () => {
  const pptxBuffer = await buildMinimalPptx();
  const zipBuffer = await convertPptxToOutZipInternal(pptxBuffer, {
    renderBackgrounds: async (_inputPptx, outDir) => {
      const backgroundsDir = path.join(outDir, "backgrounds");
      await fs.mkdir(backgroundsDir, { recursive: true });
      await fs.writeFile(
        path.join(backgroundsDir, "slide-1.png"),
        Buffer.from(ONE_BY_ONE_PNG_BASE64, "base64"),
      );
    },
  });

  const outZip = await JSZip.loadAsync(zipBuffer);
  assert.ok(outZip.file("doc.json"), "doc.json should exist");
  assert.ok(
    outZip.file("backgrounds/slide-1.png"),
    "background should exist",
  );
  assert.ok(
    outZip.file("assets/images/slide-1-img-1.png"),
    "image asset should exist",
  );
});

async function buildMinimalPptx(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
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
</Relationships>`,
  );
  zip.file(
    "ppt/media/image1.png",
    Buffer.from(ONE_BY_ONE_PNG_BASE64, "base64"),
  );

  return zip.generateAsync({ type: "nodebuffer" });
}
