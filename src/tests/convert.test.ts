import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { convertPptxToOutZipWithDependencies } from "../convert";

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

async function main(): Promise<void> {
  const inputBuffer = await buildMinimalPptx();

  const { zipBuffer } = await convertPptxToOutZipWithDependencies(inputBuffer, {
    renderBackgrounds: async (_inputPptx, outDir) => {
      const backgroundsDir = path.join(outDir, "backgrounds");
      await fs.mkdir(backgroundsDir, { recursive: true });
      const payload = "background".repeat(200);
      await fs.writeFile(path.join(backgroundsDir, "slide-1.png"), payload);
    },
  });

  const outZip = await JSZip.loadAsync(zipBuffer);
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

  const docRaw = await outZip.file("doc.json")?.async("string");
  if (!docRaw) {
    throw new Error("doc.json missing from out.zip.");
  }
  const doc = JSON.parse(docRaw) as { slides?: Array<{ id: string }> };
  if (!doc.slides || doc.slides.length !== 1) {
    throw new Error("Expected exactly one slide in doc.json.");
  }

  console.log("conversion test passed");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
