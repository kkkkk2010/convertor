import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import JSZip from "jszip";
import { readXml } from "../pptx/read";
import { parseSlide } from "../pptx/parse_slide";
import { DocJson } from "../types";

type SlideXml = Record<string, unknown>;
type RelsXml = Record<string, unknown>;

async function main(): Promise<void> {
  const zip = new JSZip();
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:cSld>
    <p:spTree>
      <p:pic>
        <p:blipFill>
          <a:blip r:embed="rId1">
            <a:extLst>
              <a:ext>
                <asvg:svgBlip r:embed="rId2"/>
              </a:ext>
            </a:extLst>
          </a:blip>
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
  <Relationship Id="rId2" Target="../media/image-2.svg" Type="http://schemas.microsoft.com/office/2016/relationships/image"/>
</Relationships>`;

  zip.file("ppt/slides/slide1.xml", slideXml);
  zip.file("ppt/slides/_rels/slide1.xml.rels", relsXml);
  zip.file("ppt/media/image-1.png", Buffer.from("png-preview"));
  zip.file("ppt/media/image-2.svg", Buffer.from("<svg></svg>"));

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  const archive = await JSZip.loadAsync(buffer);

  const parsedSlide = await readXml<SlideXml>(archive, "ppt/slides/slide1.xml");
  const parsedRels = await readXml<RelsXml>(
    archive,
    "ppt/slides/_rels/slide1.xml.rels",
  );

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-svg-test-"));
  const imagesDir = path.join(tempDir, "assets/images");
  await fs.mkdir(imagesDir, { recursive: true });

  const { elements } = await parseSlide(parsedSlide, {
    slideIndex: 1,
    rels: parsedRels,
    zipReadFile: async (zipPath: string) => {
      const file = archive.file(zipPath);
      if (!file) {
        throw new Error(`Missing media file: ${zipPath}`);
      }
      const data = await file.async("nodebuffer");
      return Buffer.from(data);
    },
    zipFileExists: (zipPath: string) => Boolean(archive.file(zipPath)),
    imagesDir,
    theme: {},
  });

  const doc: DocJson = {
    schemaVersion: 1,
    slideSize: { width: 10, height: 7.5, unit: "in" },
    slides: [
      {
        id: "s1",
        background: { type: "image", src: "backgrounds/slide-1.png" },
        elements,
      },
    ],
  };

  await fs.writeFile(
    path.join(tempDir, "doc.json"),
    JSON.stringify(doc, null, 2),
    "utf-8",
  );

  const outZip = new JSZip();
  const assetFiles = await fs.readdir(imagesDir);
  for (const asset of assetFiles) {
    const content = await fs.readFile(path.join(imagesDir, asset));
    outZip.file(`assets/images/${asset}`, content);
  }
  outZip.file("doc.json", JSON.stringify(doc, null, 2));

  const outBuffer = await outZip.generateAsync({ type: "nodebuffer" });
  const outArchive = await JSZip.loadAsync(outBuffer);
  const outFiles = Object.keys(outArchive.files);

  const svgAssets = outFiles.filter((file) => file.endsWith(".svg"));
  if (svgAssets.length !== 1) {
    throw new Error(`Expected 1 svg asset, got ${svgAssets.length}`);
  }
  const pngAssets = outFiles.filter(
    (file) => file.startsWith("assets/images/") && file.endsWith(".png"),
  );
  if (pngAssets.length !== 0) {
    throw new Error(`Expected no png assets, got ${pngAssets.join(", ")}`);
  }
  const imageElement = elements.find((el) => el.type === "image");
  if (!imageElement || !imageElement.src.endsWith(".svg")) {
    throw new Error("Expected doc.json image src to reference svg asset.");
  }

  await fs.rm(tempDir, { recursive: true, force: true });
  console.log("svgBlip test passed");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
