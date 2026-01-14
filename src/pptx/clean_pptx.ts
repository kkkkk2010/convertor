import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export async function createCleanPptx(
  zip: JSZip,
  slidePaths: string[],
  outPath: string,
): Promise<void> {
  for (const slidePath of slidePaths) {
    const file = zip.file(slidePath);
    if (!file) {
      continue;
    }
    const xml = await file.async("string");
    const cleaned = cleanSlideXml(xml);
    zip.file(slidePath, cleaned);
  }

  await fs.mkdir(path.dirname(outPath), { recursive: true });
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(outPath, buffer);
}

function cleanSlideXml(xml: string): string {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const pics = Array.from(doc.getElementsByTagName("p:pic"));
  for (const node of pics) {
    node.parentNode?.removeChild(node);
  }

  const shapes = Array.from(doc.getElementsByTagName("p:sp"));
  for (const sp of shapes) {
    const hasTextBody = sp.getElementsByTagName("p:txBody").length > 0;
    if (hasTextBody) {
      sp.parentNode?.removeChild(sp);
    }
  }

  return new XMLSerializer().serializeToString(doc);
}
