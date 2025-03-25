import JSZip from "jszip";
import fs from "fs";
import { getFileJson } from "./utils/xml";
import { modifyCoreProperties } from "./rels/core";
import { modifyExtendedProperties } from "./rels/app";
import { modifyThumbnail } from "./rels/thumbnail";
import { modifyDocument } from "./rels/document";

const relationshipTypes: Record<
  string,
  (zip: JSZip, path: string, options: SweepOptions) => Promise<void>
> = {
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties":
    modifyCoreProperties,
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail":
    modifyThumbnail,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
    modifyDocument,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties":
    modifyExtendedProperties,
};

export type SweepOptions = {
  remove?: {
    totalTime?: boolean;
    core?: {
      title?: boolean;
      creator?: boolean;
      lastModifiedBy?: boolean;
      revision?: boolean;
      created?: boolean;
      modified?: boolean;
    };
    thumbnail?: boolean;
    comments?: {
      legacy?: boolean;
      modern?: boolean;
    };
    notes?: boolean;
    authors?: boolean;
    view?: boolean;
    image?: {
      metadata?: boolean;
    };
  };
};

export async function pptxSweep(
  sourcePath: string,
  destinationPath: string,
  options: SweepOptions
) {
  // 2. Read the source file
  const buffer = fs.readFileSync(sourcePath);

  // 3. Create a zip object
  const zip = await JSZip.loadAsync(buffer);

  const rels = await getFileJson(zip, "_rels/.rels");

  const relationships: { Id: string; Type: string; Target: string }[] =
    rels.Relationships.Relationship.map((r: any) => r._attributes);

  for (const relationship of relationships) {
    const modifyFunction = relationshipTypes[relationship.Type];

    if (modifyFunction) {
      await modifyFunction(zip, relationship.Target, options);
    }
  }

  // 4. Create a buffer of the updated zip
  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });

  // 5. Write the updated buffer to the destination file
  fs.writeFileSync(destinationPath, updatedBuffer);
}
