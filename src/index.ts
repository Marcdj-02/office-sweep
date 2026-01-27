import JSZip from "jszip";
import fs from "fs";
import path from "path";
import { cwd } from "process";
import { getFileJson } from "./utils/xml";
import { modifyCoreProperties } from "./rels/core";
import { modifyExtendedProperties } from "./rels/app";
import { modifyThumbnail } from "./rels/thumbnail";
import { modifyDocument } from "./rels/document";
import { ModifyReturn, Image } from "./types";
import { generateHash } from "./utils/random";

const relationshipTypes: Record<
  string,
  (
    zip: JSZip,
    path: string,
    options: SweepOptions
  ) => Promise<ModifyReturn | void>
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
    destinationFilePath: string;
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
      hanging?: boolean;
    };
  };
  extract?: {
    destinationFolderPath: string;
    images?: boolean;
  };
};

export async function pptxSweep(sourcePath: string, options: SweepOptions) {
  // 2. Read the source file
  const buffer = fs.readFileSync(sourcePath);

  // 3. Create a zip object
  const zip = await JSZip.loadAsync(buffer);

  const rels = await getFileJson(zip, "_rels/.rels");

  const relationships: { Id: string; Type: string; Target: string }[] =
    rels.Relationships.Relationship.map((r: any) => r._attributes);

  let images: Image[] = [];

  for (const relationship of relationships) {
    const modifyFunction = relationshipTypes[relationship.Type];

    if (modifyFunction) {
      const result = await modifyFunction(zip, relationship.Target, options);

      if (result) {
        images = [...images, ...result.images];
      }
    }
  }

  // 4. Create a buffer of the updated zip
  const updatedBuffer = await zip.generateAsync({ type: "nodebuffer" });

  if (options.remove) {
    // 5. Write the updated buffer to the destination file
    fs.writeFileSync(options.remove.destinationFilePath, updatedBuffer);

    return {};
  } else if (options.extract) {
    // 6. Write all extract materials to the destination folder

    let finalImages: Array<{
      name: string;
      path: string;
      slideIndexes: number[];
    }> = [];

    if (options.extract.images) {
      const imagesPath = path.join(
        cwd(),
        options.extract.destinationFolderPath,
        "images"
      );

      // 6.1 Make an images directory
      fs.mkdirSync(imagesPath, { recursive: true });

      for (const image of images) {
        const imageData = await zip
          .file(image.internalPath)
          ?.async("nodebuffer");
        if (imageData) {
          fs.writeFileSync(path.join(imagesPath, image.name), imageData);
        }

        finalImages.push({
          name: image.name,
          path: path.join(imagesPath, image.name),
          slideIndexes: image.slideIndexes,
        });
      }
    }

    return {
      success: true,
      images: finalImages,
    };
  }

  return { success: false, error: "Invalid options" };
}

export async function unzip(
  sourcePath: string,
  destinationPath: string
): Promise<void> {
  const zip = new JSZip();
  const buffer = fs.readFileSync(sourcePath);

  await zip.loadAsync(buffer);

  // Ensure the destination directory exists
  fs.mkdirSync(destinationPath);

  const files = Object.keys(zip.files);
  for (const file of files) {
    const content = await zip.file(file)?.async("nodebuffer");
    if (content) {
      const destFilePath = `${destinationPath}/${file}`;
      fs.mkdirSync(
        destFilePath.substring(0, destFilePath.lastIndexOf("/")),

        { recursive: true }
      );
      fs.writeFileSync(destFilePath, content);
    }
  }
}
