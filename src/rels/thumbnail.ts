import JSZip from "jszip";
import { SweepOptions } from "../office";
import { ModifyReturn } from "../types";

export async function modifyThumbnail(
  zip: JSZip,
  thumbnailPath: string,
  options: SweepOptions
): Promise<void> {
  if (options.remove?.ppt?.thumbnail) {
    // 1. Remove the thumbnail file
    zip.remove(thumbnailPath);

    // 2. Construct pattern for removing the relationship
    const pattern = new RegExp(
      `<Relationship[^>]*[^>]*Target="${thumbnailPath}"[^>]*>`
    );

    //3. Read the relationships file
    const rels = await zip.file("_rels/.rels")?.async("string");
    if (!rels) {
      throw new Error(`File not found: _rels/.rels`);
    }

    // 4. Remove the relationship
    const newRels = rels.replace(pattern, "");

    // 5. Update the relationships file
    zip.file("_rels/.rels", newRels);
  }
}
