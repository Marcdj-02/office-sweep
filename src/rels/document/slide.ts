import JSZip from "jszip";
import { SweepOptions } from "../../office";
import { addRelativePath, getRelsPath } from "../../utils/paths";
import { getFileJson } from "../../utils/xml";
import { modifyNotesSlide } from "./slide/notesSlide";
import { modifyModernComments } from "./slide/comments/modernComments";
import { modifyLegacyComments } from "./slide/comments/legacyComments";
import { modifyImage } from "./slide/image";
import { ModifyReturn, Image } from "../../types";

const relationshipTypes: Record<
  string,
  | ((
      zip: JSZip,
      referencingRelsPath: string,
      path: string,
      options: SweepOptions
    ) => Promise<ModifyReturn | void>)
  | undefined
> = {
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
    modifyImage,
  "http://schemas.microsoft.com/office/2018/10/relationships/comments":
    modifyModernComments,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
    modifyLegacyComments,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
    modifyNotesSlide,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
    undefined,
};

export async function modifySlide(
  zip: JSZip,
  referencingRelsPath: string,
  slidePath: string,
  options: SweepOptions
): Promise<ModifyReturn> {
  const relsPath = getRelsPath(slidePath);

  const rels = await getFileJson(zip, relsPath);

  const relationships: { Id: string; Type: string; Target: string }[] =
    Array.isArray(rels.Relationships.Relationship)
      ? rels.Relationships.Relationship.map((r: any) => r._attributes)
      : [rels.Relationships.Relationship._attributes];

  let images: Image[] = [];

  for (const relationship of relationships) {
    if (relationship.Target === "NULL"){
      continue;
    }
    
    const modifyFunction = relationshipTypes[relationship.Type];

    if (modifyFunction) {
      const result = await modifyFunction(
        zip,
        relsPath,
        addRelativePath(slidePath, relationship.Target),
        options
      );

      if (result) {
        images = [...images, ...result.images];
      }
    }
  }

  if (options.remove?.ppt?.comments) {
    let slideContent = await zip.file(slidePath)?.async("string");
    if (!slideContent) {
      throw new Error(`File not found: ${slidePath}`);
    }

    const replacePatterns = [
      /<p188:commentRel[^>]*?\/>/,
    ];

    for (const pattern of replacePatterns) {
      slideContent = slideContent.replace(pattern, "");
    }

    zip.file(slidePath, slideContent);
  }

  return { images };
}
