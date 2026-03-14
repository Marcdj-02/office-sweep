import JSZip from "jszip";
import { SweepOptions } from "../office";
import { Image, ModifyReturn } from "../types";
import { getRelsPath } from "../utils/paths";
import { getFileJson } from "../utils/xml";
import { modifyAuthors } from "./document/authors";
import { modifyCommentAuthors } from "./document/commentAuthors";
import { modifySlide } from "./document/slide";
import { modifyLegacyComments } from "./document/slide/comments/legacyComments";
import { modifyImage } from "./document/slide/image";
import { modifyViewProps } from "./document/viewProps";

const relationshipTypes: Record<
  string,
  ((
    zip: JSZip,
    referencingRelsPath: string,
    path: string,
    options: SweepOptions
  ) => Promise<ModifyReturn | void>) | null
> = {
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide":
    modifySlide,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster":
    null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles":
    null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
    null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps":
    modifyViewProps,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
    null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps":
    null,
  "http://schemas.microsoft.com/office/2018/10/relationships/authors":
    modifyAuthors,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster":
    null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors":
    modifyCommentAuthors,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
    modifyImage,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
    modifyLegacyComments,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings": null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable": null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering": null,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles": null,
};

export async function modifyDocument(
  zip: JSZip,
  documentPath: string,
  options: SweepOptions
): Promise<ModifyReturn> {
  const json = await getFileJson(zip, documentPath);

  const documentType = (() => {
    if (documentPath.includes("presentation.xml")) return "ppt";
    return "word";
  })()

  const slideIdList = (() => {
    if (documentType !== 'ppt') return [];

    return json["p:presentation"]["p:sldIdLst"]["p:sldId"].map(
      (s: any) => s["_attributes"]["r:id"]
    )
  })()

  const relsPath = getRelsPath(documentPath);

  const rels = await getFileJson(zip, relsPath);

  const relationships: { Id: string; Type: string; Target: string }[] =
    rels.Relationships.Relationship.map((r: any) => r._attributes);

  let images: Image[] = [];

  for (const relationship of relationships) {
    const modifyFunction = relationshipTypes[relationship.Type];

    if (modifyFunction) {
      const slideIndex = slideIdList.findIndex(
        (id: string) => id === relationship.Id
      );

      const result = await modifyFunction(
        zip,
        relsPath,
        `${documentType}/${relationship.Target}`,
        options
      );

      if (result) {
        for (const image of result.images) {
          const existingImage = images.find(
            (i: Image) => i.internalPath === image.internalPath
          );
          if (existingImage) {
            if (slideIndex === -1) {
              continue;
            }

            existingImage.slideIndexes.push(slideIndex);
          } else {
            const newImage = slideIndex === -1 ? image : { ...image, slideIndexes: [slideIndex] };
            
            images.push(newImage);
          }
        }
      }
    }
  }

  if (options.remove?.word?.comments) {
    if (documentType !== 'word') {
      throw new Error("Cannot remove word comments on non-word document");
    }

    let documentContent = await zip.file(documentPath)?.async("string");
    if (!documentContent) {
      throw new Error(`File not found: ${documentPath}`);
    }

    const replacePatterns = [
      /<w:commentRangeStart[^>]*?\/>/,
      /<w:commentRangeEnd[^>]*?\/>/,
      /<w:commentReference[^>]*?\/>/,
    ];

    for (const pattern of replacePatterns) {
      documentContent = documentContent.replace(pattern, "");
    }

    console.log(documentContent);

    zip.file(documentPath, documentContent);
  }

  return { images };
}
