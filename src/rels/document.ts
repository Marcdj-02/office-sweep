import JSZip from "jszip";
import { SweepOptions } from "..";
import { getRelsPath } from "../utils/paths";
import { getFileJson } from "../utils/xml";
import { modifySlide } from "./document/slide";
import { modifyNotesMaster } from "./document/notesMaster";
import { modifyTableStyles } from "./document/tableStyles";
import { modifyTheme } from "./document/theme";
import { modifyViewProps } from "./document/viewProps";
import { modifySlideMaster } from "./document/slideMaster";
import { modifyPresProps } from "./document/presProps";
import { modifyHandoutMaster } from "./document/handoutMaster";
import { modifyAuthors } from "./document/authors";
import { modifyCommentAuthors } from "./document/commentAuthors";
import { ModifyReturn, Image } from "../types";
import fs from "fs";

const relationshipTypes: Record<
  string,
  (
    zip: JSZip,
    referencingRelsPath: string,
    path: string,
    options: SweepOptions
  ) => Promise<ModifyReturn | void>
> = {
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide":
    modifySlide,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster":
    modifyNotesMaster,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles":
    modifyTableStyles,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
    modifyTheme,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps":
    modifyViewProps,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
    modifySlideMaster,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps":
    modifyPresProps,
  "http://schemas.microsoft.com/office/2018/10/relationships/authors":
    modifyAuthors,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster":
    modifyHandoutMaster,
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors":
    modifyCommentAuthors,
};

export async function modifyDocument(
  zip: JSZip,
  documentPath: string,
  options: SweepOptions
): Promise<ModifyReturn> {
  const json = await getFileJson(zip, documentPath);

  const slideIdList = json["p:presentation"]["p:sldIdLst"]["p:sldId"].map(
    (s: any) => s["_attributes"]["r:id"]
  );

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
        `ppt/${relationship.Target}`,
        options
      );

      if (result) {
        for (const image of result.images) {
          const existingImage = images.find(
            (i: Image) => i.internalPath === image.internalPath
          );
          if (existingImage) {
            existingImage.slideIndexes.push(slideIndex);
          } else {
            images.push({ ...image, slideIndexes: [slideIndex] });
          }
        }
      }
    }
  }

  return { images };
}
