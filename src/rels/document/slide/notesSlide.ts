import JSZip from "jszip";
import { SweepOptions } from "../../../office";
import { getRelsPath } from "../../../utils/paths";

export async function modifyNotesSlide(
  zip: JSZip,
  referencingRelsPath: string,
  notesSlidePath: string,
  options: SweepOptions
): Promise<void> {
  if (!options.remove?.ppt?.notes) {
    return;
  }

  let referencingRelsPathFileContent = await zip
    .file(referencingRelsPath)
    ?.async("string");
  if (!referencingRelsPathFileContent) {
    throw new Error(`File not found: ${referencingRelsPath}`);
  }

  const pattern = new RegExp(
    `<Relationship[^>]*[^>]*Target="[./a-zA-Z]*?${notesSlidePath
      .split("/")
      .at(-1)}"[^>]*>`
  );

  referencingRelsPathFileContent = referencingRelsPathFileContent.replace(
    pattern,
    ""
  );

  zip.file(referencingRelsPath, referencingRelsPathFileContent);

  const relsPath = getRelsPath(notesSlidePath);

  zip.remove(notesSlidePath);
  zip.remove(relsPath);
}
