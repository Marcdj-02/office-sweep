import JSZip from "jszip";
import { SweepOptions } from "../../../../office";
import { getRelsPath } from "../../../../utils/paths";

export async function modifyLegacyComments(
  zip: JSZip,
  referencingRelsPath: string,
  commentPath: string,
  options: SweepOptions
): Promise<void> {
  if (
    !options.remove?.ppt?.comments?.legacy &&
    !options.remove?.word?.comments
  ) {
    return;
  }

  let referencingRelsPathFileContent = await zip
    .file(referencingRelsPath)
    ?.async("string");
  if (!referencingRelsPathFileContent) {
    throw new Error(`File not found: ${referencingRelsPath}`);
  }

  const pattern = new RegExp(
    `<Relationship[^>]*[^>]*Target="[./a-zA-Z]*?${commentPath
      .split("/")
      .at(-1)}"[^>]*>`
  );

  referencingRelsPathFileContent = referencingRelsPathFileContent.replace(
    pattern,
    ""
  );

  zip.file(referencingRelsPath, referencingRelsPathFileContent);

  const relsPath = getRelsPath(commentPath);

  zip.remove(commentPath);
  zip.remove(relsPath);
}
