import JSZip from "jszip";
import { SweepOptions } from "../../office";

export async function modifyCommentAuthors(
  zip: JSZip,
  referencingRelsPath: string,
  commentAuthorsPath: string,
  options: SweepOptions
): Promise<void> {
  if (options.remove?.ppt?.authors || options.remove?.ppt?.comments) {
    const referencingRelsFileContent = await zip
      .file(referencingRelsPath)
      ?.async("text");
    if (!referencingRelsFileContent) {
      throw new Error(`File not found: ${referencingRelsPath}`);
    }

    const updatedReferencingRelsFileContent =
      referencingRelsFileContent.replace(
        new RegExp(
          `<Relationship[^>]*?Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"[^>]*?\/>`
        ),
        ""
      );

    zip.file(referencingRelsPath, updatedReferencingRelsFileContent);

    zip.remove(commentAuthorsPath);
  }
}
