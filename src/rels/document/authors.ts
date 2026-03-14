import JSZip from "jszip";
import { SweepOptions } from "../../office";

export async function modifyAuthors(
  zip: JSZip,
  referencingRelsPath: string,
  authorsPath: string,
  options: SweepOptions
): Promise<void> {
  if (options.remove?.ppt?.authors) {
    const referencingRelsFileContent = await zip
      .file(referencingRelsPath)
      ?.async("text");
    if (!referencingRelsFileContent) {
      throw new Error(`File not found: ${referencingRelsPath}`);
    }

    const updatedReferencingRelsFileContent =
      referencingRelsFileContent.replace(
        new RegExp(
          `<Relationship[^>]*?Type="http://schemas.microsoft.com/office/2018/10/relationships/authors"[^>]*?\/>`
        ),
        ""
      );

    zip.file(referencingRelsPath, updatedReferencingRelsFileContent);

    zip.remove(authorsPath);
  }
}
