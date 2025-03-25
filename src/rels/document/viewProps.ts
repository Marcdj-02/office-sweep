import JSZip from "jszip";
import { SweepOptions } from "../..";

export async function modifyViewProps(
  zip: JSZip,
  referencingRelsPath: string,
  viewPropsPath: string,
  options: SweepOptions
): Promise<void> {
  if (options.remove?.view) {
    const referencingRelsFileContent = await zip
      .file(referencingRelsPath)
      ?.async("text");
    if (!referencingRelsFileContent) {
      throw new Error(`File not found: ${referencingRelsPath}`);
    }

    const updatedReferencingRelsFileContent =
      referencingRelsFileContent.replace(
        new RegExp(
          `<Relationship[^>]*?Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"[^>]*?\/>`
        ),
        ""
      );

    zip.file(referencingRelsPath, updatedReferencingRelsFileContent);

    zip.remove(viewPropsPath);
  }
}
