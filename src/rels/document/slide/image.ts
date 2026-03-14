import JSZip from "jszip";
import { SweepOptions } from "../../../office";
import path from "path";
import { PNG } from "pngjs";
import jpeg from "jpeg-js";
import { ModifyReturn } from "../../../types";

export async function modifyImage(
  zip: JSZip,
  referencingRelsPath: string,
  imagePath: string,
  options: SweepOptions
): Promise<ModifyReturn> {
  const imageData = await zip.file(imagePath)?.async("nodebuffer");
  if (!imageData) {
    throw new Error(`File not found: ${imagePath}`);
  }

  const extension = path.extname(imagePath).toLowerCase();

  if (options.remove?.ppt?.image?.metadata) {
    switch (extension) {
      case ".jpg":
      case ".jpeg":
        const jpegImageData = jpeg.decode(imageData, {
          maxResolutionInMP: 1000000,
          maxMemoryUsageInMB: 1000000,
        });
        if (!jpegImageData) {
          throw new Error(`Failed to decode JPEG image: ${imagePath}`);
        }
        const jpegBuffer = jpeg.encode(jpegImageData, 100).data;
        zip.file(imagePath, jpegBuffer);
        break;
      case ".png":
        const pngImageData = PNG.sync.read(imageData);
        if (!pngImageData) {
          throw new Error(`Failed to decode PNG image: ${imagePath}`);
        }
        const pngBuffer = PNG.sync.write(pngImageData);
        zip.file(imagePath, pngBuffer);
        break;
      default:
        // Do nothing for unsupported formats
        break;
    }
  }

  return {
    images: [
      {
        slideIndexes: [], // Will be overwritten by the slide index
        name: path.basename(imagePath),
        internalPath: imagePath,
      },
    ],
  };
}
