import JSZip from "jszip";
import { SweepOptions } from "../../..";
import path from "path";
import { ExifTool } from "exiftool-vendored";
import { PNG } from "pngjs";
import jpeg from "jpeg-js";

const TEMP_DIR = "/temp";

export async function modifyImage(
  zip: JSZip,
  referencingRelsPath: string,
  imagePath: string,
  options: SweepOptions
): Promise<void> {
  if (!options.remove?.image?.metadata) return;

  const imageData = await zip.file(imagePath)?.async("nodebuffer");
  if (!imageData) {
    throw new Error(`File not found: ${imagePath}`);
  }

  const extension = path.extname(imagePath).toLowerCase();
  console.log(`Image extension: ${extension}`);

  switch (extension) {
    case ".jpg":
    case ".jpeg":
      const jpegImageData = jpeg.decode(imageData);
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
