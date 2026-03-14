import JSZip from "jszip";
import fs from "fs";

export async function unzip(
    sourcePath: string,
    destinationPath: string
  ): Promise<void> {
    const zip = new JSZip();
    const buffer = fs.readFileSync(sourcePath);
  
    await zip.loadAsync(buffer);
  
    // Ensure the destination directory exists
    fs.mkdirSync(destinationPath);
  
    const files = Object.keys(zip.files);
    for (const file of files) {
      const content = await zip.file(file)?.async("nodebuffer");
      if (content) {
        const destFilePath = `${destinationPath}/${file}`;
        fs.mkdirSync(
          destFilePath.substring(0, destFilePath.lastIndexOf("/")),
  
          { recursive: true }
        );
        fs.writeFileSync(destFilePath, content);
      }
    }
  }
  