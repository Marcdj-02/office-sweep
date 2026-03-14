// import { sweep } from "./dist/index.js";
const { pptxSweep } = require("../../dist/index.js");
const fs = require("fs");

const inputFolder = "./input";
const outputFolder = "./output/removed";

if (fs.existsSync(outputFolder)) {
  fs.rmSync(outputFolder, { recursive: true });
}

fs.mkdirSync(outputFolder);

// Read powerpoint directory for files names
const files = fs.readdirSync(inputFolder);

for (let i = 0; i < files.length; i++) {
  const file = files[i];

  pptxSweep(`${inputFolder}/${file}`, {
    remove: {
      destinationFilePath: `${outputFolder}/${file}`,
      core: {
        title: true,
        creator: true,
        lastModifiedBy: true,
        revision: true,
        created: true,
        modified: true,
      },
      notes: true,
      comments: {
        modern: true,
        legacy: true,
      },
      image: {
        metadata: true,
      },
      authors: true,
      view: true,
    },
  });
}
