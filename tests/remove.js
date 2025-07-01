// import { sweep } from "./dist/index.js";
const { pptxSweep } = require("../dist/index.js");
const fs = require("fs");

// Read powerpoint directory for files names
const files = fs.readdirSync("./powerpoints");

for (let i = 0; i < files.length; i++) {
  const file = files[i];
  const output = file.replace(".pptx", "-swept2.pptx");

  pptxSweep(`./powerpoints/${file}`, `./powerpoints-swept/${output}`, {
    remove: {
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
