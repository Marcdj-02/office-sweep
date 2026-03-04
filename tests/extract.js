// import { sweep } from "./dist/index.js";
const { pptxSweep } = require("../dist/index.js");
const fs = require("fs");

const file = "null-ref.pptx";

async function main() {
  const result = await pptxSweep(`./powerpoints/${file}`, {
    extract: {
      destinationFolderPath:
        "./powerpoints-extracted/a4660e69-82ef-493e-8c4b-ea88e2433005",
      images: true,
    },
  });

  console.log(result);
}

main();