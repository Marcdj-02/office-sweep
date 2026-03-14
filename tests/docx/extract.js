// import { sweep } from "./dist/index.js";
const { officeSweep } = require("../../dist/index.js");
const fs = require("fs");

const file = "3.docx";

async function main() {
  const result = await officeSweep(`./input/${file}`, {
    extract: {
      destinationFolderPath:
        `./output/extracted/${file}`,
      images: true,
    },
  });

  console.log(result);
}

main();