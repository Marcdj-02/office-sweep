const { unzip } = require("../../dist/index.js");
const fs = require("fs");

const inputFolder = "./input";
const outputFolder = "./output/unzipped";

if (fs.existsSync(outputFolder)) {
  fs.rmSync(outputFolder, { recursive: true });
}

fs.mkdirSync(outputFolder);

// Read powerpoint directory for files names
const files = fs.readdirSync(inputFolder);

for (let i = 0; i < files.length; i++) {
  const file = files[i];
  const output = file.substring(0, file.lastIndexOf("."));

  unzip(`${inputFolder}/${file}`, `${outputFolder}/${output}`);
}
