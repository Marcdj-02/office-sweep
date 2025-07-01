const { unzip } = require("../dist/index.js");
const fs = require("fs");

// Read powerpoint directory for files names
const files = fs.readdirSync("./powerpoints");

for (let i = 0; i < files.length; i++) {
  const file = files[i];
  const output = file.replace(".pptx", "");

  unzip(`./powerpoints/${file}`, `./powerpoints-unzipped/${output}`);
}
