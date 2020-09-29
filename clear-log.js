const fs = require("fs");

try {
  fs.writeFileSync("./log/replace-log.txt", "", "utf8");
  console.log("clear log file (´-ω-`)");
}
catch (err) {
  throw err;
}