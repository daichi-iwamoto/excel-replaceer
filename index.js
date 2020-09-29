// モジュールのインストール
const xlsx = require("xlsx");
const utils = xlsx.utils;
const fs = require("fs");

// エクセルファイルの読み込み
const testX = xlsx.readFileSync("./test/test.xlsx");

// シートの読み込み
const sheet = testX.Sheets["Sheet1"];

// セルの範囲の取得
const range = sheet["!ref"];

// セルの範囲を数値化
const rangeN = utils.decode_range(range);

// log start
try {
  const start_date = new Date();
  const log_start = `----- Start : ${start_date.getFullYear()}/${start_date.getMonth()}/${start_date.getDate()} ${start_date.getHours()}:${start_date.getMinutes()}:${start_date.getSeconds()}:${start_date.getMilliseconds()} -----\n`;
  fs.appendFileSync("./log/replace-log.txt", log_start, "utf8");
} catch (err) {
  throw err;
}

// ループ処理
for (let r = rangeN.s.r; r <= rangeN.e.r; r++) {
  // 置換ファイル取得
  let address = utils.encode_cell({ c: 0, r: r });
  let cell = sheet[address];

  // 置換ファイルの読み込み
  let data;
  try {
    data = fs.readFileSync("./test/" + cell.v, "utf-8");
  } 
  // 置換ファイルがない場合
  catch (err) {
    try {
      console.log(`【 ${cell.v} 】file not found (+_+)`);
      fs.appendFileSync("./log/replace-log.txt", `【 ${cell.v} 】file not found\n`, "utf8");
    } catch (err) {
      throw err;
    }
    continue;
  }

  // 変更前の内容を取得
  let B_address = utils.encode_cell({ c: 1, r: r });
  let B_cell = sheet[B_address];

  // 変更後の内容を取得
  let A_address = utils.encode_cell({ c: 2, r: r });
  let A_cell = sheet[A_address];

  // 置換
  const beforeTxt = data;
  const afterTxt = beforeTxt.replace(
    "<title>" + B_cell.v + "</title>",
    "<title>" + A_cell.v + "</title>"
  );

  // 置換が行われていた場合
  if (beforeTxt !== afterTxt) {
    // ファイルの上書き
    try {
      fs.writeFileSync("./test/" + cell.v, afterTxt, "utf8");
      try {
        console.log(`【 ${cell.v} 】done (/・ω・)/`);
        fs.appendFileSync("./log/replace-log.txt", `【 ${cell.v} 】done\n`, "utf8");
      } catch (err) {
        throw err;
      }
    }
    
    // ファイルの上書きに失敗した場合
    catch (err) {
      try {
        console.log(`【 ${cell.v} 】file write err (゜-゜)!?`);
        fs.appendFileSync("./log/replace-log.txt", `【 ${cell.v} 】file write err\n`, "utf8");
      } catch (err) {
        throw err;
      }
    }
  }
  
  // 置換が行われなかった場合
  else {;
    try {
      console.log(`【 ${cell.v} 】no replace (>_<)`);
      fs.appendFileSync("./log/replace-log.txt", `【 ${cell.v} 】no replace\n`, "utf8");
    } catch (err) {
      throw err;
    }
  }
}

// log end
try {
  const end_date = new Date();
  const log_end = `----- End : ${end_date.getFullYear()}/${end_date.getMonth()}/${end_date.getDate()} ${end_date.getHours()}:${end_date.getMinutes()}:${end_date.getSeconds()}:${end_date.getMilliseconds()} -----\n\n`;
  fs.appendFileSync("./log/replace-log.txt", log_end, "utf8");
} catch (err) {
  throw err;
}