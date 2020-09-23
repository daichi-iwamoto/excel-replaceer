// モジュールのインストール
const xlsx = require('xlsx');
const utils = xlsx.utils;

// エクセルファイルの読み込み
const testX = xlsx.readFile('./test/test.xlsx');

// シートの読み込み
const sheet = testX.Sheets['Sheet1'];

// セルの範囲の取得
const range = sheet['!ref'];
console.log(range);

// セルの範囲を数値化
const rangeN = utils.decode_range(range);
console.log(rangeN);

// セルの値の取得 1 : セル名で取得
const cell01 = sheet['A1'];
console.log(cell01);

// セルの値取得 2 : utilsを使用してアドレス指定
const address02 = utils.encode_cell({r:0, c:0});
const cell02 = sheet[address02];
console.log(cell02);

// 範囲分ループで取得
for (let r = rangeN.s.r; r <= rangeN.e.r; r++) {
    for (let c = rangeN.s.c; c <= rangeN.e.c; c++) {
        let address = utils.encode_cell({c:c, r:r});
        let cell = sheet[address];
        console.log(`${address} : ${cell.v}`);
    }
}
