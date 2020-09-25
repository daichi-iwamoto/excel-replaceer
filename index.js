// モジュールのインストール
const xlsx = require('xlsx');
const utils = xlsx.utils;

const fs = require('fs');

// エクセルファイルの読み込み
const testX = xlsx.readFile('./test/test.xlsx');

// シートの読み込み
const sheet = testX.Sheets['Sheet1'];

// セルの範囲の取得
const range = sheet['!ref'];
// console.log(range);

// セルの範囲を数値化
const rangeN = utils.decode_range(range);

// ループ処理
for (let r = rangeN.s.r; r <= rangeN.e.r; r++) {
    // ファイル名取得
    let address = utils.encode_cell({c:0, r:r});
    let cell = sheet[address];

    // htmlの読み込み
    fs.readFile('./test/' + cell.v, 'utf-8', (err, data) => {
        // エラー処理
        if (err) {
            console.log(`【 ${cell.v} 】ファイル読み込みエラー`);
            throw err;
        }

        // 置換処理
        // 変更前の内容を取得
        let B_address = utils.encode_cell({c:1, r:r});
        let B_cell = sheet[B_address];

        // 変更後の内容を取得
        let A_address = utils.encode_cell({c:2, r:r}); 
        let A_cell = sheet[A_address];

        // 置換
        const beforeTxt = data;
        const afterTxt = beforeTxt.replace(new RegExp(B_cell.v,"g"), A_cell.v);
        
        // ファイルの上書き
        fs.writeFile('./test/' + cell.v, afterTxt, (err) => {
            if (err) {
                console.log(`【 ${cell.v} 】ファイル置換エラー`);
                throw err;
            }

            console.log(`【 ${cell.v} 】success !`);
        });
    });
}
