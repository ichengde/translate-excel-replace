const translate = require('./api');
const XLSX = require('xlsx');

const excelFilePath = './in.xlsx'; //process.argv.slice(2)[0];
let workbook = XLSX.readFile(excelFilePath);

console.log(process.argv.slice(2));
let work_array = process.argv.slice(2);
let first_sheet_name = workbook.SheetNames[0];

/* Get worksheet */
let worksheet = workbook.Sheets[first_sheet_name];

let isFinish = work_array.map((col) => {
    return work(col);
});

Promise.all(isFinish).then(() => {
    save();
});

function work(colE) {
    let max = worksheet['!ref'].match(/\d+$/g);
    let result = [];
    let row = 2;
    let desired_value = '';


    for (let i = 2; i <= max; i++) {
        let desired_cell = worksheet[colE + i];
        console.log('query' + colE + i);
        if (desired_cell) {

            desired_value = desired_cell.v;
            result.push(translate(desired_value, { from: 'en', to: 'zh-TW' }));
        } else {
            result.push('');
        }
    }

    return Promise.all(result).then((afterTranslate) => {
        for (let i = 2; i <= max; i++) {
            let desired_cell = worksheet[colE + i];
            desired_cell.v = afterTranslate[i - 2].text;
            console.log('get ' + colE + i);
        }
        return Promise.resolve();
    }).catch((err) => {
        if (err.code == 'BAD_NETWORK') {
            console.log('中断,网络太差');
        }
    });
}


function save() {
    XLSX.writeFile(workbook, 'out.xlsx');
}