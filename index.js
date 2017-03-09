const translate = require('./lib/api')
const XLSX = require('xlsx')
const config = require('./config')
const excelFilePath = './in.xlsx'; // process.argv.slice(2)[0]
let workbook = XLSX.readFile(excelFilePath)

console.log(process.argv.slice(2))
let work_array = process.argv.slice(2)
let first_sheet_name = workbook.SheetNames[0]

/* Get worksheet */
let worksheet = workbook.Sheets[first_sheet_name]

let max = worksheet['!ref'].match(/\d+$/g)
let result = []
let row = 2
let queryCount = 0
let rowPos = 2
console.log('config.maxSpeed:' + config.maxSpeed);
let startTime = new Date();
work_array.forEach((col) => {
    work(col)
})

function queryOne(cellPos) {
    return new Promise(function(resolve, reject) {
        let desired_cell = worksheet[cellPos]

        if (desired_cell) {
            translate(desired_cell.v, { from: 'en', to: 'zh-TW' }).then((translateResult) => {
                desired_cell.v = translateResult.text
                    // console.log(translateResult.text)

                resolve()
            }).catch((err) => {
                if (err.code == 'BAD_NETWORK') {
                    console.log('中断,网络太差')
                }
                reject()
            })
        }
    })
}


function save() {
    console.log('进行保存');
    let endTime = new Date();
    let runTime = endTime.getTime() - startTime.getTime()
    console.log('runTime:' + (runTime / 1000) + 's');
    XLSX.writeFile(workbook, 'out.xlsx')
}

function run(colE, i) {
    let pSet = []
    let limitCondition = 0;
    let c = 0;

    if (i + config.maxSpeed < max)
        limitCondition = i + config.maxSpeed;
    else
        limitCondition = max;

    console.log('process:' + limitCondition);

    for (c = i; c <= limitCondition; c++) {
        pSet.push(queryOne(colE + c))
    }
    Promise.all(pSet).then(() => {
        if (i >= limitCondition) {
            save()
            return;
        } else {
            run(colE, limitCondition)
        }
    })
}

function work(colE) {


    run(colE, 2)
}