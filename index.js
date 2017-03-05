const translate = require('./api')
const XLSX = require('xlsx')

const excelFilePath = './in.xlsx'; // process.argv.slice(2)[0]
let workbook = XLSX.readFile(excelFilePath)

console.log(process.argv.slice(2))
let work_array = process.argv.slice(2)
let first_sheet_name = workbook.SheetNames[0]

/* Get worksheet */
let worksheet = workbook.Sheets[first_sheet_name]

work_array.forEach((col) => {
  work(col)
})

function catOne (cellPos) {
  let desired_cell = worksheet[cellPos]

  if (desired_cell) {
    console.log(desired_cell.v)
  }
}

function queryOne (cellPos) {
  return new Promise(function (resolve, reject) {
    let desired_cell = worksheet[cellPos]

    if (desired_cell) {
      translate(desired_cell.v, { from: 'en', to: 'zh-TW' }).then((translateResult) => {
        desired_cell.v = translateResult
        console.log(translateResult.text)

        resolve(cellPos)
      }).catch((err) => {
        if (err.code == 'BAD_NETWORK') {
          console.log('中断,网络太差')
        }
        reject()
      })
    }
  })
}

function work (colE) {
  let max = worksheet['!ref'].match(/\d+$/g)
  let result = []
  let row = 2
  let queryCount = 0
  let rowPos = 2

  function run (colE, i) {
    let pSet = []
    for (let c = i;i + 50 < max ? c < i + 50 : c < max;c++) {
      pSet.push(queryOne(colE + c))
    }
    Promise.all(pSet).then(() => {
      if (i + 50 >= max) {
        save()
      }
      run(colE, i + 50)
    })
  }

  run(colE, 2)
}

function save () {
  XLSX.writeFile(workbook, 'out.xlsx')
}
