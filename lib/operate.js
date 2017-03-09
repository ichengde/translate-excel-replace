function catOne(cellPos) {
    let desired_cell = worksheet[cellPos]

    if (desired_cell) {
        console.log(desired_cell.v)
    }
}
module.exports = catOne;