const xlsx = require('xlsx');

class ExcelReader {
    constructor() {
        this.workbook = xlsx.readFile("./assets/item_database.xlsm");
    }

    getManufactures() {
        this.worksheet = this.workbook.Sheets["Manufacturers"];
        return this.worksheet['A10'];
    }
}

module.exports = ExcelReader
