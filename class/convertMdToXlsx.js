'use strict';

const ConvertMd = require('./convertMd');
const fsExtra = require('fs-extra');
const Excel = require('exceljs');

class ConvertMdToXlsx extends ConvertMd {
    static convertToXlsx(writeFile, readFile) {
        (async () => {
            console.log('start converting [' + readFile + '] to [' + writeFile + ']...');
            var workbook = new Excel.Workbook();
    
            var sheet1 = workbook.addWorksheet('入力チェック');
            var sheet2 = workbook.addWorksheet('表示チェック');
            var sheet3 = workbook.addWorksheet('画面制御');
            var sheet4 = workbook.addWorksheet('画面遷移');
            var sheet5 = workbook.addWorksheet('APIマッピング');
            var sheet6 = workbook.addWorksheet('権限チェック');
    
            const contents = fsExtra.readFileSync(readFile);
        
            const contentsStr = contents.toString();
            const contentsList = contentsStr.split('\n');
            //console.log('contentsList = ' + contentsList);
    
            const indexList = this.createIndexListForSeparatingContentsList(contentsList);
            
            // Separate contentsList to list for each sheet
            const contentsListInput = contentsList.slice(0, indexList[0]);
            const contentsListDisplay = contentsList.slice(indexList[0], indexList[1]);
            const contentsListControl = contentsList.slice(indexList[1], indexList[2]);
            const contentsListMoving = contentsList.slice(indexList[2], indexList[3]);
            const contentsListAPI = contentsList.slice(indexList[3], indexList[4]);
            const contentsListAuth = contentsList.slice(indexList[4]);
    
            this.removeEmptyRows(contentsListInput);
            this.removeEmptyRows(contentsListDisplay);
            this.removeEmptyRows(contentsListControl);
            this.removeEmptyRows(contentsListMoving);
            this.removeEmptyRows(contentsListAPI);
            this.removeEmptyRows(contentsListAuth);
    
            const indexListForInput = this.createIndexList(contentsListInput);
            const indexListForDisplay = this.createIndexList(contentsListDisplay);
            const indexListForControl = this.createIndexList(contentsListControl);
            const indexListForMoving = this.createIndexList(contentsListMoving);
            const indexListForAPI = this.createIndexList(contentsListAPI);
            const indexListForAuth = this.createIndexList(contentsListAuth);
    
            const recordsForInput = this.createRecords(indexListForInput, contentsListInput);
            const recordsForDisplay = this.createRecords(indexListForDisplay, contentsListDisplay);
            const recordsForControl = this.createRecords(indexListForControl, contentsListControl);
            const recordsForMoving = this.createRecords(indexListForMoving, contentsListMoving);
            const recordsForAPI = this.createRecords(indexListForAPI, contentsListAPI);
            const recordsForAuth = this.createRecords(indexListForAuth, contentsListAuth);
    
            this.writeRecords(sheet1, recordsForInput);
            this.writeRecords(sheet2, recordsForDisplay);
            this.writeRecords(sheet3, recordsForControl);
            this.writeRecords(sheet4, recordsForMoving);
            this.writeRecords(sheet5, recordsForAPI);
            this.writeRecords(sheet6, recordsForAuth);
    
            const sheetInput = workbook.getWorksheet('入力チェック');
            const sheetDisplay = workbook.getWorksheet('表示チェック');
            const sheetControl = workbook.getWorksheet('画面制御');
            const sheetMoving = workbook.getWorksheet('画面遷移');
            const sheetAPI = workbook.getWorksheet('APIマッピング');
            const sheetAuth = workbook.getWorksheet('権限チェック');
    
            const lengthOfInput = recordsForInput.length;
            const lengthOfDisplay = recordsForDisplay.length;
            const lengthOfControl = recordsForControl.length;
            const lengthOfMoving = recordsForMoving.length;
            const lengthOfAPI = recordsForAPI.length;
            const lengthOfAuth = recordsForAuth.length;
    
            this.formatSheet(sheetInput, lengthOfInput);
            this.formatSheet(sheetDisplay, lengthOfDisplay);
            this.formatSheet(sheetControl, lengthOfControl);
            this.formatSheet(sheetMoving, lengthOfMoving);
            this.formatSheet(sheetAPI, lengthOfAPI);
            this.formatSheet(sheetAuth, lengthOfAuth);
    
            await workbook.xlsx.writeFile(writeFile);
    
            //console.log('\n');
            console.log('=============================');
            console.log('Completely suceeded! you got a converted excel file [' + writeFile + ']');
        })().catch(e => {
            console.log('ERROR: ' + e);
        });
    }

    static createIndexListForSeparatingContentsList (contentsList) {
        const indexList = [];
        var point = 0;
        for (let i = 0; i < 5; i += 1) {
            const result = contentsList.indexOf('===', point);
            if (result !== -1) {
                indexList.push(result);
                point = result + 1;
            }
        }
        //console.log('indexList[raw data] = ' + indexList);
        return indexList;
    }

    static writeRecords (sheet, records) {
        var row = 1;

        for (let element of records) {
            for (let column = 1; column <= 12; column += 1) {
                const index = parseInt(column) - 1;
                sheet.getCell(row, column).value = element[index];
            }
            row += 1;
        }
    }

    static formatSheet (worksheet, lengthOfRow) {
        // a width of each column
        const columnArray = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
        const columnWidthArray = [12.5, 15, 15, 15, 10, 56.88, 38.25, 15, 15, 15, 15, 15, 15];

        const borderStyles = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
        };
        const fillStyles = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'006495ed'}
        };
        const fontStyles = {
            bold: true
        };

        for (let index in columnArray) {
            var column = worksheet.getColumn(columnArray[index]);
            column.width = columnWidthArray[index];
            column.alignment = {
                vertical: 'top',
                horizontal: 'left',
                wrapText: true
            };
            column.border = borderStyles;

            var cellOfRow1 = worksheet.getCell(columnArray[index]+1);
            cellOfRow1.fill = fillStyles;
            cellOfRow1.font = fontStyles;
        }

        const limit = Math.max(10, lengthOfRow);

        for (let index = 1; index <= limit; index += 1) {
            var row = worksheet.getRow(index);
            row.height = 37.5;
        }
    }
}

module.exports = ConvertMdToXlsx;
module.exports.ConvertMdToXlsx = ConvertMdToXlsx;
module.exports.default = ConvertMdToXlsx;