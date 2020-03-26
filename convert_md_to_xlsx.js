const fsExtra = require('fs-extra');
const Excel = require('exceljs');

if (process.argv.length !== 4) {
    console.log('Error! : Please check parameters');
    return;
}

const readFile = process.argv[2];
const writeFile = process.argv[3];

const convertMdToXlsx = (writeFile, readFile) => {
    (async () => {
        console.log('start converting [' + readFile + '] to [' + writeFile + ']...');
        var workbook = new Excel.Workbook();

        /*workbook.creator = "Me";
        workbook.lastModifiedBy = "Her";
        workbook.created = new Date(1985, 8, 30);
        workbook.modified = new Date();*/

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

        const contentsListInput = contentsList.slice(0, indexList[0]);
        const contentsListDisplay = contentsList.slice(indexList[0], indexList[1]);
        const contentsListControl = contentsList.slice(indexList[1], indexList[2]);
        const contentsListMoving = contentsList.slice(indexList[2], indexList[3]);
        const contentsListAPI = contentsList.slice(indexList[3], indexList[4]);
        const contentsListAuth = contentsList.slice(indexList[4]);

        removeEmptyRows(contentsListInput);
        removeEmptyRows(contentsListDisplay);
        removeEmptyRows(contentsListControl);
        removeEmptyRows(contentsListMoving);
        removeEmptyRows(contentsListAPI);
        removeEmptyRows(contentsListAuth);

        const indexListForInput = createIndexList(contentsListInput);
        const indexListForDisplay = createIndexList(contentsListDisplay);
        const indexListForControl = createIndexList(contentsListControl);
        const indexListForMoving = createIndexList(contentsListMoving);
        const indexListForAPI = createIndexList(contentsListAPI);
        const indexListForAuth = createIndexList(contentsListAuth);

        const recordsForInput = createRecords(indexListForInput, contentsListInput);
        const recordsForDisplay = createRecords(indexListForDisplay, contentsListDisplay);
        const recordsForControl = createRecords(indexListForControl, contentsListControl);
        const recordsForMoving = createRecords(indexListForMoving, contentsListMoving);
        const recordsForAPI = createRecords(indexListForAPI, contentsListAPI);
        const recordsForAuth = createRecords(indexListForAuth, contentsListAuth);

        writeRecords(sheet1, recordsForInput);
        writeRecords(sheet2, recordsForDisplay);
        writeRecords(sheet3, recordsForControl);
        writeRecords(sheet4, recordsForMoving);
        writeRecords(sheet5, recordsForAPI);
        writeRecords(sheet6, recordsForAuth);

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

        formatSheet(sheetInput, lengthOfInput);
        formatSheet(sheetDisplay, lengthOfDisplay);
        formatSheet(sheetControl, lengthOfControl);
        formatSheet(sheetMoving, lengthOfMoving);
        formatSheet(sheetAPI, lengthOfAPI);
        formatSheet(sheetAuth, lengthOfAuth);

        await workbook.xlsx.writeFile(writeFile);

        //console.log('\n');
        console.log('=============================');
        console.log('Completely suceeded! you got a converted excel file [' + writeFile + ']');
    })().catch(e => {
        console.log('ERROR: ' + e);
    });
};

const removeEmptyRows = (targetList) => {
    for (let index in targetList) {
        if (targetList[index] === '') {
            targetList.splice(index, 1);
        }
    }
    //console.log('targetList = ' + targetList);
}

const createIndexList = (targetList) => {
    const indexList = [];
    var index = 0;
    while (index < targetList.length) {
        const result = targetList[index].indexOf('#');
        if (result !== -1) {
            indexList.push(index);
            index += 4;
        } else {
            index += 1;
        }
    }
    //console.log('indexList = ' + indexList);
    //console.log('=============================');
    return indexList;
};

const createRecords = (indexList, targetList) => {
    const records = [
        [
            "項目番号",
            "大項目",
            "中項目",
            "小項目",
            "種別",
            "試験手順",
            "確認項目",
            "試験実施日",
            "試験実施者",
            "試験バージョン",
            "試験結果",
            "問処番号",
            "備考",
        ]
    ];

    for (let i in indexList) {
        const next = parseInt(i) + 1;
        var strOfWays = "";
        var strOfItems = "";
        const startIndex = indexList[i] + 4;
        const subList = targetList.slice(startIndex, indexList[next]);

        for (let j = 0; j < subList.length; j += 1) {
            if (subList[j].search(/\d/) !== -1) {
                strOfWays += subList[j] + "\n";
            }
            if (subList[j].search("-") !== -1) {
                const subStr = subList[j].slice(1); 
                strOfItems += subStr + "\n";
            }
        }
        strOfWays.replace(/\n/g,"\r\n");
        strOfItems.replace(/\n/g,"\r\n");
        
        const initialIndex = indexList[i];
        records.push([
            parseInt(i)+1,
            targetList[initialIndex].slice(2),
            targetList[initialIndex+1].slice(3),
            targetList[initialIndex+2].slice(4),
            targetList[initialIndex+3].slice(5),
            strOfWays,
            strOfItems,
            "",
            "",
            "",
            "",
            "",
            ""
        ]);
    }
    //console.log('\n');
    //console.log('records = ' + records);
    return records;
};

const writeRecords = (sheet, records) => {
    var row = 1;
    //console.log('[writeRecords] records = ' + records);

    for (element of records) {
        for (let column = 1; column <= 12; column += 1) {
            const index = parseInt(column) - 1;
            sheet.getCell(row, column).value = element[index];
        }
        row += 1;
    }
}

const formatSheet = (worksheet, lengthOfRow) => {
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

convertMdToXlsx(writeFile, readFile);