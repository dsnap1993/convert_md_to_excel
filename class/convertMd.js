'use strict';

class ConvertMd {
    static removeEmptyRows (targetList) {
        for (let index in targetList) {
            if (targetList[index] === '') {
                targetList.splice(index, 1);
            }
        }
    }

    static createIndexList (targetList) {
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
        return indexList;
    }

    static createRecords (indexList, targetList) {
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
        return records;
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

module.exports = ConvertMd;
module.exports.ConvertMd = ConvertMd;
module.exports.default = ConvertMd;