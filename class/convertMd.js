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
}

module.exports = ConvertMd;
module.exports.ConvertMd = ConvertMd;
module.exports.default = ConvertMd;