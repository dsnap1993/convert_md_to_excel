//const {createObjectCsvWriter} = require('csv-writer');
const fsExtra = require('fs-extra');
const iconv  = require('iconv-lite');
const util = require('util');

if (process.argv.length !== 4) {
    console.log('Error! : Please check parameters');
    return;
}

const readFile = process.argv[2];
const writeFile = process.argv[3];

const convertMdToCsv = (writeFile, readFile) => {
    (async () => {
        const contents = fsExtra.readFileSync(readFile);
    
        const contentsStr = contents.toString();
        const contentsList = contentsStr.split('\n');
        //console.log('contentsList[before] = ' + contentsList);

        for (let index in contentsList) {
            if (contentsList[index] === '') {
                contentsList.splice(index, 1);
            }
        }
        console.log('contentsList[after] = ' + contentsList);

        const indexList = [];
        var index = 0;
        while (index < contentsList.length) {
            const result = contentsList[index].indexOf('#');
            if (result !== -1) {
                indexList.push(index);
                index += 4;
            } else {
                index += 1;
            }
        }
        console.log('indexList = ' + indexList);
        console.log('=============================');

        //const records = [];
        const records = [
            [
                '項目番号',
                '大項目',
                '中項目',
                '小項目',
                '試験手順',
                '確認項目',
                '試験実施日',
                '試験実施者',
                '試験バージョン',
                '試験結果',
                '問処番号',
                '備考',
            ]
        ];

        for (let i in indexList) {
            const next = parseInt(i) + 1;
            var strOfWays = '"';
            var strOfItems = '"';
            const startIndex = indexList[i] + 4;
            const subList = contentsList.slice(startIndex, indexList[next]);

            for (let j = 0; j < subList.length; j += 1) {
                if (subList[j].search(/\d/) !== -1) {
                    strOfWays += subList[j] + '\n';
                }
                if (subList[j].search('-') !== -1) {
                    const subStr = subList[j].slice(1); 
                    strOfItems += subStr + '\n';
                }
            }
            strOfWays += '"';
            strOfItems += '"';
            
            const initialIndex = indexList[i];
            /*records.push({
                number: contentsList[initialIndex].slice(2),
                section_1: contentsList[initialIndex+1].slice(3),
                section_2: contentsList[initialIndex+2].slice(4),
                section_3: contentsList[initialIndex+3].slice(5),
                ways: strOfWays,
                items: strOfItems,
                date: '',
                personnel: '',
                version: '',
                result: '',
                ticket_number: '',
                comment: ''
            });*/
            records.push([
                contentsList[initialIndex].slice(2),
                contentsList[initialIndex+1].slice(3),
                contentsList[initialIndex+2].slice(4),
                contentsList[initialIndex+3].slice(5),
                strOfWays,
                strOfItems,
                '',
                '',
                '',
                '',
                '',
                ''
            ]);
            console.log('\n');
        }
    
        console.log('=============================');
        console.log('records = ' + JSON.stringify(records));

        fsExtra.writeFileSync(writeFile, "");
        var fd = await util.promisify(fsExtra.open)(writeFile, "w");
        const bufOfSpace = iconv.encode('\r\n','Shift_JIS');

        for (row of records) {
            for (element of row) {
                var buf = iconv.encode(element+',',"Shift_JIS");
                await util.promisify(fsExtra.write)(fd, buf, 0, buf.length);
            }
            await util.promisify(fsExtra.write)(fd, bufOfSpace, 0, bufOfSpace.length);
        }
    
        /*const csvWriter = createObjectCsvWriter({
            path: writeFile,
            header: [
                {id: 'number', title: '項目番号'},
                {id: 'section_1', title: '大項目'},
                {id: 'section_2', title: '中項目'},
                {id: 'section_3', title: '小項目'},
                {id: 'ways', title: '試験手順'},
                {id: 'items', title: '確認項目'},
                {id: 'date', title: '試験実施日'},
                {id: 'personnel', title: '試験実施者'},
                {id: 'version', title: '試験バージョン'},
                {id: 'result', title: '試験結果'},
                {id: 'ticket_number', title: '問処番号'},
                {id: 'comment', title: '備考'},
            ],
            encoding:'utf8',
            append :false,
        });
        //Write CSV file
        await csvWriter.writeRecords(records);*/

        console.log('\n');
        console.log('=============================');
        console.log('Completely suceeded! you got a converted csv file [' + writeFile + ']');
    })().catch(e => {
        console.log('ERROR: ' + e);
    });
};

convertMdToCsv(writeFile, readFile);