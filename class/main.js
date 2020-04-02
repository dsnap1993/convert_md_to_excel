const ConvertMdToXlsx = require('./convertMdToXlsx');
const ValidateParams = require('./validateParams');

const readFile = process.argv[2];
const writeFile = process.argv[3];

const message = {};
message.complete_params = ValidateParams.checkToCompleteParams(process.argv.length);
if (message.complete_params !== '') {
    message.read_file = ValidateParams.checkValidFile('md', readFile);
    message.write_file = ValidateParams.checkValidFile('xlsx', writeFile);
}

var result = true;
for (let key in message) {
    if (message[key] !== undefined) {
        console.log(message[key]);
        result = false;
    }
}

if (!result) {
    return;
}

ConvertMdToXlsx.convertToXlsx(writeFile, readFile);
