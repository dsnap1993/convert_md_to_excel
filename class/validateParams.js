'use strict';

class ValidateParams {
    static checkToCompleteParams (lengthOfArgv) {
        var message = '';
        if (lengthOfArgv !== 4) {
            message += 'Error! : Please check parameters.';
        }
        return message;
    }

    static checkValidFile (extension, file) {
        var message = '';
        const pattern = extension + '$';
        var regexp = new RegExp(pattern);
        if (!regexp.test(file)) {
            message += 'Error! : invalid file [' + file + ']. A file with ' + extension + ' is only vaild.';
        }
        return message;
    }
}

module.exports = ValidateParams;
module.exports.ValidateParams = ValidateParams;
module.exports.default = ValidateParams;