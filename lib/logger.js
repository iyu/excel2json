/**
 * @fileOverview logger
 * @name logger.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 * https://github.com/yuhei-a/excel2json
 */
var path = require('path');

var COLOR = {
    BLACK: '\u001b[30m',
    RED: '\u001b[31m',
    GREEN: '\u001b[32m',
    YELLOW: '\u001b[33m',
    BLUE: '\u001b[34m',
    MAGENTA: '\u001b[35m',
    CYAN: '\u001b[36m',
    WHITE: '\u001b[37m',

    RESET: '\u001b[0m'
};

module.exports = {
    _getDateLine: function() {
        var date = new Date();
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        month = month > 10 ? month : '0' + month;
        var day = date.getDate();
        day = day > 10 ? day : '0' + day;
        var time = date.toLocaleTimeString();
        return '[' + year + '-' + month + '-' + day + ' ' + time + ']';
    },
    _prepareStackTrace: function(err, stack) {
        var stackLine = stack[1];
        var filename = path.relative('./', stackLine.getFileName());
        return '(' + filename + ':' + stackLine.getLineNumber() + ')';
    },
    _getFileLineNumber: function() {
        var obj = {};
        var original = Error.prepareStackTrace;
        Error.prepareStackTrace = this._prepareStackTrace;
        Error.captureStackTrace(obj, this._getFileLineNumber);
        var stack = obj.stack;
        Error.prepareStackTrace = original;

        return stack;
    },
    info: function() {
        Array.prototype.unshift.call(arguments, COLOR.RESET);
        Array.prototype.unshift.call(arguments, '[INFO]');
        Array.prototype.unshift.call(arguments, this._getDateLine());
        Array.prototype.unshift.call(arguments, COLOR.GREEN);
        Array.prototype.push.call(arguments, this._getFileLineNumber());
        console.log.apply(null, arguments);
    },
    error: function() {
        Array.prototype.unshift.call(arguments, COLOR.RESET);
        Array.prototype.unshift.call(arguments, '[ERROR]');
        Array.prototype.unshift.call(arguments, this._getDateLine());
        Array.prototype.unshift.call(arguments, COLOR.RED);
        Array.prototype.push.call(arguments, this._getFileLineNumber());
        console.log.apply(null, arguments);
    }
};
