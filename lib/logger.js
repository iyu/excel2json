/**
 * @fileOverview logger
 * @name logger.js
 * @author Yuhei Aihara
 * https://github.com/iyu/excel2json
 */
/* eslint no-console:off */

'use strict';

const path = require('path');

const originalLogger = {
  COLOR: {
    BLACK: '\u001b[30m',
    RED: '\u001b[31m',
    GREEN: '\u001b[32m',
    YELLOW: '\u001b[33m',
    BLUE: '\u001b[34m',
    MAGENTA: '\u001b[35m',
    CYAN: '\u001b[36m',
    WHITE: '\u001b[37m',
    RESET: '\u001b[0m',
  },
  _getDateLine: function _getDateLine() {
    const date = new Date();
    const year = date.getFullYear();
    let month = date.getMonth() + 1;
    month = month > 10 ? month : `0${month}`;
    let day = date.getDate();
    day = day > 10 ? day : `0${day}`;
    const time = date.toLocaleTimeString();
    return `[${year}-${month}-${day} ${time}]`;
  },
  _prepareStackTrace: function _prepareStackTrace(err, stack) {
    const stackLine = stack[1];
    const filename = path.relative('./', stackLine.getFileName());
    return `(${filename}:${stackLine.getLineNumber()})`;
  },
  _getFileLineNumber: function _getFileLineNumber() {
    const obj = {};
    const original = Error.prepareStackTrace;
    Error.prepareStackTrace = this._prepareStackTrace;
    Error.captureStackTrace(obj, this._getFileLineNumber);
    const stack = obj.stack;
    Error.prepareStackTrace = original;

    return stack;
  },
  info: function info() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[INFO]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.GREEN);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.info.apply(null, arguments);
  },
  debug: function debug() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[DEBUG]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.BLUE);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.log.apply(null, arguments);
  },
  error: function error() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[ERROR]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.RED);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.error.apply(null, arguments);
  },
};
Object.freeze(originalLogger);

module.exports = originalLogger;
