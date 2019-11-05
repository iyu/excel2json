/**
 * @fileOverview logger
 * @name logger
 * @author Yuhei Aihara
 * https://github.com/iyu/excel2json
 */
/* eslint no-console:off */

import path from 'path';

export interface ILogger {
  info(...args: any[]): void;
  debug(...args: any[]): void;
  error(...args: any[]): void;
}

class OriginalLogger implements ILogger {
  public COLOR = {
    BLACK: '\u001b[30m',
    RED: '\u001b[31m',
    GREEN: '\u001b[32m',
    YELLOW: '\u001b[33m',
    BLUE: '\u001b[34m',
    MAGENTA: '\u001b[35m',
    CYAN: '\u001b[36m',
    WHITE: '\u001b[37m',
    RESET: '\u001b[0m',
  }

  private _getDateLine() {
    const date = new Date();
    const year = date.getFullYear();
    let month: string | number = date.getMonth() + 1;
    month = month > 10 ? month : `0${month}`;
    let day: string | number = date.getDate();
    day = day > 10 ? day : `0${day}`;
    const time = date.toLocaleTimeString();
    return `[${year}-${month}-${day} ${time}]`;
  }

  private _prepareStackTrace(err: Error, stack: any) {
    const stackLine = stack[1];
    const filename = path.relative('./', stackLine.getFileName());
    return `(${filename}:${stackLine.getLineNumber()})`;
  }

  private _getFileLineNumber() {
    const obj: any = {};
    const original = Error.prepareStackTrace;
    Error.prepareStackTrace = this._prepareStackTrace;
    Error.captureStackTrace(obj, this._getFileLineNumber);
    const { stack } = obj;
    Error.prepareStackTrace = original;

    return stack;
  }

  info(...args: any[]) {
    Array.prototype.unshift.call(args, this.COLOR.RESET);
    Array.prototype.unshift.call(args, '[INFO]');
    Array.prototype.unshift.call(args, this._getDateLine());
    Array.prototype.unshift.call(args, this.COLOR.GREEN);
    Array.prototype.push.call(args, this._getFileLineNumber());
    console.info.call(null, ...args);
  }

  debug(...args: any[]) {
    Array.prototype.unshift.call(args, this.COLOR.RESET);
    Array.prototype.unshift.call(args, '[DEBUG]');
    Array.prototype.unshift.call(args, this._getDateLine());
    Array.prototype.unshift.call(args, this.COLOR.BLUE);
    Array.prototype.push.call(args, this._getFileLineNumber());
    console.log.call(null, ...args);
  }

  error(...args: any[]) {
    Array.prototype.unshift.call(args, this.COLOR.RESET);
    Array.prototype.unshift.call(args, '[ERROR]');
    Array.prototype.unshift.call(args, this._getDateLine());
    Array.prototype.unshift.call(args, this.COLOR.RED);
    Array.prototype.push.call(args, this._getFileLineNumber());
    console.error.call(null, ...args);
  }
}

export default new OriginalLogger();
