/**
 * @fileOverview excel2json main
 * @name index
 * @author Yuhei Aihara
 * https://github.com/iyu/excel2json
 */

import _ from 'lodash';

import excelParser from './excel/parser';
import logger, { ILogger } from './logger';

interface Opts {
  // Cell with a custom sheet option.
  option_cell?: string; // eslint-disable-line camelcase
  // Line with a data attribute.
  attr_line?: number; // eslint-disable-line camelcase
  // Line with a data.
  data_line?: number; // eslint-disable-line camelcase
  // ref key
  ref_key?: string; // eslint-disable-line camelcase
  // Custom logger.
  logger?: ILogger;
}

interface DefaultOpts {
  option_cell: string; // eslint-disable-line camelcase
  attr_line: number; // eslint-disable-line camelcase
  data_line: number; // eslint-disable-line camelcase
  ref_key: string; // eslint-disable-line camelcase
  logger?: ILogger;
}

interface Cell {
  cell: string;
  column: string;
  row: number;
  value: string;
}

interface CellOpts {
  attr_line: number; // eslint-disable-line camelcase
  data_line: number; // eslint-disable-line camelcase
  ref_key: string; // eslint-disable-line camelcase
  format?: { [key:string]: { type: string|null; key: string; keys: string[] } };
  name?: string;
  type?: string;
  key?: string;
}

interface ParseResult {
  num: number;
  name: string;
  opts: CellOpts;
  list: any[];
}
interface ParseError {
  num: number;
  name: string;
  error: Error;
}

class Excel2Json {
  public opts: DefaultOpts = {
    option_cell: 'A1',
    attr_line: 2,
    data_line: 4,
    ref_key: '_id',
  }

  public logger: ILogger = logger

  private _parser: { [key: string]: (d: string) => number|boolean|string|any; } = {
    number: (d: string) => {
      if (d.length >= 18) {
        // IEEE754
        return Number(Number(d).toFixed(8));
      }
      return Number(d);
    },
    num: (d: string) => {
      return this._parser.number(d);
    },
    boolean: (d: string) => {
      return !!d && d.toLowerCase() !== 'false' && d !== '0';
    },
    bool: (d: string) => {
      return this._parser.boolean(d);
    },
    date: (d: string) => {
      return Math.round(
        (
          ((Number(d) - 25569) * 24)
          + (new Date().getTimezoneOffset() / 60)
        ) * 3600000,
      );
    },
    auto: (d: string) => {
      return Number.isFinite(Number(d)) ? d : this._parser.number(d);
    },
  }

  /**
   * setup
   * @param {Object} options
   * @example
   * var options = {
   *     option_cell: 'A1',
   *     attr_line: 2,
   *     data_line: 4,
   *     ref_key: '_id',
   *     logger: CustomLogger
   * };
   */
  setup(options: Opts) {
    _.assign(this.opts, options);

    if (this.opts.logger) {
      this.logger = this.opts.logger;
    }

    this.logger.info('excel2json setup');
    return this;
  }

  /**
   * format
   * @param {Array} cells
   * @private
   * @example
   * var cells = [
   *     { cell: 'A1', value: '{}' }, { cell: 'A4', value: '_id' },,,
   * ]
   */
  _format(cells: Cell[]) {
    let beforeRow: number;
    let idx: { [key: string]: { type: string; value: number } } = {};
    const list: any[] = [];

    const opts: CellOpts = {
      attr_line: this.opts.attr_line,
      data_line: this.opts.data_line,
      ref_key: this.opts.ref_key,
      format: undefined,
    };

    _.forEach(cells, (cell) => {
      if (beforeRow !== cell.row) {
        _.forEach(idx, (i) => {
          if (i.type !== 'format') {
            i.value += 1;
          }
        });
        beforeRow = cell.row;
      }
      if (cell.cell === this.opts.option_cell) {
        const _opts = JSON.parse(cell.value);
        _.assign(opts, _opts);
        return;
      }

      if (cell.row === opts.attr_line) {
        const type = cell.value.match(/:(\w+)$/);
        const key = cell.value.replace(/:\w+$/, '');
        const keys = key.split('.');

        opts.format = opts.format || {};
        opts.format[cell.column] = {
          type: type && type[1],
          key,
          keys,
        };
        return;
      }

      const format = opts.format && opts.format[cell.column];
      let data: any;
      let _idx: any;

      if (cell.row < opts.data_line || !format) {
        return;
      }
      if (format.type && format.type.toLowerCase() === 'index') {
        _idx = parseInt(cell.value, 10);
        if (!idx[format.key] || idx[format.key].value !== _idx) {
          idx[format.key] = {
            type: 'format',
            value: _idx,
          };
          _.forEach(idx, (i, key) => {
            if (new RegExp(`^${format.key}.+$`).test(key)) {
              idx[key].value = 0;
            }
          });
        }
        return;
      }
      if (format.key === opts.ref_key || format.key === '__ref') {
        idx = {};
        list.push({});
      }

      data = _.last(list);
      _.forEach(format.keys, (_key, i) => {
        const isArray = /^#/.test(_key);
        const isSplitArray = /^\$/.test(_key);
        let __key;
        if (isArray) {
          _key = _key.replace(/^#/, '');
          data[_key] = data[_key] || [];
        }
        if (isSplitArray) {
          _key = _key.replace(/^\$/, '');
        }

        if (i + 1 !== format.keys.length) {
          if (isArray) {
            __key = format.keys.slice(0, i + 1).join('.');
            _idx = idx[__key];
            if (!_idx) {
              idx[__key] = {
                type: 'normal',
                value: data[_key].length ? data[_key].length - 1 : 0,
              };
              _idx = idx[__key];
            }
            data[_key][_idx.value] = data[_key][_idx.value] || {};
            data = data[_key][_idx.value];
            return;
          }
          data[_key] = data[_key] || {};
          data = data[_key];
          return;
        }

        if (isArray) {
          __key = format.keys.slice(0, i + 1).join('.');
          _idx = idx[__key];
          if (!_idx) {
            idx[__key] = {
              type: 'normal',
              value: data[_key].length ? data[_key].length - 1 : 0,
            };
            _idx = idx[__key];
          }
          data = data[_key];
          _key = _idx.value;
        }

        if (data[_key]) {
          return;
        }

        const type = format.type && format.type.toLowerCase();
        if (type && this._parser[type]) {
          data[_key] = isSplitArray ? cell.value.split(',').map(this._parser[type]) : this._parser[type](cell.value);
        } else {
          data[_key] = isSplitArray ? cell.value.split(',') : cell.value;
        }
      });
    });

    return {
      opts,
      list,
    };
  }

  /**
   * find origin data
   * @param dataMap
   * @param opts
   * @param data
   * @private
   */
  _findOrigin(dataMap: any, opts: CellOpts, data: any) {
    let origin = dataMap[data.__ref];
    if (!origin || !opts.key) {
      this.logger.error('not found origin.', JSON.stringify(data));
      return;
    }

    const keys = opts.key.split('.');
    const __in = data.__in ? data.__in.split('.') : [];
    for (let i = 0; i < keys.length; i++) {
      if (/^#/.test(keys[i])) {
        const key = keys[i].replace(/^#/, '');
        const index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
        if (!index) {
          this.logger.error('not found index.', JSON.stringify(data));
          return;
        }
        origin[key] = origin[key] || [];
        origin = origin[key];
        origin[index] = origin[index] || {};
        origin = origin[index];
      } else if (keys[i] === '$') {
        origin = origin[__in[i]];
      } else if (i + 1 === keys.length) {
        origin[keys[i]] = origin[keys[i]] || (opts.type === 'array' ? [] : {});
        origin = origin[keys[i]];
      } else {
        origin[keys[i]] = origin[keys[i]] || {};
        origin = origin[keys[i]];
      }
      if (!origin) {
        this.logger.error('not found origin parts.', JSON.stringify(data));
        return;
      }
    }

    if (opts.type === 'array') {
      if (!Array.isArray(origin)) {
        this.logger.error('is not Array.', JSON.stringify(data));
        return;
      }
      origin.push({});
      origin = origin[origin.length - 1];
    } else if (opts.type === 'map') {
      if (!data.__key) {
        this.logger.error('not found __key.', JSON.stringify(data));
        return;
      }
      origin[data.__key] = {};
      origin = origin[data.__key];
    }

    return origin;
  }

  /**
   * excel parser main script
   * @param {String} filepath
   * @param {Array} sheets
   * @param {Function} callback
   */
  async parse(
    filepath: string,
    sheets: [],
    callback: (err: Error | null, result?: ParseResult[], errList?: ParseError[]) => void,
  ) {
    let excelData;
    try {
      excelData = await excelParser.execute(filepath, sheets);
    } catch (e) {
      return callback(e);
    }
    let errList: ParseError[] | undefined;
    const promises = _.map(excelData, (sheetData) => {
      let result: any;
      try {
        result = this._format(sheetData.cells);
      } catch (e) {
        this.logger.error('invalid sheet format.', sheetData.num, sheetData.name);
        errList = errList || [];
        errList.push({
          num: sheetData.num,
          name: sheetData.name,
          error: e,
        });
        return Promise.resolve();
      }

      return new Promise((resolve) => {
        setImmediate(() => {
          resolve({
            num: sheetData.num,
            name: sheetData.name,
            opts: result.opts,
            list: result.list,
          });
        });
      });
    });

    let result: ParseResult[];
    try {
      result = await Promise.all(promises) as ParseResult[];
    } catch (e) {
      return callback(e);
    }

    callback(null, _.compact(result), errList);
  }

  /**
   * sheetDatas to json
   * @param {Array} sheetDatas
   * @param {Function} callback
   */
  toJson(sheetDatas: ParseResult[], callback: Function) {
    const collectionMap: { [key:string]: any } = {};
    const optionMap: { [key:string]: CellOpts; } = {};
    const errors: { [key:string]: any } = {};
    for (let i = 0; i < sheetDatas.length; i++) {
      const sheetData = sheetDatas[i];
      const { opts } = sheetData;
      const name = opts.name || sheetData.name;
      const refKey = opts.ref_key;
      collectionMap[name] = collectionMap[name] || {};
      const dataMap = collectionMap[name];
      if (!optionMap[name]) {
        optionMap[name] = opts;
      } else {
        optionMap[name] = _.assign({}, opts, optionMap[name]);
        _.assign(optionMap[name].format, opts.format);
      }
      for (let j = 0; j < sheetData.list.length; j++) {
        const data = sheetData.list[j];
        if (!opts.type || opts.type === 'origin') {
          dataMap[data[refKey]] = data;
        } else {
          const origin = this._findOrigin(dataMap, opts, data);
          if (origin) {
            delete data.__ref;
            delete data.__in;
            delete data.__key;
            _.assign(origin, data);
          } else {
            errors[name] = errors[name] || [];
            errors[name].push(data);
          }
        }
      }
    }

    callback(_.isEmpty(errors) ? null : errors, collectionMap, optionMap);
  }
}

export default new Excel2Json();
