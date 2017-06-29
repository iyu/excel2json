/**
 * @fileOverview excel2json main
 * @name index.js
 * @author Yuhei Aihara
 * https://github.com/iyu/excel2json
 */

'use strict';

const _ = require('lodash');
const async = require('neo-async');

const excelParser = require('./excel/parser');
const logger = require('./logger');

class Excel2Json {
  constructor() {
    this.opts = {
      // Cell with a custom sheet option.
      option_cell: 'A1',
      // Line with a data attribute.
      attr_line: 2,
      // Line with a data.
      data_line: 4,
      // ref key
      ref_key: '_id',
      // Custom logger.
      logger: undefined,
    };
    this.logger = logger;

    this._parser = {
      number: (d) => {
        if (d.length >= 18) {
          // IEEE754
          return Number(Number(d).toFixed(8));
        }
        return Number(d);
      },
      num: (d) => {
        return this._parser.number(d);
      },
      boolean: (d) => {
        return !!d && d.toLowerCase() !== 'false' && d !== '0';
      },
      bool: (d) => {
        return this._parser.boolean(d);
      },
      date: (d) => {
        return Math.round(
          (
            ((Number(d) - 25569) * 24) +
            (new Date().getTimezoneOffset() / 60)
          ) * 3600000);
      },
      auto: (d) => {
        return isNaN(d) ? d : this._parser.number(d);
      },
    };
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
  setup(options) {
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
  _format(cells) {
    const opts = {};
    let beforeRow;
    let idx = {};
    const list = [];

    _.assign(opts, {
      attr_line: this.opts.attr_line,
      data_line: this.opts.data_line,
      ref_key: this.opts.ref_key,
    });

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
      let data;
      let _idx;

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
        if (this._parser[type]) {
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
  _findOrigin(dataMap, opts, data) {
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
  parse(filepath, sheets, callback) {
    async.angelFall([
      (next) => {
        excelParser.execute(filepath, sheets, next);
      },
      (excelData, next) => {
        let errList;
        async.map(excelData, (sheetData, done) => {
          let result;
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
            return done();
          }

          async.setImmediate(() => {
            done(null, {
              num: sheetData.num,
              name: sheetData.name,
              opts: result.opts,
              list: result.list,
            });
          });
        }, (err, result) => {
          if (err) {
            return next(err);
          }

          next(null, _.compact(result), errList);
        });
      },
    ], (err, list, errList) => {
      if (err) {
        return callback(err);
      }

      callback(null, list, errList);
    });
  }

  /**
   * sheetDatas to json
   * @param {Array} sheetDatas
   * @param {Function} callback
   */
  toJson(sheetDatas, callback) {
    const collectionMap = {};
    const optionMap = {};
    const errors = {};
    for (let i = 0; i < sheetDatas.length; i++) {
      const sheetData = sheetDatas[i];
      const opts = sheetData.opts;
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

module.exports = new Excel2Json();
