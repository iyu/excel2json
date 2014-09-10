/**
 * @fileOverview excel2json main
 * @name index.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 * https://github.com/yuhei-a/excel2json
 */
var _ = require('lodash'),
    async = require('async');

var excelParser = require('./excel/parser'),
    logger = require('./logger');

function Excel2Json() {
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
        logger: undefined
    };
    this.logger = logger;
}

module.exports = new Excel2Json();

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
Excel2Json.prototype.setup = function(options) {
    this.logger.info('excel2json setup');
    _.extend(this.opts, options);

    if (this.opts.logger) {
        this.logger = this.opts.logger;
    }

    return this;
};

/**
 * parser
 * @private
 */
Excel2Json.prototype._parser = {
    number: function(d) {
        return Number(d);
    },
    boolean: function(d) {
        return !!d && d.toLowerCase() !== 'false';
    },
    date: function(d) {
        return Math.round(((Number(d) - 25569) * 24 + new Date().getTimezoneOffset() / 60) * 3600000);
    }
};

/**
 * format
 * @param {Array} cells
 * @private
 * @example
 * var cells = [
 *     { cell: 'A1', value: '{}' }, { cell: 'A4', value: '_id' },,,
 * ]
 */
Excel2Json.prototype._format = function(cells) {
    var _this = this,
        opts = {},
        idx = {},
        list = [];

    _.extend(opts, {
        attr_line: this.opts.attr_line,
        data_line: this.opts.data_line,
        ref_key: this.opts.ref_key
    });

    _.each(cells, function(cell) {
        if (cell.cell === _this.opts.option_cell) {
            var _opts;
            try {
                _opts = JSON.parse(cell.value) || {};
            } catch (e) {
                _opts = {};
            }
            _.extend(opts, _opts);
            return;
        }

        if (cell.row === opts.attr_line) {
            var type = cell.value.match(/:(\w+)$/);
            var keys = cell.value.replace(/:\w+$/, '').split('.');
            opts.format = opts.format || {};
            opts.format[cell.column] = {
                type: type && type[1],
                keys: keys
            };
            return;
        }

        var format = opts.format && opts.format[cell.column];
        if (cell.row < opts.data_line || !format) {
            return;
        }
        if (format.type && format.type.toLowerCase() === 'index') {
            idx[format.keys.join('.')] = parseInt(cell.value, 10);
            return;
        }

        if (cell.column === 'A') {
            list.push({});
        }

        var data = _.last(list);
        _.each(format.keys, function(_key, i) {
            var isArray = /^#/.test(_key),
                isSplitArray = /^\$/.test(_key);
            if (isArray) {
                _key = _key.replace(/^#/, '');
                data[_key] = data[_key] || [];
            }
            if (isSplitArray) {
                _key = _key.replace(/^\$/, '');
            }

            if (i + 1 !== format.keys.length) {
                if (isArray) {
                    var _idx = idx[format.keys.slice(0, i + 1).join('.')];
                    _idx = typeof _idx === 'number' ? _idx : data[_key].length;
                    data = data[_key][_idx] = data[_key][_idx] || {};
                    return;
                }
                data = data[_key] = data[_key] || {};
                return;
            }

            if (isArray) {
                var __key = data[_key].length;
                data = data[_key];
                _key = __key;
            }

            var type = format.type && format.type.toLowerCase();
            if (type === 'number' || type === 'num') {
                data[_key] = isSplitArray ? cell.value.split(',').map(_this._parser.number) : _this._parser.number(cell.value);
            } else if (type === 'boolean' || type === 'bool') {
                data[_key] = isSplitArray ? cell.value.split(',').map(_this._parser.boolean) : _this._parser.boolean(cell.value);
            } else if (type === 'date') {
                data[_key] = isSplitArray ? cell.value.split(',').map(_this._parser.date) : _this._parser.date(cell.value);
            } else {
                data[_key] = isSplitArray ? cell.value.split(',') : cell.value;
            }
        });
    });

    return {
        opts: opts,
        list: list
    };
};

/**
 * find origin data
 * @param dataMap
 * @param opts
 * @param data
 * @private
 */
Excel2Json.prototype._findOrigin = function(dataMap, opts, data) {
    var origin = dataMap[data.__ref];
    if (!origin || !opts.key) {
        this.logger.error('not found origin.', JSON.stringify(data));
        return;
    }

    var keys = opts.key.split('.');
    var __in = data.__in ? data.__in.split('.') : [];
    for (var i = 0; i < keys.length; i++) {
        if (/^#/.test(keys[i])) {
            var key = keys[i].replace(/^#/, '');
            var index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
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
        origin = origin[data.__key] = {};
    }

    return origin;
};

/**
 * excel parser main script
 * @param {String} filepath
 * @param {Array} sheets
 * @param {Function} callback
 */
Excel2Json.prototype.parse = function(filepath, sheets, callback) {
    var _this = this,
        list,
        excelData;

    async.series([
        function(next) {
            excelParser.execute(filepath, sheets, function(err, result) {
                if (err) {
                    return next(err);
                }

                excelData = result;
                next();
            });
        },
        function(next) {
            async.map(excelData, function(sheetData, _next) {
                var result;
                try {
                    result = _this._format(sheetData.cells);
                } catch (e) {
                    _this.logger.error(e.stack);
                    return _next(e);
                }

                async.setImmediate(function() {
                    _next(null, {
                        num: sheetData.num,
                        name: sheetData.name,
                        opts: result.opts,
                        list: result.list
                    });
                });
            }, function(err, result) {
                if (err) {
                    return next(err);
                }

                list = result;
                next();
            });
        }
    ], function(err) {
        if (err) {
            return callback(err);
        }

        callback(null, list);
    });
};

/**
 * sheetDatas to json
 * @param {Array} sheetDatas
 * @param {Function} callback
 */
Excel2Json.prototype.toJson = function(sheetDatas, callback) {
    var collectionMap = {};
    var optionMap = {};
    var errors = {};
    for (var i = 0; i < sheetDatas.length; i++) {
        var sheetData = sheetDatas[i];
        var opts = sheetData.opts;
        var name = opts.name || sheetData.name;
        var refKey = opts.ref_key;
        var dataMap = collectionMap[name] = collectionMap[name] || {};
        if (!optionMap[name]) {
            optionMap[name] = opts;
        } else {
            optionMap[name] = _.extend({}, opts, optionMap[name]);
            _.extend(optionMap[name].format, opts.format);
        }
        for (var j = 0; j < sheetData.list.length; j++) {
            var data = sheetData.list[j];
            if (!opts.type || opts.type === 'origin') {
                dataMap[data[refKey]] = data;
            } else {
                var origin = this._findOrigin(dataMap, opts, data);
                if (origin) {
                    delete data.__ref;
                    delete data.__in;
                    delete data.__key;
                    _.extend(origin, data);
                } else {
                    errors[name] = errors[name] || [];
                    errors[name].push(data);
                }
            }
        }
    }

    callback(Object.keys(errors).length ? errors : null, collectionMap, optionMap);
};
