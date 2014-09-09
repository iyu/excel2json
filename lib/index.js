var _ = require('lodash');
var async = require('async');

var excelParser = require('./excelParser');

function Excel2Json() {
    // Cell with a custom sheet option.
    this.optionCel = 'A1';
    this._optionCel = [ 0, 0 ];
    // Line with a data attribute.
    this.attrLine = '2';
    this._attrLine = 1;
    // Line with a data.
    this.dataLine = '4';
    this._dataLine = 3;
    // refKey
    this.refKey = '_id';

    return this;
}

module.exports = new Excel2Json();

/**
 *
 * @param _dataMap
 * @param option
 * @param content
 */
function findOrigin(_dataMap, option, content) {
    var origin = _dataMap[content.__ref];
    if (!origin || !option.key) {
        console.error('not found origin.', JSON.stringify(content));
        return;
    }

    var keys = option.key.split('.');
    var __in = content.__in ? content.__in.split('.') : [];
    for (var i = 0; i < keys.length; i++) {
        if (/^#/.test(keys[i])) {
            var key = keys[i].replace(/^#/, '');
            var index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
            if (!index) {
                console.error('not found index.', JSON.stringify(content));
                return;
            }
            origin[key] = origin[key] || [];
            origin = origin[key];
            origin[index] = origin[index] || {};
            origin = origin[index];
        } else if (keys[i] === '$') {
            origin = origin[__in[i]];
        } else if (i + 1 === keys.length) {
            origin[keys[i]] = origin[keys[i]] || (option.type === 'array' ? [] : {});
            origin = origin[keys[i]];
        } else {
            origin = origin[keys[i]];
        }
        if (!origin) {
            console.error('not found origin parts.', JSON.stringify(content));
            return;
        }
    }

    if (option.type === 'array') {
        if (!Array.isArray(origin)) {
            console.error('is not Array.', JSON.stringify(content));
            return;
        }
        origin.push({});
        origin = origin[origin.length - 1];
    } else if (option.type === 'map') {
        if (!content.__key) {
            console.error('not found __key.', JSON.stringify(content));
            return;
        }
        origin = origin[content.__key] = {};
    } else {
        console.error(option);
        return;
    }
    return origin;
}

/**
 * Conversion to index from the cell name
 * @param {String} celname
 * @example
 * var celname = 'A2';
 * var ret = celname2index(celname);
 * console.log(ret);
 * // > [0,1]
 */
Excel2Json.prototype.celname2index = function(celname) {
    var cel = celname.match(/(\D+)(\d+)/);
    if (!cel || !cel[1] || !cel[2]) {
        return;
    }
    var column = cel[1];
    var line = cel[2];
    var result = [ -1, line - 1 ];
    for (var i = 0; i < column.length; i++) {
        result[0] += Math.pow(26, column.length - i - 1) * (column[i].charCodeAt() - 64);
    }

    return result;
};

/**
 * Setup options
 * @param {Object} options
 */
Excel2Json.prototype.setup = function(options) {
    options = options || {};
    this.optionCel = options.optionCel || this.optionCel;
    this.attrLine = options.attrLine || this.attrLine;
    this.dataLine = options.dataLine || this.dataLine;
    this.refKey = options.refKey || this.refKey;

    this._optionCel = this.celname2index(this.optionCel);
    this._attrLine = this.attrLine - 1;
    this._dataLine = this.dataLine - 1;

    return this;
};

/**
 * format
 * @param {Array} options
 * @param {Array} lineData
 * @param {Function} callback
 */
Excel2Json.prototype.format = function(options, lineData, callback) {
    var data, err, isOrigin = false;
    try {
        data = {};
        var index = {};
        for (var i = 0; i < options.length; i++) {
            var opt = options[i];
            if (!opt || lineData[i] === '') {
                continue;
            }
            if (!/#/.test(opt)) {
                isOrigin = true;
            }
            var _lineData = lineData[i];
            if (/:number$/.test(opt)) {
                _lineData = Number(_lineData);
            } else if (/:boolean$/.test(opt)) {
                _lineData = _lineData === 'true';
            } else if (/:date$/.test(opt)) {
                _lineData = Math.round(((Number(_lineData) - 25569) * 24 + new Date().getTimezoneOffset() / 60) * 3600000);
            } else if (/:index$/.test(opt)) {
                index[opt.replace(':index', '')] = Number(_lineData);
                continue;
            }
            var opts = opt.replace(/:.+$/, '').split('.');
            var _data = data;
            for (var j = 0; j < opts.length; j++) {
                var key = opts[j].replace(/^#/, '');
                var isArray = /^#/.test(opts[j]);
                var _index;
                if (j === opts.length - 1) {
                    if (isArray) {
                        if (opts[j] !== '#') {
                            _data[key] = [_lineData];
                        } else {
                            _index = index[opts.slice(0, j).join('.')] || 0;
                            _data[_index] = [_lineData];
                        }
                    } else {
                        _data[key] = _lineData;
                    }
                } else {
                    if (isArray) {
                        if (opts[j + 1] !== '#') {
                            _index = index[opts.slice(0, j + 1).join('.')] || 0;
                            _data[key] = _data[key] || [];
                            _data[key][_index] = _data[key][_index] || {};
                            _data = _data[key][_index];
                        } else {
                            _data[key] = _data[key] || [];
                            _data = _data[key];
                        }
                    } else {
                        _data[key] = _data[key] || {};
                        _data = _data[key];
                    }
                }
            }
        }
    } catch(e) {
        err = new Error('format error.');
    } finally {
        callback(err, data, isOrigin);
    }
};

/**
 * extend
 * @param {Object} originData
 * @param {Object} subData
 */
Excel2Json.prototype.extend = function(originData, subData) {
    function extend(_originData, _subData) {
        if (Array.isArray(_subData)) {
            for (var i = 0; i < _subData.length; i++) {
                if (_subData[i] === undefined) {
                    continue;
                }
                if (!_originData[i]) {
                    _originData[i] = _subData[i];
                    continue;
                }
                if (Array.isArray(_subData[i])) {
                    extend(_originData[i], _subData[i]);
                    continue;
                }
                if (typeof _subData[i] === 'object') {
                    var isObject = true;
                    for (var _key in _subData[i]) {
                        if (typeof _subData[i][_key] !== 'object') {
                            isObject = false;
                            break;
                        }
                    }
                    if (isObject) {
                        extend(_originData[i], _subData[i]);
                        continue;
                    }
                }
                _originData.push(_subData[i]);
            }

            return _originData;
        }

        for (var key in _subData) {
            if (!_originData.hasOwnProperty(key)) {
                _originData[key] = _subData[key];
                continue;
            }
            if (typeof _subData[key] === 'object') {
                extend(_originData[key], _subData[key]);
            }
        }

        return _originData;
    }

    return extend(originData, subData);
};

/**
 * excel parser main script
 * @param {String} filepath
 * @param {Array} sheets
 * @param {Function} callback
 */
Excel2Json.prototype.parse = function(filepath, sheets, callback) {
    var _this = this;

    var list = [];
    var excelData;
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
            async.eachSeries(excelData, function(sheetData, _next) {
                var sheetOptions, attrOptions;
                try {
                    sheetOptions = JSON.parse(sheetData.contents[_this._optionCel[0]][_this._optionCel[1]]) || {};
                    attrOptions = sheetData.contents[sheetOptions.attrLine ? sheetOptions.attrLine - 1 : _this._attrLine] || [];
                } catch (e) {
                    console.error(e.stack);
                    return _next();
                }

                var data = sheetData.contents.slice(sheetOptions.dataLine ? sheetOptions.dataLine - 1 : _this._dataLine);
                var _list = [];
                async.eachSeries(data, function(lineData, __next) {
                    _this.format(attrOptions, lineData, function(err, result, isOrigin) {
                        if (err) {
                            return __next(err);
                        }

                        if (isOrigin) {
                            _list.push(result);
                            return setImmediate(function() {
                                __next();
                            });
                        }

                        var originData = _list[_list.length - 1];
                        _this.extend(originData, result);
                        __next();
                    });
                }, function(err) {
                    if (err) {
                        console.error(err.stack);
                        return _next();
                    }

                    list.push({
                        num: sheetData.num,
                        name: sheetData.name,
                        option: sheetOptions,
                        contents: _list
                    });
                    _next();
                });
            }, next);
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
    var errors = {};
    for (var i = 0; i < sheetDatas.length; i++) {
        var sheetData = sheetDatas[i];
        var name = sheetData.option.name || sheetData.name;
        var refKey = sheetData.option.refKey || this.refKey;
        var dataMap = collectionMap[name] = collectionMap[name] || {};
        for (var j = 0; j < sheetData.contents.length; j++) {
            var content = sheetData.contents[j];
            if (!sheetData.option.type || sheetData.option.type === 'origin') {
                dataMap[content[refKey]] = content;
            } else {
                var origin = findOrigin(dataMap, sheetData.option, content);
                if (origin) {
                    delete content.__ref;
                    delete content.__in;
                    delete content.__key;
                    _.extend(origin, content);
                } else {
                    errors[name] = errors[name] || [];
                    errors[name].push(content);
                }
            }
        }
    }

    callback(Object.keys(errors).length ? errors : null, collectionMap);
};
