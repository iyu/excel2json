var async = require('async');
var excel = require('excel');

function Parser() {
    // Cell with a custom sheet option.
    this.optionCel = 'A1';
    this._optionCel = [ 0, 0 ];
    // Line with a data attribute.
    this.attrLine = '2';
    this._attrLine = 1;
    // Line with a data.
    this.dataLine = '4';
    this._dataLine = 3;

    return this;
}

module.exports = new Parser();

/**
 * Conversion to index from the cell name
 * @param {String} celname
 * @example
 * var celname = 'A2';
 * var ret = celname2index(celname);
 * console.log(ret);
 * // > [0,1]
 */
Parser.prototype.celname2index = function(celname) {
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
Parser.prototype.setup = function(options) {
    options = options || {};
    this.optionCel = options.optionCel || this.optionCel;
    this.attrLine = options.attrLine || this.attrLine;
    this.dataLine = options.dataLine || this.dataLine;

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
Parser.prototype.format = function(options, lineData, callback) {
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
Parser.prototype.extend = function(originData, subData) {
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
                continue;
            }
        }

        return _originData;
    }

    return extend(originData, subData);
};

/**
 * excel parser main script
 * @param {String} filepath
 * @param {Function} callback
 */
Parser.prototype.parse = function(filepath, callback) {
    var self = this;
    var list = [];

    var excelData, sheetOptions, attrOptions;
    async.series([
        function(next) {
            excel(filepath, function(err, result) {
                if (err) {
                    return next(err);
                }

                excelData = result;
                sheetOptions = JSON.parse(excelData[self._optionCel[0]][self._optionCel[1]]) || {};
                attrOptions = excelData[sheetOptions.attrLine ? sheetOptions.attrLine - 1 : self._attrLine] || [];

                next();
            });
        },
        function(next) {
            var _excelData = excelData.slice(sheetOptions.dataLine ? sheetOptions.dataLine - 1 : self._dataLine);
            async.each(_excelData, function(lineData, _next) {
                self.format(attrOptions, lineData, function(err, result, isOrigin) {
                    if (err) {
                        return _next(err);
                    }

                    if (isOrigin) {
                        list.push(result);
                        return setImmediate(function() {
                            _next();
                        });
                    }

                    var originData = list[list.length - 1];
                    self.extend(originData, result);
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
