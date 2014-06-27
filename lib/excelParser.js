var fs = require('fs');

var _ = require('lodash'),
    async = require('async'),
    libxmljs = require('libxmljs'),
    JSZip = require('node-zip');

var XMLNS = { a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' };

function ExcelParser() {
}

module.exports = new ExcelParser();

/**
 * @param {Array} cells
 */
function calculateDimensions (cells) {
    var comparator = function (a, b) { return a - b; },
        allRows = _.map(cells, function (cell) { return cell.row; }).sort(comparator),
        allCols = _.map(cells, function (cell) { return cell.column; }).sort(comparator),
        minRow = allRows[0],
        maxRow = _.last(allRows),
        minCol = allCols[0],
        maxCol = _.last(allCols);

    return [
        { row: minRow, column: minCol },
        { row: maxRow, column: maxCol }
    ];
}

/**
 * @param {String} col
 */
function colToInt(col) {
    var letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    col = col.trim().split('');

    var n = 0;
    for (var i = 0; i < col.length; i++) {
        n *= 26;
        n += letters.indexOf(col[i]);
    }

    return n;
}

/**
 * @param {String} cell
 */
function CellCoords(cell) {
    var cells = cell.split(/([0-9]+)/);

    this.row = parseInt(cells[1]);
    this.column = colToInt(cells[0]);
}

/**
 * @param {Object} cellNode
 */
function Cell(cellNode) {
    var na = {
            value: function() { return ''; },
            text:  function() { return ''; }
        },
        r = cellNode.attr('r').value(),
        type = (cellNode.attr('t') || na).value(),
        value = (cellNode.get('a:v', XMLNS) || na ).text(),
        coords = new CellCoords(r);

    this.column = coords.column;
    this.row = coords.row;
    this.value = value;
    this.type = type;
}

/**
 * @param {String} path
 * @param {Array} sheets
 * @param {Function} callback
 */
ExcelParser.prototype.extractFiles = function(path, sheets, callback) {
    var files = {
        strings: {},
        book: {},
        sheets: []
    };

    fs.readFile(path, 'binary', function(err, data) {
        if (err || !data) {
            return callback(err || new Error('data not exists'));
        }

        var zip,
            stringsRaw,
            stringsContents,
            bookRaw,
            bookContents,
            sheetNum,
            raw,
            contents;
        try {
            zip = new JSZip(data, { base64: false });
        } catch (e) {
            return callback(e);
        }

        stringsRaw = zip && zip.files && zip.files['xl/sharedStrings.xml'];
        stringsContents = stringsRaw && (typeof stringsRaw.asText === 'function') && stringsRaw.asText();
        if (!stringsContents) {
            return callback(new Error('xl/sharedStrings.xml not exists (maybe not xlsx file)'));
        }
        files.strings.contents = stringsContents;

        bookRaw = zip && zip.files && zip.files['xl/workbook.xml'];
        bookContents = bookRaw && (typeof bookRaw.asText === 'function') && bookRaw.asText();
        if (!bookContents) {
            return callback(new Error('xl/workbook.xml not exists (maybe not xlsx file)'));
        }
        files.book.contents = bookContents;

        if (sheets && sheets.length) {
            for (var i = 0; i < sheets.length; i++) {
                sheetNum = sheets[i];
                raw = zip.files['xl/worksheets/sheet' + sheetNum + '.xml'];
                contents = raw && (typeof raw.asText === 'function') && raw.asText();
                if (!contents) {
                    return callback(new Error('sheet ' + sheetNum + ' not exists'));
                }

                files.sheets.push({
                    num: sheetNum,
                    contents: contents
                });
            }
        } else {
            sheetNum = 1;
            while (true) {
                raw = zip.files['xl/worksheets/sheet' + sheetNum + '.xml'];
                contents = raw && (typeof raw.asText === 'function') && raw.asText();
                if (!contents) {
                    break;
                }

                files.sheets.push({
                    num: sheetNum,
                    contents: contents
                });
                sheetNum++;
            }
        }

        callback(null, files);
    });
};

/**
 * @param {Object} files
 * @param {Function} callback
 */
ExcelParser.prototype.extractData = function(files, callback) {
    var strings,
        book,
        sheetNames,
        sheets;

    try {
        strings = libxmljs.parseXml(files.strings.contents);
        book = libxmljs.parseXml(files.book.contents);
        sheetNames = _.map(book.find('//a:sheets//a:sheet', XMLNS), function(tag) {
            return tag.attr('name').value();
        });

        //sheets and sheetNames were retained the arrangement.
        sheets = _.map(files.sheets, function(sheetObj) {
            return {
                num: sheetObj.num,
                name: sheetNames[sheetObj.num - 1],
                xml: libxmljs.parseXml(sheetObj.contents)
            };
        });
    } catch (e) {
        return callback(e);
    }

    async.mapSeries(sheets, function(sheetObj, next) {
        var sheet = sheetObj.xml,
            cellNodes,
            cells,
            d,
            onedata = [];

        async.series([
            function(_next) {
                cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', XMLNS);
                async.setImmediate(_next);
            },
            function(_next) {
                var count = 0;
                async.mapSeries(cellNodes, function(node, __next) {
                    // use setImmediate every 100 times
                    count = (count + 1) % 100;
                    if (count === 0) {
                        async.setImmediate(function() {
                            __next(null, new Cell(node));
                        });
                    } else {
                        __next(null, new Cell(node));
                    }
                }, function(err, results) {
                    cells = results;
                    _next();
                });
            },
            function(_next) {
                d = sheet.get('//a:dimension/@ref', XMLNS);
                if (d) {
                    d = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
                } else {
                    d = calculateDimensions(cells);
                }
                async.setImmediate(_next);
            },
            function(_next) {
                var cols = d[1].column - d[0].column + 1,
                    rows = d[1].row - d[0].row + 1;
                _(rows).times(function() {
                    var _row = [];
                    _(cols).times(function() { _row.push(''); });
                    onedata.push(_row);
                });
                async.setImmediate(_next);
            },
            function(_next) {
                _(cells).each(function(cell) {
                    var value = cell.value;

                    if (cell.type === 's') {
                        var tmp = '';
                        _(strings.find('//a:si[' + (parseInt(value) + 1) + ']//a:t', XMLNS)).each(function(t) {
                            if (t.get('..').name() !== 'rPh') {
                                tmp += t.text();
                            }
                        });
                        value = tmp;
                    }

                    onedata[cell.row - d[0].row][cell.column - d[0].column] = value;

                });
                async.setImmediate(_next);
            }
        ], function() {
            next(null, {
                num: sheetObj.num,
                name: sheetObj.name,
                contents: onedata
            });
        });
    }, callback);
};

/**
 * @param {String} path
 * @param {Array} sheets
 * @param {Function} callback
 */
ExcelParser.prototype.execute = function(path, sheets, callback) {
    var _this = this;
    this.extractFiles(path, sheets, function(err, files) {
        if (err) {
            return callback(err);
        }

        _this.extractData(files, callback);
    });
};
