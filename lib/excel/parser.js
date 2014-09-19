/**
 * @fileOverview excel parser
 * @name parser.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 */
var child_process = require('child_process'),
    fs = require('fs'),
    path = require('path'),
    zlib = require('zlib');

var _ = require('lodash'),
    async = require('async'),
    libxmljs = require('libxmljs'),
    JSZip = require('node-zip');

var cellConverter = require('./cell'),
    logger = require('../logger');

var XML_NS = { a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' };

function ExcelParser() {
    this.processes = 5;
}

module.exports = new ExcelParser();

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
            logger.error(e.stack);
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
    var _this = this,
        book,
        strings,
        sheetNames,
        sheets;

    try {
        book = libxmljs.parseXml(files.book.contents);
        strings = libxmljs.parseXml(files.strings.contents);
        sheetNames = _.map(book.find('//a:sheets//a:sheet', XML_NS), function(tag) {
            return tag.attr('name').value();
        });

        //sheets and sheetNames were retained the arrangement.
        sheets = _.map(files.sheets, function(sheetObj) {
            return {
                num: sheetObj.num,
                name: sheetNames[sheetObj.num - 1],
                contents: sheetObj.contents,
                xml: libxmljs.parseXml(sheetObj.contents)
            };
        });
    } catch (e) {
        logger.error(e.stack);
        return callback(e);
    }

    async.mapSeries(sheets, function(sheetObj, next) {
        var sheet = sheetObj.xml,
            cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', XML_NS),
            nodes,
            tasks = [];

        if (cellNodes.length < 20000) {
            return next(null, {
                num: sheetObj.num,
                name: sheetObj.name,
                cells: cellConverter(cellNodes, strings, XML_NS)
            });
        }

        nodes = cellNodes.length / _this.processes | 0;
        _.times(_this.processes, function(i) {
            tasks.push({
                start: nodes * i,
                end: i + 1 === _this.processes ? cellNodes.length : nodes * (i + 1)
            });
        });

        async.series({
            strings: function(_next) {
                zlib.deflate(files.strings.contents, _next);
            },
            sheets: function(_next) {
                zlib.deflate(sheetObj.contents, _next);
            }
        }, function(err, sendData) {
            if (err) {
                return next(err);
            }
            sendData.strings = sendData.strings.toString('base64');
            sendData.sheets = sendData.sheets.toString('base64');
            sendData.ns = XML_NS;

            async.map(tasks, function(task, _next) {
                var cellConverter = child_process.fork(path.join(__dirname, './cell')),
                    err,
                    result = [];

                cellConverter.on('message', function(data) {
                    err = data.err;
                    if (data.result) {
                        result = data.result;
                    }
                    cellConverter.send({ exit: true });
                });
                cellConverter.on('exit', function(code) {
                    if (code !== 0) {
                        return _next(err || code);
                    }
                    _next(err, result);
                });
                cellConverter.send(_.extend({
                    start: task.start,
                    end: task.end
                }, sendData));
            }, function(err, result) {
                if (err) {
                    return next(err);
                }

                next(null, {
                    num: sheetObj.num,
                    name: sheetObj.name,
                    cells: _.flatten(result)
                });
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
