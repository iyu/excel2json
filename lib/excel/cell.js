/**
 * @fileOverview cell converter
 * @name cell.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 */
var zlib = require('zlib');

var _ = require('lodash'),
    async = require('async'),
    libxmljs = require('libxmljs');

var logger = require('../logger');

/**
 * @param {String} cell
 */
function getCellCoord(cell) {
    var cells = cell.match(/^(\D+)(\d+)$/);

    return {
        cell: cell,
        column: cells[1],
        row: parseInt(cells[2], 10)
    };
}

module.exports = function(cellNodes, strings, ns) {
    var result = [];
    _.each(cellNodes, function(cellNode) {
        var coord = getCellCoord(cellNode.attr('r').value()),
            type = cellNode.attr('t'),
            id = cellNode.get('a:v', ns),
            value;

        if (!id) {
            // empty cell
            return;
        }

        if (type && type.value() === 's') {
            value = '';
            _.each(strings.find('//a:si[' + (parseInt(id.text(), 10) + 1) + ']//a:t', ns), function(t) {
                if (t.get('..').name() !== 'rPh') {
                    value += t.text();
                }
            });
        } else {
            value = id.text();
        }

        if (value === '') {
            // empty cell
            return;
        }

        result.push({
            cell: coord.cell,
            column: coord.column,
            row: coord.row,
            value: value
        });
    });
    return result;
};

if (require.main === module) {
    process.on('message', function(data) {
        if (data.exit) {
            process.exit();
        }

        var ns = data.ns,
            start = data.start,
            end = data.end;

        async.series({
            strings: function(next) {
                zlib.unzip(new Buffer(data.strings, 'base64'), next);
            },
            sheet: function(next) {
                zlib.unzip(new Buffer(data.sheets, 'base64'), next);
            }
        }, function(err, unzip) {
            if (err) {
                logger.error(err.stack);
                process.send({ err: err });
            }

            var strings = libxmljs.parseXml(unzip.strings),
                sheet = libxmljs.parseXml(unzip.sheet),
                cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);

            var result = module.exports(cellNodes.slice(start, end), strings, ns);

            process.send({ result: result });
        });
    });

    process.on('uncaughtException', function(err) {
        logger.error(err.stack);
        process.send({ err: err });
        process.exit(1);
    });
}
