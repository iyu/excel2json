/**
 * @fileOverview
 * @name cell.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 */
var zlib = require('zlib');

var _ = require('lodash'),
    async = require('async'),
    libxmljs = require('libxmljs');

var na = {
    value: function() { return ''; },
    text:  function() { return ''; }
};

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
function getCellCoord(cell) {
    var cells = cell.split(/([0-9]+)/);

    return {
        row: parseInt(cells[1]),
        column: colToInt(cells[0])
    };
}

module.exports = function(cellNodes, strings, ns) {
    return _.map(cellNodes, function(cellNode) {
        var coord = getCellCoord(cellNode.attr('r').value()),
            type = (cellNode.attr('t') || na).value(),
            id = (cellNode.get('a:v', ns) || na).text(),
            value = '';

        if (type === 's') {
            _(strings.find('//a:si[' + (parseInt(id) + 1) + ']//a:t', ns)).each(function(t) {
                if (t.get('..').name() !== 'rPh') {
                    value += t.text();
                }
            });
        } else {
            value = id;
        }

        return {
            row: coord.row,
            column: coord.column,
            value: value
        };
    });
};

if (require.main === module) {
    process.on('message', function(data) {
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
                console.error(err.stack);
                process.send({ err: err });
                process.exit();
            }

            var strings = libxmljs.parseXml(unzip.strings),
                sheet = libxmljs.parseXml(unzip.sheet),
                cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);

            var result = module.exports(cellNodes.slice(start, end), strings, ns);

            process.send({ result: result });
            process.exit();
        });
    });

    process.on('uncaughtException', function(err) {
        console.error(err.stack);
        process.send({ err: err });
        process.exit(1);
    });
}
