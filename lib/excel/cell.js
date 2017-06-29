/**
 * @fileOverview cell converter
 * @name cell.js
 * @author Yuhei Aihara
 */

'use strict';

const zlib = require('zlib');

const _ = require('lodash');
const async = require('neo-async');
const libxmljs = require('libxmljs');

const logger = require('../logger');

/**
 * @param {String} cell
 */
function getCellCoord(cell) {
  const cells = cell.match(/^(\D+)(\d+)$/);

  return {
    cell,
    column: cells[1],
    row: parseInt(cells[2], 10),
  };
}

module.exports = (cellNodes, strings, ns) => {
  const result = [];
  _.forEach(cellNodes, (cellNode) => {
    const coord = getCellCoord(cellNode.attr('r').value());
    const type = cellNode.attr('t');
    const id = cellNode.get('a:v', ns);
    let value;

    if (!id) {
      // empty cell
      return;
    }

    if (type && type.value() === 's') {
      value = '';
      _.forEach(strings.find(`//a:si[${parseInt(id.text(), 10) + 1}]//a:t`, ns), (t) => {
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
      value,
    });
  });
  return result;
};

if (require.main === module) {
  process.on('message', (data) => {
    if (data.exit) {
      process.exit();
    }

    const ns = data.ns;
    const start = data.start;
    const end = data.end;

    async.parallel({
      strings: (next) => {
        zlib.unzip(new Buffer(data.strings, 'base64'), next);
      },
      sheet: (next) => {
        zlib.unzip(new Buffer(data.sheets, 'base64'), next);
      },
    }, (err, unzip) => {
      if (err) {
        logger.error(err.stack);
        process.send({ err });
      }

      const strings = libxmljs.parseXml(unzip.strings);
      const sheet = libxmljs.parseXml(unzip.sheet);
      const cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);

      const result = module.exports(cellNodes.slice(start, end), strings, ns);

      process.send({ result });
    });
  });

  process.on('uncaughtException', (err) => {
    logger.error(err.stack);
    process.send({ err });
    process.exit(1);
  });
}
