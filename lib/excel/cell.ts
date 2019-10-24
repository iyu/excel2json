/**
 * @fileOverview cell converter
 * @name cell
 * @author Yuhei Aihara
 */

import zlib from 'zlib';

import _ from 'lodash';
import libxmljs from 'libxmljs';

import logger from '../logger';

/**
 * @param {String} cell
 */
function getCellCoord(cell: string) {
  const cells = cell.match(/^(\D+)(\d+)$/) || [];

  return {
    cell,
    column: cells[1],
    row: parseInt(cells[2], 10),
  };
}

export default (cellNodes: any, strings: any, ns: any) => {
  const result: any = [];
  _.forEach(cellNodes, (cellNode) => {
    const coord = getCellCoord(cellNode.attr('r').value());
    const type = cellNode.attr('t');
    const id = cellNode.get('a:v', ns);
    let value: string;

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
  process.on('message', async (data) => {
    if (data.exit) {
      process.exit();
    }

    const { ns, start, end } = data;
    const unzip: any = {};
    try {
      unzip.strings = await zlib.unzip(Buffer.from(data.strings, 'base64'), (ret) => { return ret; });
      unzip.sheet = await zlib.unzip(Buffer.from(data.sheets, 'base64'), (ret) => { return ret; });
    } catch (err) {
      logger.error(err.stack);
      if (process.send) {
        process.send({ err });
      }
    }

    const strings = libxmljs.parseXml(unzip.strings);
    const sheet = libxmljs.parseXml(unzip.sheet);
    const cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c');

    const result = module.exports(cellNodes.slice(start, end), strings, ns);

    if (process.send) {
      process.send({ result });
    }
  });

  process.on('uncaughtException', (err) => {
    logger.error(err.stack);
    if (process.send) {
      process.send({ err });
    }
    process.exit(1);
  });
}
