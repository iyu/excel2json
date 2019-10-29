/**
 * @fileOverview cell converter
 * @name cell
 * @author Yuhei Aihara
 */

import zlib from 'zlib';

import _ from 'lodash';
import libxmljs, { Document, Element, StringMap } from 'libxmljs';

import logger from '../logger';

interface Coord {
  // 'A1'
  cell: string;
  // 'A'
  column: string;
  // 1
  row: number;
}

export interface Cell extends Coord {
  // value
  value: string;
}

interface CustomElement extends Element {
  // eslint-disable-next-line camelcase
  get(xpath: string, ns_uri?: string): Element | null;
  // eslint-disable-next-line camelcase
  get(xpath: string, ns_uri?: StringMap): Element | null;
}

export interface CustomDocument extends Document {
  find(xpath: string, namespaces?: StringMap): CustomElement[];
}

export interface ProcessSendData {
  start?: number;
  end?: number;
  strings?: string; // base64
  sheets?: string; // base64
  ns?: StringMap;
  exit?: boolean;
}

const getCellCoord = (cell: string): Coord => {
  const cells = cell.match(/^(\D+)(\d+)$/) || [];

  return {
    cell,
    column: cells[1],
    row: parseInt(cells[2], 10),
  };
};

export default (cellNodes: CustomElement[], strings: CustomDocument, ns: StringMap): Cell[] => {
  const result: Cell[] = [];
  _.forEach(cellNodes, (cellNode) => {
    const attr = cellNode.attr('r');
    const coord = getCellCoord(attr ? attr.value() : '');
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
        const elm = t.get('..');
        if (elm && elm.name() !== 'rPh') {
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
  process.on('message', async (data: ProcessSendData) => {
    if (data.exit) {
      process.exit();
    }

    const { ns, start, end } = data;
    const unzip: { strings: string, sheets: string } = {
      strings: await new Promise((resolve) => {
        zlib.unzip(Buffer.from(data.strings || '', 'base64'), (err, buf) => {
          if (err) {
            logger.error(err.stack);
            if (process.send) {
              process.send({ err });
            }
            return;
          }
          resolve(buf.toString());
        });
      }),
      sheets: await new Promise((resolve) => {
        zlib.unzip(Buffer.from(data.sheets || '', 'base64'), (err, buf) => {
          if (err) {
            logger.error(err.stack);
            if (process.send) {
              process.send({ err });
            }
            return;
          }
          resolve(buf.toString());
        });
      }),
    };

    const strings = libxmljs.parseXml(unzip.strings);
    const sheets = libxmljs.parseXml(unzip.sheets) as CustomDocument;
    const cellNodes = sheets.find('/a:worksheet/a:sheetData/a:row/a:c', ns);

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
