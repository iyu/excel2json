/**
 * @fileOverview excel parser
 * @name parser
 * @author Yuhei Aihara
 */

import cp from 'child_process';
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';

import _ from 'lodash';
import libxmljs from 'libxmljs';
import JSZip from 'jszip';

import cellConverter, { CustomDocument, ProcessSendData, Cell } from './cell';
import logger from '../logger';

interface Files {
  strings: { contents: string; };
  book: { contents: string; };
  sheets: Array<{ num: number; contents: string; }>
}

interface Data {
  num: number;
  name: string;
  cells: Cell[];
}

const XML_NS = { a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' };

async function asText(zip: JSZip, files: Files, sheetNum: number) {
  const raw = zip.files[`xl/worksheets/sheet${sheetNum}.xml`];
  const contents = raw && await raw.async('text');
  if (!contents) {
    return;
  }

  files.sheets.push({
    num: sheetNum,
    contents,
  });
  await asText(zip, files, sheetNum + 1);
}

class ExcelParser {
  public processes = 5;

  /**
   * @param {String} filepath
   * @param {Array} sheets
   */
  async extractFiles(filepath: string, sheets: number[]): Promise<Files> {
    const files: Files = {
      strings: { contents: '' },
      book: { contents: '' },
      sheets: [],
    };

    return new Promise((resolve, reject) => {
      fs.readFile(filepath, 'binary', async (err, data) => {
        if (err || !data) {
          return reject(err || new Error('data not exists'));
        }

        let zip: JSZip;
        try {
          zip = await JSZip.loadAsync(data, { base64: false });
        } catch (e) {
          logger.error(e.stack);
          return reject(e);
        }

        const stringsRaw = zip && zip.files && zip.files['xl/sharedStrings.xml'];
        const stringsContents = stringsRaw && await stringsRaw.async('text');
        if (!stringsContents) {
          return reject(new Error('xl/sharedStrings.xml not exists (maybe not xlsx file)'));
        }
        files.strings.contents = stringsContents;

        const bookRaw = zip && zip.files && zip.files['xl/workbook.xml'];
        const bookContents = bookRaw && await bookRaw.async('text');
        if (!bookContents) {
          return reject(new Error('xl/workbook.xml not exists (maybe not xlsx file)'));
        }
        files.book.contents = bookContents;

        if (sheets && sheets.length) {
          await Promise.all(_.map(sheets, (sheetNum) => {
            return (async () => {
              const raw = zip.files[`xl/worksheets/sheet${sheetNum}.xml`];
              const contents = raw && await raw.async('text');
              if (!contents) {
                return reject(new Error(`sheet ${sheetNum} not exists`));
              }

              files.sheets.push({
                num: sheetNum,
                contents,
              });
            })();
          }));
        } else {
          await asText(zip, files, 1);
        }

        resolve(files);
      });
    });
  }

  /**
   * @param {Object} files
   */
  async extractData(files: Files): Promise<Data[]> {
    let strings: CustomDocument;
    let sheetNames: string[] = [];
    let sheets: { num: number; name: string; contents: string; xml: CustomDocument; }[];

    try {
      const book = libxmljs.parseXml(files.book.contents) as CustomDocument;
      strings = libxmljs.parseXml(files.strings.contents) as CustomDocument;
      sheetNames = _.map(book.find('//a:sheets//a:sheet', XML_NS), (tag) => {
        const name = tag.attr('name');
        return name ? name.value() : '';
      });

      // sheets and sheetNames were retained the arrangement.
      sheets = _.map(files.sheets, (sheetObj) => {
        return {
          num: sheetObj.num,
          name: sheetNames[sheetObj.num - 1],
          contents: sheetObj.contents,
          xml: libxmljs.parseXml(sheetObj.contents) as CustomDocument,
        };
      });
    } catch (e) {
      logger.error(e.stack);
      throw e;
    }

    const result: Data[] = [];
    await Promise.all(_.map(sheets, (sheetObj) => {
      return (async () => {
        const sheet = sheetObj.xml;
        const cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', XML_NS);
        const tasks: { start: number, end: number }[] = [];

        if (cellNodes.length < 20000) {
          result.push({
            num: sheetObj.num,
            name: sheetObj.name,
            cells: cellConverter(cellNodes, strings, XML_NS),
          });
          return;
        }

        const nodes = Math.floor(cellNodes.length / this.processes);
        _.times(this.processes, (i) => {
          tasks.push({
            start: nodes * i,
            end: i + 1 === this.processes ? cellNodes.length : nodes * (i + 1),
          });
        });

        const stringsBuf = await new Promise((resolve, reject) => {
          zlib.deflate(files.strings.contents, (err, buf) => {
            if (err) {
              return reject(err);
            }
            resolve(buf);
          });
        }) as Buffer;
        const sheetsBuf = await new Promise((resolve, reject) => {
          zlib.deflate(sheetObj.contents, (err, buf) => {
            if (err) {
              return reject(err);
            }
            resolve(buf);
          });
        }) as Buffer;
        const sendData: ProcessSendData = {
          strings: stringsBuf.toString('base64'),
          sheets: sheetsBuf.toString('base64'),
          ns: XML_NS,
        };

        const cells = await Promise.all(_.map(tasks, (task) => {
          const _cellConverter = cp.fork(path.join(__dirname, './cell'));
          let _err: Error;
          let _result: Cell[];

          return new Promise((resolve, reject) => {
            _cellConverter.on('message', (data) => {
              _err = data.err;
              if (data.result) {
                _result = data.result;
              }
              _cellConverter.send({ exit: true });
            });
            _cellConverter.on('exit', (code) => {
              if (code !== 0) {
                return reject(_err || code);
              }
              resolve(_result);
            });
            _cellConverter.send(_.assign({
              start: task.start,
              end: task.end,
            }, sendData));
          });
        }));
        result.push({
          num: sheetObj.num,
          name: sheetObj.name,
          cells: _.flatten(cells) as Cell[],
        });
      })();
    }));
    return result;
  }

  async execute(filePath: string, sheets: number[]): Promise<Data[]> {
    const files = await this.extractFiles(filePath, sheets);
    const result = await this.extractData(files);
    return result;
  }
}

export default new ExcelParser();
