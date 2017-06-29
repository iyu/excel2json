/**
 * @fileOverview excel parser
 * @name parser.js
 * @author Yuhei Aihara
 */

'use strict';

const cp = require('child_process');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const _ = require('lodash');
const async = require('neo-async');
const libxmljs = require('libxmljs');
const JSZip = require('node-zip');

const cellConverter = require('./cell');
const logger = require('../logger');

const XML_NS = { a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' };

function asText(zip, files, sheetNum) {
  const raw = zip.files[`xl/worksheets/sheet${sheetNum}.xml`];
  const contents = raw && (typeof raw.asText === 'function') && raw.asText();
  if (!contents) {
    return;
  }

  files.sheets.push({
    num: sheetNum,
    contents,
  });
  asText(zip, files, sheetNum + 1);
}

class ExcelParser {
  constructor() {
    this.processes = 5;
  }

  /**
   * @param {String} filepath
   * @param {Array} sheets
   * @param {Function} callback
   */
  extractFiles(filepath, sheets, callback) {
    const files = {
      strings: {},
      book: {},
      sheets: [],
    };

    fs.readFile(filepath, 'binary', (err, data) => {
      if (err || !data) {
        return callback(err || new Error('data not exists'));
      }

      let zip;
      try {
        zip = new JSZip(data, { base64: false });
      } catch (e) {
        logger.error(e.stack);
        return callback(e);
      }

      const stringsRaw = zip && zip.files && zip.files['xl/sharedStrings.xml'];
      const stringsContents = stringsRaw && (typeof stringsRaw.asText === 'function') && stringsRaw.asText();
      if (!stringsContents) {
        return callback(new Error('xl/sharedStrings.xml not exists (maybe not xlsx file)'));
      }
      files.strings.contents = stringsContents;

      const bookRaw = zip && zip.files && zip.files['xl/workbook.xml'];
      const bookContents = bookRaw && (typeof bookRaw.asText === 'function') && bookRaw.asText();
      if (!bookContents) {
        return callback(new Error('xl/workbook.xml not exists (maybe not xlsx file)'));
      }
      files.book.contents = bookContents;

      if (sheets && sheets.length) {
        for (let i = 0; i < sheets.length; i++) {
          const sheetNum = sheets[i];
          const raw = zip.files[`xl/worksheets/sheet${sheetNum}.xml`];
          const contents = raw && (typeof raw.asText === 'function') && raw.asText();
          if (!contents) {
            return callback(new Error(`sheet ${sheetNum} not exists`));
          }

          files.sheets.push({
            num: sheetNum,
            contents,
          });
        }
      } else {
        asText(zip, files, 1);
      }

      callback(null, files);
    });
  }

  /**
   * @param {Object} files
   * @param {Function} callback
   */
  extractData(files, callback) {
    let book;
    let strings;
    let sheetNames;
    let sheets;

    try {
      book = libxmljs.parseXml(files.book.contents);
      strings = libxmljs.parseXml(files.strings.contents);
      sheetNames = _.map(book.find('//a:sheets//a:sheet', XML_NS), (tag) => {
        return tag.attr('name').value();
      });

      // sheets and sheetNames were retained the arrangement.
      sheets = _.map(files.sheets, (sheetObj) => {
        return {
          num: sheetObj.num,
          name: sheetNames[sheetObj.num - 1],
          contents: sheetObj.contents,
          xml: libxmljs.parseXml(sheetObj.contents),
        };
      });
    } catch (e) {
      logger.error(e.stack);
      return callback(e);
    }

    async.mapSeries(sheets, (sheetObj, next) => {
      const sheet = sheetObj.xml;
      const cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', XML_NS);
      const tasks = [];

      if (cellNodes.length < 20000) {
        return next(null, {
          num: sheetObj.num,
          name: sheetObj.name,
          cells: cellConverter(cellNodes, strings, XML_NS),
        });
      }

      const nodes = Math.floor(cellNodes.length / this.processes);
      _.times(this.processes, (i) => {
        tasks.push({
          start: nodes * i,
          end: i + 1 === this.processes ? cellNodes.length : nodes * (i + 1),
        });
      });

      async.parallel({
        strings: (_next) => {
          zlib.deflate(files.strings.contents, _next);
        },
        sheets: (_next) => {
          zlib.deflate(sheetObj.contents, _next);
        },
      }, (err, sendData) => {
        if (err) {
          return next(err);
        }
        sendData.strings = sendData.strings.toString('base64');
        sendData.sheets = sendData.sheets.toString('base64');
        sendData.ns = XML_NS;

        async.map(tasks, (task, done) => {
          const _cellConverter = cp.fork(path.join(__dirname, './cell'));
          let _err;
          let result;

          _cellConverter.on('message', (data) => {
            _err = data.err;
            if (data.result) {
              result = data.result;
            }
            cellConverter.send({ exit: true });
          });
          _cellConverter.on('exit', (code) => {
            if (code !== 0) {
              return done(_err || code);
            }
            done(_err, result);
          });
          _cellConverter.send(_.assign({
            start: task.start,
            end: task.end,
          }, sendData));
        }, (_err, result) => {
          if (_err) {
            return next(_err);
          }

          next(null, {
            num: sheetObj.num,
            name: sheetObj.name,
            cells: _.flatten(result),
          });
        });
      });
    }, callback);
  }

  /**
   * @param {String} filepath
   * @param {Array} sheets
   * @param {Function} callback
   */
  execute(filepath, sheets, callback) {
    this.extractFiles(filepath, sheets, (err, files) => {
      if (err) {
        return callback(err);
      }

      this.extractData(files, callback);
    });
  }
}

module.exports = new ExcelParser();
