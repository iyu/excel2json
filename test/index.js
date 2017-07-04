'use strict';

const should = require('should');

const excel2json = require('../');
const data = require('./data/test.json');

describe('#parse', () => {
  const filepath = 'test/data/index.xlsx';
  it(filepath, (done) => {
    excel2json.parse(filepath, [], (err, result, errList) => {
      should.not.exist(err);
      should.exist(result);
      should.not.exist(errList);
      result.should.have.length(3);
      result[0].should.property('num', 1);
      result[0].should.property('name', 'Sheet1');
      result[0].should.property('opts', {
        attr_line: 2,
        data_line: 4,
        ref_key: '_id',
        format: {
          A: { type: null, key: '_id', keys: ['_id'] },
          B: { type: 'index', key: '#array', keys: ['#array'] },
          C: { type: null, key: '#array.key', keys: ['#array', 'key'] },
          D: { type: 'number', key: '#array.num', keys: ['#array', 'num'] },
          E: { type: 'number', key: '#array.#list', keys: ['#array', '#list'] },
          F: { type: null, key: 'key', keys: ['key'] },
        },
      });
      result[0].should.property('list', [
        {
          _id: 'aaa',
          array: [
            {
              key: 'a',
              num: 1,
              list: [1, 2],
            },
            {
              key: 'b',
              num: 2,
              list: [1, 2, 3],
            },
            {
              key: 'c',
              num: 3,
              list: [1, 2, 3, 4],
            },
          ],
          key: 'hoge',
        },
        {
          _id: 'bbb',
          key: 'fuga',
        },
      ]);
      result[1].should.property('list', [
        {
          __ref: 'aaa',
          bool: true,
          arr: [
            { code: 'a' },
            { code: 'b' },
          ],
        },
        {
          __ref: 'aaa',
          bool: false,
          arr: [
            { code: 'c' },
            { code: 'd' },
          ],
        },
        {
          __ref: 'bbb',
          bool: true,
          arr: [
            { code: 'a' },
            { code: 'b' },
          ],
        },
        {
          __ref: 'bbb',
          bool: false,
          arr: [
            { code: 'c' },
            { code: 'd' },
          ],
        },
      ]);
      result[2].should.property('list', [
        {
          _id: 'a',
          date: new Date('2014/06/01 05:00').getTime(),
        },
        {
          _id: 'b',
          date: new Date('2014/07/01 05:00').getTime(),
        },
      ]);

      excel2json.toJson(result, (_err, _result) => {
        should.not.exist(_err);
        should.exist(_result);
        Object.keys(_result).should.have.length(2);
        _result.should.have.property('Sheet1');
        _result.Sheet1.should.have.property('bbb', {
          _id: 'bbb',
          key: 'fuga',
          obj: {
            array: [
              {
                bool: true,
                arr: [
                  { code: 'a' },
                  { code: 'b' },
                ],
              },
              {
                bool: false,
                arr: [
                  { code: 'c' },
                  { code: 'd' },
                ],
              },
            ],
          },
        });
        _result.should.have.property('Sheet2');
        done();
      });
    });
  });
});

describe('#toJson', () => {
  const filepath = 'test/data/test.xlsx';
  it(filepath, (done) => {
    excel2json.parse(filepath, [], (err, result, errList) => {
      should.not.exist(err);
      should.exist(result);
      should.not.exist(errList);

      excel2json.toJson(result, (_err, _result) => {
        should.not.exist(_err);
        should.exist(_result);
        _result.should.eql(data);
        done();
      });
    });
  });
});
