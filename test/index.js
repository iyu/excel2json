var should = require('should');

var excel2json = require('../');

describe('#parse', function() {
    var filepath = 'test/data/index.xlsx';
    it(filepath, function(done) {
        excel2json.parse(filepath, [], function(err, result) {
            should.not.exist(err);
            should.exist(result);
            result.should.have.length(2);
            result[0].should.property('num', 1);
            result[0].should.property('name', 'Sheet1');
            result[0].should.property('opts', {
                attr_line: 2,
                data_line: 4,
                ref_key: '_id',
                format: {
                    A: { type: null, keys: [ '_id' ] },
                    B: { type: 'index', keys: [ '#array' ] },
                    C: { type: null, keys: [ '#array', 'key' ] },
                    D: { type: 'number', keys: [ '#array', 'num' ] },
                    E: { type: 'number', keys: [ '#array', '#list' ] },
                    F: { type: null, keys: [ 'key' ] }
                }
            });
            result[0].should.property('list', [
                {
                    _id: 'aaa',
                    array: [
                        {
                            key: 'a',
                            num: 1,
                            list: [ 1, 2 ]
                        },
                        {
                            key: 'b',
                            num: 2,
                            list: [ 1, 2, 3 ]
                        },
                        {
                            key: 'c',
                            num: 3,
                            list: [ 1, 2, 3, 4 ]
                        }
                    ],
                    key: 'hoge'
                },
                {
                    _id: 'bbb',
                    key: 'fuga'
                }
            ]);
            result[1].should.property('list', [
                {
                    _id: 'a',
                    date: new Date('2014/06/01 05:00').getTime()
                },
                {
                    _id: 'b',
                    date: new Date('2014/07/01 05:00').getTime()
                }
            ]);
            done();
        });
    });
});

describe('#toJson', function() {
    var filepath = 'test/data/test.xlsx';
    it(filepath, function(done) {
        excel2json.parse(filepath, [], function(err, result) {
            should.not.exist(err);
            should.exist(result);

            excel2json.toJson(result, function(_err, _result) {
                should.not.exist(_err);
                should.exist(_result);
                _result.should.eql(require('./data/test.json'));
                done();
            });
        });
    });
});
