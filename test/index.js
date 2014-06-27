var should = require('should');

var excel2json = require('../');

var filepath = 'test/data/index.xlsx';

describe('#celname2index', function() {
    it('"AA1" be [26,0]', function(done) {
        var result = excel2json.celname2index('AA1');
        should.exist(result);
        result[0].should.equal(26);
        result[1].should.equal(0);
        done();
    });

    it('"ABC123" be [730,122]', function(done) {
        var result = excel2json.celname2index('ABC123');
        should.exist(result);
        result[0].should.equal(730);
        result[1].should.equal(122);
        done();
    });
});

describe('#format', function() {
    it('origin data test', function(done) {
        var options = [ '_id', 'obj.code', 'obj.value:number', 'obj.#list.code', 'obj.#list._obj.#_list:number' ];
        var data = [ 'test1', 'test1', '1', 'test1_1', '1'];
        var expected = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    {
                        code: 'test1_1',
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        excel2json.format(options, data, function(err, result, isOrigin) {
            should.not.exist(err);
            should.exist(result);
            should.exist(isOrigin);
            result.should.eql(expected);
            isOrigin.should.be.true;
            done();
        });
    });

    it('sub data test', function(done) {
        var options = [ '_id', 'obj.code', 'obj.value:number', 'obj.#list.code', 'obj.#list._obj.#_list:number' ];
        var data = [ '', '', '', '', '1'];
        var expected = {
            obj: {
                list: [
                    {
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        excel2json.format(options, data, function(err, result, isFirst) {
            should.not.exist(err);
            should.exist(result);
            should.exist(isFirst);
            result.should.eql(expected);
            isFirst.should.be.false;
            done();
        });
    });
});

describe('#extend', function() {
    it('extend array to array', function(done) {
        var originData = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    {
                        code: 'test1_1',
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        var subData = {
            obj: {
                list: [
                    {
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        var expected = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    {
                        code: 'test1_1',
                        _obj: {
                            _list: [ 1, 1 ]
                        }
                    }
                ]
            }
        };
        var result = excel2json.extend(originData, subData);
        result.should.eql(expected);
        done();
    });

    it('extend array to array', function(done) {
        var originData = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    {
                        code: 'test1_1',
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        var subData = {
            obj: {
                list: [
                    undefined,
                    {
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        var expected = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    {
                        code: 'test1_1',
                        _obj: {
                            _list: [ 1 ]
                        }
                    },
                    {
                        _obj: {
                            _list: [ 1 ]
                        }
                    }
                ]
            }
        };
        var result = excel2json.extend(originData, subData);
        result.should.eql(expected);
        done();
    });

    it('extend array to array', function(done) {
        var originData = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    [ 1 ]
                ]
            }
        };
        var subData = {
            obj: {
                list: [
                    [ 1 ]
                ]
            }
        };
        var expected = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    [ 1, 1 ]
                ]
            }
        };
        var result = excel2json.extend(originData, subData);
        result.should.eql(expected);
        done();
    });

    it('extend array to array', function(done) {
        var originData = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    [ 1 ]
                ]
            }
        };
        var subData = {
            obj: {
                list: [
                    undefined,
                    [ 1 ]
                ]
            }
        };
        var expected = {
            _id: 'test1',
            obj: {
                code: 'test1',
                value: 1,
                list: [
                    [ 1 ],
                    [ 1 ]
                ]
            }
        };
        var result = excel2json.extend(originData, subData);
        result.should.eql(expected);
        done();
    });
});

describe('#format and #extend', function() {
    var expected = [
        {
            _id: 'test1',
            list: [
                [ 1, 2 ],
                [ 3, 4 ]
            ],
            arr: [
                {
                    code: 'test1_1',
                    list: [
                        [ 1, 2 ],
                        [ 3, 4 ]
                    ],
                    arr: [
                        {
                            code: 'test1_1_1',
                            is: true
                        },
                        {
                            code: 'test1_1_2',
                            is: false
                        }
                    ]
                },
                {
                    code: 'test1_2',
                    list: [
                        [ 1, 2 ],
                        [ 3, 4 ]
                    ],
                    arr: [
                        {
                            code: 'test1_2_1',
                            is: true
                        },
                        {
                            code: 'test1_2_2',
                            is: false
                        }
                    ]
                }
            ]
        },
        {
            _id: 'test2',
            list: [
                [ 1, 2 ],
                [ 3, 4 ]
            ],
            arr: [
                {
                    code: 'test2_1',
                    list: [
                        [ 1, 2 ],
                        [ 3, 4 ]
                    ],
                    arr: [
                        {
                            code: 'test2_1_1',
                            is: true
                        },
                        {
                            code: 'test2_1_2',
                            is: false
                        }
                    ]
                },
                {
                    code: 'test2_2',
                    list: [
                        [ 1, 2 ],
                        [ 3, 4 ]
                    ],
                    arr: [
                        {
                            code: 'test2_2_1',
                            is: true
                        },
                        {
                            code: 'test2_2_2',
                            is: false
                        }
                    ]
                }
            ]
        }
    ];
    var attrOptions = [ '_id',      '#list:index',  '#list.#:number',  '#arr:index',   '#arr.code',    '#arr.#list:index', '#arr.#list.#:number',  '#arr.#arr.code',   '#arr.#arr.is:boolean'  ];
    var excelData = [
                      [ 'test1',    '0',            '1',                '0',            'test1_1',      '0',                '1',                    'test1_1_1',        'true'                  ],
                      [ '',         '0',            '2',                '0',            '',             '0',                '2',                    'test1_1_2',        'false'                 ],
                      [ '',         '1',            '3',                '0',            '',             '1',                '3',                    '',                 ''                      ],
                      [ '',         '1',            '4',                '0',            '',             '1',                '4',                    '',                 ''                      ],
                      [ '',         '',             '',                 '1',            'test1_2',      '0',                '1',                    'test1_2_1',        'true'                  ],
                      [ '',         '',             '',                 '1',            '',             '0',                '2',                    'test1_2_2',        'false'                 ],
                      [ '',         '',             '',                 '1',            '',             '1',                '3',                    '',                 ''                      ],
                      [ '',         '',             '',                 '1',            '',             '1',                '4',                    '',                 ''                      ],
                      [ 'test2',    '0',            '1',                '0',            'test2_1',      '0',                '1',                    'test2_1_1',        'true'                  ],
                      [ '',         '0',            '2',                '0',            '',             '0',                '2',                    'test2_1_2',        'false'                 ],
                      [ '',         '1',            '3',                '0',            '',             '1',                '3',                    '',                 ''                      ],
                      [ '',         '1',            '4',                '0',            '',             '1',                '4',                    '',                 ''                      ],
                      [ '',         '',             '',                 '1',            'test2_2',      '0',                '1',                    'test2_2_1',        'true'                  ],
                      [ '',         '',             '',                 '1',            '',             '0',                '2',                    'test2_2_2',        'false'                 ],
                      [ '',         '',             '',                 '1',            '',             '1',                '3',                    '',                 ''                      ],
                      [ '',         '',             '',                 '1',            '',             '1',                '4',                    '',                 ''                      ]
    ];
    var formatData = [
        { _id: 'test1', list: [ [ 1 ] ], arr: [ { code: 'test1_1', list: [ [ 1 ] ], arr: [ { code: 'test1_1_1', is: true } ] } ] },
        {               list: [ [ 2 ] ], arr: [ {                  list: [ [ 2 ] ], arr: [ { code: 'test1_1_2', is: false } ] } ] },
        {               list: [ , [ 3 ] ], arr: [ {       list: [ , [ 3 ] ] } ] },
        {               list: [ , [ 4 ] ], arr: [ {       list: [ , [ 4 ] ] } ] },
        {                                           arr: [ , { code: 'test1_2', list: [ [ 1 ] ], arr: [ { code: 'test1_2_1', is: true } ] } ] },
        {                                           arr: [ , {                  list: [ [ 2 ] ], arr: [ { code: 'test1_2_2', is: false } ] } ] },
        {                                           arr: [ , {                  list: [ , [ 3 ] ] } ] },
        {                                           arr: [ , {                  list: [ , [ 4 ] ] } ] },
        { _id: 'test2', list: [ [ 1 ] ], arr: [ { code: 'test2_1', list: [ [ 1 ] ], arr: [ { code: 'test2_1_1', is: true } ] } ] },
        {               list: [ [ 2 ] ], arr: [ {                  list: [ [ 2 ] ], arr: [ { code: 'test2_1_2', is: false } ] } ] },
        {               list: [ , [ 3 ] ], arr: [ {                list: [ , [ 3 ] ] } ] },
        {               list: [ , [ 4 ] ], arr: [ {                list: [ , [ 4 ] ] } ] },
        {                                           arr: [ , { code: 'test2_2', list: [ [ 1 ] ], arr: [ { code: 'test2_2_1', is: true } ] } ] },
        {                                           arr: [ , {                  list: [ [ 2 ] ], arr: [ { code: 'test2_2_2', is: false } ] } ] },
        {                                           arr: [ , {                  list: [ , [ 3 ] ] } ] },
        {                                           arr: [ , {                  list: [ , [ 4 ] ] } ] }
    ];

    describe('#format attrOptions: ' + JSON.stringify(attrOptions), function() {
        excelData.forEach(function(data, idx) {
            it('#format data: ' + JSON.stringify(data), function(done) {
                excel2json.format(attrOptions, data, function(err, result, isOrigin) {
                    should.not.exist(err);
                    should.exist(result);
                    should.exist(isOrigin);
                    result.should.eql(formatData[idx]);
                    isOrigin.should.eql(idx === 0 || idx === 8);
                    done();
                });
            });
        });
    });

    describe('#extend', function() {
        var _expected = [
            { _id: 'test1', list: [ [ 1 ] ], arr: [ { code: 'test1_1', list: [ [ 1 ] ], arr: [ { code: 'test1_1_1', is: true } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] },
                                                            { code: 'test1_2', list: [ [ 1 ] ], arr: [ { code: 'test1_2_1', is: true } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] },
                                                            { code: 'test1_2', list: [ [ 1, 2 ] ], arr: [ { code: 'test1_2_1', is: true }, { code: 'test1_2_2', is: false } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] },
                                                            { code: 'test1_2', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test1_2_1', is: true }, { code: 'test1_2_2', is: false } ] } ] },
            { _id: 'test1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_1_1', is: true }, { code: 'test1_1_2', is: false } ] },
                                                            { code: 'test1_2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test1_2_1', is: true }, { code: 'test1_2_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1 ] ], arr: [ { code: 'test2_1', list: [ [ 1 ] ], arr: [ { code: 'test2_1_1', is: true } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] },
                                                            { code: 'test2_2', list: [ [ 1 ] ], arr: [ { code: 'test2_2_1', is: true } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] },
                                                            { code: 'test2_2', list: [ [ 1, 2 ] ], arr: [ { code: 'test2_2_1', is: true }, { code: 'test2_2_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] },
                                                            { code: 'test2_2', list: [ [ 1, 2 ], [ 3 ] ], arr: [ { code: 'test2_2_1', is: true }, { code: 'test2_2_2', is: false } ] } ] },
            { _id: 'test2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_1_1', is: true }, { code: 'test2_1_2', is: false } ] },
                                                            { code: 'test2_2', list: [ [ 1, 2 ], [ 3, 4 ] ], arr: [ { code: 'test2_2_1', is: true }, { code: 'test2_2_2', is: false } ] } ] },
        ];
        var list = [];
        formatData.forEach(function(data, idx) {
            it('#extend formatData: ' + JSON.stringify(data), function(done) {
                if (data._id) {
                    list.push(data);
                } else {
                    excel2json.extend(list[list.length - 1], data);
                }
                _expected[idx].should.eql(list[list.length - 1]);
                done();
            });
        });
    });

    it('Combined test #format and #extend', function(done) {
        var list = [];
        excelData.forEach(function(data, idx) {
            excel2json.format(attrOptions, data, function(err, result, isOrigin) {
                should.not.exist(err);
                should.exist(result);
                should.exist(isOrigin);
                result.should.eql(formatData[idx]);
                isOrigin.should.eql(idx === 0 || idx === 8);

                if (isOrigin) {
                    list.push(result);
                } else {
                    excel2json.extend(list[list.length - 1], result);
                }
            });
        });
        list.should.eql(expected);
        done();
    });
});


describe('#parse', function() {
    it(filepath, function(done) {
        excel2json.parse(filepath, [], function(err, result) {
            should.not.exist(err);
            should.exist(result);
            result.should.have.length(2);
            result[0].should.property('num', 1);
            result[0].should.property('contents', [
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
            result[1].should.property('contents', [
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
