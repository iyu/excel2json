[![NPM version][npm-image]][npm-url]
[![Downloads][downloads-image]][downloads-url]

Excel2Json
==========

Can be converted to JSON format any Excel data.

example Excel data

|   | A      | B        | C                | D |
|:-:|:-------|:---------|:-----------------|---|
| 1 | {}     |          |                  |   |
| 2 | _id    | obj.code | obj.value:number |   |
| 3 |        |          |                  |   |
| 4 | first  | one      | 1                |   |
| 5 | second | two      | 2                |   |
| 6 |        |          |                  |   |

converted to Object
```js
[
    {
        _id: 'first',
        obj: {
            code: 'one',
            value: 1
        }
    }, 
    {
        _id: 'second',
        obj: {
            code: 'two',
            value: 2
        }
    }
]
```

## Installation
```
npm install excel2json
```

## Usage
### Quick start
example sheet.xlsx

|   | A              | B        | C                | D |
|:-:|:---------------|:---------|:-----------------|---|
| 1 | {name: 'Test'} |          |                  |   |
| 2 | _id            | obj.code | obj.value:number |   |
| 3 |                |          |                  |   |
| 4 | first          | one      | 1                |   |
| 5 | second         | two      | 2                |   |
| 6 |                |          |                  |   |

Sheet1
```js
var excel2json = require('excel2json');

var filename = './sheet.xlsx';
var sheets = [1];
excel2json.parse(filename, sheets, function(err, data) {
    console.log(data);
    // [{
    //    num: 1,                    // sheet number
    //    name: 'Sheet1',            // sheet name
    //    option: {                  // option extend sheet option (ex: A1)
    //        name: 'Test'
    //        attr_line:
    //        data_line:
    //        ref_key: '_id',
    //        format: {
    //            A: { type: null, key: '_id', keys: [ '_id' ] },
    //            B: { type: null, key: 'obj.code', keys: [ 'obj', 'code' ] },
    //            C: { type: 'number', key: 'obj.value', keys: [ 'obj', 'value' ] }
    //        }
    //    },
    //    contents: [
    //        { _id: 'first', obj: { code: 'one', value: 1 } },
    //        { _id: 'second', obj: { code: 'two', value: 2 } }
    //    ]
    // }]

    excel2json.toJson(data, function(err, json) {
        console.log(json);
        // {
        //    Test: {
        //        first: {
        //            _id: 'first',
        //            obj: { code: 'one', value: 1 }
        //        },
        //        second: {
        //            _id: 'second',
        //            obj: { code: 'two', value: 2 }
        //        }
        //    }
        // }
    });
});
```

### Setup
Setup options.
```js
var excel2json = require('excel2json');

excel2json.setup({
    option_cell: 'A1', // Cell with a custom sheet option. It is not yet used now. (default: 'A1'
    attr_line: '2',    // Line with a data attribute. (default: '2'
    data_line: '4',    // Line with a data. (default: '4'
    ref_key: '_id'     // ref key. (default: '_id'
    logger: Logger     // custom logger.
});
```

### Sheet option
sheet option. setting with optionCell (default: 'A1'
* `name`
* `type`
* `key`
* `attr_line`
* `data_line`
* `ref_key`


### Attribute
Specify the key name.

**Special character**
* `#` Use when the array.
* `$` Use when the split array.
* `:number` or `:num` Use when the parameters of type `Number`.
* `:boolean` or `:bool` Use when the parameters of type `Boolean`.
* `:date` Use when the parameters of unix time.
* `:index` Use when the array of array.

### Example
An example of a complex format.


[test.xlsx](https://github.com/iyu/excel2json/raw/master/test/data/test.xlsx) > [test.json](https://github.com/iyu/excel2json/blob/master/test/data/test.json)
```js
var excel2json = require('excel2json');

excel2json.parse('test.xlsx', [], function(err, sheetDatas) {
    excel2json.toJson(sheetDatas, function(err, result) {
        fs.writeFileSync('test.json', JSON.stringify(result, null, 4));
    });
});
```

## Contribution
1. Fork it ( [https://github.com/iyu/excel2json/fork](https://github.com/iyu/excel2json/fork) )
2. Create a feature branch
3. Commit your changes
4. Rebase your local changes against the master branch
5. Run test suite with the `npm test; npm run lint` command and confirm that it passes
5. Create new Pull Request

[npm-image]: https://img.shields.io/npm/v/excel2json.svg?style=flat-square
[npm-url]: https://www.npmjs.com/package/excel2json
[downloads-image]: https://img.shields.io/npm/dm/excel2json.svg?style=flat-square
[downloads-url]: https://www.npmjs.com/package/excel2json
