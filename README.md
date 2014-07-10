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
```
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
```
var excel2json = require('excel2json');

var filename = './sheet.xlsx';
var sheets = [1];
excel2json.parse(filename, sheets, function(err, data) {
    console.log(data);
    // [{
    //    num: 1,                    // sheet number
    //    name: 'Sheet1',            // sheet name
    //    option: { name: 'Test' },  // sheet option (A1)
    //    contents: [
    //        { _id: 'first', obj: { code: 'one', value: 1 } }, { _id: 'second', obj: { code: 'two', value: 2 } }
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
```
var excel2json = require('excel2json');

excel2json.setup({
    optionCell: 'A1', // Cell with a custom sheet option. It is not yet used now. (default: 'A1'
    attrLine: '2',    // Line with a data attribute. (default: '2'
    dataLine: '4',    // Line with a data. (default: '4'
    refKey: '_id'     // ref key. (default: '_id'
});
```

### Sheet option
sheet option. setting with optionCell (default: 'A1'
* `name`
* `type`
* `key`
* `attrLine`
* `dataLine`
* `refKey`


### Attribute
Specify the key name.

**Special character**
* `#` Use when the array.
* `:number` Use when the parameters of type `Number`. 
* `:boolean` Use when the parameters of type `Boolean`.
* `:date` Use when the parameters of unix time.
* `:index` Use when the array of array.

### Example
An example of a complex format.


[test.xlsx](https://github.com/yuhei-a/excel2json/raw/master/test/data/test.xlsx) > [test.json](https://github.com/yuhei-a/excel2json/blob/master/test/data/test.json)
```
var excel2json = require('excel2json');

excel2json.parse('test.xlsx', [], function(err, sheetDatas) {
    excel2json.toJson(sheetDatas, function(err, result) {
        fs.writeFileSync('test.json', JSON.stringifi(result, null, 4));
    });
});
```

## Test
Run `npm test` and `npm run-script jshint`
