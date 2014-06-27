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

|   | A      | B        | C                | D |
|:-:|:-------|:---------|:-----------------|---|
| 1 | {}     |          |                  |   |
| 2 | _id    | obj.code | obj.value:number |   |
| 3 |        |          |                  |   |
| 4 | first  | one      | 1                |   |
| 5 | second | two      | 2                |   |
| 6 |        |          |                  |   |
Sheet1
```
var excel2json = require('excel2json');

var filename = './sheet.xlsx';
var sheets = [1];
excel2json.parse(filename, sheets, function(err, data) {
    console.log(data);
    // [{
        num: 1,         // sheet number
        name: 'Sheet1', // sheet name
        option: {},     // sheet option (A1)
        contents: [
            { _id: 'first', obj: { code: 'one', value: 1 } }, { _id: 'second', obj: { code: 'two', value: 2 } }
        ]
    }]
});
```

### Setup
Setup options.
```
var excel2json = require('excel2json');

excel2json.setup({
    optionCell: 'A1', // Cell with a custom sheet option. It is not yet used now. (default: 'A1'
    attrLine: '2',    // Line with a data attribute. (default: '2'
    dataLine: '4'     // Line with a data. (default: '4'
});
```

### Attribute
Specify the key name.

**Special character**
* `#` Use when the array.
* `:number` Use when the parameters of type `Number`. 
* `:boolean` Use when the parameters of type `Boolean`.
* `:date` Use when the parameters of unix time.
* `:index` Use when the array of array.

### An example of a complex format
Data to be expected.
```
[
    {
        _id: 'test1',
        arr: [
            {
                code: 'test1_1',
                list: [ [ 1, 2 ], [ 3, 4 ] ],
                arr: [
                    { code: 'test1_1_1', is: true },
                    { code: 'test1_1_2', is: false }
                ]
            },
            {
                code: 'test1_2',
                list: [ [ 5, 6, 7 ], [ 8 ] ],
                arr: [
                    { code: 'test1_2_1', is: true },
                    { code: 'test1_2_2', is: false }
                ]
            }
        ]
    },
    {
        _id: 'test2',
        arr: [
            {
                code: 'test2_1',
                list: [ [ 1 ], [ 2 ], [ 3 ], [ 4 ] ],
                arr: [
                    { code: 'test2_1_1', is: true },
                    { code: 'test2_1_2', is: false }
                ]
            },
            {
                code: 'test2_2',
                list: [ [ 5, 6, 7 ], [ 8 ] ],
                arr: [
                    { code: 'test2_2_1', is: true },
                    { code: 'test2_2_2', is: false }
                ]
            }
        ]
    }
]
```

Necessary Excel data.

|    | A      | B          | C         | D                | E                   | F              | G                    |
|:--:|:-------|:-----------|:----------|:-----------------|:--------------------|:---------------|:---------------------|
| 1  | {}     |            |           |                  |                     |                |                      |
| 2  | _id    | #arr:index | #arr.code | #arr.#list:index | #arr.#list.#:number | #arr.#arr.code | #arr.#arr.is:boolean |
| 3  |        |            |           |                  |                     |                |                      |
| 4  | test1  | 0          | test1_1   | 0                | 1                   | test1_1_1      | true                 |
| 5  |        | 0          |           | 0                | 2                   | test1_1_2      | false                |
| 6  |        | 0          |           | 1                | 3                   |                |                      |
| 7  |        | 0          |           | 1                | 4                   |                |                      |
| 8  |        | 1          | test1_2   | 0                | 5                   | test1_2_1      | true                 |
| 9  |        | 1          |           | 0                | 6                   | test1_2_2      | false                |
| 10 |        | 1          |           | 0                | 7                   |                |                      |
| 11 |        | 1          |           | 1                | 8                   |                |                      |
| 12 | test2  | 0          | test2_1   | 0                | 1                   | test2_1_1      | true                 |
| 13 |        | 0          |           | 1                | 2                   | test2_1_2      | false                |
| 14 |        | 0          |           | 2                | 3                   |                |                      |
| 15 |        | 0          |           | 3                | 4                   |                |                      |
| 16 |        | 1          | test2_2   | 0                | 5                   | test2_2_1      | true                 |
| 17 |        | 1          |           | 0                | 6                   | test2_2_2      | false                |
| 18 |        | 1          |           | 0                | 7                   |                |                      |
| 19 |        | 1          |           | 1                | 8                   |                |                      |

## Test
Run `npm test` and `npm run-script jshint`
