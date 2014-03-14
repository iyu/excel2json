excel2json
==========

excel2json

## Install

## Use
example sheet.xlsx

|   | A      | B        | C                | D |
|:-:|:-------|:---------|:-----------------|---|
| 1 | {}     |          |                  |   |
| 2 | _id    | obj.code | obj.value:number |   |
| 3 |        |          |                  |   |
| 4 | first  | one      | 1                |   |
| 5 | second | two      | 2                |   |
| 6 |        |          |                  |   |

```
var excel2json = require('excel2json');


excel2json('sheet.xlsx', function(err, data) {
    console.log(data);
    // [ { _id: 'first', { code: 'one', value: 1 } }, { _id: 'second', { code: 'two', value: 2 } } ]
});
```


```
{
    _id: 'first'
    obj: {
        list: [
            { code: 'one' },
            { code: 'two' }
        ]
    },
    list: [
        { list2: [ 1, 2 ] },
        { list2: [ 3, 4 ] }
    ]
}
```
|   | A      | B              | C           | D                   | E |
|:-:|:-------|:---------------|:------------|:--------------------|---|
| 1 | {}     |                |             |                     |   |
| 2 | _id    | obj.#list.code | #list:index | #list.#list2:number |   |
| 3 |        |                |             |                     |   |
| 4 | first  | one            | 0           | 1                   |   |
| 5 |        | two            | 0           | 2                   |   |
| 6 |        |                | 1           | 3                   |   |
| 7 |        |                | 1           | 4                   |   |
| 8 |        |                |             |                     |   |

```
{
    _id: 'first'
    list: [
        {
            value: 10,
            list2: [
                {
                    bool: true    
                },
                {
                    bool: false
                }
            ]
        },
        {
            value: 20,
            list2: [
                {
                    bool: true    
                },
                {
                    bool: false
                }
            ]
        }
    ]
}
```
|   | A      | B           | C                  | D                         | E |
|:-:|:-------|:------------|:-------------------|:--------------------------|---|
| 1 | {}     |             |                    |                           |   |
| 2 | _id    | #list:index | #list.value:number | #list.#list2.bool:boolean |   |
| 3 |        |             |                    |                           |   |
| 4 | first  | 0           | 10                 | true                      |   |
| 5 |        | 0           |                    | false                     |   |
| 6 |        | 1           | 20                 | true                      |   |
| 7 |        | 1           |                    | false                     |   |
| 8 |        |             |                    |                           |   |

## Test
Run `npm test` and `npm run-script jshint`
