# node-docx-tables
Node library to extract tables from docx (office open xml) documents

## Install
``` bash
npm i --save docx-tables
```

## Usage
``` JavaScript
const docxTables = require('docx-tables')

docxTables({
  file: 'path/to/the/docx/file'
}).then((data) => {
  // .docx table data
  console.log(data)
}).catch((error) => {
  console.error(error)
})
```

### Example

| a | b
|-|-|
| c | d

A .docx file containing above table will result following JSON output:
```
{ '0':
   [ { position: { row: 0, col: 0 }, data: 'a\n' },
     { position: { row: 0, col: 1 }, data: 'b\n' } ],
  '1':
   [ { position: { row: 1, col: 0 }, data: 'c\n' },
     { position: { row: 1, col: 1 }, data: 'd\n' } ] }
```

# License
[MIT](https://opensource.org/licenses/MIT)

## Author
[saikksub](https://github.com/saikksub)
