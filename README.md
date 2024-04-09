# xlsx-streaming-writer
<a href="https://www.npmjs.com/package/semos-cloud-xlsx-streaming-writer">NPM site</a>
### Introduction
This is a Node.js package for streaming data into Excel file.
</br></br>
This package is used to quickly generate a very large excel xlsx files with some simple formating.
</br>
Best way to use it is by sending data batches to the stream as you won't encounter a memory problem.
</br></br>
This was rewritten from https://www.jsdelivr.com/package/npm/xlsx-stream-writer and changed to work with custom formating.
</br>
</br>
It uses JSZip to compress resulting structure
</br>
</br>
### Usage 
#### You can add rows with style:
```javascript
const XlsxStreamWriter = require("semos-cloud-xlsx-streaming-writer");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const styles = {
    header: {fill: '005CB7', format: '0.00', border: 1, font: 1}, //font: 1 - white, 13, calibri, bold
    evenRow: {fill: 'FFFFFF', format: '0.00', border: 1, font: 0}, //font: 0 - black, 10, calibri, normal
    oddRow: {fill: 'E4E4E6', format: '0.00', border: 1, font: 0},
  }

const xlsx = new XlsxStreamWriter(styles);
xlsx.addRows(rows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```
#### Or add readable stream of rows with style:
```javascript
const XlsxStreamWriter = require("semos-cloud-xlsx-streaming-writer");
const { wrapRowsInStream } = require("../src/helpers");
const fs = require("fs");

const rows = [
  ["Name", "Location"],
  ["Alpha", "Adams"],
  ["Bravo", "Boston"],
  ["Charlie", "Chicago"],
];

const styles = {
    header: {fill: '005CB7', format: '0.00', border: 1, font: 1}, //font: 1 - white, 13, calibri, bold
    evenRow: {fill: 'FFFFFF', format: '0.00', border: 1, font: 0}, //font: 0 - black, 10, calibri, normal
    oddRow: {fill: 'E4E4E6', format: '0.00', border: 1, font: 0},
  }

const streamOfRows = wrapRowsInStream(rows);

const xlsx = new XlsxStreamWriter(styles);
xlsx.addRows(streamOfRows);

xlsx.getFile().then(buffer => {
  fs.writeFileSync("result.xlsx", buffer);
});
```
### Generate EXCEL with batch streming
```javascript
async function main(req, localization) {
        let pageNumber = 0;
        const rs = wrapRowsInStream([]);
        rs.setMaxListeners(1000000);
        try {
            while (true) {
                const dataChunk = await fetchDataChunk(pageNumber, req, localization);
                if(pageNumber === 1){
                  break;
                }else{
                  await writeToExcelStream(dataChunk, rs);
                }
                pageNumber++;
            }
        } catch (error) {
            console.error('Error occurred:', error);
        } finally {
          
            const styles = {
              header: {fill: '005CB7', format: '0.00', border: 1, font: 1}, //font: 1 - white, 13, calibri, bold
              evenRow: {fill: 'FFFFFF', format: '0.00', border: 1, font: 0}, //font: 0 - black, 10, calibri, normal
              oddRow: {fill: 'E4E4E6', format: '0.00', border: 1, font: 0},
            }
            const xlsx = new XlsxStreamWriter(styles);
            xlsx.addRows(rs);

            xlsx.getFile().then(buffer => {
              fs.writeFileSync("package.xlsx", buffer);
            });

            console.log('Excel file generated successfully.');
        }
      }

      await main(req, localization);
```
### Important
Note that for creating styles evenRow is <b>MANDATORY</b> but header and oddRow are optional</br>
#### This package offers minimal customization options.
</br>
