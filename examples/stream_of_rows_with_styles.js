const XlsxStreamWriter = require("../src/xlsx-stream-writer");
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
