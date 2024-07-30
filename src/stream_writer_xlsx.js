const Readable = require("stream-browserify").Readable;
const PassThrough = require("stream-browserify").PassThrough;
const JSZip = require("jszip");
const xmlParts = require("./parts");
const xmlBlobs = require("./blobs");
const { getCellAddress, wrapRowsInStream, escapeXml } = require("./helpers");
const getStyles = require('./styles').getStyles;
const fs = require('fs');
// const { crc32 } = require("crc");

const defaultOptions = {
  inlineStrings: false,
  styles: [],
  styleIdFunc: (value, columnId, rowId) => rowId === 0 ? 4 : rowId % 2 === 0 ? 3 : 2,
};

class XlsxStreamWriter {
  constructor(dir = 'excels', styles = undefined, options = undefined) {
    this.options = Object.assign({}, defaultOptions, options);
    this.dir = dir;
    this.sharedStringsArr = [];
    this.sharedStringsMap = {};
    this.rowIndex = 0;
    this.uniqueCount = 0;
    this.addRowsComplete = true;
    this.finalizeSharedStringsComplete = false;
    this.sheetFile = `${this.dir}/sheet1.xml`;
    this.sharedStringsFile = `${this.dir}/sharedStrings.xml`;
    this.finalSharedStringsFile = `${this.dir}/tmp_sharedStrings.xml`;

    if (!styles) {
      styles = {
        header: { fill: '005CB7', format: '0.00', border: 1, font: 1 }, //font: 1 - white, 13, calibri, bold
        evenRow: { fill: 'FFFFFF', format: '0.00', border: 1, font: 0 }, //font: 0 - black, 10, calibri, normal
        oddRow: { fill: 'E4E4E6', format: '0.00', border: 1, font: 0 },
      }
    }

    if (styles) {
      if (styles.evenRow) {
        this.options.styles.push(styles.evenRow);
      }
      if (styles.oddRow) {
        this.options.styles.push(styles.oddRow);
      } else {
        this.options.styles.push(styles.evenRow);
      }
      if (styles.header) {
        this.options.styles.push(styles.header);
      } else {
        this.options.styles.push(styles.evenRow);
      }
    }

    this.xlsx = {
      "[Content_Types].xml": cleanUpXml(xmlBlobs.contentTypes),
      "_rels/.rels": cleanUpXml(xmlBlobs.rels),
      "xl/workbook.xml": cleanUpXml(xmlBlobs.workbook),
      "xl/styles.xml": cleanUpXml(getStyles(this.options.styles)),
      "xl/_rels/workbook.xml.rels": cleanUpXml(xmlBlobs.workbookRels),
    };

    this._initializeFiles();
  }

  _initializeFiles() {
    if (fs.existsSync(this.dir)) {
      fs.rmdirSync(this.dir, { recursive: true });
    }
    fs.mkdirSync(this.dir, { recursive: true });
    fs.writeFileSync(this.sheetFile, xmlParts.sheetHeader, 'utf8');
  }

  /**
   * Add rows to xlsx.
   * @param {Array | Readable} rowsOrStream array of arrays or readable stream of arrays
   * @return {undefined}
   */
  async addRows(rowsOrStream) {
    return new Promise(async (resolve, reject) => {
      try {
        while (!this.addRowsComplete) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
        this.addRowsComplete = false;
        let rowsStream;

        if (rowsOrStream instanceof Readable) {
          rowsStream = rowsOrStream;
        } else if (Array.isArray(rowsOrStream)) {
          rowsStream = wrapRowsInStream(rowsOrStream);
        } else {
          throw new Error("Argument must be an array of arrays or a readable stream of arrays");
        }

        const rowsToXml = this._getRowsToXmlTransformStream();
        const tsToString = this._getToStringTransforStream();
        const writeStream = fs.createWriteStream(this.sheetFile, { flags: 'a' });

        const handleStreamError = (err) => {
          this.addRowsComplete = false;
          console.error(err);
          reject(err);
        };

        rowsStream.on('error', handleStreamError);
        rowsToXml.on('error', handleStreamError);
        tsToString.on('error', handleStreamError);
        writeStream.on('error', handleStreamError);

        writeStream.on('finish', () => {
          this.addRowsComplete = true;
          resolve(true);
        });

        if (this.options.inlineStrings) {
          rowsStream.pipe(rowsToXml).pipe(tsToString).pipe(writeStream);
        } else {
          rowsStream.pipe(rowsToXml).pipe(writeStream);
        }

      } catch (err) {
        reject({ error: true, message: 'Error occurred while retrieving birthday data.' });
      }
    });
  };

  _getToStringTransforStream() {
    const ts = PassThrough();
    ts._transform = (data, encoding, callback) => {
      ts.push(data.toString(), "utf8");
      callback();
    };
    return ts;
  }

  _getRowsToXmlTransformStream() {
    const ts = PassThrough({ objectMode: true });
    ts._transform = (data, encoding, callback) => {
      const rowXml = this._getRowXml(data, this.rowIndex);
      ts.push(rowXml.toString(), "utf8");
      this.rowIndex++;
      callback();
    };
    return ts;
  }

  _getRowXml(row, rowIndex) {
    let rowXml = xmlParts.getRowStart(rowIndex);
    row.forEach((cellValue, colIndex) => {
      const cellAddress = getCellAddress(rowIndex + 1, colIndex + 1);
      const styleId = this.options.styleIdFunc(cellValue, colIndex, rowIndex);
      rowXml += this._getCellXml(cellValue, cellAddress, styleId);
    });
    rowXml += xmlParts.rowEnd;
    return rowXml;
  }

  _getCellXml(value, address, styleId = 0) {
    let cellXml;
    if (Number.isNaN(value) || value === null || typeof value === "undefined")
      cellXml = xmlParts.getStringCellXml("", address, styleId);
    else if (typeof value === "number")
      cellXml = xmlParts.getNumberCellXml(value, address, styleId);
    else cellXml = this._getStringCellXml(value, address, styleId);
    return cellXml;
  }

  _getStringCellXml(value, address, styleId) {
    const stringValue = String(value);
    return this.options.inlineStrings
      ? xmlParts.getInlineStringCellXml(escapeXml(String(value)), address, styleId)
      : xmlParts.getStringCellXml(this._lookupString(stringValue), address, styleId);
  }

  _lookupString(value) {
    let sharedStringIndex = this.sharedStringsMap[value];
    if (typeof sharedStringIndex !== "undefined") return sharedStringIndex;
    sharedStringIndex = this.sharedStringsArr.length;
    this.sharedStringsMap[value] = sharedStringIndex;
    this.sharedStringsArr.push(value);
    this.uniqueCount++;
    fs.appendFileSync(this.sharedStringsFile, xmlParts.getSharedStringXml(escapeXml(value)), 'utf8');
    return sharedStringIndex;
  }

  _finalizeSharedStringsFile() {
    const sharedStringsFooter = xmlParts.sharedStringsFooter;
    const header = xmlParts.getSharedStringsHeader(this.uniqueCount);

    return new Promise((resolve, reject) => {
      fs.writeFileSync(this.finalSharedStringsFile, header);
      const readStream = fs.createReadStream(this.sharedStringsFile, { highWaterMark: 1024 * 1024 });
      const writeStream = fs.createWriteStream(this.finalSharedStringsFile, { flags: 'a' });

      readStream.on('data', (chunk) => {
        writeStream.write(chunk);
      });
      readStream.on('end', () => {
        writeStream.end(sharedStringsFooter, () => {
          this.finalizeSharedStringsComplete = true;
          resolve();
        });
      });
      readStream.on('error', (err) => {
        reject(err);
      });
      writeStream.on('error', (err) => {
        reject(err);
      });
      
    });
  }

  async getFile() {
    while (!this.addRowsComplete) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    this._finalizeSharedStringsFile();
    while (!this.finalizeSharedStringsComplete) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    const zip = new JSZip();
    // add all static files
    Object.keys(this.xlsx).forEach(key => zip.file(key, this.xlsx[key]));
    // add footer to the sheet file
    fs.appendFileSync(this.sheetFile, xmlParts.sheetFooter);
    // add "xl/worksheets/sheet1.xml"
    zip.file("xl/worksheets/sheet1.xml", fs.readFileSync(this.sheetFile));
    // add "xl/sharedStrings.xml"
    zip.file("xl/sharedStrings.xml", fs.readFileSync(this.finalSharedStringsFile));

    const isBrowser =
      typeof window !== "undefined" &&
      {}.toString.call(window) === "[object Window]";

    return new Promise((resolve, reject) => {
      if (isBrowser) {
        zip
          .generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: {
              level: 4,
            },
            streamFiles: true,
          })
          .then(resolve)
          .catch(reject);
      } else {
        zip
          .generateAsync({
            type: "nodebuffer",
            platform: process.platform,
            compression: "DEFLATE",
            compressionOptions: {
              level: 4,
            },
            streamFiles: true,
          })
          .then(resolve)
          .catch(reject);
      }
    });
  }
}

function cleanUpXml(xml) {
  return xml.replace(/>\s+</g, "><").trim();
}

module.exports = XlsxStreamWriter;
