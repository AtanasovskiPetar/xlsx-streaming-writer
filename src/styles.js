const helpers = require("./helpers");

const replaceRegex = /\s+/g;
const replaceReSec = />\s+</g;

const header = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`;

const bottom = "</styleSheet>";

const getFillXmlHeader = numFills => `<fills count="${numFills}">`;
const fillXmlDefault = [
  `<fill>
  <patternFill patternType="none"/>
  </fill>`,
  `<fill>
  <patternFill patternType="gray125"/>
  </fill>`,
];

const getFillXml = fillColor =>
  `<fill><patternFill patternType="solid"><fgColor rgb="${fillColor}"/><bgColor indexed="64"/></patternFill></fill>`;

const fillXmlBottom = "</fills>";

const fontsXml = `
  <fonts count="2" x14ac:knownFonts="1">
    <font>
      <sz val="10"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="13"/>
      <color theme="2"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>`;

const bordersXml = `
  <borders count="2">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left style="thin">
        <color auto="1"/>
      </left>
      <right style="thin">
        <color auto="1"/>
      </right>
      <top style="thin">
        <color auto="1"/>
      </top>
      <bottom style="thin">
        <color auto="1"/>
      </bottom>
      <vertical style="thin">
        <color auto="1"/>
      </vertical>
      <horizontal style="thin">
        <color auto="1"/>
      </horizontal>
    </border>
  </borders>`;


const getCellXfXml = ({ numFmtId, fillId, border, font }) =>
  `<xf numFmtId="${numFmtId === undefined ? 0 : numFmtId}" fontId="${font}" fillId="${
    fillId === undefined ? 0 : fillId
  }" borderId="${border}" xfId="0">
    <alignment horizontal="center" vertical="center"/>
  </xf>`;

const cellXfXmlDefault = [
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0">
      <alignment horizontal="center" vertical="center"/>
    </xf>`,
    `<xf numFmtId="0" fontId="0" fillId="1" borderId="1" xfId="0">
			<alignment horizontal="center" vertical="center"/>
		</xf>`
  ];

function getCellXfsBlock(cellXfs) {
  return `<cellXfs count="${cellXfs.length}">${cellXfs.join("")}</cellXfs>`;
}

const restXml = `<cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
                 defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
             xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
    </extLst>`;

const compact = xml =>
  xml
    .replace(replaceRegex, " ")
    .replace(replaceReSec, "><")
    .trim();

/**
 * @param { Array<Object> } styles
 * each style could have { fill, format }
 * Numbering Formats
   Fonts
   Fills
   Borders
   Cell Style Formats
   Cell Formats <== cell styleindex is referring to one of these
   ...the rest
 * @returns { String } styles.xml string
 * */

function getStyles(styles) {
  const NUM_FORMATS_START = 166;
  const numFormatsXml = [];
  const numFormatsIndex = {};
  const fillsXml = fillXmlDefault;
  const fillsIndex = {};
  const cellXfsXml = cellXfXmlDefault;
  styles.forEach(style => {
    const { fill, format, border, font } = style;
    if (format !== undefined) {
      if (numFormatsIndex[format] === undefined) {
        const formatIndex = numFormatsXml.length + NUM_FORMATS_START;
        numFormatsIndex[format] = formatIndex;
        numFormatsXml.push(
          getFormatXml(helpers.escapeXmlExtended(format), formatIndex),
        );
      }
    }
    if (fill !== undefined) {
      if (fillsIndex[fill] === undefined) {
        fillsIndex[fill] = fillsXml.length;
        fillsXml.push(getFillXml(helpers.escapeXmlExtended(fill)));
      }
    }
    cellXfsXml.push(
      getCellXfXml({
        numFmtId: numFormatsIndex[format],
        fillId: fillsIndex[fill],
        border: border,
        font: font,
      }),
    );
  });

  let xml = "";
  xml += header;
  xml += getNumFormatsXmlBlock(numFormatsXml);
  xml += fontsXml;
  xml += getFillXmlBlock(fillsXml);
  xml += bordersXml;
  xml += getCellXfsBlock(cellXfsXml);
  xml += restXml;
  xml += bottom;
  return compact(xml);
}

const getFormatXml = (format, length) =>
  `<numFmt numFmtId="${length}" formatCode="${format}"/>`;

function getNumFormatsXmlBlock(formats) {
  if (!Array.isArray(formats) || !formats.length) return "";
  return `<numFmts count="${formats.length}">${formats.join("")}</numFmts>`;
}

function getFillXmlBlock(fillsXml) {
  return getFillXmlHeader(fillsXml.length) + fillsXml.join("") + fillXmlBottom;
}

module.exports = { getStyles };
