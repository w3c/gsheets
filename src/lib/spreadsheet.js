/* eslint-env node */

"use strict";

const config = require("../config.json");
const { GoogleSpreadsheet } = require('google-spreadsheet');

function toMarkdown(options) {
  const str = s => s.replace(/\|/g, "--").replace(/\n/g, "<br />");
  let text = "";
  text += `# ${this.title}\n`;
  for (let sheet_index = 0; sheet_index < this.sheetCount; sheet_index++) {
    const sheet = this.sheetsByIndex[sheet_index];
    text += `\n## ${sheet.title}\n\n`;
    if (sheet.headerValues) {
      text += `| ${sheet.headerValues.map(str).join(' | ')} |\n`;
      text += `| ${sheet.headerValues.map(h => "--").join(' | ')} |\n`;
    } else {
      text += `| ${sheet.rows[0]._rawData.map(h => " ").join(' | ')} |\n`;
      text += `| ${sheet.rows[0]._rawData.map(h => "--").join(' | ')} |\n`;
    }
    if (sheet.rows) {
      sheet.rows.forEach(row => {
        for (let col = 0; col < row.cells.length; col++) {
          const cell = row.cells[col];
          let content = str(cell.formattedValue || "");
          if (cell.hyperlink) {
            content = `[${cell.hyperlink}](${content})`;
          }
          text += `| ${content} `;
        }
        text += " |\n";
      });
    }
  }
  if (!options || options.nosource !== false) {
    text += `\n<hr>\n\nGenerated from a [Google spreadsheet](https://docs.google.com/spreadsheets/d/${this.spreadsheetId}/)\n`;
  }
  return text;
}

function toJSON() {

  const ngs = {};
  ngs.title = this.title;
  ngs.sheets = [];
  for (let sheet_index = 0; sheet_index < this.sheetCount; sheet_index++) {
    const sheet = this.sheetsByIndex[sheet_index];
    const ngss = {};
    ngss.title = sheet.title;
    ngss.headerValues = sheet.headerValues;
    ngss.sheedId = sheet.sheetId;
    if (sheet.rows) {
      ngss.rows = sheet.rows.map(row => {
        const ngssr = [];
        for (let col = 0; col < row.cells.length; col++) {
          const cell = row.cells[col];
          let content = cell.formattedValue || "";
          if (cell.hyperlink) {
            content = cell.hyperlink;
          }
          ngssr.push(content);
        }
        return ngssr;
      });
      ngss.data = sheet.rows.map(row => {
        const d = {};
        const headers = row.headerValues || sheet.headerValues || [];
        headers.forEach(header => {
          if (header !== "") d[header] = row[header];
        });
        return d;
      });
      ngs.sheets.push(ngss);
    }
  }
  return ngs;
}


function toHTML(options) {
  const str = s => {
    s = s.trim().replace(/\n/g, "<br />");
    if (s.startsWith("https://")) {
      let index = s.indexOf(' ');
      if (index !== -1) {
        return `<a href='${s.substring(0, index)}'>${s.substring(index)}</a>`;
      } else {
        return `<a href='${s}'>${s.substring(8)}</a>`;
      }
    }
    return s;
  };
  let text = `<html><head><meta charset=utf-8><title>${this.title}</title>\n`;
  text+="<style>body {margin: 5ex;}th,td{border:1px solid black;padding:.5ex;}th{border:2px solid black;}td.empty{border:none;}table{border-collapse:collapse}"
  + 'h2 > a.self-link::before{ content: "§"; color: inherit; font-size: 83%; margin-left: -1em; opacity: 0.5; padding-right: 1ex; text-decoration: none}</style>\n';
  text+= "</head><body>\n";
  text += "<h1>";
  if (!options || options.nologo !== false) {
    text += '<svg xmlns="http://www.w3.org/2000/svg" style="display:inline" version="1.1" width="40px" height="40px" x="0" y="0" viewBox="0 0 40 40" preserveAspectRatio="none">'
+'<defs>'
+'<linearGradient id="a" x1="50.005%" x2="50.005%" y1="8.586%" y2="100.014%">'
+'<stop stop-color="#263238" stop-opacity=".2" offset="0%"/><stop stop-color="#263238" stop-opacity=".02" offset="100%"/>'
+'</linearGradient>'
+'<radialGradient id="b" cx="3.168%" cy="2.718%" r="161.248%" fx="3.168%" fy="2.718%" gradientTransform="matrix(1 0 0 .72222 0 .008)"><stop stop-color="#FFF" offset="0%"/>'
+'<stop stop-color="#FFF" stop-opacity="0" offset="100%"/></radialGradient>'
+'</defs>'
+'<g fill="none" fill-rule="evenodd">'
+'<path fill="#0F9D58" d="M9.5 2H24l9 9v24.5c0 1.3807119-1.1192881 2.5-2.5 2.5h-21C8.11928813 38 7 36.8807119 7 35.5v-31C7 3.11928813 8.11928813 2 9.5 2z"/>'
+'<path fill="#263238" fill-opacity=".1" d="M7 35c0 1.3807119 1.11928813 2.5 2.5 2.5h21c1.3807119 0 2.5-1.1192881 2.5-2.5v.5c0 1.3807119-1.1192881 2.5-2.5 2.5h-21C8.11928813 38 7 36.8807119 7 35.5V35z"/>'
+'<path fill="#FFF" fill-opacity=".2" d="M9.5 2H24v.5H9.5C8.11928813 2.5 7 3.61928813 7 5v-.5C7 3.11928813 8.11928813 2 9.5 2z"/><path fill="url(#a)" fill-rule="nonzero" d="M17.5 8l8.5 8.5V9" transform="translate(7 2)"/>'
+'<path fill="#87CEAC" d="M24 2l9 9h-6.5C25.1192881 11 24 9.88071187 24 8.5V2z"/><path fill="#F1F1F1" d="M13 18h14v14H13V18zm2 2v2h4v-2h-4zm0 4v2h4v-2h-4zm0 4v2h4v-2h-4zm6-8v2h4v-2h-4zm0 4v2h4v-2h-4zm0 4v2h4v-2h-4z"/>'
+'<path fill="white" fill-opacity=".1" d="M2.5 0H17l9 9v24.5c0 1.3807119-1.1192881 2.5-2.5 2.5h-21C1.11928813 36 0 34.8807119 0 33.5v-31C0 1.11928813 1.11928813 0 2.5 0z" transform="translate(7 2)"/>'
+'</g></svg>'
+`<span  style="vertical-align: top">${this.title}</span>`;
  } else {
    text += this.title;
  }
  text += '</h1>\n';
  for (let sheet_index = 0; sheet_index < this.sheetCount; sheet_index++) {
    const sheet = this.sheetsByIndex[sheet_index];
    text += `\n<h2 id='sheet${sheet.sheedId}'><a class='self-link' aria-label='§' href='#sheet${sheet.sheetId}'></a>${sheet.title}</h2>\n\n<table><thead><tr>`;
    if (sheet.headerValues) {
      text += `<th>${sheet.headerValues.map(str).join('</th><th>')}</th>\n`;
    } else {
      text += `<th>${sheet.rows[0]._rawData.map(h => "&nbsp;").join('</th><th>')}</th>\n`;
    }
    text += "</thead>\n<tbody>\n";
    if (sheet.rows) {
      sheet.rows.forEach(row => {
        text += `<tr>`;
        for (let col = 0; col < row.cells.length; col++) {
          const cell = row.cells[col];
          const rs = cell.effectiveFormat;
          let style = "";
          if (rs) {
            if (rs.textFormat.italic) {
              style+="font-style:italic;";
            }
            if (rs.textFormat.bold) {
              style+="font-weight:bold;";
            }
            if (rs.textFormat.strikethrough) {
              style+="text-decoration:line-through;";
            } else if (rs.textFormat.underline) {
              style+="text-decoration:underline;";
            }
            function color(gsc) {
              if (!gsc.red) { // weird, I get { green: 1, blue: 1 }
                gsc.red = gsc.green;
              }
              if (gsc.red === 1 && gsc.green === 1 && gsc.blue === 1) {
                return "white";
              }
              if (gsc.red === 0 && gsc.green === 0 && gsc.blue === 0) {
                return "black";
              }
              return `rgb(${(gsc.red*100).toPrecision(3)}% ${(gsc.green*100).toPrecision(3)}% ${(gsc.blue*100).toPrecision(3)}%)`;
            }
            if (rs.backgroundColor) {
              const c = color(rs.backgroundColor);
              if (c !== "white") {
                style+=`background-color:${c};`;
              }
            }
            if (rs.textFormat.foregroundColor && rs.textFormat.foregroundColor.red) {
              const c = color(rs.textFormat.foregroundColor);
              if (c !== "black") {
                style+=`color:${c};`;
              }
            }
          }
          if (style !== "") {
            style = ` style="${style}"`;
          }
          let content = cell.formattedValue || "&nbsp;";
          if (cell.hyperlink) {
            content = `<a href="${cell.hyperlink}">${content}</a>`;
          }
          text += `<td${style}>${content}</td>`;
        }
        if (row.cells.length === 0) {
          text += "<td class='empty'>&nbsp;</td>";
        }
        text += `</tr>\n`;
      });
    }
    text += "</tbody>\n</table>\n";
  }
  if (!options || options.nosource !== false) {
    text += `<hr><p>Generated from a <a href="https://docs.google.com/spreadsheets/d/${this.spreadsheetId}/">google spreadsheet</a>.</p>\n`;
  }
  text += "</body></html>";
  return text;
}

async function fetch_spreadsheet(id) {
  // if it's a link
  const GOOGLE_URL = 'https://docs.google.com/spreadsheets/d/';
  if (id.startsWith(GOOGLE_URL)) {
    id = id.substring(GOOGLE_URL.length);
    let idx = id.indexOf('/');
    if (idx !== -1) {
      id = id.substring(0, idx);
    }
  }

  const doc = new GoogleSpreadsheet(id);
  doc.useApiKey(config.googleKey);

  await doc.loadInfo();

  doc.toJSON = toJSON;
  doc.toHTML = toHTML;
  doc.toMarkdown = toMarkdown;
  doc.date = new Date();
  // we for ce the load of all of the cells and enhance the object a bit...
  const ropes = [];
  for (let sheet_index = 0; sheet_index < doc.sheetCount; sheet_index++) {
    const sheet = doc.sheetsByIndex[sheet_index];
    await sheet.loadCells();
    const p = sheet.getRows().then(rows => {
      sheet.rows = rows.map(row => {
        row.cells = [];
        if (row._rawData) {
          for (let idx = 0; idx < row._rawData.length; idx++) {
            const cell = sheet.getCell(row.rowNumber-1, idx);
            row.cells.push(cell._rawData);
          }
        }
        return row;
      });
      // return undefined
    });
    ropes.push(p);
  }
  return Promise.all(ropes).then(() => {
    return doc;
  });
}

module.exports = fetch_spreadsheet;
