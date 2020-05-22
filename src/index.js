/* eslint-env node */

"use strict";

const fs = require('fs').promises;

const monitor = require('./lib/monitor.js');
const spreadsheet = require('./lib/spreadsheet.js');
const { Repository } = require('./lib/github.js');
const spreadsheets = require('../spreadsheets.json');

const GH = "https://github.com/\([^/]+/[^/]+\)/blob/\([^/]+\)/\(.*\)";

async function save_spreadsheet(location, doc, options) {
  let content;
  if (location.endsWith(".json")) {
    content = JSON.stringify(doc, null, " ");
  } else if (location.endsWith(".md")) {
    content = doc.toMarkdown(options);
  } else if (location.endsWith(".html")) {
    content = doc.toHTML(options);
  }
  if (location.startsWith("https://github.com/")) {
    let branch;
    let path;
    let repo;
    let match = location.match(GH);
    if (match) {
      repo = match[1];
      branch = match[2];
      path = match[3];
    } else {
      monitor.error(`not a valid location ${location}`);
      return;
    }
    repo = new Repository(repo);
    repo.createContent(path, "Spreadsheet snapshot", content, branch).then(res => {
      switch (res.status) {
        case 200:
          monitor.log(`Updated ${doc.spreadsheetId} into ${location}`);
          break;
        case 201:
          monitor.log(`Created ${doc.spreadsheetId} into ${location}`);
          break;
        default:
          monitor.error(`Unexpected status ${res.status} ${doc.spreadsheetId}`);
      }
    }).catch(err => {
      monitor.error("We got an error");
      console.error(err);
    })
  } else {
    monitor.error(`not a valid location ${location}`);
  }
}

async function run() {
  let doc;

  for (let index = 0; index < spreadsheets.length; index++) {
    const entry = spreadsheets[index];
    if (entry.id === null || entry.location === null) {
      monitor.log(`invalid entry ${entry.id} ${entry.location}`);
      continue;
    }
    const doc = await spreadsheet(entry.id).catch(err => {
      monitor.log(err);
      return undefined;
    });
    if (!doc) continue;
    if (Array.isArray(entry.location)) {
      entry.location.forEach(location => save_spreadsheet(location, doc, entry.options));
    } else { // String
      save_spreadsheet(entry.location, doc, entry.options);
    }
  }
}

run();
