// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
//
// Copyright (C) 2020, Sol Boucher <sol@vanguardcampaign.org>
// Copyright (C) 2020, The Vanguard Campaign Corps Mods (vanguardcampaign.org)

const OUTPUT_SHEET = 'MiniVAN';

const PRIORITY_OFFSET = -7;
const COUNTY_OFFSET = -6;
const PRECINCT_OFFSET = -5;
const COUNT_OFFSET = -2;
const TURFPACKET_OFFSET = -1;

const COMPLETION_FORMULA = '=sumif(\'MiniVAN Commit\'!$A:$A, "List "&indirect("R[0]C[-1]", FALSE), \'MiniVAN Commit\'!$F:$F)/indirect("R[0]C[-2]", FALSE)';

const SORT_PRIORITY = 1;
const SORT_COUNTY = 2;
const SORT_PRECINCT = 3;
const SORT_TURF = 4;
const SORT_ORDER = [SORT_PRIORITY, SORT_COUNTY, SORT_PRECINCT, SORT_TURF];

const PDFTOTEXT_DEPLOYMENT = 'https://orgvanize-pdftotext.herokuapp.com';

function TURFPACKET(url_skip) {
  if(url_skip.length != 1 || url_skip[0].length <= 1)
    throw 'Expected a horizontal cell range';
  url_skip = url_skip[0];
  
  var url = url_skip[0];
  var skip = url_skip[url_skip.length - 1];
  if(!url.startsWith('http') || String(skip).length)
    return;
  
  var file = url.match(/([^\/]+)(\/[^\/]*)?$/);
  if(!file)
    throw 'Unparseable URL';
  file = file[1];
  
  var pdf = fetch('https://drive.google.com/uc?id=' + file).getBytes();
  var txt = fetch(PDFTOTEXT_DEPLOYMENT + '/?layout&l=2', pdf).getDataAsString();
  if(txt.startsWith('\n'))
    throw txt.split('\n')[1];
  
  txt = txt.replace(/\f/g, '');
  return txt.split('\n').filter(function(elem) {
    return elem.match(/^[0-9-]+  +Turf [0-9]+/);
  }).join('\n');
}

function onEdit(event) {
  var range = event.range;
  if(range.getNumRows() != 1 || range.getNumColumns() != 1 || !range.isChecked())
    return;
  
  var sheet = event.source.getActiveSheet();
  var parse = rowcell(sheet, range, TURFPACKET_OFFSET);
  if(!parse.getFormula().toUpperCase().startsWith('=TURFPACKET('))
    return;
  
  var turfs = parse.getValue();
  if(!turfs || turfs.startsWith('#')) {
    range.uncheck();
    return;
  }
  
  var priority = rowcell(sheet, range, PRIORITY_OFFSET).getValue();
  var county = rowcell(sheet, range, COUNTY_OFFSET).getValue();
  var precinct = rowcell(sheet, range, PRECINCT_OFFSET).getValue();
  turfs = turfs.split('\n').map(function(elem) {
    elem = elem.replaceAll(/ +/g, ' ').split(' ');
    
    var turf = elem[2];
    var doors = elem[4];
    var list = elem[0];
    return [priority, county, precinct, turf, doors, list, COMPLETION_FORMULA];
  });
  rowcell(sheet, range, COUNT_OFFSET).setValue(turfs.length);
  
  var shot = event.source.getSheetByName(OUTPUT_SHEET);
  for(var turf of turfs)
    shot.appendRow(turf);
  
  var data = shot.getRange(3, 1, shot.getLastRow() - 2, shot.getLastColumn());
  data.sort(SORT_ORDER);
}

function rowcell(sheet, range, offset) {
  return sheet.getRange(range.getRow(), range.getColumn() + offset);
}

function fetch(resource, payload, authorization) {
  var config = {};
  if(payload) {
    config.method = 'post';
    config.payload = payload;
  }
  if(authorization)
    config.headers = {
      Authorization: 'Bearer ' + authorization,
    };
  return UrlFetchApp.fetch(resource, config).getBlob();
}
