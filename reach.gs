// @OnlyCurrentDoc

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
// Copyright (C) 2022 Sol Boucher <sol@vanguardcampaign.org>
// Copyright (C) 2022 The Vanguard Campaign Corps Mods

const CALLBACK_TIMEOUT_S = 3600;
const ENDPOINT = 'https://api.reach.vote/api/v1/imports/tags/';
const OFFSET_READY = 2;
const OFFSET_STARTED = 3;
const OFFSET_FINISHED = 4;
const OFFSET_LOG = 5;
const SHEET_CONFIG = 'Configuration';
const SHEET_TAGS = 'Types';

function reachOut(url_tag_ready) {
  if(url_tag_ready.length != 1 || url_tag_ready[0].length != 3)
    throw 'Expected single 3 cell row';

  const [url, tag, ready] = url_tag_ready[0];
  if(!ready)
    return;

  const file = url.replace(/^.+\/([^/]+)\/[^/]+$/, '$1');
  const hits = _findCells(file);
  if(hits.length != 1)
    throw 'Duplicate file link';

  const row = hits[0].getRow();
  const col = hits[0].getColumn();
  const token = _generateAccessToken();
  const file_url = _createCallback(
    '_doGet',
    {file, row, col},
    token,
  );
  const callback_url = _createCallback(
    '_doPost',
    {row, col},
    token,
    CALLBACK_TIMEOUT_S,
  );
  let job_id = _importTag(tag, file_url, callback_url, true);
  if(!job_id) {
    UrlFetchApp.fetch(callback_url + '&regenerate=true', {method: 'post'});
    job_id = _importTag(tag, file_url, callback_url);
    if(!job_id)
      return 'Reach upload failed';
  }
  return job_id;
}

function _importTag(tag_name, file_url, callback_url, recover) {
  const tags = _getConfig(SHEET_TAGS);
  const tag_id = tags[tag_name];
  if(!tag_id)
    return false;

  const config = _getConfig(SHEET_CONFIG);
  const response = UrlFetchApp.fetch(
    ENDPOINT + tag_id + '?callback_url=' + encodeURIComponent(callback_url),
    {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + config.reach_bearer_token,
      },
      payload: '{"file_url": "' + file_url + '"}',
      muteHttpExceptions: Boolean(recover),
    },
  );
  if(response.getResponseCode() >= 400)
    return null;

  return JSON.parse(response.getContentText()).data.id;
}

function _doGet(event) {
  const args = event.parameter;
  _getCell(Number(args.row), Number(args.col) + OFFSET_STARTED).setFormula('=TRUE');

  let file = UrlFetchApp.fetch('https://drive.google.com/uc?id=' + args.file).getContentText();
  file = _fromUtf16(file).slice(1);
  file = file.replaceAll(/.+/gm, '$&,Voterfile VAN ID');
  file = file.replace(/.+/, 'person_id,person_id_type');
  return ContentService.createTextOutput(file);
}

function _doPost(event) {
  const args = event.parameter;
  if(args.regenerate == "true") {
    _regenerateBearerToken();
    return ContentService.createTextOutput();
  }

  const row = Number(args.row);
  const col = Number(args.col);
  const log = _getCell(row, col + OFFSET_LOG);
  const job_id = log.getValue();
  _getCell(row, col + OFFSET_FINISHED).setFormula('=TRUE');

  const config = _getConfig(SHEET_CONFIG);
  const response = JSON.parse(UrlFetchApp.fetch(ENDPOINT + job_id, {
    headers: {
      'Authorization': 'Bearer ' + config.reach_bearer_token,
    },
  }).getContentText());
  log.setValue(JSON.stringify(response, null, '\t'));
  return ContentService.createTextOutput();
}

function _fromUtf16(string) {
  return Array.from(string).filter(function(ignore, index) {
    return index % 2 == 0;
  }).join('');
}

function _createCallback(method, args, access_token, timeout) {
  if(!access_token)
    access_token = _generateAccessToken();

  const state = ScriptApp.newStateToken().withMethod(method);
  if(args)
    for(const [key, val] of Object.entries(args))
      state.withArgument(key, val);
  if(timeout)
    state.withTimeout(timeout);
  return 'https://script.google.com/d/' + ScriptApp.getScriptId() + '/usercallback'
    + '?state=' + state.createToken()
    + '&access_token=' + access_token
  ;
}

function _generateAccessToken() {
  const config = _getConfig(SHEET_CONFIG);
  const refresh = 'client_id=' + config.google_client_id
    + '&client_secret=' + config.google_client_secret
    + '&refresh_token=' + config.google_refresh_token
  ;
  if(refresh.includes('undefined'))
    return null;

  const oauth = JSON.parse(UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: 'grant_type=refresh_token&' + refresh,
  }).getContentText());
  if(!oauth.access_token) {
    console.error(oauth.error + ': ' + oauth.error_description);
    return null;
  }
  return oauth.access_token;
}

function _regenerateBearerToken() {
  const config = _getConfig(SHEET_CONFIG);
  const oauth = JSON.parse(UrlFetchApp.fetch('https://api.reach.vote/oauth/token', {
    method: 'post',
    payload: 'grant_type=password'
      + '&username=' + config.reach_username
      + '&password=' + config.reach_password
    ,
  }));
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const cell = _findCells(config.reach_bearer_token, sheet);
  cell[0].setValue(oauth.access_token);
}

function _getCell(row, col) {
  return SpreadsheetApp.getActiveSheet().getRange(row, col);
}

function _findCells(text, sheet) {
  if(!sheet)
    sheet = SpreadsheetApp.getActiveSheet();

  return sheet.getDataRange().createTextFinder(text).findAll();
}

function _getConfig(sheet_name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  if(!sheet)
    return null;

  const config = sheet.getDataRange().getValues();
  if(!config[0] || config[0].length != 2)
    return null;

  return Object.fromEntries(config);
}
