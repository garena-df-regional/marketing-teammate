"""
Build the Google Apps Script (Code.gs) from the current index.html template.
Run this whenever the HTML template changes, then copy Code.gs into Apps Script.

The HTML template is embedded as base64 to avoid any backtick / ${} escaping
issues inside Apps Script. The client-side JS stays byte-for-byte identical to
the Python-generated site.
"""
import base64, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, 'index.html')
OUT_DIR = os.path.join(HERE, 'apps-script')
os.makedirs(OUT_DIR, exist_ok=True)

with open(INDEX, encoding='utf-8') as f:
    html = f.read()

# Replace the baked-in data object with a placeholder.
# index.html contains:  const RAW = {....};\n\n// ── PIC badge helper
m = re.search(r'const RAW = .*?;\n\n// ── PIC badge helper', html, re.DOTALL)
if not m:
    raise SystemExit("Could not locate 'const RAW = ...;' anchor in index.html")
template = html[:m.start()] + 'const RAW = __DATA_JSON__;\n\n// ── PIC badge helper' + html[m.end():]

template_b64 = base64.b64encode(template.encode('utf-8')).decode('ascii')

# ─── GAS source ───────────────────────────────────────────────────────────────
GAS = r'''/**
 * DF REG MKT Teammate — Google Sheet → GitHub Pages publisher
 *
 * One-time setup:
 *   1. Extensions → Apps Script, paste this whole file into Code.gs, Save.
 *   2. Reload the Google Sheet. A "DF MKT Site" menu appears.
 *   3. DF MKT Site → 🔑 Set GitHub Token  → paste the fine-grained token.
 *   4. DF MKT Site → 🔄 Update Website     → publishes to GitHub Pages.
 *
 * To change the GitHub target, edit CONFIG below.
 * The HTML template is embedded (base64). To change the site design, edit
 * generate_site.py, re-run it, then re-run build_apps_script.py and repaste.
 */

var CONFIG = {
  owner:  'garena-df-regional',
  repo:   'marketing-teammate',
  branch: 'main',
  path:   'index.html'
};

var STD = ['scenario','pic','steps','prepare','timeline','qa','links'];

// Sheets with bespoke layouts; everything else is a card (Tab\d+) or generic table.
var SPECIAL = {'Index':'index', 'Reg & Local Contactor':'contactor', 'Social Media Link':'sm'};
var SPECIAL_LABEL = {index:'Index', contactor:'Contactor', sm:'Social Media Links'};

// Decide a sheet's render type, nav label, and stable id from its name.
function classify(name) {
  if (SPECIAL[name]) {
    var t = SPECIAL[name];
    return {type: t, label: SPECIAL_LABEL[t], id: '_' + t};
  }
  var m = name.match(/^\s*Tab\s*(\d+)\s*(.*)$/i);
  if (m) {
    var rest = (m[2] || '').replace(/^\s+|\s+$/g, '');
    return {type: 'card', label: 'Tab ' + m[1] + (rest ? ' · ' + rest : ''), id: 'tab' + m[1]};
  }
  var slug = name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'tab';
  return {type: 'generic', label: name, id: 'g_' + slug};
}

var TEMPLATE_B64 = '__TEMPLATE_B64__';

// ─── Menu ─────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DF MKT Site')
    .addItem('🔄 Update Website', 'updateWebsite')
    .addSeparator()
    .addItem('🔑 Set GitHub Token', 'setGitHubToken')
    .addToUi();
}

function setGitHubToken() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Set GitHub Token',
    'Paste the fine-grained GitHub token (Contents: Read and write on ' +
    CONFIG.owner + '/' + CONFIG.repo + '):',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var token = res.getResponseText().trim();
  if (!token) { ui.alert('No token entered.'); return; }
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  ui.alert('Token saved. You can now use "Update Website".');
}

// ─── Main ─────────────────────────────────────────────────────────────────────
function updateWebsite() {
  var ui = SpreadsheetApp.getUi();
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    ui.alert('No GitHub token set yet.\n\nRun "DF MKT Site → 🔑 Set GitHub Token" first.');
    return;
  }
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Building site…', 'DF MKT Site', 30);
    var data = buildData();
    pushFile(CONFIG.path, buildHtml(data), token);     // index.html (human UI)
    pushFile('content.md', buildMarkdown(data), token); // plain text for AI
    var url = 'https://' + CONFIG.owner.toLowerCase() + '.github.io/' + CONFIG.repo + '/';
    SpreadsheetApp.getActiveSpreadsheet().toast('Done. Live in ~1 min.', 'DF MKT Site', 8);
    ui.alert('Website updated.\n\n' + url + '\n\nChanges go live in about 1 minute.');
  } catch (e) {
    ui.alert('Update failed:\n\n' + e.message);
    throw e;
  }
}

// ─── Read the spreadsheet into the same data shape as generate_site.py ─────────
// NAV + tabs are built dynamically in sheet order, so new sheets appear with no
// code change: `TabN ...` → card page, anything else → generic table page.
function buildData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nav = [], tabs = {}, firstCard = false;
  ss.getSheets().forEach(function (sheet) {
    var cl = classify(sheet.getName());
    if      (cl.type === 'card')      tabs[cl.id] = parseCard(sheet);
    else if (cl.type === 'index')     tabs[cl.id] = parseIndex(sheet);
    else if (cl.type === 'contactor') tabs[cl.id] = readContactor(sheet);
    else if (cl.type === 'sm')        tabs[cl.id] = readSM(sheet);
    else                              tabs[cl.id] = parseGeneric(sheet);
    if (cl.type === 'card' && !firstCard) { nav.push({sep: true}); firstCard = true; }
    nav.push({id: cl.id, label: cl.label, type: cl.type});
  });
  return {nav: nav, tabs: tabs};
}

// Standard 7-column scenario tab (row1=title, row2=header, row3+=data)
function parseCard(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var rng = sheet.getRange(3, 1, lastRow - 2, 7);
  var vals = rng.getValues(), rich = rng.getRichTextValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v.some(function (x) { return x !== '' && x !== null; })) continue;
    var row = {};
    for (var j = 0; j < STD.length; j++) row[STD[j]] = clean(v[j]);
    for (var k = 0; k < 7; k++) {
      var url = extractLink(rich[i][k], sheet, 3 + i, 1 + k);
      if (url) row[STD[k] + '_url'] = url;
    }
    rows.push(row);
  }
  return rows;
}

function parseIndex(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var vals = sheet.getRange(3, 1, lastRow - 2, Math.max(3, sheet.getLastColumn())).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r.some(function (x) { return clean(x); })) continue;
    out.push({tab: clean(r[0]), name: clean(r[1]), desc: clean(r[2])});
  }
  return out;
}

// Generic table: first non-empty row = header, rest = rows (value + hyperlink per cell)
function parseGeneric(sheet) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return {headers: [], rows: []};
  var rng = sheet.getRange(1, 1, lastRow, lastCol);
  var vals = rng.getValues(), rich = rng.getRichTextValues();
  var headers = [], rows = [];
  for (var i = 0; i < vals.length; i++) {
    var raw = vals[i].map(clean);
    while (raw.length && raw[raw.length - 1] === '') raw.pop();
    if (!raw.some(function (x) { return x; })) continue;
    if (!headers.length) { headers = raw; continue; }
    var rowcells = [];
    for (var j = 0; j < headers.length; j++) {
      rowcells.push({text: clean(vals[i][j]), url: extractLink(rich[i][j], sheet, i + 1, j + 1) || ''});
    }
    rows.push(rowcells);
  }
  return {headers: headers, rows: rows};
}

function readContactor(sheet) {
  if (!sheet) return {team: [], contactors: []};
  var vals = sheet.getDataRange().getValues();
  var team = [], contactors = [], sec = null;
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r.some(function (x) { return x !== '' && x !== null; })) continue;
    var f = clean(r[0]);
    if (f.indexOf('Regional Team') !== -1) { sec = null; continue; }
    if (f === 'Member')                    { sec = 'team'; continue; }
    if (f.indexOf('Local MKT') !== -1)     { sec = null; continue; }
    if (f === 'Region')                    { sec = 'loc'; continue; }
    if (sec === 'team')      team.push({name: clean(r[0]), resp: clean(r[1]), email: clean(r[2])});
    else if (sec === 'loc')  contactors.push({region: clean(r[0]), contact: clean(r[1]), email: clean(r[2])});
  }
  return {team: team, contactors: contactors};
}

function readSM(sheet) {
  var sm = {off_hdr: [], off: [], non_hdr: [], non: [], esp_hdr: [], esp: []};
  if (!sheet) return sm;
  var vals = sheet.getDataRange().getValues();
  var sec = null;
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var first = clean(r[0]);
    if (first === 'Official Channels') { sec = 'off'; continue; }
    if (first === 'Non-Official')      { sec = 'non'; continue; }
    if (first === 'Esports Channel')   { sec = 'esp'; continue; }
    if (!r.some(function (x) { return x !== '' && x !== null; })) continue;

    if (sec === 'off') {
      if (first === 'Channel') sm.off_hdr = r.map(clean).filter(function (x) { return x; });
      else sm.off.push(padRow(r, sm.off_hdr.length));
    } else if (sec === 'non') {
      if (first === 'Channel') sm.non_hdr = r.map(clean).filter(function (x) { return x; });
      else sm.non.push(padRow(r, sm.non_hdr.length));
    } else if (sec === 'esp') {
      if (first === 'Channel') {
        sm.esp_hdr = r.map(clean).filter(function (x) { return x; });
      } else {
        var rowVals = r.map(function (x) { return clean(x) ? clean(x) : '—'; });
        while (rowVals.length && rowVals[rowVals.length - 1] === '—') rowVals.pop();
        if (rowVals.some(function (x) { return x !== '—'; })) sm.esp.push(rowVals);
      }
    }
  }
  return sm;
}

function padRow(r, n) {
  var out = [];
  for (var i = 0; i < n; i++) out.push(clean(r[i]) ? clean(r[i]) : '—');
  return out;
}

// Extract a hyperlink from a cell: rich-text link, first run link, or =HYPERLINK()
function extractLink(richValue, sheet, row, col) {
  try {
    if (richValue) {
      var u = richValue.getLinkUrl();
      if (u) return u;
      var runs = richValue.getRuns();
      for (var i = 0; i < runs.length; i++) {
        var ru = runs[i].getLinkUrl();
        if (ru) return ru;
      }
    }
  } catch (e) {}
  try {
    var formula = sheet.getRange(row, col).getFormula();
    var m = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);
    if (m) return m[1];
  } catch (e) {}
  return null;
}

function clean(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

// ─── Build the HTML from template + data ───────────────────────────────────────
function buildHtml(data) {
  var template = Utilities.newBlob(Utilities.base64Decode(TEMPLATE_B64)).getDataAsString('UTF-8');
  var json = JSON.stringify(data);
  return template.split('__DATA_JSON__').join(json);
}

// AI-friendly plain Markdown (mirrors generate_site.py build_markdown)
function buildMarkdown(data) {
  var nav = data.nav, tabs = data.tabs;
  var L = ['# DF REG MKT Teammate — Knowledge Base', ''];
  L.push('> Knowledge base for Local Marketing teams collaborating with DF Regional Marketing.');
  L.push('> Each section is a scenario: who the Regional PIC is, the process, what to prepare, timeline, Q&A, and related links.');
  L.push('');
  nav.forEach(function (item) {
    if (item.sep) return;
    var nid = item.id, label = item.label, typ = item.type;
    if (typ === 'index') {
      L.push('## ' + label);
      tabs[nid].forEach(function (r) { L.push('- **' + r.tab + '. ' + r.name + '** — ' + r.desc); });
      L.push('');
    } else if (typ === 'contactor') {
      var cc = tabs[nid];
      L.push('## ' + label);
      L.push('### Regional Team');
      cc.team.forEach(function (m) { L.push('- **' + m.name + '** (' + m.email + '): ' + m.resp); });
      L.push('### Local MKT Contactors');
      cc.contactors.forEach(function (m) { L.push('- **' + m.region + '**: ' + m.contact + ' (' + m.email + ')'); });
      L.push('');
    } else if (typ === 'sm') {
      var sm = tabs[nid];
      L.push('## ' + label);
      var smBlock = function (title, hdr, rows) {
        if (!rows || !rows.length) return;
        L.push('### ' + title);
        rows.forEach(function (row) {
          var pairs = [];
          for (var i = 0; i < hdr.length; i++) {
            if (i < row.length && row[i] && row[i] !== '—') pairs.push(hdr[i] + ': ' + row[i]);
          }
          if (pairs.length) L.push('- ' + pairs.join(' | '));
        });
      };
      smBlock('Official Channels', sm.off_hdr, sm.off);
      smBlock('Non-Official Channels', sm.non_hdr, sm.non);
      if (sm.esp && sm.esp.length) {
        L.push('### Esports Channels');
        sm.esp.forEach(function (row) {
          var cells = row.filter(function (x) { return x && x !== '—'; });
          if (cells.length) L.push('- ' + cells.join(' | '));
        });
      }
      L.push('');
    } else if (typ === 'generic') {
      var g = tabs[nid] || {headers: [], rows: []};
      L.push('## ' + label);
      g.rows.forEach(function (row) {
        var pairs = [];
        for (var i = 0; i < g.headers.length; i++) {
          var cell = row[i] || {};
          var val = cell.url || cell.text || '';
          if (val) pairs.push(g.headers[i] + ': ' + val);
        }
        if (pairs.length) L.push('- ' + pairs.join(' | '));
      });
      L.push('');
    } else {
      var rows = tabs[nid] || [];
      L.push('## ' + label);
      if (!rows.length) { L.push('_(No content yet.)_'); L.push(''); return; }
      rows.forEach(function (r) {
        L.push('### ' + (r.scenario || '(Untitled)'));
        if (r.pic)      L.push('- **Regional PIC**: ' + r.pic);
        if (r.steps)    L.push('- **Process Steps**: ' + r.steps);
        if (r.prepare)  L.push('- **What to Prepare**: ' + r.prepare);
        if (r.timeline) L.push('- **Timeline**: ' + r.timeline);
        if (r.qa)       L.push('- **Common Q&A**: ' + r.qa);
        if (r.links) {
          if (r.links_url) L.push('- **Related Links**: [' + r.links + '](' + r.links_url + ')');
          else L.push('- **Related Links**: ' + r.links);
        }
        L.push('');
      });
    }
  });
  return L.join('\n');
}

// ─── Push one file to GitHub via the contents API ──────────────────────────────
function pushFile(path, content, token) {
  var base = 'https://api.github.com/repos/' + CONFIG.owner + '/' + CONFIG.repo +
             '/contents/' + path;
  var headers = {
    Authorization: 'token ' + token,
    Accept: 'application/vnd.github+json'
  };

  // Get current file SHA (required to update an existing file)
  var sha = null;
  var getRes = UrlFetchApp.fetch(base + '?ref=' + CONFIG.branch, {
    method: 'get', headers: headers, muteHttpExceptions: true
  });
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  } else if (getRes.getResponseCode() !== 404) {
    throw new Error('GitHub GET ' + path + ' failed (' + getRes.getResponseCode() + '): ' +
                    getRes.getContentText());
  }

  var payload = {
    message: 'Update ' + path + ' from Google Sheet (' + new Date().toISOString() + ')',
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: CONFIG.branch
  };
  if (sha) payload.sha = sha;

  var putRes = UrlFetchApp.fetch(base, {
    method: 'put', headers: headers, contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var code = putRes.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('GitHub PUT ' + path + ' failed (' + code + '): ' + putRes.getContentText());
  }
}
'''

GAS = GAS.replace('__TEMPLATE_B64__', template_b64)

out = os.path.join(OUT_DIR, 'Code.gs')
with open(out, 'w', encoding='utf-8') as f:
    f.write(GAS)

print('Generated: ' + out)
print('Code.gs size: {:,} bytes (template b64: {:,})'.format(len(GAS), len(template_b64)))
