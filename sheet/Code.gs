/**
 * ff-manager draft sheet builder.
 *
 * Setup: create a blank Google Sheet -> Extensions -> Apps Script -> paste this
 * file -> run buildSheet() (authorize when asked) -> import out/blended.csv into
 * the Raw tab (File > Import > Replace data at selected cell, A1) and idp.csv
 * into RawIDP. Paste Yahoo XRank/ADP into the Yahoo tab (cols B:E). Set your
 * draft slot on Params. Use the "Draft Tools" menu during the draft.
 *
 * buildSheet() is safe to rerun: it rewrites formula tabs (Params, Master,
 * Board) but never touches data you own (Raw, RawIDP, Yahoo B:E, Targets, Log).
 *
 * Architecture: nothing is ever deleted. The draft macro appends to Log; every
 * panel is a QUERY over Master that excludes logged players, so formulas never
 * break and targets persist because they're keyed by fpid, not row.
 */

var POSITIONS = ['QB', 'RB', 'WR', 'TE'];
var BLOCK_COLS = 12;          // fpid,Player,Tm,Bye,Pts,XRk,ADP,PAR,Diff,P1,P2,Tgt
var BLOCK_GAP = 1;
var BOARD_DATA_ROW = 3;
var BOARD_ROWS = 34;          // players shown per panel
var MASTER_ROWS = 1200;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Draft Tools')
    .addItem('Mark drafted (other team)', 'markDrafted')
    .addItem('Draft to MY team', 'draftToMe')
    .addItem('Undo last pick', 'undoLastPick')
    .addSeparator()
    .addItem('Rebuild formulas & formatting', 'buildSheet')
    .addToUi();
}

function buildSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDataTabs_(ss);
  buildParams_(ss);
  buildMaster_(ss);
  buildBoard_(ss);
  ss.toast('Sheet built. Import blended.csv into Raw, paste Yahoo ranks, set your slot on Params.');
}

// ---------- data tabs (never overwritten if they hold data) ----------

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureDataTabs_(ss) {
  var raw = getOrCreate_(ss, 'Raw');
  if (raw.getLastRow() === 0) {
    raw.getRange('A1').setValue('Import out/blended.csv here: File > Import > Upload > Replace data at selected cell (A1)');
  }
  var idp = getOrCreate_(ss, 'RawIDP');
  if (idp.getLastRow() === 0) {
    idp.getRange('A1').setValue('Import out/idp.csv here the same way');
  }

  var yahoo = getOrCreate_(ss, 'Yahoo');
  yahoo.getRange('A1:F1').setValues([['key (auto)', 'Name', 'Pos', 'XRank', 'ADP', 'reconciled? (auto)']]);
  // Helper columns only; B:E belong to the user's paste.
  yahoo.getRange('A2').setFormula(
    '=IF($B2="","",REGEXREPLACE(REGEXREPLACE(LOWER($B2),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$","")&"|"&UPPER($C2))');
  yahoo.getRange('F2').setFormula(
    '=IF($B2="","",IF(COUNTIF(Master!$X$2:$X,$A2)>0,"ok","NO MATCH — fix name or add alias"))');
  yahoo.getRange('A2:A2').autoFill(yahoo.getRange('A2:A500'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  yahoo.getRange('F2:F2').autoFill(yahoo.getRange('F2:F500'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  var targets = getOrCreate_(ss, 'Targets');
  targets.getRange('A1:C1').setValues([['fpid', 'Player (auto)', 'tag']]);
  if (targets.getLastRow() < 2) {
    targets.getRange('B2').setFormula('=IF($A2="","",IFNA(VLOOKUP($A2,Master!$A:$B,2,FALSE),"fpid not found"))');
    targets.getRange('B2:B2').autoFill(targets.getRange('B2:B100'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['high target', 'target', 'mild target', 'downside'], true).build();
  targets.getRange('C2:C100').setDataValidation(rule);

  var log = getOrCreate_(ss, 'Log');
  log.getRange('A1:D1').setValues([['pick#', 'fpid', 'Player', 'mine']]);
}

// ---------- Params ----------

function buildParams_(ss) {
  var p = getOrCreate_(ss, 'Params');
  p.clear();
  p.getRange('A1:B6').setValues([
    ['Teams', 10],
    ['My draft slot', 5],
    ['Rounds', 20],
    ['Sigma min (picks)', 4],
    ['Sigma k (× rank)', 0.18],
    ['XRank weight (vs Yahoo ADP)', 0.7],
  ]);
  p.getRange('A8:B13').setValues([
    ['QB', 12], ['RB', 26], ['WR', 40], ['TE', 12], ['K', 10], ['DST', 10],
  ]);
  p.getRange('C7').setValue('replacement rank ->  baseline pts:');
  for (var i = 0; i < 6; i++) {
    p.getRange(8 + i, 4).setFormula(
      '=IFERROR(LARGE(FILTER(Master!$I$2:$I,Master!$C$2:$C=$A' + (8 + i) + '),$B' + (8 + i) + '),0)');
  }
  p.getRange('A15:A18').setValues([['Current pick'], ['My next pick'], ['My pick after'], ['Max PAR remaining']]);
  p.getRange('B15').setFormula('=COUNTA(Log!$B$2:$B)+1');
  p.getRange('B16').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>=$B$15)),999)');
  p.getRange('B17').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>$B$16)),999)');
  p.getRange('B18').setFormula('=IFERROR(MAX(FILTER(Master!$T$2:$T,Master!$Q$2:$Q=FALSE)),"")');
  p.getRange('E1:F1').setValues([['round', 'my pick']]);
  for (var r = 1; r <= 20; r++) {
    p.getRange(1 + r, 5).setValue(r);
    p.getRange(1 + r, 6).setFormula(
      '=IF(ISODD($E' + (1 + r) + '),($E' + (1 + r) + ')*$B$1-$B$1+$B$2,($E' + (1 + r) + ')*$B$1-$B$2+1)');
  }
  p.getRange('A20').setValue('Snake picks assume you mark EVERY pick in the draft (yours and others’) via Draft Tools.');
}

// ---------- Master ----------

function buildMaster_(ss) {
  var m = getOrCreate_(ss, 'Master');
  m.clear();
  m.getRange('A1:X1').setValues([[
    'fpid', 'Player', 'Pos', 'Tm', 'Bye', 'FfcADP', 'FPpts', 'WWOpts', 'Pts', 'src', 'wdiff', 'delta7',
    'norm', 'XRank', 'YADP', 'BehRank', 'Drafted', 'Mine', 'Tag', 'PAR', 'DiffTop', 'P1', 'P2', 'key',
  ]]);
  m.getRange('A2').setFormula('={Raw!A2:L' + MASTER_ROWS + '}');

  var f = {
    M: '=IF($A2="","",REGEXREPLACE(REGEXREPLACE(LOWER($B2),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))',
    N: '=IF($A2="","",IFNA(VLOOKUP($X2,Yahoo!$A:$E,4,FALSE),""))',
    O: '=IF($A2="","",IFNA(VLOOKUP($X2,Yahoo!$A:$E,5,FALSE),""))',
    P: '=IF($A2="","",IF($N2<>"",IF($O2<>"",Params!$B$6*$N2+(1-Params!$B$6)*$O2,$N2),' +
       'IF($O2<>"",$O2,IF($F2<>"",$F2,400))))',
    Q: '=IF($A2="","",COUNTIF(Log!$B:$B,$A2)>0)',
    R: '=IF($A2="","",COUNTIFS(Log!$B:$B,$A2,Log!$D:$D,TRUE)>0)',
    S: '=IF($A2="","",IFNA(VLOOKUP($A2,Targets!$A:$C,3,FALSE),""))',
    T: '=IF($A2="","",ROUND($I2-IFNA(VLOOKUP($C2,Params!$A$8:$D$13,4,FALSE),0),1))',
    U: '=IF($A2="","",ROUND(MAXIFS($I:$I,$C:$C,$C2,$Q:$Q,FALSE)-$I2,1))',
    V: '=IF($A2="","",IF($Q2,0,LET(sd,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'snow,1-NORMDIST(Params!$B$15,$P2,sd,TRUE),snext,1-NORMDIST(Params!$B$16,$P2,sd,TRUE),' +
       'IF(snow<0.0001,1,ROUND(snext/snow,3)))))',
    W: '=IF($A2="","",IF($Q2,0,LET(sd,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'snow,1-NORMDIST(Params!$B$15,$P2,sd,TRUE),sthen,1-NORMDIST(Params!$B$17,$P2,sd,TRUE),' +
       'IF(snow<0.0001,1,ROUND(sthen/snow,3)))))',
    X: '=IF($A2="","",$M2&"|"&$C2)',
  };
  for (var col in f) {
    m.getRange(col + '2').setFormula(f[col]);
    m.getRange(col + '2:' + col + '2').autoFill(m.getRange(col + '2:' + col + MASTER_ROWS),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  m.setFrozenRows(1);
}

// ---------- Board ----------

function colLetter_(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function blockStart_(i) { return 1 + i * (BLOCK_COLS + BLOCK_GAP); }

function buildBoard_(ss) {
  var b = getOrCreate_(ss, 'Board');
  b.clear();
  b.clearConditionalFormatRules();
  var headers = ['fpid', 'Player', 'Tm', 'Bye', 'Pts', 'XRk', 'ADP', 'PAR', 'Diff', 'P1', 'P2', 'Tgt'];
  var rules = [];

  POSITIONS.forEach(function (pos, i) {
    var bs = blockStart_(i);
    b.getRange(1, bs).setValue(pos).setFontWeight('bold').setFontSize(12);
    b.getRange(2, bs, 1, BLOCK_COLS).setValues([headers]).setFontWeight('bold');
    b.getRange(BOARD_DATA_ROW, bs).setFormula(
      '=IFERROR(QUERY(Master!$A$2:$X,"select A,B,D,E,I,N,O,T,U,V,W,S where C=\'' + pos +
      '\' and Q=false order by I desc limit ' + BOARD_ROWS + '",0),"")');
    b.setColumnWidth(bs, 40);
    b.setColumnWidth(bs + 1, 150);
    b.getRange(1, bs, BOARD_ROWS + 2, 1).setFontColor('#bbbbbb').setFontSize(8);

    var dataRange = b.getRange(BOARD_DATA_ROW, bs, BOARD_ROWS, BLOCK_COLS);
    var tgt = colLetter_(bs + 11), par = colLetter_(bs + 7);
    var p1Range = b.getRange(BOARD_DATA_ROW, bs + 9, BOARD_ROWS, 2);
    b.getRange(BOARD_DATA_ROW, bs + 9, BOARD_ROWS, 2).setNumberFormat('0%');

    [['high target', '#f4cccc'],   // pale red
     ['target', '#fce5cd'],        // pale orange
     ['mild target', '#fff2cc'],   // pale yellow
     ['downside', '#efefef'],      // pale grey
    ].forEach(function (t) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([dataRange])
        .whenFormulaSatisfied('=$' + tgt + BOARD_DATA_ROW + '="' + t[0] + '"')
        .setBackground(t[1]).build());
    });
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([b.getRange(BOARD_DATA_ROW, bs + 7, BOARD_ROWS, 1)])
      .whenFormulaSatisfied('=$' + par + BOARD_DATA_ROW + '=Params!$B$18')
      .setBackground('#b4a7d6').build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([p1Range])
      .setGradientMinpoint('#e06666').setGradientMaxpoint('#93c47d').build());
  });

  // Status block to the right of the panels.
  var sc = blockStart_(POSITIONS.length);
  b.getRange(1, sc, 4, 2).setValues([
    ['Current pick', ''], ['My next', ''], ['Then', ''], ['Max PAR', ''],
  ]);
  b.getRange(1, sc + 1).setFormula('=Params!$B$15');
  b.getRange(2, sc + 1).setFormula('=Params!$B$16');
  b.getRange(3, sc + 1).setFormula('=Params!$B$17');
  b.getRange(4, sc + 1).setFormula('=Params!$B$18');
  b.getRange(1, sc, 4, 2).setFontWeight('bold');

  b.setConditionalFormatRules(rules);
  b.setFrozenRows(2);
}

// ---------- draft macros ----------

function markDrafted() { logPick_(false); }
function draftToMe() { logPick_(true); }

function logPick_(mine) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var fpid, name;

  if (sheet.getName() === 'Board') {
    if (cell.getRow() < BOARD_DATA_ROW) { ss.toast('Select a player row first.'); return; }
    for (var i = 0; i < POSITIONS.length; i++) {
      var bs = blockStart_(i);
      if (cell.getColumn() >= bs && cell.getColumn() < bs + BLOCK_COLS) {
        fpid = sheet.getRange(cell.getRow(), bs).getValue();
        name = sheet.getRange(cell.getRow(), bs + 1).getValue();
      }
    }
  } else if (sheet.getName() === 'Master') {
    fpid = sheet.getRange(cell.getRow(), 1).getValue();
    name = sheet.getRange(cell.getRow(), 2).getValue();
  } else {
    ss.toast('Draft from the Board or Master tab.');
    return;
  }
  if (!fpid) { ss.toast('No player on the selected row.'); return; }

  var log = ss.getSheetByName('Log');
  if (log.createTextFinder(String(fpid)).matchEntireCell(true).findNext()) {
    ss.toast(name + ' is already drafted.');
    return;
  }
  log.appendRow([log.getLastRow(), fpid, name, mine]);
  ss.toast('Pick ' + log.getLastRow() + ': ' + name + (mine ? '  ->  MY TEAM' : ''), '', 3);
}

function undoLastPick() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName('Log');
  if (log.getLastRow() <= 1) { ss.toast('Nothing to undo.'); return; }
  var name = log.getRange(log.getLastRow(), 3).getValue();
  log.deleteRow(log.getLastRow());
  ss.toast('Undid: ' + name);
}
