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

// Panel order and sourcing. 'master' panels (12 cols: fpid,Player,Tm,Bye,Pts,
// XRk,ADP,PAR,Diff,P1,P2,Tgt) come from Master; 'idp' panels (4 cols: Player,
// Tm,Pts,Tgt) come from MasterIDP.
var BLOCKS = [
  { pos: 'QB',  title: 'QB',  type: 'master', rows: 40 },
  { pos: 'RB',  title: 'RB',  type: 'master', rows: 70 },
  { pos: 'WR',  title: 'WR',  type: 'master', rows: 90 },
  { pos: 'TE',  title: 'TE',  type: 'master', rows: 50 },
  { pos: 'LB',  title: 'LB',  type: 'idp',    rows: 30 },
  { pos: 'DB',  title: 'DB',  type: 'idp',    rows: 30 },
  { pos: 'DL',  title: 'DL',  type: 'idp',    rows: 30 },
  { pos: 'DST', title: 'DEF', type: 'master', rows: 14 },
  { pos: 'K',   title: 'K',   type: 'master', rows: 14 },
];
var BLOCK_GAP = 1;
var TITLE_ROW = 6;        // summary block lives in rows 1-4
var HEADER_ROW = 7;
var BOARD_DATA_ROW = 8;
var PARAMS_POS_ROW = { QB: 8, RB: 9, WR: 10, TE: 11, K: 12, DST: 13, LB: 14, DB: 15, DL: 16 };

// Conditional-format tuning. Rank columns (XRk/ADP): colour is strongest at
// the best rank, half-faded by CF_RANK_MID percentile, white by CF_RANK_FADE.
// PAR: purple starts appearing above CF_PAR_MID percent of the panel's range.
var CF_RANK_MID = 8;
var CF_RANK_FADE = 30;
var CF_PAR_MID = 88;  // percentile where PAR starts picking up purple

// Raw data published by the pipeline (see refresh.sh) — pulled by refreshData().
var DATA_URLS = {
  Raw: 'https://raw.githubusercontent.com/epgreig/ff-manager/master/out/blended.csv',
  RawIDP: 'https://raw.githubusercontent.com/epgreig/ff-manager/master/out/idp.csv',
};

// master block: [#, fpid(hidden), Player, Tm, Bye, XRk, ADP, PAR, P1, P2, Tgt, Diff]
function blockWidth_(b) { return b.type === 'master' ? 12 : 4; }

function blockStarts_() {
  var starts = [], c = 1;
  BLOCKS.forEach(function (b) { starts.push(c); c += blockWidth_(b) + BLOCK_GAP; });
  starts.push(c); // status block position
  return starts;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Draft Tools')
    .addItem('Mark drafted (other team)', 'markDrafted')
    .addItem('Draft to MY team', 'draftToMe')
    .addItem('Undo last pick', 'undoLastPick')
    .addSeparator()
    .addItem('Refresh data from GitHub', 'refreshData')
    .addItem('Rebuild formulas & formatting', 'buildSheet')
    .addToUi();
}

function buildSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDataTabs_(ss);
  buildParams_(ss);
  buildMaster_(ss);
  buildMasterIdp_(ss);
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

  // Type a player NAME in column A and pick a tag in B; C and D are derived.
  // Matching is by name only, so a name shared across positions (Josh Allen
  // QB / LB) tags both — rare, and visible as colour on two panels.
  var targets = getOrCreate_(ss, 'Targets');
  targets.getRange('A1:D1')
    .setValues([['Player (type name)', 'tag', 'matched?', 'key (auto)']]).setFontWeight('bold');
  targets.getRange('C2').setFormula(
    '=IF($A2="","",IF(COUNTIF(Master!$M:$M,$D2)+COUNTIF(MasterIDP!$H:$H,$D2)>0,"ok","NO MATCH — check spelling"))');
  targets.getRange('D2').setFormula(
    '=IF($A2="","",REGEXREPLACE(REGEXREPLACE(LOWER($A2),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))');
  targets.getRange('C2:D2').copyTo(targets.getRange('C3:D200'));
  targets.setColumnWidth(1, 170);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['high target', 'target', 'mild target', 'downside'], true).build();
  targets.getRange('B2:B200').setDataValidation(rule);

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
  p.getRange('A7:D7').setValues([['Pos', 'repl rank', 'x before my next', 'baseline pts']])
    .setFontWeight('bold');
  p.getRange('A8:C16').setValues([
    ['QB', 22, 2], ['RB', 26, 4], ['WR', 40, 5], ['TE', 12, 2], ['K', 10, 0], ['DST', 10, 0],
    ['LB', 12, 0], ['DB', 12, 0], ['DL', 12, 0],
  ]);
  p.getRange('H7:I7').setValues([['1-gap', 'x-gap']]).setFontWeight('bold');
  for (var i = 0; i < 9; i++) {
    var r = 8 + i;
    // Offence/K/DEF baselines come from Master, the three IDP slots from MasterIDP.
    p.getRange(r, 4).setFormula(r >= 14
      ? '=IFERROR(LARGE(FILTER(MasterIDP!$D$2:$D,MasterIDP!$B$2:$B=$A' + r + '),$B' + r + '),0)'
      : '=IFERROR(LARGE(FILTER(Master!$I$2:$I,Master!$C$2:$C=$A' + r + '),$B' + r + '),0)');
    if (r >= 14) continue;  // gaps only apply to the drafted-early positions
    p.getRange(r, 8).setFormula(
      '=IFERROR(LET(v,SORT(FILTER(Master!$T$2:$T,Master!$C$2:$C=$A' + r +
      ',Master!$Q$2:$Q=FALSE),1,FALSE),ROUND(INDEX(v,1)-INDEX(v,2),0)),"")');
    p.getRange(r, 9).setFormula(
      '=IFERROR(LET(v,SORT(FILTER(Master!$T$2:$T,Master!$C$2:$C=$A' + r +
      ',Master!$Q$2:$Q=FALSE),1,FALSE),ROUND(INDEX(v,1)-INDEX(v,MIN($C' + r + '+1,ROWS(v))),0)),"")');
  }
  p.getRange('A19:A22').setValues([['Current pick'], ['My next pick'], ['My pick after'], ['Max PAR remaining']]);
  p.getRange('B19').setFormula('=COUNTA(Log!$B$2:$B)+1');
  p.getRange('B20').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>=$B$19)),999)');
  p.getRange('B21').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>$B$20)),999)');
  p.getRange('B22').setFormula('=IFERROR(MAX(FILTER(Master!$T$2:$T,Master!$Q$2:$Q=FALSE)),"")');
  p.getRange('E1:F1').setValues([['round', 'my pick']]);
  for (var r = 1; r <= 20; r++) {
    p.getRange(1 + r, 5).setValue(r);
    p.getRange(1 + r, 6).setFormula(
      '=IF(ISODD($E' + (1 + r) + '),($E' + (1 + r) + ')*$B$1-$B$1+$B$2,($E' + (1 + r) + ')*$B$1-$B$2+1)');
  }
  p.getRange('A24').setValue('Snake picks assume you mark EVERY pick in the draft (yours and others’) via Draft Tools.');
}

// ---------- Master ----------

function buildMaster_(ss) {
  var m = getOrCreate_(ss, 'Master');
  m.clear();
  m.getRange('A1:X1').setValues([[
    'fpid', 'Player', 'Pos', 'Tm', 'Bye', 'FfcADP', 'FPpts', 'WWOpts', 'Pts', 'src', 'wdiff', 'delta7',
    'norm', 'XRank', 'YADP', 'BehRank', 'Drafted', 'Mine', 'Tag', 'PAR', 'DiffTop', 'P1', 'P2', 'key',
  ]]);
  var raw = ss.getSheetByName('Raw');
  var n = raw ? Math.max(0, raw.getLastRow() - 1) : 0;
  if (n === 0) { m.setFrozenRows(1); return; }
  // FILTER (not an array reference) so unused rows stay genuinely empty:
  // padded '' cells would make QUERY type these columns as text and the
  // Board's "where C='QB' and Q=false" would match nothing.
  m.getRange('A2').setFormula('=FILTER(Raw!A2:L, Raw!A2:A<>"")');

  var f = {
    M: '=IF($A2="","",REGEXREPLACE(REGEXREPLACE(LOWER($B2),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))',
    N: '=IF($A2="","",IFNA(VLOOKUP($X2,Yahoo!$A:$E,4,FALSE),""))',
    O: '=IF($A2="","",IFNA(VLOOKUP($X2,Yahoo!$A:$E,5,FALSE),""))',
    P: '=IF($A2="","",IF($N2<>"",IF($O2<>"",Params!$B$6*$N2+(1-Params!$B$6)*$O2,$N2),' +
       'IF($O2<>"",$O2,IF($F2<>"",$F2,400))))',
    Q: '=IF($A2="",FALSE,COUNTIF(Log!$B:$B,$A2)>0)',
    R: '=IF($A2="",FALSE,COUNTIFS(Log!$B:$B,$A2,Log!$D:$D,TRUE)>0)',
    S: '=IF($A2="","",IFNA(INDEX(Targets!$B:$B,MATCH($M2,Targets!$D:$D,0)),""))',
    T: '=IF($A2="","",ROUND($I2-IFNA(VLOOKUP($C2,Params!$A$8:$D$13,4,FALSE),0),0))',
    U: '=IF($A2="","",ROUND(MAXIFS($I:$I,$C:$C,$C2,$Q:$Q,FALSE)-$I2,0))',
    V: '=IF($A2="","",IF($Q2,0,LET(sd,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'snow,1-NORMDIST(Params!$B$19,$P2,sd,TRUE),snext,1-NORMDIST(Params!$B$20,$P2,sd,TRUE),' +
       'IF(snow<0.0001,1,ROUND(snext/snow,3)))))',
    W: '=IF($A2="","",IF($Q2,0,LET(sd,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'snow,1-NORMDIST(Params!$B$19,$P2,sd,TRUE),sthen,1-NORMDIST(Params!$B$21,$P2,sd,TRUE),' +
       'IF(snow<0.0001,1,ROUND(sthen/snow,3)))))',
    X: '=IF($A2="","",$M2&"|"&$C2)',
  };
  for (var col in f) {
    m.getRange(col + '2').setFormula(f[col]);
    if (n > 1) m.getRange(col + '2').copyTo(m.getRange(col + '3:' + col + (n + 1)));
  }
  m.setFrozenRows(1);
}

// ---------- MasterIDP (mirror of RawIDP + drafted/tag, keyed by name) ----------

function buildMasterIdp_(ss) {
  var m = getOrCreate_(ss, 'MasterIDP');
  m.clear();
  m.getRange('A1:H1').setValues([['Player', 'Pos', 'Tm', 'Pts', 'Drafted', 'Tag', 'PAR', 'norm']]);
  var rawIdp = ss.getSheetByName('RawIDP');
  var n = rawIdp ? Math.max(0, rawIdp.getLastRow() - 1) : 0;
  if (n === 0) { m.setFrozenRows(1); return; }
  m.getRange('A2').setFormula('=FILTER(RawIDP!A2:D, RawIDP!A2:A<>"")');
  var f = {
    E: '=IF($A2="",FALSE,COUNTIF(Log!$B:$B,$A2)>0)',
    F: '=IF($A2="","",IFNA(INDEX(Targets!$B:$B,MATCH($H2,Targets!$D:$D,0)),""))',
    G: '=IF($A2="","",ROUND($D2-IFNA(VLOOKUP($B2,Params!$A$14:$D$16,4,FALSE),0),0))',
    H: '=IF($A2="","",REGEXREPLACE(REGEXREPLACE(LOWER($A2),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))',
  };
  for (var col in f) {
    m.getRange(col + '2').setFormula(f[col]);
    if (n > 1) m.getRange(col + '2').copyTo(m.getRange(col + '3:' + col + (n + 1)));
  }
  m.setFrozenRows(1);
}

// ---------- Board ----------

function colLetter_(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function buildBoard_(ss) {
  var b = getOrCreate_(ss, 'Board');
  b.clear();
  b.clearConditionalFormatRules();
  b.setFrozenRows(0);

  // Nine panels need ~100 columns; a fresh sheet only has 26, and any
  // setColumnWidth past the end throws "Those columns are out of bounds".
  var starts0 = blockStarts_();
  var needCols = starts0[BLOCKS.length] + 6;
  if (b.getMaxColumns() < needCols) {
    b.insertColumnsAfter(b.getMaxColumns(), needCols - b.getMaxColumns());
  }
  // Hidden state survives clear(); an older layout's hidden column would
  // otherwise stay hidden and swallow a real one (RB lost its names this way).
  b.showColumns(1, b.getMaxColumns());
  var needRows = BOARD_DATA_ROW + Math.max.apply(null, BLOCKS.map(function (x) { return x.rows; })) + 4;
  if (b.getMaxRows() < needRows) {
    b.insertRowsAfter(b.getMaxRows(), needRows - b.getMaxRows());
  }
  var masterHeaders = ['#', 'id', 'Player', 'Tm', 'Bye', 'XRk', 'ADP', 'PAR', 'P1', 'P2', 'Tgt', 'Diff'];
  var idpHeaders = ['Player', 'Tm', 'PAR', 'Tgt'];
  var starts = blockStarts_();
  var rules = [];
  var tagColors = [
    ['high target', '#f4cccc'],   // pale red
    ['target', '#fce5cd'],        // pale orange
    ['mild target', '#fff2cc'],   // pale yellow
    ['downside', '#efefef'],      // pale grey
  ];

  // Summary block (top-left) + status block beside it. Panel columns are
  // narrow, so every label/value spans a merged range to stay readable.
  b.getRange(1, 1, 4, 26).breakApart();
  var bestPos =
    'LET(col,IF({0}=1,Params!$H$8:$H$13,Params!$I$8:$I$13),' +
    'INDEX(Params!$A$8:$A$13,MATCH(MAX(col),col,0)))';
  var summary = [
    ['Best avail (rank)',
     '=IFERROR(INDEX(Master!$B:$B,MATCH(MINIFS(Master!$P:$P,Master!$Q:$Q,FALSE),Master!$P:$P,0)),"")',
     '=IFERROR(ROUND(MINIFS(Master!$P:$P,Master!$Q:$Q,FALSE),0),"")'],
    ['Max PAR',
     '=IFERROR(INDEX(Master!$B:$B,MATCH(Params!$B$22,Master!$T:$T,0)),"")',
     '=Params!$B$22'],
    ['Biggest 1-gap',
     '=IFERROR(LET(pos,' + bestPos.replace('{0}', '1') +
       ',pos&": "&INDEX(Master!$B:$B,MATCH(MAXIFS(Master!$T:$T,Master!$C:$C,pos,Master!$Q:$Q,FALSE),Master!$T:$T,0))),"")',
     '=IFERROR(MAX(Params!$H$8:$H$13),"")'],
    ['Biggest x-gap',
     '=IFERROR(LET(pos,' + bestPos.replace('{0}', '2') +
       ',pos&": "&INDEX(Master!$B:$B,MATCH(MAXIFS(Master!$T:$T,Master!$C:$C,pos,Master!$Q:$Q,FALSE),Master!$T:$T,0))),"")',
     '=IFERROR(MAX(Params!$I$8:$I$13),"")'],
  ];
  summary.forEach(function (row, i) {
    b.getRange(i + 1, 1, 1, 3).merge().setValue(row[0]).setFontWeight('bold');
    b.getRange(i + 1, 4, 1, 5).merge().setFormula(row[1]);
    b.getRange(i + 1, 9).setFormula(row[2]);
  });
  var status = [
    ['Current pick', '=Params!$B$19'],
    ['My next pick', '=Params!$B$20'],
    ['Then', '=Params!$B$21'],
    ['My picks so far', '=COUNTIF(Log!$D:$D,TRUE)'],
  ];
  status.forEach(function (row, i) {
    b.getRange(i + 1, 15, 1, 3).merge().setValue(row[0]).setFontWeight('bold');
    b.getRange(i + 1, 18, 1, 2).merge().setFormula(row[1]);
  });

  BLOCKS.forEach(function (blk, i) {
    var bs = starts[i];
    var width = blockWidth_(blk);
    b.setColumnWidth(bs + width, 12);  // spacer between panels
    b.getRange(TITLE_ROW, bs).setValue(blk.title).setFontWeight('bold').setFontSize(12);
    var lastDataRow = BOARD_DATA_ROW + blk.rows - 1;

    if (blk.type === 'master') {
      b.getRange(HEADER_ROW, bs, 1, width).setValues([masterHeaders]).setFontWeight('bold');
      // QUERY spills cols bs+1..bs+10: fpid,Player,Tm,Bye,XRk,ADP,PAR,P1,P2,Tgt
      b.getRange(BOARD_DATA_ROW, bs + 1).setFormula(
        '=IFERROR(QUERY(Master!$A$2:$X,"select A,B,D,E,N,O,T,V,W,S where C=\'' + blk.pos +
        '\' and Q=false order by I desc limit ' + blk.rows + '",0),"")');

      var playerL = colLetter_(bs + 2), ptsL = colLetter_(bs + 7);  // PAR column
      b.getRange(BOARD_DATA_ROW, bs).setFormula(
        '=IF($' + playerL + BOARD_DATA_ROW + '="","",ROW()-' + (BOARD_DATA_ROW - 1) + ')');
      // Diff shown at row 2 (1-gap) and row x+1 (the x-gap), like the old sheet.
      b.getRange(BOARD_DATA_ROW, bs + 11).setFormula(
        '=IF($' + ptsL + BOARD_DATA_ROW + '="","",IF(OR(ROW()=' + (BOARD_DATA_ROW + 1) +
        ',ROW()=' + BOARD_DATA_ROW + '+Params!$C$' + PARAMS_POS_ROW[blk.pos] + '),' +
        'ROUND($' + ptsL + '$' + BOARD_DATA_ROW + '-$' + ptsL + BOARD_DATA_ROW + ',0),""))');
      [bs, bs + 11].forEach(function (c) {
        b.getRange(BOARD_DATA_ROW, c, 1, 1).autoFill(
          b.getRange(BOARD_DATA_ROW, c, blk.rows, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      });

      // [#, fpid, Player, Tm, Bye, XRk, ADP, PAR, P1, P2, Tgt, Diff]
      [26, 0, 132, 28, 26, 38, 38, 38, 38, 38, 22, 36].forEach(function (w, k) {
        if (w) b.setColumnWidth(bs + k, w);
      });
      b.hideColumns(bs + 1);
      b.hideColumns(bs + 10);  // Tgt: the row colour conveys it
      b.getRange(BOARD_DATA_ROW, bs + 3, blk.rows, 2).setFontSize(8);      // Tm, Bye
      b.getRange(BOARD_DATA_ROW, bs + 7, blk.rows, 1).setNumberFormat('0');  // PAR
      b.getRange(BOARD_DATA_ROW, bs + 11, blk.rows, 1).setNumberFormat('0'); // Diff
      b.getRange(BOARD_DATA_ROW, bs + 8, blk.rows, 2).setNumberFormat('0%');

      // Color scales, three-point like the old sheet: rank columns are
      // "lowest is best", so colour saturates at the min and fades to white
      // by CF_RANK_FADE percentile; PAR is "highest is best" and only the top
      // end (above CF_PAR_MID percent of the range) picks up purple.
      var IT = SpreadsheetApp.InterpolationType;
      [[bs + 5, '#6d9eeb', '#c9daf8'], [bs + 6, '#93c47d', '#d9ead3']].forEach(function (c) {
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .setRanges([b.getRange(BOARD_DATA_ROW, c[0], blk.rows, 1)])
          .setGradientMinpointWithValue(c[1], IT.MIN, '')
          .setGradientMidpointWithValue(c[2], IT.PERCENTILE, String(CF_RANK_MID))
          .setGradientMaxpointWithValue('#ffffff', IT.PERCENTILE, String(CF_RANK_FADE))
          .build());
      });
      // Anchor white AT the cutoff, not at the minimum: with white at MIN the
      // ramp runs across the whole column and tints every cell.
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([b.getRange(BOARD_DATA_ROW, bs + 7, blk.rows, 1)])
        .setGradientMinpointWithValue('#ffffff', IT.PERCENTILE, String(CF_PAR_MID))
        .setGradientMaxpointWithValue('#b4a7d6', IT.MAX, '')
        .build());
      // Survival odds: pale red = likely gone (act now), fading to nothing
      // when he is safe. No green — a high number needs no attention.
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([b.getRange(BOARD_DATA_ROW, bs + 8, blk.rows, 2)])
        .setGradientMinpointWithValue('#f4cccc', IT.NUMBER, '0')
        .setGradientMidpointWithValue('#fff2cc', IT.NUMBER, '0.5')
        .setGradientMaxpointWithValue('#ffffff', IT.NUMBER, '1')
        .build());
    } else {
      b.getRange(HEADER_ROW, bs, 1, width).setValues([idpHeaders]).setFontWeight('bold');
      b.getRange(BOARD_DATA_ROW, bs).setFormula(
        '=IFERROR(QUERY(MasterIDP!$A$2:$H,"select A,C,G,F where B=\'' + blk.pos +
        '\' and E=false order by D desc limit ' + blk.rows + '",0),"")');
      [132, 28, 38, 22].forEach(function (w, k) { b.setColumnWidth(bs + k, w); });
      b.hideColumns(bs + 3);  // Tgt
      b.getRange(BOARD_DATA_ROW, bs + 1, blk.rows, 1).setFontSize(8);       // Tm
      b.getRange(BOARD_DATA_ROW, bs + 2, blk.rows, 1).setNumberFormat('0'); // PAR
    }

    var dataRange = b.getRange(BOARD_DATA_ROW, bs, blk.rows, width);
    var tgtCol = blk.type === 'master' ? bs + 10 : bs + 3;
    var tgt = colLetter_(tgtCol);
    tagColors.forEach(function (t) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([dataRange])
        .whenFormulaSatisfied('=$' + tgt + BOARD_DATA_ROW + '="' + t[0] + '"')
        .setBackground(t[1]).build());
    });
  });

  b.getRange(5, 1, 1, 8).merge().setFormula(
    '=IF(COUNTA(Raw!$A$2:$A)=0,"NO DATA LOADED — run Draft Tools > Refresh data from GitHub, then Rebuild formulas & formatting","")')
    .setFontColor('#cc0000').setFontWeight('bold');

  b.setConditionalFormatRules(rules);
  b.setFrozenRows(HEADER_ROW);
}

function refreshData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var tab in DATA_URLS) {
    var text = UrlFetchApp.fetch(DATA_URLS[tab]).getContentText();
    var values = Utilities.parseCsv(text);
    var sh = ss.getSheetByName(tab);
    sh.clearContents();
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
  buildMaster_(ss);
  buildMasterIdp_(ss);
  buildBoard_(ss);
  ss.toast('Raw + RawIDP refreshed from GitHub; formulas resized to the new data.');
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
    var starts = blockStarts_();
    for (var i = 0; i < BLOCKS.length; i++) {
      if (cell.getColumn() >= starts[i] && cell.getColumn() < starts[i] + blockWidth_(BLOCKS[i])) {
        if (BLOCKS[i].type === 'idp') {
          // IDP blocks have no fpid; their first column (the name) is the id.
          fpid = name = sheet.getRange(cell.getRow(), starts[i]).getValue();
        } else {
          // master blocks: [#, fpid, Player, ...]
          fpid = sheet.getRange(cell.getRow(), starts[i] + 1).getValue();
          name = sheet.getRange(cell.getRow(), starts[i] + 2).getValue();
        }
      }
    }
  } else if (sheet.getName() === 'Master') {
    fpid = sheet.getRange(cell.getRow(), 1).getValue();
    name = sheet.getRange(cell.getRow(), 2).getValue();
  } else if (sheet.getName() === 'MasterIDP') {
    fpid = name = sheet.getRange(cell.getRow(), 1).getValue();
  } else {
    ss.toast('Draft from the Board, Master, or MasterIDP tab.');
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
