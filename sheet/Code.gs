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
  { pos: 'QB',  title: 'QB',  type: 'master', rows: 30 },
  { pos: 'RB',  title: 'RB',  type: 'master', rows: 70 },
  { pos: 'WR',  title: 'WR',  type: 'master', rows: 90 },
  { pos: 'TE',  title: 'TE',  type: 'master', rows: 30 },
  { pos: 'LB',  title: 'LB',  type: 'idp',    rows: 30 },
  { pos: 'DB',  title: 'DB',  type: 'idp',    rows: 30 },
  { pos: 'DL',  title: 'DL',  type: 'idp',    rows: 30 },
  { pos: 'DST', title: 'DEF', type: 'master', rows: 20 },
  { pos: 'K',   title: 'K',   type: 'master', rows: 20 },
];
var BLOCK_GAP = 1;
var TITLE_ROW = 9;        // summary block lives in rows 1-7, banner on 8
var HEADER_ROW = 10;
var BOARD_DATA_ROW = 11;
var EBA_ROW = 30;         // Params: expected-best-available table header
// BoardBackup is a frozen snapshot of Board — values only, no formulas — so it
// survives damage to Master, Log or any formula the live board depends on.
var BACKUP_NAME = 'BoardBackup';
var PARAMS_POS_ROW = { QB: 8, RB: 9, WR: 10, TE: 11, K: 12, DST: 13, LB: 14, DB: 15, DL: 16 };

// Conditional-format tuning. Rank columns (XRk/ADP) shade roughly the top
// CF_RANK_TOP_N players still on the board, board-wide rather than per panel,
// so the colours are comparable across positions. PAR likewise shades the top
// CF_PAR_TOP_N players anywhere on the board.
var CF_RANK_TOP_N = 12;   // players shaded on XRk / ADP, board-wide
var CF_PAR_TOP_N = 25;    // players shaded on PAR, board-wide

// Raw data published by the pipeline (see refresh.sh) — pulled by refreshData().
var DATA_URLS = {
  Raw: 'https://raw.githubusercontent.com/epgreig/ff-manager/master/out/blended.csv',
  RawIDP: 'https://raw.githubusercontent.com/epgreig/ff-manager/master/out/idp.csv',
};

// master block: [#, fpid*, Player, Tm, Bye, XRk, ADP, PAR, PS, PL, P2*, Tgt*,
//                 Risk, Diff, cS*, eS*, cL*, eL*, c2*, e2*]  (* = hidden)
// The hidden tail drives the expected-best-available model over each snake gap:
// cum is P(every better player is gone), eba the player's share of the value
// you can still expect to get. Risk is P(gone) x (PAR - E[best]) — the points
// you expect to lose by passing, netting off the fallback you would take
// instead, which is where PAR and the survival odds become one number.
function blockWidth_(b) { return b.type === 'master' ? 20 : 4; }

function blockStarts_() {
  var starts = [], c = 1;
  BLOCKS.forEach(function (b) { starts.push(c); c += blockWidth_(b) + BLOCK_GAP; });
  starts.push(c); // status block position
  return starts;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Draft Tools')
    .addItem('Mark drafted', 'markDrafted')
    .addItem('Undo last pick', 'undoLastPick')
    .addSeparator()
    .addItem('Refresh data from GitHub', 'refreshData')
    .addItem('Re-snapshot backup board', 'snapshotBackup')
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
  snapshotBackup_(ss);
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

  // Paste the Yahoo player list straight into YahooRaw (headers on row 2, data
  // from row 3). Column A there packs three things into one cell —
  // "Jahmyr Gibbs\nDet - RB" plus an optional injury tag — so the Yahoo tab
  // unpacks it: name, position, expert Rank (col B) and ADP (col G, else F).
  var yraw = getOrCreate_(ss, 'YahooRaw');
  if (yraw.getLastRow() === 0) {
    yraw.getRange('A1').setValue('Paste the Yahoo player table here (its own header rows included)');
  }
  var cell = 'INDEX(YahooRaw!$A:$A,ROW()+1)';
  var yahoo = getOrCreate_(ss, 'Yahoo');
  yahoo.getRange('A1:F1').setValues([
    ['key (auto)', 'Name (auto)', 'Pos (auto)', 'XRank (auto)', 'ADP (auto)', 'matched? (auto)'],
  ]).setFontWeight('bold');
  yahoo.getRange('B2').setFormula('=IFERROR(REGEXEXTRACT(' + cell + ',"^[^\\n]+"),"")');
  yahoo.getRange('C2').setFormula(
    '=IFERROR(LET(p,UPPER(REGEXEXTRACT(' + cell + ',"\\n[^\\n]*-\\s*([A-Za-z]+)")),' +
    'IF(p="DEF","DST",p)),"")');
  yahoo.getRange('D2').setFormula('=IFERROR(N(INDEX(YahooRaw!$B:$B,ROW()+1)),"")');
  yahoo.getRange('E2').setFormula(
    '=IFERROR(LET(all,N(INDEX(YahooRaw!$G:$G,ROW()+1)),pre,N(INDEX(YahooRaw!$F:$F,ROW()+1)),' +
    'IF(all>0,all,IF(pre>0,pre,""))),"")');
  // Column G is the manual escape hatch: type the FantasyPros spelling there
  // and the key is rebuilt from it. Defences key on nickname alone.
  yahoo.getRange('A2').setFormula(
    '=IF($B2="","",LET(nm,IF($G2<>"",$G2,$B2),IF($C2="DST",' +
    'LOWER(REGEXEXTRACT(TRIM(nm),"(\\S+)$"))&"|DST",' +
    'REGEXREPLACE(REGEXREPLACE(LOWER(nm),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$","")&"|"&$C2)))');
  yahoo.getRange('F2').setFormula(
    '=IF($B2="","",IF(COUNTIF(Master!$X$2:$X,$A2)>0,"ok",' +
    '"NO MATCH — put the FantasyPros spelling in column G"))');
  yahoo.getRange('G1').setValue('fix: FantasyPros name').setFontWeight('bold');
  yahoo.getRange('A2:F2').copyTo(yahoo.getRange('A3:F400'));
  yahoo.setColumnWidth(7, 170);
  yahoo.setColumnWidth(1, 160);
  yahoo.setColumnWidth(2, 150);
  yahoo.setColumnWidth(6, 170);

  // One column per tag: type names under the heading you mean. Columns F:I
  // hold normalised copies for lookup, K lists anything that failed to match.
  var targets = getOrCreate_(ss, 'Targets');
  targets.getRange('A1:D1')
    .setValues([['high target', 'medium target', 'mild target', 'averse']])
    .setFontWeight('bold');
  targets.getRange('F1:I1').setValues([['h (auto)', 'm (auto)', 'l (auto)', 'a (auto)']])
    .setBackground('#efefef');
  ['A', 'B', 'C', 'D'].forEach(function (col, k) {
    targets.getRange(2, 6 + k).setFormula(
      '=ARRAYFORMULA(IF($' + col + '$2:$' + col + '$200="","",' +
      'REGEXREPLACE(REGEXREPLACE(LOWER(' + '$' + col + '$2:$' + col + '$200' + '),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$","")))');
    targets.setColumnWidth(1 + k, 150);
  });
  targets.getRange('K1').setValue('unmatched — check spelling').setFontWeight('bold');
  targets.getRange('K2').setFormula(
    '=IFERROR(LET(nm,TOCOL($A$2:$D$200,1),bad,FILTER(nm,MAP(nm,LAMBDA(x,' +
    'COUNTIF(Master!$M:$M,REGEXREPLACE(REGEXREPLACE(LOWER(x),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))+' +
    'COUNTIF(MasterIDP!$H:$H,REGEXREPLACE(REGEXREPLACE(LOWER(x),"[.\'’-]","")," (jr|sr|ii|iii|iv|v)$",""))=0))),' +
    'TEXTJOIN(", ",TRUE,bad)),"")');
  targets.setColumnWidth(11, 400);

  var log = getOrCreate_(ss, 'Log');
  log.getRange('A1:C1').setValues([['pick#', 'fpid', 'Player']]);
}

// ---------- Params ----------

function buildParams_(ss) {
  var p = getOrCreate_(ss, 'Params');
  p.clear();
  p.getRange('A1:B6').setValues([
    ['Teams', 12],
    ['My draft slot', 5],
    ['Rounds', 20],
    ['Sigma min (picks)', 4],
    ['Sigma k (× rank)', 0.18],
    ['XRank weight (vs Yahoo ADP)', 0.7],
  ]);
  p.getRange('A7:D7').setValues([['Pos', 'repl rank', 'x before my next', 'baseline pts']])
    .setFontWeight('bold');
  // Replacement ranks = what is actually free on the waiver wire after the
  // draft, NOT the number of starters. Ethan's numbers, from the best players
  // who went undrafted last year, smoothed, adjusted for the third WR slot.
  // DEF sits at 6 because the league drafts twelve defences but not the same
  // twelve this board ranks highest, so a top-6 board DEF survives.
  p.getRange('A8:C16').setValues([
    ['QB', 20, 2], ['RB', 56, 4], ['WR', 64, 6], ['TE', 16, 2], ['K', 8, 0], ['DST', 6, 0],
    ['LB', 8, 0], ['DB', 8, 0], ['DL', 8, 0],
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
  p.getRange('A19:A26').setValues([['Current pick'], ['My next pick'], ['My pick after'],
    ['My third pick'], ['Max PAR remaining'],
    ['Short gap (picks)'], ['Long gap (picks)'], ['Both gaps']]);
  // In a snake, the distance between your picks alternates between two values
  // that always sum to 2 x teams: 2*slot-1 turning back on you, and the rest.
  p.getRange('B24').setFormula('=MIN(2*$B$2-1,2*$B$1-2*$B$2+1)');
  p.getRange('B25').setFormula('=MAX(2*$B$2-1,2*$B$1-2*$B$2+1)');
  p.getRange('B26').setFormula('=2*$B$1');
  p.getRange('B19').setFormula('=COUNTA(Log!$B$2:$B)+1');
  p.getRange('B20').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>=$B$19)),999)');
  p.getRange('B21').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>$B$20)),999)');
  p.getRange('B22').setFormula('=IFERROR(MIN(FILTER($F$2:$F$21,$F$2:$F$21>$B$21)),999)');
  p.getRange('B23').setFormula('=IFERROR(MAX(FILTER(Master!$T$2:$T,Master!$Q$2:$Q=FALSE)),"")');
  // Tag nudges: a deliberate thumb on the scale, not a re-projection.
  p.getRange('D18:E18').setValues([['tag', 'pts adj']]).setFontWeight('bold');
  p.getRange('D19:E22').setValues([
    ['high target', 0.04], ['medium target', 0.04], ['mild target', 0.02], ['averse', -0.02],
  ]);
  p.getRange('E19:E22').setNumberFormat('0%');
  p.getRange('E1:F1').setValues([['round', 'my pick']]);
  for (var r = 1; r <= 20; r++) {
    p.getRange(1 + r, 5).setValue(r);
    p.getRange(1 + r, 6).setFormula(
      '=IF(ISODD($E' + (1 + r) + '),($E' + (1 + r) + ')*$B$1-$B$1+$B$2,($E' + (1 + r) + ')*$B$1-$B$2+1)');
  }
  // Expected best available, read off the Board's hidden EBA columns.
  // "wait cost" = what the top man is worth now minus what you expect to be
  // able to take at your next pick — i.e. the price of passing on the position.
  p.getRange(EBA_ROW, 1, 1, 8).setValues([
    ['Pos', 'top PAR', 'E[best] short', 'wait S', 'E[best] long', 'wait L', 'E[best] 2gaps', 'wait 2'],
  ]).setFontWeight('bold');
  var starts = blockStarts_();
  var er = EBA_ROW + 1;
  BLOCKS.forEach(function (blk, i) {
    if (blk.type !== 'master' || ['K', 'DST'].indexOf(blk.pos) >= 0) return;
    var bs = starts[i];
    var first = BOARD_DATA_ROW, last = BOARD_DATA_ROW + blk.rows - 1;
    var parL = colLetter_(bs + 7);
    p.getRange(er, 1).setValue(blk.pos);
    p.getRange(er, 2).setFormula('=IFERROR(Board!$' + parL + '$' + first + ',"")');
    [15, 17, 19].forEach(function (off, k) {   // eS / eL / e2 contribution columns
      var eL = colLetter_(bs + off);
      p.getRange(er, 3 + 2 * k).setFormula(
        '=IFERROR(ROUND(SUM(Board!$' + eL + '$' + first + ':$' + eL + '$' + last + '),1),"")');
      p.getRange(er, 4 + 2 * k).setFormula(
        '=IFERROR(ROUND($B' + er + '-' + colLetter_(3 + 2 * k) + er + ',1),"")');
    });
    er++;
  });

  p.getRange('A28').setValue('Snake picks assume you mark EVERY pick in the draft (yours and others’) via Draft Tools.');
}

// ---------- Master ----------

function buildMaster_(ss) {
  var m = getOrCreate_(ss, 'Master');
  m.clear();
  m.getRange('A1:Z1').setValues([[
    'fpid', 'Player', 'Pos', 'Tm', 'Bye', 'FfcADP', 'FPpts', 'WWOpts', 'Pts', 'src', 'wdiff', 'delta7',
    'norm', 'XRank', 'YADP', 'BehRank', 'Drafted', '(unused)', 'Tag', 'PAR', 'DiffTop', 'PS', 'PL', 'key',
    'P2', 'AdjPts',
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
    S: '=IF($A2="","",IFS(COUNTIF(Targets!$F:$F,$M2)>0,"high target",' +
       'COUNTIF(Targets!$G:$G,$M2)>0,"medium target",' +
       'COUNTIF(Targets!$H:$H,$M2)>0,"mild target",' +
       'COUNTIF(Targets!$I:$I,$M2)>0,"averse",TRUE,""))',
    // PAR runs off the tag-adjusted points, but the replacement baseline stays
    // on the raw projections — a preference should move a player, not the bar.
    T: '=IF($A2="","",ROUND($Z2-IFNA(VLOOKUP($C2,Params!$A$8:$D$13,4,FALSE),0),0))',
    Z: '=IF($A2="","",ROUND($I2*(1+IFNA(VLOOKUP($S2,Params!$D$19:$E$22,2,FALSE),0)),1))',
    U: '=IF($A2="","",ROUND(MAXIFS($I:$I,$C:$C,$C2,$Q:$Q,FALSE)-$I2,0))',
    // Survival across the snake gaps, conditional on being here now.
    // LET names must not look like cell refs (s3 -> #NAME), hence sdev/base/targ.
    V: '=IF($A2="","",IF($Q2,0,LET(sdev,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'base,1-NORMDIST(Params!$B$19,$P2,sdev,TRUE),' +
       'targ,1-NORMDIST(Params!$B$19+Params!$B$24,$P2,sdev,TRUE),' +
       'IF(base<0.0001,1,ROUND(targ/base,3)))))',
    W: '=IF($A2="","",IF($Q2,0,LET(sdev,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'base,1-NORMDIST(Params!$B$19,$P2,sdev,TRUE),' +
       'targ,1-NORMDIST(Params!$B$19+Params!$B$25,$P2,sdev,TRUE),' +
       'IF(base<0.0001,1,ROUND(targ/base,3)))))',
    // Defences are keyed on the nickname only: Yahoo writes "Patriots" where
    // FantasyPros writes "New England Patriots".
    X: '=IF($A2="","",IF($C2="DST",LOWER(REGEXEXTRACT(TRIM($B2),"(\\S+)$"))&"|DST",' +
       '$M2&"|"&$C2))',
    Y: '=IF($A2="","",IF($Q2,0,LET(sdev,MAX(Params!$B$4,Params!$B$5*$P2),' +
       'base,1-NORMDIST(Params!$B$19,$P2,sdev,TRUE),' +
       'targ,1-NORMDIST(Params!$B$19+Params!$B$26,$P2,sdev,TRUE),' +
       'IF(base<0.0001,1,ROUND(targ/base,3)))))',
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
    F: '=IF($A2="","",IFS(COUNTIF(Targets!$F:$F,$H2)>0,"high target",' +
       'COUNTIF(Targets!$G:$G,$H2)>0,"medium target",' +
       'COUNTIF(Targets!$H:$H,$H2)>0,"mild target",' +
       'COUNTIF(Targets!$I:$I,$H2)>0,"averse",TRUE,""))',
    G: '=IF($A2="","",ROUND($D2*(1+IFNA(VLOOKUP($F2,Params!$D$19:$E$22,2,FALSE),0))' +
       '-IFNA(VLOOKUP($B2,Params!$A$14:$D$16,4,FALSE),0),0))',
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
  var masterHeaders = ['#', 'id', 'Player', 'Tm', 'Bye', 'XRk', 'ADP', 'PAR', 'PS', 'PL',
                       'P2', 'Tgt', 'Risk', 'Diff', 'cS', 'eS', 'cL', 'eL', 'c2', 'e2'];
  var idpHeaders = ['Player', 'Tm', 'PAR', 'Tgt'];
  var starts = blockStarts_();
  var rules = [];
  var xrkRanges = [], adpRanges = [], parRanges = [], probRanges = [], urgRanges = [];
  var masterRows = 0;
  var IT = SpreadsheetApp.InterpolationType;
  var tagColors = [
    ['high target', '#f4cccc'],   // pale red
    ['medium target', '#fce5cd'], // pale orange
    ['mild target', '#fff2cc'],   // pale yellow
    ['averse', '#efefef'],        // pale grey
  ];

  // Summary block (top-left) + status block beside it. Panel columns are
  // narrow, so every label/value spans a merged range to stay readable.
  b.getRange(1, 1, 8, Math.min(60, b.getMaxColumns())).breakApart();
  var bestPos =
    'LET(col,IF({0}=1,Params!$H$8:$H$13,Params!$I$8:$I$13),' +
    'INDEX(Params!$A$8:$A$13,MATCH(MAX(col),col,0)))';
  // Costliest wait: which position's top man is worth most over what you
  // expect to still be able to take at that pick. Column D names the player.
  var waitRow = function (n) {
    var col = colLetter_(4 + 2 * (n - 1));  // D / F / H on the Params EBA table
    var lo = EBA_ROW + 1, hi = EBA_ROW + 4;
    var rng = 'Params!$' + col + '$' + lo + ':$' + col + '$' + hi;
    var pos = 'INDEX(Params!$A$' + lo + ':$A$' + hi + ',MATCH(MAX(' + rng + '),' + rng + ',0))';
    return ['Costliest ' + ['S', 'L', '2'][n - 1] + '-wait',
      '=IFERROR(LET(p,' + pos + ',p&": "&INDEX(Master!$B:$B,' +
        'MATCH(MAXIFS(Master!$T:$T,Master!$C:$C,p,Master!$Q:$Q,FALSE),Master!$T:$T,0))),"")',
      '=IFERROR(ROUND(MAX(' + rng + '),0),"")'];
  };
  var summary = [
    ['Highest X-rank',
     '=IFERROR(INDEX(Master!$B:$B,MATCH(MINIFS(Master!$N:$N,Master!$Q:$Q,FALSE,Master!$N:$N,">0"),Master!$N:$N,0)),"")',
     '=IFERROR(MINIFS(Master!$N:$N,Master!$Q:$Q,FALSE,Master!$N:$N,">0"),"")'],
    ['Highest ADP',
     '=IFERROR(INDEX(Master!$B:$B,MATCH(MINIFS(Master!$O:$O,Master!$Q:$Q,FALSE,Master!$O:$O,">0"),Master!$O:$O,0)),"")',
     '=IFERROR(MINIFS(Master!$O:$O,Master!$Q:$Q,FALSE,Master!$O:$O,">0"),"")'],
    ['Biggest 1-gap',
     '=IFERROR(LET(pos,' + bestPos.replace('{0}', '1') +
       ',pos&": "&INDEX(Master!$B:$B,MATCH(MAXIFS(Master!$T:$T,Master!$C:$C,pos,Master!$Q:$Q,FALSE),Master!$T:$T,0))),"")',
     '=IFERROR(MAX(Params!$H$8:$H$13),"")'],
    ['Biggest x-gap',
     '=IFERROR(LET(pos,' + bestPos.replace('{0}', '2') +
       ',pos&": "&INDEX(Master!$B:$B,MATCH(MAXIFS(Master!$T:$T,Master!$C:$C,pos,Master!$Q:$Q,FALSE),Master!$T:$T,0))),"")',
     '=IFERROR(MAX(Params!$I$8:$I$13),"")'],
    waitRow(1),
    waitRow(2),
    waitRow(3),
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
    ['Picks until mine', '=IFERROR(MAX(0,Params!$B$20-Params!$B$19),"")'],
  ];
  var sc = starts0[1];  // second panel: label spans #/id/Player, value Tm..ADP
  status.forEach(function (row, i) {
    b.getRange(i + 1, sc, 1, 3).merge().setValue(row[0]).setFontWeight('bold');
    b.getRange(i + 1, sc + 3, 1, 4).merge().setFormula(row[1]);
  });

  BLOCKS.forEach(function (blk, i) {
    var bs = starts[i];
    var width = blockWidth_(blk);
    b.setColumnWidth(bs + width, 12);  // spacer between panels
    b.getRange(TITLE_ROW, bs).setValue(blk.title).setFontWeight('bold').setFontSize(12);
    var lastDataRow = BOARD_DATA_ROW + blk.rows - 1;

    if (blk.type === 'master') {
      b.getRange(HEADER_ROW, bs, 1, width).setValues([masterHeaders]).setFontWeight('bold');
      // QUERY spills cols bs+1..bs+11: fpid,Player,Tm,Bye,XRk,ADP,PAR,P1,P2,P3,Tgt
      b.getRange(BOARD_DATA_ROW, bs + 1).setFormula(
        '=IFERROR(QUERY(Master!$A$2:$Z,"select A,B,D,E,N,O,T,V,W,Y,S where C=\'' + blk.pos +
        '\' and Q=false order by Z desc limit ' + blk.rows + '",0),"")');

      var playerL = colLetter_(bs + 2), ptsL = colLetter_(bs + 7);  // PAR column
      b.getRange(BOARD_DATA_ROW, bs).setFormula(
        '=IF($' + playerL + BOARD_DATA_ROW + '="","",ROW()-' + (BOARD_DATA_ROW - 1) + ')');
      // Risk = P(gone) x (his PAR - what you expect to get at the position
      // anyway). Losing a player does not cost his whole PAR, only his edge
      // over the fallback; and that edge is only forfeited in the worlds where
      // he is actually gone. Negative means waiting looks better than taking
      // him. E[best] comes from the Params table, which only covers QB/RB/WR/TE.
      var parC = colLetter_(bs + 7), plC = colLetter_(bs + 9);
      var ebaRow = { QB: 0, RB: 1, WR: 2, TE: 3 }[blk.pos];
      if (ebaRow !== undefined) {
        b.getRange(BOARD_DATA_ROW, bs + 12).setFormula(
          '=IF(OR($' + parC + BOARD_DATA_ROW + '="",$' + plC + BOARD_DATA_ROW + '=""),"",' +
          'IFERROR(ROUND((1-$' + plC + BOARD_DATA_ROW + ')*($' + parC + BOARD_DATA_ROW +
          '-Params!$E$' + (EBA_ROW + 1 + ebaRow) + '),0),""))');
      }
      // Diff shown at row 2 (1-gap) and row x+1 (the x-gap), like the old sheet.
      b.getRange(BOARD_DATA_ROW, bs + 13).setFormula(
        '=IF($' + ptsL + BOARD_DATA_ROW + '="","",IF(OR(ROW()=' + (BOARD_DATA_ROW + 1) +
        ',ROW()=' + BOARD_DATA_ROW + '+Params!$C$' + PARAMS_POS_ROW[blk.pos] + '),' +
        'ROUND($' + ptsL + '$' + BOARD_DATA_ROW + '-$' + ptsL + BOARD_DATA_ROW + ',0),""))');
      // Expected best available: walking down a panel already sorted by PAR,
      // cum = P(every better player is gone), so PAR * P(available) * cum is
      // this player's share of what you expect to still be there when you pick.
      var parL = colLetter_(bs + 7);
      var probL = [colLetter_(bs + 8), colLetter_(bs + 9), colLetter_(bs + 10)];
      var cumL = [colLetter_(bs + 14), colLetter_(bs + 16), colLetter_(bs + 18)];
      var prev = BOARD_DATA_ROW - 1, first = BOARD_DATA_ROW;
      [[bs + 14, probL[0], cumL[0]], [bs + 16, probL[1], cumL[1]],
       [bs + 18, probL[2], cumL[2]]].forEach(function (spec) {
        b.getRange(first, spec[0]).setFormula(
          '=IF($' + playerL + first + '="","",IF(ROW()=' + first + ',1,$' + spec[2] + prev +
          '*(1-$' + spec[1] + prev + ')))');
        b.getRange(first, spec[0] + 1).setFormula(
          '=IF($' + playerL + first + '="","",$' + parL + first + '*$' + spec[1] + first +
          '*$' + colLetter_(spec[0]) + first + ')');
      });
      [bs, bs + 12, bs + 13, bs + 14, bs + 15, bs + 16, bs + 17, bs + 18,
       bs + 19].forEach(function (c) {
        b.getRange(BOARD_DATA_ROW, c, 1, 1).autoFill(
          b.getRange(BOARD_DATA_ROW, c, blk.rows, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      });

      // [#, fpid, Player, Tm, Bye, XRk, ADP, PAR, PS, PL, P2, Tgt, Urg, Diff]
      [24, 0, 114, 26, 22, 30, 30, 32, 26, 26, 0, 0, 30, 30].forEach(function (w, k) {
        if (w) b.setColumnWidth(bs + k, w);
      });
      b.hideColumns(bs + 1);
      b.hideColumns(bs + 10, 2);  // P2 feeds the model; Tgt drives name colour
      b.hideColumns(bs + 14, 6);  // EBA working columns
      b.getRange(BOARD_DATA_ROW, bs, blk.rows, 1).setNumberFormat('0');       // #
      b.getRange(BOARD_DATA_ROW, bs + 5, blk.rows, 2).setNumberFormat('0');   // XRk, ADP
      b.getRange(BOARD_DATA_ROW, bs + 3, blk.rows, 2).setFontSize(8);      // Tm, Bye
      b.getRange(BOARD_DATA_ROW, bs + 8, blk.rows, 2).setFontSize(8);      // PS, PL
      b.getRange(BOARD_DATA_ROW, bs + 12, blk.rows, 1).setNumberFormat('0');  // Risk
      b.getRange(BOARD_DATA_ROW, bs + 7, blk.rows, 1).setNumberFormat('0');  // PAR
      b.getRange(BOARD_DATA_ROW, bs + 13, blk.rows, 1).setNumberFormat('0'); // Diff
      b.getRange(BOARD_DATA_ROW, bs + 8, blk.rows, 2).setNumberFormat('0%');

      // Gather ranges; colour rules are applied once across every panel below,
      // so a rank or PAR colour means the same thing in each of them.
      xrkRanges.push(b.getRange(BOARD_DATA_ROW, bs + 5, blk.rows, 1));
      adpRanges.push(b.getRange(BOARD_DATA_ROW, bs + 6, blk.rows, 1));
      parRanges.push(b.getRange(BOARD_DATA_ROW, bs + 7, blk.rows, 1));
      probRanges.push(b.getRange(BOARD_DATA_ROW, bs + 8, blk.rows, 2));
      urgRanges.push(b.getRange(BOARD_DATA_ROW, bs + 12, blk.rows, 1));
      masterRows += blk.rows;
    } else {
      b.getRange(HEADER_ROW, bs, 1, width).setValues([idpHeaders]).setFontWeight('bold');
      b.getRange(BOARD_DATA_ROW, bs).setFormula(
        '=IFERROR(QUERY(MasterIDP!$A$2:$H,"select A,C,G,F where B=\'' + blk.pos +
        '\' and E=false order by G desc limit ' + blk.rows + '",0),"")');
      [132, 28, 38, 22].forEach(function (w, k) { b.setColumnWidth(bs + k, w); });
      b.hideColumns(bs + 3);  // Tgt
      b.getRange(BOARD_DATA_ROW, bs + 1, blk.rows, 1).setFontSize(8);       // Tm
      b.getRange(BOARD_DATA_ROW, bs + 2, blk.rows, 1).setNumberFormat('0'); // PAR
    }

    // Tag colour marks the name cell only, not the whole row.
    var nameRange = b.getRange(BOARD_DATA_ROW, bs + (blk.type === 'master' ? 2 : 0), blk.rows, 1);
    var tgtCol = blk.type === 'master' ? bs + 11 : bs + 3;
    var tgt = colLetter_(tgtCol);
    tagColors.forEach(function (t) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([nameRange])
        .whenFormulaSatisfied('=$' + tgt + BOARD_DATA_ROW + '="' + t[0] + '"')
        .setBackground(t[1]).build());
    });
  });

  b.getRange(8, 1, 1, 8).merge().setFormula(
    '=IF(COUNTA(Raw!$A$2:$A)=0,"NO DATA LOADED — run Draft Tools > Refresh data from GitHub, then Rebuild formulas & formatting","")')
    .setFontColor('#cc0000').setFontWeight('bold');

  // One rule per metric spanning every position panel: percentiles are taken
  // over the whole board, so the top-ranked players anywhere get the colour
  // instead of each panel shading its own local leaders.
  var rankFade = Math.max(1, Math.round(100 * CF_RANK_TOP_N / masterRows));
  var rankMid = Math.max(1, Math.round(rankFade / 3));
  var parStart = Math.min(99, 100 - Math.round(100 * CF_PAR_TOP_N / masterRows));
  [[xrkRanges, '#6d9eeb', '#c9daf8'], [adpRanges, '#93c47d', '#d9ead3']].forEach(function (c) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges(c[0])
      .setGradientMinpointWithValue(c[1], IT.MIN, '')
      .setGradientMidpointWithValue(c[2], IT.PERCENTILE, String(rankMid))
      .setGradientMaxpointWithValue('#ffffff', IT.PERCENTILE, String(rankFade))
      .build());
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges(parRanges)
    .setGradientMinpointWithValue('#ffffff', IT.PERCENTILE, String(parStart))
    .setGradientMaxpointWithValue('#b4a7d6', IT.MAX, '')
    .build());
  // Survival odds: pale red = likely gone (act now), fading out when he is
  // safe. Anchored at 0/0.5/1, so it already reads the same in every panel.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges(probRanges)
    .setGradientMinpointWithValue('#f4cccc', IT.NUMBER, '0')
    .setGradientMidpointWithValue('#fff2cc', IT.NUMBER, '0.5')
    .setGradientMaxpointWithValue('#ffffff', IT.NUMBER, '1')
    .build());
  // Risk: zero (and below, where waiting wins) stays white; red as the
  // expected cost of passing grows.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges(urgRanges)
    .setGradientMinpointWithValue('#ffffff', IT.MIN, '')
    .setGradientMidpointWithValue('#ffffff', IT.NUMBER, '0')
    .setGradientMaxpointWithValue('#e06666', IT.MAX, '')
    .build());

  b.setConditionalFormatRules(rules);
  b.setFrozenRows(HEADER_ROW);
}

/**
 * Freeze the current Board into BoardBackup: same values and look, but every
 * formula replaced by its result, so the snapshot cannot be broken by later
 * edits to Master, Log or Params. Rerun any time from the Draft Tools menu.
 */
function snapshotBackup_(ss) {
  var board = ss.getSheetByName('Board');
  if (!board) return;
  var old = ss.getSheetByName(BACKUP_NAME);
  if (old) ss.deleteSheet(old);
  var stray = ss.getSheetByName('Copy of Board');
  if (stray) ss.deleteSheet(stray);

  var bak = board.copyTo(ss).setName(BACKUP_NAME);
  var rng = bak.getDataRange();
  rng.setValues(rng.getValues());   // formulas -> plain values
  bak.clearConditionalFormatRules();  // static sheet: colours are already baked

  var sc = blockStarts_()[2];
  bak.getRange(1, sc, 6, 7).breakApart().clearContent();
  bak.getRange(1, sc, 6, 7).merge()
    .setValue('BACKUP — DO NOT DRAFT FROM THIS SHEET')
    .setBackground('#f4cccc').setFontColor('#990000').setFontSize(28)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(true);
  ss.setActiveSheet(board);
}

function snapshotBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  snapshotBackup_(ss);
  ss.toast('BoardBackup re-snapshotted from the current Board.');
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
  snapshotBackup_(ss);
  ss.toast('Raw + RawIDP refreshed from GitHub; formulas resized to the new data.');
}

// ---------- draft macros ----------

function markDrafted() { logPick_(); }

function logPick_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var fpid, name;

  if (sheet.getName() === 'BoardBackup') {
    ss.toast('That is the backup board — draft from the Board tab.');
    return;
  }
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
  log.appendRow([log.getLastRow(), fpid, name]);
  ss.toast('Pick ' + log.getLastRow() + ': ' + name, '', 3);
}

function undoLastPick() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName('Log');
  if (log.getLastRow() <= 1) { ss.toast('Nothing to undo.'); return; }
  var name = log.getRange(log.getLastRow(), 3).getValue();
  log.deleteRow(log.getLastRow());
  ss.toast('Undid: ' + name);
}
