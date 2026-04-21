/**
 * Airbnb Portfolio Dashboard - Google Apps Script
 *
 * Jak spustit:
 *   1. Otevrete novy Google Spreadsheet (sheets.new)
 *   2. Rozsireni > Apps Script
 *   3. Nahradte veskeры existujici kod timto souborem (Ctrl+A, vlozit)
 *   4. Kliknete na trojuhelnik (Run) > funkce "createDashboard"
 *   5. Schvalte pozadovana opravneni
 */

// --- KONFIGURACE ---

var APARTMENTS = [
  'Byt 1', 'Byt 2', 'Byt 3', 'Byt 4', 'Byt 5',
  'Byt 6', 'Byt 7', 'Byt 8', 'Byt 9', 'Byt 10'
];

var PLATFORMS = ['Airbnb', 'Prima rezervace', 'Booking.com', 'Jine'];

var FIXED_TYPES = [
  'Sprava nemovitosti',
  'Pojisteni',
  'Hypoteka / najem',
  'Dane a poplatky',
  'Internet / sluzby',
  'Ostatni fixni'
];

var VAR_TYPES = [
  'Oprava',
  'Udrzba',
  'Uklid',
  'Vybaveni / doplneni',
  'Pradelna',
  'Zahrada / exteriér',
  'Ostatni'
];

var MONTHS = [
  'Leden', 'Unor', 'Brezen', 'Duben', 'Kveten', 'Cerven',
  'Cervenec', 'Srpen', 'Zari', 'Rijen', 'Listopad', 'Prosinec'
];

var MONTH_ARRAY = '{"Leden","Unor","Brezen","Duben","Kveten","Cerven","Cervenec","Srpen","Zari","Rijen","Listopad","Prosinec"}';

// --- HLAVNI FUNKCE ---

function createDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sheetOrder = [
    'Byty',
    'Trzby',
    'Fixni naklady',
    'Variabilni naklady',
    'Dashboard',
    'Detail bytu'
  ];

  var sheets = {};
  var i, name, s;

  for (i = 0; i < sheetOrder.length; i++) {
    name = sheetOrder[i];
    s = ss.getSheetByName(name);
    if (s === null) {
      s = ss.insertSheet(name);
    } else {
      s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).breakApart();
      s.clearContents();
      s.clearFormats();
      s.clearConditionalFormatRules();
    }
    sheets[name] = s;
  }

  var allSheets = ss.getSheets();
  for (i = 0; i < allSheets.length; i++) {
    var sName = allSheets[i].getName();
    var found = false;
    for (var j = 0; j < sheetOrder.length; j++) {
      if (sheetOrder[j] === sName) { found = true; break; }
    }
    if (!found && allSheets[i].getLastRow() === 0) {
      try { ss.deleteSheet(allSheets[i]); } catch (e) {}
    }
  }

  buildBytySheet(sheets['Byty']);
  buildTrzbySheet(sheets['Trzby']);
  buildFixniSheet(sheets['Fixni naklady']);
  buildVariabilniSheet(sheets['Variabilni naklady']);
  buildDashboardSheet(sheets['Dashboard']);
  buildDetailSheet(sheets['Detail bytu']);

  for (i = 0; i < sheetOrder.length; i++) {
    ss.setActiveSheet(sheets[sheetOrder[i]]);
    ss.moveActiveSheet(i + 1);
  }

  ss.setActiveSheet(sheets['Dashboard']);

  ui.alert(
    'Dashboard byl vytvoren!',
    'Pred prvnim pouzitim upravte nazvy bytu v zalozce "Byty".\n\n' +
    'Spravce zadava data do:\n' +
    '- Trzby (po kazde rezervaci)\n' +
    '- Fixni naklady (jednou mesicne)\n' +
    '- Variabilni naklady (prubezne)\n\n' +
    'Vy sledujete vysledky v:\n' +
    '- Dashboard (prehled vsech bytu)\n' +
    '- Detail bytu (detail jednoho bytu)',
    ui.ButtonSet.OK
  );
}

// --- SHEET: Byty ---

function buildBytySheet(s) {
  s.setTabColor('#5c6bc0');

  setHeaders(s, 1, ['#', 'Nazev bytu', 'Adresa', 'Poznamka'], '#283593');
  setColWidths(s, [40, 160, 280, 220]);

  for (var i = 0; i < APARTMENTS.length; i++) {
    s.getRange(i + 2, 1, 1, 4).setValues([[i + 1, APARTMENTS[i], '', '']]);
  }

  s.getRange('A2:A11').setHorizontalAlignment('center');
  s.setFrozenRows(1);

  s.getRange('A13').setValue(
    'Upravte "Nazev bytu" podle potreby. ' +
    'Stejny nazev se pouziva v rozbalovacich nabidkach - musi se presne shodovat.'
  ).setFontColor('#666666').setFontStyle('italic').setWrap(true);
}

// --- SHEET: Trzby ---

function buildTrzbySheet(s) {
  s.setTabColor('#2e7d32');

  var headers = [
    'Datum prijezdu', 'Datum odjezdu', 'Byt', 'Platforma',
    'Jmeno hosta', 'Noci *', 'Trzba (Kc)', 'Provize %',
    'Provize Kc *', 'Cista trzba *'
  ];
  setHeaders(s, 1, headers, '#1b5e20');

  s.getRange('F1').setBackground('#81c784');
  s.getRange('I1').setBackground('#81c784');
  s.getRange('J1').setBackground('#81c784');

  s.getRange('A2:B1000').setNumberFormat('DD.MM.YYYY');
  s.getRange('G2:G1000').setNumberFormat('#,##0');
  s.getRange('H2:H1000').setNumberFormat('0.00');
  s.getRange('I2:J1000').setNumberFormat('#,##0');

  var rows = 500;
  var noci = [];
  var prov = [];
  var cista = [];
  for (var r = 2; r <= rows + 1; r++) {
    noci.push(['=IF(OR(A' + r + '="",B' + r + '=""),"",B' + r + '-A' + r + ')']);
    prov.push(['=IF(G' + r + '="","",ROUND(G' + r + '*H' + r + '/100,0))']);
    cista.push(['=IF(G' + r + '="","",G' + r + '-I' + r + ')']);
  }
  s.getRange(2, 6, rows, 1).setFormulas(noci);
  s.getRange(2, 9, rows, 1).setFormulas(prov);
  s.getRange(2, 10, rows, 1).setFormulas(cista);

  addDropdown(s, 'C2:C1000', APARTMENTS);
  addDropdown(s, 'D2:D1000', PLATFORMS);

  setColWidths(s, [125, 125, 140, 155, 165, 55, 110, 80, 115, 125]);
  s.setFrozenRows(1);
}

// --- SHEET: Fixni naklady ---

function buildFixniSheet(s) {
  s.setTabColor('#c62828');

  setHeaders(s, 1,
    ['Rok', 'Mesic', 'Byt', 'Typ nakladu', 'Castka (Kc)', 'Poznamka'],
    '#b71c1c'
  );

  s.getRange('A2:A1000').setNumberFormat('0');
  s.getRange('E2:E1000').setNumberFormat('#,##0');

  addDropdown(s, 'B2:B1000', MONTHS);
  addDropdown(s, 'C2:C1000', APARTMENTS);
  addDropdown(s, 'D2:D1000', FIXED_TYPES);
  s.getRange('A2:A1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  setColWidths(s, [65, 115, 145, 210, 125, 260]);
  s.setFrozenRows(1);
}

// --- SHEET: Variabilni naklady ---

function buildVariabilniSheet(s) {
  s.setTabColor('#e65100');

  setHeaders(s, 1,
    ['Datum', 'Byt', 'Kategorie', 'Popis', 'Castka (Kc)', 'Cislo dokladu', 'Poznamka'],
    '#bf360c'
  );

  s.getRange('A2:A1000').setNumberFormat('DD.MM.YYYY');
  s.getRange('E2:E1000').setNumberFormat('#,##0');

  addDropdown(s, 'B2:B1000', APARTMENTS);
  addDropdown(s, 'C2:C1000', VAR_TYPES);

  setColWidths(s, [115, 145, 175, 290, 125, 125, 210]);
  s.setFrozenRows(1);
}

// --- SHEET: Dashboard ---

function buildDashboardSheet(s) {
  s.setTabColor('#6a1b9a');

  var now = new Date();
  var year = now.getFullYear();
  var monthName = MONTHS[now.getMonth()];

  s.getRange('A1').setValue('PREHLED PORTFOLIA - AIRBNB')
    .setFontSize(18).setFontWeight('bold').setFontColor('#4a148c');

  s.getRange('A3').setValue('ROK');
  s.getRange('B3').setValue(year).setFontWeight('bold').setFontSize(14);
  s.getRange('D3').setValue('MESIC');
  s.getRange('E3').setValue(monthName).setFontWeight('bold').setFontSize(14);
  s.getRange('A3:F3').setBackground('#f3e5f5');

  addDropdown(s, 'E3', MONTHS);
  s.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  s.getRange('O3').setFormula('=MATCH(E3,' + MONTH_ARRAY + ',0)');
  s.hideColumns(15);

  var tableHeaders = [
    'BYT',
    'Trzba (Kc)', 'Fixni nakl.', 'Var. nakl.', 'Zisk (Kc)', 'Obsazenost',
    'Trzba min.m.', 'Zisk min.m.', 'D trzba',
    'Trzba loni', 'Zisk loni', 'D trzba (loni)'
  ];
  setHeaders(s, 5, tableHeaders, '#4a148c');
  setColWidths(s, [160, 120, 120, 120, 120, 95, 120, 120, 95, 120, 120, 105]);

  var totalRow = 6 + APARTMENTS.length;
  var apt, r, q;

  for (var i = 0; i < APARTMENTS.length; i++) {
    r = 6 + i;
    apt = APARTMENTS[i];
    q = '"' + apt + '"';

    s.getRange(r, 1).setValue(apt);

    s.getRange(r, 2).setFormula(
      '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$B$3)*(MONTH(Trzby!$A$2:$A$1001)=$O$3)*(Trzby!$C$2:$C$1001=' + q + ')*(Trzby!$J$2:$J$1001))'
    );

    s.getRange(r, 3).setFormula(
      '=SUMPRODUCT((\'Fixni naklady\'!$A$2:$A$1001=$B$3)*(\'Fixni naklady\'!$B$2:$B$1001=$E$3)*(\'Fixni naklady\'!$C$2:$C$1001=' + q + ')*(\'Fixni naklady\'!$E$2:$E$1001))'
    );

    s.getRange(r, 4).setFormula(
      '=SUMPRODUCT((YEAR(\'Variabilni naklady\'!$A$2:$A$1001)=$B$3)*(MONTH(\'Variabilni naklady\'!$A$2:$A$1001)=$O$3)*(\'Variabilni naklady\'!$B$2:$B$1001=' + q + ')*(\'Variabilni naklady\'!$E$2:$E$1001))'
    );

    s.getRange(r, 5).setFormula('=' + colLetter(2) + r + '-' + colLetter(3) + r + '-' + colLetter(4) + r);

    s.getRange(r, 6).setFormula(
      '=IFERROR(SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$B$3)*(MONTH(Trzby!$A$2:$A$1001)=$O$3)*(Trzby!$C$2:$C$1001=' + q + ')*(Trzby!$F$2:$F$1001))/DAY(EOMONTH(DATE($B$3,$O$3,1),0)),0)'
    );

    s.getRange(r, 7).setFormula(
      '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=IF($O$3=1,$B$3-1,$B$3))*(MONTH(Trzby!$A$2:$A$1001)=IF($O$3=1,12,$O$3-1))*(Trzby!$C$2:$C$1001=' + q + ')*(Trzby!$J$2:$J$1001))'
    );

    var pY  = 'IF($O$3=1,$B$3-1,$B$3)';
    var pMN = 'INDEX(' + MONTH_ARRAY + ',1,IF($O$3=1,12,$O$3-1))';
    var pMI = 'IF($O$3=1,12,$O$3-1)';

    s.getRange(r, 8).setFormula(
      '=' + colLetter(7) + r +
      '-SUMPRODUCT((\'Fixni naklady\'!$A$2:$A$1001=' + pY + ')*(\'Fixni naklady\'!$B$2:$B$1001=' + pMN + ')*(\'Fixni naklady\'!$C$2:$C$1001=' + q + ')*(\'Fixni naklady\'!$E$2:$E$1001))' +
      '-SUMPRODUCT((YEAR(\'Variabilni naklady\'!$A$2:$A$1001)=' + pY + ')*(MONTH(\'Variabilni naklady\'!$A$2:$A$1001)=' + pMI + ')*(\'Variabilni naklady\'!$B$2:$B$1001=' + q + ')*(\'Variabilni naklady\'!$E$2:$E$1001))'
    );

    s.getRange(r, 9).setFormula('=' + colLetter(2) + r + '-' + colLetter(7) + r);

    s.getRange(r, 10).setFormula(
      '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$B$3-1)*(MONTH(Trzby!$A$2:$A$1001)=$O$3)*(Trzby!$C$2:$C$1001=' + q + ')*(Trzby!$J$2:$J$1001))'
    );

    s.getRange(r, 11).setFormula(
      '=' + colLetter(10) + r +
      '-SUMPRODUCT((\'Fixni naklady\'!$A$2:$A$1001=$B$3-1)*(\'Fixni naklady\'!$B$2:$B$1001=$E$3)*(\'Fixni naklady\'!$C$2:$C$1001=' + q + ')*(\'Fixni naklady\'!$E$2:$E$1001))' +
      '-SUMPRODUCT((YEAR(\'Variabilni naklady\'!$A$2:$A$1001)=$B$3-1)*(MONTH(\'Variabilni naklady\'!$A$2:$A$1001)=$O$3)*(\'Variabilni naklady\'!$B$2:$B$1001=' + q + ')*(\'Variabilni naklady\'!$E$2:$E$1001))'
    );

    s.getRange(r, 12).setFormula('=' + colLetter(2) + r + '-' + colLetter(10) + r);
  }

  s.getRange(totalRow, 1).setValue('CELKEM').setFontWeight('bold').setFontSize(12);

  for (var c = 2; c <= 12; c++) {
    var cl = colLetter(c);
    var formula;
    if (c === 6) {
      formula = '=AVERAGE(' + cl + '6:' + cl + (totalRow - 1) + ')';
    } else {
      formula = '=SUM(' + cl + '6:' + cl + (totalRow - 1) + ')';
    }
    s.getRange(totalRow, c).setFormula(formula).setFontWeight('bold');
  }

  var moneyCols = [2, 3, 4, 5, 7, 8, 9, 10, 11, 12];
  for (var mi = 0; mi < moneyCols.length; mi++) {
    s.getRange(6, moneyCols[mi], APARTMENTS.length + 1, 1).setNumberFormat('#,##0 "Kc"');
  }
  s.getRange(6, 6, APARTMENTS.length + 1, 1).setNumberFormat('0%');

  applyPosNegFormatting(s, 6, 5, APARTMENTS.length + 1, 1);
  applyPosNegFormatting(s, 6, 9, APARTMENTS.length + 1, 1);
  applyPosNegFormatting(s, 6, 12, APARTMENTS.length + 1, 1);

  for (var ri = 0; ri < APARTMENTS.length; ri++) {
    var bg = (ri % 2 === 0) ? '#fafafa' : '#f3e5f5';
    s.getRange(6 + ri, 1, 1, 12).setBackground(bg);
  }
  s.getRange(totalRow, 1, 1, 12).setBackground('#e8eaf6');

  s.setFrozenRows(5);
  s.setFrozenColumns(1);
}

// --- SHEET: Detail bytu ---

function buildDetailSheet(s) {
  s.setTabColor('#00695c');

  var now = new Date();

  s.getRange('A1').setValue('DETAIL BYTU')
    .setFontSize(18).setFontWeight('bold').setFontColor('#004d40');

  s.getRange('A3').setValue('BYT');
  s.getRange('B3').setValue('Byt 1').setFontWeight('bold').setFontSize(13);
  s.getRange('D3').setValue('ROK');
  s.getRange('E3').setValue(now.getFullYear()).setFontWeight('bold').setFontSize(13);
  s.getRange('G3').setValue('MESIC');
  s.getRange('H3').setValue(MONTHS[now.getMonth()]).setFontWeight('bold').setFontSize(13);
  s.getRange('A3:I3').setBackground('#e0f2f1');

  addDropdown(s, 'B3', APARTMENTS);
  addDropdown(s, 'H3', MONTHS);
  s.getRange('E3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  s.getRange('K3').setFormula('=MATCH(H3,' + MONTH_ARRAY + ',0)');
  s.hideColumns(11);

  s.getRange('A5').setValue('SHRNUT MESICE')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  var summaryDefs = [
    ['Cista trzba',       '#,##0 "Kc"', '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$E$3)*(MONTH(Trzby!$A$2:$A$1001)=$K$3)*(Trzby!$C$2:$C$1001=$B$3)*(Trzby!$J$2:$J$1001))', false],
    ['Pocet noci',        '0',          '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$E$3)*(MONTH(Trzby!$A$2:$A$1001)=$K$3)*(Trzby!$C$2:$C$1001=$B$3)*(Trzby!$F$2:$F$1001))', false],
    ['Obsazenost',        '0%',         '=IFERROR(B7/DAY(EOMONTH(DATE($E$3,$K$3,1),0)),0)', false],
    ['Pocet rezervaci',   '0',          '=SUMPRODUCT((YEAR(Trzby!$A$2:$A$1001)=$E$3)*(MONTH(Trzby!$A$2:$A$1001)=$K$3)*(Trzby!$C$2:$C$1001=$B$3))', false],
    ['Prumer Kc / noc',   '#,##0 "Kc"', '=IFERROR(B6/B7,0)', false],
    ['',                  '',           '', false],
    ['Fixni naklady',     '#,##0 "Kc"', '=SUMPRODUCT((\'Fixni naklady\'!$A$2:$A$1001=$E$3)*(\'Fixni naklady\'!$B$2:$B$1001=$H$3)*(\'Fixni naklady\'!$C$2:$C$1001=$B$3)*(\'Fixni naklady\'!$E$2:$E$1001))', false],
    ['Variabilni naklady','#,##0 "Kc"', '=SUMPRODUCT((YEAR(\'Variabilni naklady\'!$A$2:$A$1001)=$E$3)*(MONTH(\'Variabilni naklady\'!$A$2:$A$1001)=$K$3)*(\'Variabilni naklady\'!$B$2:$B$1001=$B$3)*(\'Variabilni naklady\'!$E$2:$E$1001))', false],
    ['Celkove naklady',   '#,##0 "Kc"', '=B12+B13', false],
    ['',                  '',           '', false],
    ['CISTY ZISK',        '#,##0 "Kc"', '=B6-B14', true]
  ];

  for (var i = 0; i < summaryDefs.length; i++) {
    var row = 6 + i;
    var def = summaryDefs[i];
    var label = def[0];
    var fmt   = def[1];
    var f     = def[2];
    var bold  = def[3];

    s.getRange(row, 1).setValue(label);
    if (f !== '') {
      s.getRange(row, 2).setFormula(f);
    }
    if (fmt !== '') {
      s.getRange(row, 2).setNumberFormat(fmt);
    }
    if (bold) {
      s.getRange(row, 1).setFontWeight('bold').setFontSize(13);
      s.getRange(row, 2).setFontWeight('bold').setFontSize(13);
      s.getRange(row, 1, 1, 2).setBackground('#e8f5e9');
    }
  }

  applyPosNegFormatting(s, 16, 2, 1, 1);

  setColWidths(s, [185, 135, 20, 55, 80, 20, 65, 140]);

  s.getRange('D5').setValue('REZERVACE V MESICI')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  setHeaders(s, 6,
    ['Prijezd', 'Odjezd', 'Noci', 'Platforma', 'Host', 'Trzba (Kc)', 'Cista trzba (Kc)'],
    '#1b5e20', 4
  );

  s.getRange('D7').setFormula(
    '=IFERROR(' +
      'FILTER(' +
        'CHOOSE({1,2,3,4,5,6,7},' +
          'Trzby!$A$2:$A$1001,Trzby!$B$2:$B$1001,Trzby!$F$2:$F$1001,' +
          'Trzby!$D$2:$D$1001,Trzby!$E$2:$E$1001,Trzby!$G$2:$G$1001,Trzby!$J$2:$J$1001),' +
        '(YEAR(Trzby!$A$2:$A$1001)=$E$3)*(MONTH(Trzby!$A$2:$A$1001)=$K$3)*(Trzby!$C$2:$C$1001=$B$3)' +
      '),' +
      '{"Zadne rezervace","","","","","",""})'
  );

  s.getRange('D7:E106').setNumberFormat('DD.MM.YYYY');
  s.getRange('I7:J106').setNumberFormat('#,##0 "Kc"');

  s.getRange('A19').setValue('VARIABILNI NAKLADY V MESICI')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  setHeaders(s, 20, ['Datum', 'Kategorie', 'Popis', 'Castka (Kc)', 'Doklad'], '#bf360c');

  s.getRange('A21').setFormula(
    '=IFERROR(' +
      'FILTER(' +
        'CHOOSE({1,2,3,4,5},' +
          '\'Variabilni naklady\'!$A$2:$A$1001,\'Variabilni naklady\'!$C$2:$C$1001,' +
          '\'Variabilni naklady\'!$D$2:$D$1001,\'Variabilni naklady\'!$E$2:$E$1001,' +
          '\'Variabilni naklady\'!$F$2:$F$1001),' +
        '(YEAR(\'Variabilni naklady\'!$A$2:$A$1001)=$E$3)*(MONTH(\'Variabilni naklady\'!$A$2:$A$1001)=$K$3)*(\'Variabilni naklady\'!$B$2:$B$1001=$B$3)' +
      '),' +
      '{"Zadne naklady","","","",""})'
  );

  s.getRange('A21:A70').setNumberFormat('DD.MM.YYYY');
  s.getRange('D21:D70').setNumberFormat('#,##0 "Kc"');

  s.setFrozenRows(3);
}

// --- POMOCNE FUNKCE ---

function setHeaders(sheet, row, headers, bgColor, startCol) {
  if (!startCol) { startCol = 1; }
  sheet.getRange(row, startCol, 1, headers.length)
    .setValues([headers])
    .setBackground(bgColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function setColWidths(sheet, widths, startCol) {
  if (!startCol) { startCol = 1; }
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(startCol + i, widths[i]);
  }
}

function addDropdown(sheet, a1, values) {
  sheet.getRange(a1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .build()
  );
}

function colLetter(n) {
  var s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function applyPosNegFormatting(sheet, startRow, startCol, numRows, numCols) {
  var range = sheet.getRange(startRow, startCol, numRows, numCols);
  var rules = sheet.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#1b5e20')
      .setRanges([range])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#b71c1c')
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}
