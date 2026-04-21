/**
 * Airbnb Portfolio Dashboard – Google Apps Script
 *
 * Jak spustit:
 *   1. Otevřete nový Google Spreadsheet (sheets.new)
 *   2. Rozšíření > Apps Script
 *   3. Nahraďte veškerý existující kód tímto souborem (Ctrl+A, vložit)
 *   4. Klikněte na trojúhelník ▶ (Run) → funkce "createDashboard"
 *   5. Schvalte požadovaná oprávnění
 *
 * Výsledek:
 *   ⚙️ Byty            – číselník bytů (upravte názvy)
 *   Tržby              – správce zadává každou rezervaci
 *   Fixní náklady      – správce zadává jednou měsíčně
 *   Variabilní náklady – správce zadává průběžně
 *   📊 Dashboard       – váš přehled všech 10 bytů
 *   📊 Detail bytu     – detail jednoho bytu + seznam rezervací a nákladů
 */

// ─── KONFIGURACE ────────────────────────────────────────────────────────────

const CONFIG = {
  apartments: [
    'Byt 1', 'Byt 2', 'Byt 3', 'Byt 4', 'Byt 5',
    'Byt 6', 'Byt 7', 'Byt 8', 'Byt 9', 'Byt 10'
  ],
  platforms: ['Airbnb', 'Přímá rezervace', 'Booking.com', 'Jiné'],
  fixedCostTypes: [
    'Správa nemovitosti',
    'Pojištění',
    'Hypotéka / nájem',
    'Daně a poplatky',
    'Internet / služby',
    'Ostatní fixní'
  ],
  variableCostTypes: [
    'Oprava',
    'Údržba',
    'Úklid',
    'Vybavení / doplnění',
    'Prádelna',
    'Zahrada / exteriér',
    'Ostatní'
  ],
  months: [
    'Leden', 'Únor', 'Březen', 'Duben', 'Květen', 'Červen',
    'Červenec', 'Srpen', 'Září', 'Říjen', 'Listopad', 'Prosinec'
  ]
};

const MONTH_ARRAY_FORMULA =
  '{"Leden","Únor","Březen","Duben","Květen","Červen","Červenec","Srpen","Září","Říjen","Listopad","Prosinec"}';

// ─── HLAVNÍ FUNKCE ──────────────────────────────────────────────────────────

function createDashboard() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  const sheetOrder = [
    '⚙️ Byty',
    'Tržby',
    'Fixní náklady',
    'Variabilní náklady',
    '📊 Dashboard',
    '📊 Detail bytu'
  ];

  // Vytvořit nebo vymazat záložky
  const sheets = {};
  sheetOrder.forEach(name => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    else     s.clearContents().clearFormats().clearConditionalFormatRules();
    sheets[name] = s;
  });

  // Odstranit výchozí prázdné záložky (List1 / Sheet1 apod.)
  ss.getSheets().forEach(s => {
    if (!sheetOrder.includes(s.getName()) && s.getLastRow() === 0) {
      try { ss.deleteSheet(s); } catch (_) {}
    }
  });

  // Sestavit každou záložku
  buildBytySheet         (sheets['⚙️ Byty']);
  buildTrzbySheet        (sheets['Tržby']);
  buildFixniSheet        (sheets['Fixní náklady']);
  buildVariabilniSheet   (sheets['Variabilní náklady']);
  buildDashboardSheet    (sheets['📊 Dashboard']);
  buildDetailSheet       (sheets['📊 Detail bytu']);

  // Seřadit záložky
  sheetOrder.forEach((name, idx) => {
    ss.setActiveSheet(sheets[name]);
    ss.moveActiveSheet(idx + 1);
  });

  ss.setActiveSheet(sheets['📊 Dashboard']);

  ui.alert(
    '✅ Dashboard byl vytvořen!',
    'Před prvním použitím:\n' +
    '• Upravte názvy bytů v záložce ⚙️ Byty\n\n' +
    'Správce zadává data do:\n' +
    '• Tržby – po každé rezervaci\n' +
    '• Fixní náklady – jednou měsíčně\n' +
    '• Variabilní náklady – průběžně\n\n' +
    'Vy sledujete výsledky v:\n' +
    '• 📊 Dashboard – přehled všech bytů s porovnáním\n' +
    '• 📊 Detail bytu – detail jednoho bytu',
    ui.ButtonSet.OK
  );
}

// ─── SHEET: ⚙️ Byty ─────────────────────────────────────────────────────────

function buildBytySheet(s) {
  s.setTabColor('#5c6bc0');

  setHeaders(s, 1, ['#', 'Název bytu', 'Adresa', 'Poznámka'], '#283593');
  setColWidths(s, [40, 160, 280, 220]);

  CONFIG.apartments.forEach((name, i) => {
    s.getRange(i + 2, 1, 1, 4).setValues([[i + 1, name, '', '']]);
  });

  s.getRange('A2:A11').setHorizontalAlignment('center');
  s.setFrozenRows(1);

  const note = s.getRange('A13');
  note.setValue(
    'Poznámka: Upravte "Název bytu" podle potřeby. ' +
    'Stejný název se pak používá v rozbalovacích nabídkách v záložkách Tržby, ' +
    'Fixní a Variabilní náklady – musí se přesně shodovat.'
  );
  note.setFontColor('#666666').setFontStyle('italic').setWrap(true);
  s.getRange('A13:D13').merge();
}

// ─── SHEET: Tržby ───────────────────────────────────────────────────────────

function buildTrzbySheet(s) {
  s.setTabColor('#2e7d32');

  const headers = [
    'Datum příjezdu', 'Datum odjezdu', 'Byt', 'Platforma',
    'Jméno hosta', 'Nocí *', 'Tržba (Kč)', 'Provize %',
    'Provize Kč *', 'Čistá tržba *'
  ];
  setHeaders(s, 1, headers, '#1b5e20');

  // Automaticky počítané sloupce (zelené záhlaví)
  ['F1', 'I1', 'J1'].forEach(c => {
    s.getRange(c).setBackground('#81c784')
      .setNote('Tento sloupec se počítá automaticky – nevyplňujte.');
  });

  // Formáty
  s.getRange('A2:B1000').setNumberFormat('DD.MM.YYYY');
  s.getRange('G2:G1000').setNumberFormat('#,##0');
  s.getRange('H2:H1000').setNumberFormat('0.00');
  s.getRange('I2:J1000').setNumberFormat('#,##0');

  // Automatické vzorce pro 500 řádků
  const rows = 500;
  const noci = [], prov = [], cista = [];
  for (let r = 2; r <= rows + 1; r++) {
    noci .push([`=IF(OR(A${r}="",B${r}=""),"",B${r}-A${r})`]);
    prov .push([`=IF(G${r}="","",ROUND(G${r}*H${r}/100,0))`]);
    cista.push([`=IF(G${r}="","",G${r}-I${r})`]);
  }
  s.getRange(2, 6, rows, 1).setFormulas(noci);
  s.getRange(2, 9, rows, 1).setFormulas(prov);
  s.getRange(2, 10, rows, 1).setFormulas(cista);

  // Validace
  addDropdown(s, 'C2:C1000', CONFIG.apartments);
  addDropdown(s, 'D2:D1000', CONFIG.platforms);

  setColWidths(s, [125, 125, 140, 155, 165, 55, 110, 80, 115, 125]);
  s.setFrozenRows(1);
  s.getRange('A1').setNote(
    '* Sloupce označené hvězdičkou (F, I, J) se počítají samy.\n' +
    'Správce vyplňuje pouze: Datum příjezdu, Datum odjezdu, Byt, Platforma, Jméno hosta, Tržba, Provize %.'
  );
}

// ─── SHEET: Fixní náklady ────────────────────────────────────────────────────

function buildFixniSheet(s) {
  s.setTabColor('#c62828');

  setHeaders(s, 1,
    ['Rok', 'Měsíc', 'Byt', 'Typ nákladu', 'Částka (Kč)', 'Poznámka'],
    '#b71c1c'
  );

  s.getRange('A2:A1000').setNumberFormat('0');
  s.getRange('E2:E1000').setNumberFormat('#,##0');

  addDropdown(s, 'B2:B1000', CONFIG.months);
  addDropdown(s, 'C2:C1000', CONFIG.apartments);
  addDropdown(s, 'D2:D1000', CONFIG.fixedCostTypes);
  s.getRange('A2:A1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  setColWidths(s, [65, 115, 145, 210, 125, 260]);
  s.setFrozenRows(1);
}

// ─── SHEET: Variabilní náklady ───────────────────────────────────────────────

function buildVariabilniSheet(s) {
  s.setTabColor('#e65100');

  setHeaders(s, 1,
    ['Datum', 'Byt', 'Kategorie', 'Popis', 'Částka (Kč)', 'Číslo dokladu', 'Poznámka'],
    '#bf360c'
  );

  s.getRange('A2:A1000').setNumberFormat('DD.MM.YYYY');
  s.getRange('E2:E1000').setNumberFormat('#,##0');

  addDropdown(s, 'B2:B1000', CONFIG.apartments);
  addDropdown(s, 'C2:C1000', CONFIG.variableCostTypes);

  setColWidths(s, [115, 145, 175, 290, 125, 125, 210]);
  s.setFrozenRows(1);
}

// ─── SHEET: 📊 Dashboard ─────────────────────────────────────────────────────

function buildDashboardSheet(s) {
  s.setTabColor('#6a1b9a');

  const now       = new Date();
  const year      = now.getFullYear();
  const monthName = CONFIG.months[now.getMonth()];

  // Nadpis
  s.getRange('A1').setValue('PŘEHLED PORTFOLIA  –  AIRBNB')
    .setFontSize(18).setFontWeight('bold').setFontColor('#4a148c');
  // Nesloučit – merge přes A1:M1 by kolidoval se setFrozenColumns(1)

  // Výběr období
  s.getRange('A3').setValue('ROK');
  s.getRange('B3').setValue(year).setFontWeight('bold').setFontSize(14);
  s.getRange('D3').setValue('MĚSÍC');
  s.getRange('E3').setValue(monthName).setFontWeight('bold').setFontSize(14);
  s.getRange('A3:F3').setBackground('#f3e5f5');

  addDropdown(s, 'E3', CONFIG.months);
  s.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  // Skrytý pomocný sloupec O: číslo měsíce
  s.getRange('O3').setFormula(`=MATCH(E3,${MONTH_ARRAY_FORMULA},0)`);
  s.hideColumns(15);

  // Záhlaví tabulky
  const tableHeaders = [
    'BYT',
    'Tržba (Kč)', 'Fixní nákl.', 'Var. nákl.', 'Zisk (Kč)', 'Obsazenost',
    'Tržba min.m.', 'Zisk min.m.', '∆ tržba',
    'Tržba loni', 'Zisk loni', '∆ tržba (loni)'
  ];
  setHeaders(s, 5, tableHeaders, '#4a148c');
  setColWidths(s, [160, 120, 120, 120, 120, 95, 120, 120, 95, 120, 120, 105]);

  // Vzorce pro každý byt
  CONFIG.apartments.forEach((apt, i) => {
    const r = 6 + i;
    const q = `"${apt}"`;           // quoted apartment name for formula

    // A: název bytu
    s.getRange(r, 1).setValue(apt);

    // B: Čistá tržba – aktuální měsíc
    s.getRange(r, 2).setFormula(
      `=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$B$3)*(MONTH(Tržby!$A$2:$A$1001)=$O$3)*(Tržby!$C$2:$C$1001=${q})*(Tržby!$J$2:$J$1001))`
    );

    // C: Fixní náklady – aktuální měsíc
    s.getRange(r, 3).setFormula(
      `=SUMPRODUCT(('Fixní náklady'!$A$2:$A$1001=$B$3)*('Fixní náklady'!$B$2:$B$1001=$E$3)*('Fixní náklady'!$C$2:$C$1001=${q})*('Fixní náklady'!$E$2:$E$1001))`
    );

    // D: Variabilní náklady – aktuální měsíc
    s.getRange(r, 4).setFormula(
      `=SUMPRODUCT((YEAR('Variabilní náklady'!$A$2:$A$1001)=$B$3)*(MONTH('Variabilní náklady'!$A$2:$A$1001)=$O$3)*('Variabilní náklady'!$B$2:$B$1001=${q})*('Variabilní náklady'!$E$2:$E$1001))`
    );

    // E: Zisk
    s.getRange(r, 5).setFormula(`=${cr(r,2)}-${cr(r,3)}-${cr(r,4)}`);

    // F: Obsazenost (nocí / počet dnů v měsíci)
    s.getRange(r, 6).setFormula(
      `=IFERROR(SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$B$3)*(MONTH(Tržby!$A$2:$A$1001)=$O$3)*(Tržby!$C$2:$C$1001=${q})*(Tržby!$F$2:$F$1001))/DAY(EOMONTH(DATE($B$3,$O$3,1),0)),0)`
    );

    // G: Tržba – předchozí měsíc
    s.getRange(r, 7).setFormula(
      `=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=IF($O$3=1,$B$3-1,$B$3))*(MONTH(Tržby!$A$2:$A$1001)=IF($O$3=1,12,$O$3-1))*(Tržby!$C$2:$C$1001=${q})*(Tržby!$J$2:$J$1001))`
    );

    // H: Zisk – předchozí měsíc
    const pY  = 'IF($O$3=1,$B$3-1,$B$3)';
    const pMN = `INDEX(${MONTH_ARRAY_FORMULA},1,IF($O$3=1,12,$O$3-1))`;
    const pMI = 'IF($O$3=1,12,$O$3-1)';
    s.getRange(r, 8).setFormula(
      `=${cr(r,7)}` +
      `-SUMPRODUCT(('Fixní náklady'!$A$2:$A$1001=${pY})*('Fixní náklady'!$B$2:$B$1001=${pMN})*('Fixní náklady'!$C$2:$C$1001=${q})*('Fixní náklady'!$E$2:$E$1001))` +
      `-SUMPRODUCT((YEAR('Variabilní náklady'!$A$2:$A$1001)=${pY})*(MONTH('Variabilní náklady'!$A$2:$A$1001)=${pMI})*('Variabilní náklady'!$B$2:$B$1001=${q})*('Variabilní náklady'!$E$2:$E$1001))`
    );

    // I: Delta tržba vs. předchozí měsíc
    s.getRange(r, 9).setFormula(`=${cr(r,2)}-${cr(r,7)}`);

    // J: Tržba – stejný měsíc loni
    s.getRange(r, 10).setFormula(
      `=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$B$3-1)*(MONTH(Tržby!$A$2:$A$1001)=$O$3)*(Tržby!$C$2:$C$1001=${q})*(Tržby!$J$2:$J$1001))`
    );

    // K: Zisk – stejný měsíc loni
    s.getRange(r, 11).setFormula(
      `=${cr(r,10)}` +
      `-SUMPRODUCT(('Fixní náklady'!$A$2:$A$1001=$B$3-1)*('Fixní náklady'!$B$2:$B$1001=$E$3)*('Fixní náklady'!$C$2:$C$1001=${q})*('Fixní náklady'!$E$2:$E$1001))` +
      `-SUMPRODUCT((YEAR('Variabilní náklady'!$A$2:$A$1001)=$B$3-1)*(MONTH('Variabilní náklady'!$A$2:$A$1001)=$O$3)*('Variabilní náklady'!$B$2:$B$1001=${q})*('Variabilní náklady'!$E$2:$E$1001))`
    );

    // L: Delta tržba vs. loni
    s.getRange(r, 12).setFormula(`=${cr(r,2)}-${cr(r,10)}`);
  });

  // Řádek CELKEM
  const totalRow = 6 + CONFIG.apartments.length;
  s.getRange(totalRow, 1).setValue('CELKEM').setFontWeight('bold').setFontSize(12);

  for (let c = 2; c <= 12; c++) {
    const cl = colLetter(c);
    const formula = (c === 6)
      ? `=AVERAGE(${cl}6:${cl}${totalRow - 1})`
      : `=SUM(${cl}6:${cl}${totalRow - 1})`;
    s.getRange(totalRow, c).setFormula(formula).setFontWeight('bold');
  }

  // Formáty čísel
  [2, 3, 4, 5, 7, 8, 9, 10, 11, 12].forEach(c => {
    s.getRange(6, c, CONFIG.apartments.length + 1, 1).setNumberFormat('#,##0 "Kč"');
  });
  s.getRange(6, 6, CONFIG.apartments.length + 1, 1).setNumberFormat('0%');

  // Podmíněné formátování: kladné/záporné hodnoty (Zisk, Delta)
  [5, 9, 12].forEach(c => {
    applyPosNegFormatting(s, 6, c, CONFIG.apartments.length + 1, 1);
  });

  // Střídání barev řádků
  CONFIG.apartments.forEach((_, i) => {
    s.getRange(6 + i, 1, 1, 12)
      .setBackground(i % 2 === 0 ? '#fafafa' : '#f3e5f5');
  });
  s.getRange(totalRow, 1, 1, 12).setBackground('#e8eaf6');

  s.setFrozenRows(5);
  s.setFrozenColumns(1);
}

// ─── SHEET: 📊 Detail bytu ───────────────────────────────────────────────────

function buildDetailSheet(s) {
  s.setTabColor('#00695c');

  const now = new Date();

  // Nadpis
  s.getRange('A1').setValue('DETAIL BYTU')
    .setFontSize(18).setFontWeight('bold').setFontColor('#004d40');
  s.getRange('A1:I1').merge();

  // Výběr bytu / roku / měsíce
  s.getRange('A3').setValue('BYT');
  s.getRange('B3').setValue('Byt 1').setFontWeight('bold').setFontSize(13);
  s.getRange('D3').setValue('ROK');
  s.getRange('E3').setValue(now.getFullYear()).setFontWeight('bold').setFontSize(13);
  s.getRange('G3').setValue('MĚSÍC');
  s.getRange('H3').setValue(CONFIG.months[now.getMonth()]).setFontWeight('bold').setFontSize(13);
  s.getRange('A3:I3').setBackground('#e0f2f1');

  addDropdown(s, 'B3', CONFIG.apartments);
  addDropdown(s, 'H3', CONFIG.months);
  s.getRange('E3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(2020, 2035).build()
  );

  // Skrytý pomocný sloupec K: číslo měsíce
  s.getRange('K3').setFormula(`=MATCH(H3,${MONTH_ARRAY_FORMULA},0)`);
  s.hideColumns(11);

  // ── SHRNUTÍ ──────────────────────────────────────────────────────────────
  s.getRange('A5').setValue('SHRNUTÍ MĚSÍCE')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  const summary = [
    { label: 'Čistá tržba', fmt: '#,##0 "Kč"',
      f: '=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$E$3)*(MONTH(Tržby!$A$2:$A$1001)=$K$3)*(Tržby!$C$2:$C$1001=$B$3)*(Tržby!$J$2:$J$1001))' },
    { label: 'Počet nocí', fmt: '0',
      f: '=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$E$3)*(MONTH(Tržby!$A$2:$A$1001)=$K$3)*(Tržby!$C$2:$C$1001=$B$3)*(Tržby!$F$2:$F$1001))' },
    { label: 'Obsazenost', fmt: '0%',
      f: '=IFERROR(B7/DAY(EOMONTH(DATE($E$3,$K$3,1),0)),0)' },
    { label: 'Počet rezervací', fmt: '0',
      f: '=SUMPRODUCT((YEAR(Tržby!$A$2:$A$1001)=$E$3)*(MONTH(Tržby!$A$2:$A$1001)=$K$3)*(Tržby!$C$2:$C$1001=$B$3))' },
    { label: 'Průměr Kč / noc', fmt: '#,##0 "Kč"',
      f: '=IFERROR(B6/B7,0)' },
    { label: '' },
    { label: 'Fixní náklady', fmt: '#,##0 "Kč"',
      f: '=SUMPRODUCT((\'Fixní náklady\'!$A$2:$A$1001=$E$3)*(\'Fixní náklady\'!$B$2:$B$1001=$H$3)*(\'Fixní náklady\'!$C$2:$C$1001=$B$3)*(\'Fixní náklady\'!$E$2:$E$1001))' },
    { label: 'Variabilní náklady', fmt: '#,##0 "Kč"',
      f: '=SUMPRODUCT((YEAR(\'Variabilní náklady\'!$A$2:$A$1001)=$E$3)*(MONTH(\'Variabilní náklady\'!$A$2:$A$1001)=$K$3)*(\'Variabilní náklady\'!$B$2:$B$1001=$B$3)*(\'Variabilní náklady\'!$E$2:$E$1001))' },
    { label: 'Celkové náklady', fmt: '#,##0 "Kč"',
      f: '=B12+B13' },
    { label: '' },
    { label: 'ČISTÝ ZISK', fmt: '#,##0 "Kč"', f: '=B6-B14', bold: true }
  ];

  summary.forEach(({ label, fmt, f, bold }, i) => {
    const r   = 6 + i;
    const lc  = s.getRange(r, 1);
    const vc  = s.getRange(r, 2);
    lc.setValue(label);
    if (f)    vc.setFormula(f);
    if (fmt)  vc.setNumberFormat(fmt);
    if (bold) {
      lc.setFontWeight('bold').setFontSize(13);
      vc.setFontWeight('bold').setFontSize(13);
      s.getRange(r, 1, 1, 2).setBackground('#e8f5e9');
    }
  });

  applyPosNegFormatting(s, 16, 2, 1, 1); // zisk cell

  setColWidths(s, [185, 135, 20, 55, 80, 20, 65, 140]);

  // ── SEZNAM REZERVACÍ ─────────────────────────────────────────────────────
  s.getRange('D5').setValue('REZERVACE V MĚSÍCI')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  setHeaders(s, 6, ['Příjezd', 'Odjezd', 'Nocí', 'Platforma', 'Host', 'Tržba (Kč)', 'Čistá tržba (Kč)'],
    '#1b5e20', 4);

  // FILTER: vybrané sloupce z Tržby
  s.getRange('D7').setFormula(
    '=IFERROR(' +
      'FILTER(' +
        'CHOOSE({1,2,3,4,5,6,7},' +
          'Tržby!$A$2:$A$1001,Tržby!$B$2:$B$1001,Tržby!$F$2:$F$1001,' +
          'Tržby!$D$2:$D$1001,Tržby!$E$2:$E$1001,Tržby!$G$2:$G$1001,Tržby!$J$2:$J$1001),' +
        '(YEAR(Tržby!$A$2:$A$1001)=$E$3)*(MONTH(Tržby!$A$2:$A$1001)=$K$3)*(Tržby!$C$2:$C$1001=$B$3)' +
      '),' +
      '{"Žádné rezervace","","","","","",""})'
  );

  s.getRange('D7:E106').setNumberFormat('DD.MM.YYYY');
  s.getRange('I7:J106').setNumberFormat('#,##0 "Kč"');

  // ── VARIABILNÍ NÁKLADY ────────────────────────────────────────────────────
  s.getRange('A19').setValue('VARIABILNÍ NÁKLADY V MĚSÍCI')
    .setFontSize(13).setFontWeight('bold').setFontColor('#004d40');

  setHeaders(s, 20, ['Datum', 'Kategorie', 'Popis', 'Částka (Kč)', 'Doklad'], '#bf360c');

  s.getRange('A21').setFormula(
    '=IFERROR(' +
      'FILTER(' +
        'CHOOSE({1,2,3,4,5},' +
          '\'Variabilní náklady\'!$A$2:$A$1001,\'Variabilní náklady\'!$C$2:$C$1001,' +
          '\'Variabilní náklady\'!$D$2:$D$1001,\'Variabilní náklady\'!$E$2:$E$1001,' +
          '\'Variabilní náklady\'!$F$2:$F$1001),' +
        '(YEAR(\'Variabilní náklady\'!$A$2:$A$1001)=$E$3)*(MONTH(\'Variabilní náklady\'!$A$2:$A$1001)=$K$3)*(\'Variabilní náklady\'!$B$2:$B$1001=$B$3)' +
      '),' +
      '{"Žádné náklady","","","",""})'
  );

  s.getRange('A21:A70').setNumberFormat('DD.MM.YYYY');
  s.getRange('D21:D70').setNumberFormat('#,##0 "Kč"');

  s.setFrozenRows(3);
}

// ─── POMOCNÉ FUNKCE ──────────────────────────────────────────────────────────

function setHeaders(sheet, row, headers, bgColor, startCol) {
  startCol = startCol || 1;
  sheet.getRange(row, startCol, 1, headers.length)
    .setValues([headers])
    .setBackground(bgColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function setColWidths(sheet, widths, startCol) {
  startCol = startCol || 1;
  widths.forEach((w, i) => sheet.setColumnWidth(startCol + i, w));
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
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function cr(row, col) {
  return colLetter(col) + row;
}

function applyPosNegFormatting(sheet, startRow, startCol, numRows, numCols) {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  const rules = sheet.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor('#1b5e20').setRanges([range]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor('#b71c1c').setRanges([range]).build()
  );
  sheet.setConditionalFormatRules(rules);
}
