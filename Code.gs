/**
 * ====== CONFIG GLOBALE ======
 * Modifie uniquement cette section si tes noms/colonnes changent.
 */
const CFG = {
  MENU_NAME: 'Extraction',
  MENU_ITEM_TAGS: 'par Tags',
  MENU_ITEM_ACCOUNTS: 'par Compte',
  MENU_ITEM_WALLETS: 'par Wallet',

  SHEET_TRANSACTIONS: 'Transactions',
  SHEET_DROPDOWN: 'Dropdown',
  SHEET_EXPORT_PREFIX: 'Export_TMP_',

  // Colonnes 1-based
  COL_DATE: 1,          // Transactions ColA
  COL_AMOUNT: 3,        // Transactions ColC
  COL_TAG: 6,           // Transactions ColF
  COL_ACCOUNT: null,    // Transactions: colonne Compte
  COL_WALLET: null,     // Transactions: colonne Wallet
  COL_DROPDOWN_TAG: 5,  // Dropdown: colonne Tag
  COL_DROPDOWN_ACCOUNT: 1, // Dropdown: colonne Compte
  COL_DROPDOWN_WALLET: 3,  // Dropdown: colonne Wallet

  HEADER_ROW: 1,
  NULL_TAG_LABEL: 'Autre / Null',
  CURRENCY_FORMAT: '#,##0.00',
  PDF_FOLDER_FALLBACK_TO_ROOT: false,
  PDF_FOLDER_ID: '1MkEhcLCNEuvxr9WVkqFlXavVGaonwV4-',

  HEADER_ALIASES: {
    transactionTag: ['tag', 'tags', 'categorie', 'categories', 'category', 'category name'],
    transactionAccount: ['account', 'accounts', 'compte', 'comptes'],
    transactionWallet: ['wallet', 'wallets', 'portefeuille', 'portefeuilles'],
    dropdownTag: ['tag', 'tags', 'categorie', 'categories', 'category', 'category name'],
    dropdownAccount: ['account', 'accounts', 'compte', 'comptes'],
    dropdownWallet: ['wallet', 'wallets', 'portefeuille', 'portefeuilles']
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CFG.MENU_NAME)
    .addItem(CFG.MENU_ITEM_TAGS, 'openTagsExtractionModal')
    .addItem(CFG.MENU_ITEM_ACCOUNTS, 'openAccountsExtractionModal')
    .addItem(CFG.MENU_ITEM_WALLETS, 'openWalletsExtractionModal')
    .addToUi();
}

function openTagsExtractionModal() {
  openExtractionModal_({
    extractionType: 'tag',
    title: 'par Tags',
    description: 'Génère un export des dépenses par tag sur la période choisie.',
    selectionLabel: '',
    selectionPlaceholder: '',
    selectionOptions: []
  });
}

function openAccountsExtractionModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  openExtractionModal_({
    extractionType: 'account',
    title: 'par Compte',
    description: 'Filtre les transactions sur un compte précis puis génère le récapitulatif par tag.',
    selectionLabel: 'Compte',
    selectionPlaceholder: 'Choisis un compte',
    selectionOptions: getAccountsOptions_(ss)
  });
}

function openWalletsExtractionModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  openExtractionModal_({
    extractionType: 'wallet',
    title: 'par Wallet',
    description: 'Filtre les transactions sur un wallet précis puis génère le récapitulatif par tag.',
    selectionLabel: 'Wallet',
    selectionPlaceholder: 'Choisis un wallet',
    selectionOptions: getWalletsOptions_(ss)
  });
}

function openExtractionModal_(config) {
  const template = HtmlService.createTemplateFromFile('ExtractionModal');
  template.extractionType = config.extractionType;
  template.modalTitle = config.title;
  template.modalDescription = config.description;
  template.selectionLabel = config.selectionLabel || '';
  template.selectionPlaceholder = config.selectionPlaceholder || '';
  template.selectionOptions = config.selectionOptions || [];

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, config.title);
}

function runExtraction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extractionType = payload && payload.extractionType ? payload.extractionType : 'tag';
  const shouldExportPdf = !payload || payload.exportPdf !== false;
  const shouldExportSheet = !payload || payload.exportSheet !== false;

  if (!shouldExportPdf && !shouldExportSheet) {
    throw new Error('Choisis au moins un export: PDF ou Sheet.');
  }

  const range = buildDateRange_(payload);
  const tagsDict = buildTagsDict_(ss);
  const txSheet = ss.getSheetByName(CFG.SHEET_TRANSACTIONS);
  if (!txSheet) throw new Error('Sheet introuvable: ' + CFG.SHEET_TRANSACTIONS);

  const txCols = getTransactionColumns_(txSheet);
  const filter = buildFilterConfig_(extractionType, payload, txCols);
  const lastRow = txSheet.getLastRow();

  if (lastRow <= CFG.HEADER_ROW) {
    throw new Error('Aucune transaction à lire dans ' + CFG.SHEET_TRANSACTIONS);
  }

  const dataWidth = [txCols.date, txCols.amount, txCols.tag, filter.column].reduce(maxDefined_, 0);
  const data = txSheet
    .getRange(CFG.HEADER_ROW + 1, 1, lastRow - CFG.HEADER_ROW, dataWidth)
    .getValues();

  let matchedCount = 0;

  data.forEach((row) => {
    const rawDate = row[txCols.date - 1];
    const rawAmount = row[txCols.amount - 1];
    const rawTag = row[txCols.tag - 1];

    if (!(rawDate instanceof Date)) return;
    if (!isDateInRange_(rawDate, range.start, range.end)) return;

    const amount = Number(rawAmount);
    if (isNaN(amount) || amount >= 0) return;

    if (filter.column) {
      const rawFilterValue = row[filter.column - 1];
      if (!matchesFilterValue_(rawFilterValue, filter.value)) return;
    }

    let tag = String(rawTag || '').trim();
    if (!tag) tag = CFG.NULL_TAG_LABEL;
    if (!(tag in tagsDict)) tag = CFG.NULL_TAG_LABEL;

    tagsDict[tag] += amount;
    matchedCount++;
  });

  const titleParts = ['Extraction depenses par tag'];
  if (filter.label) {
    titleParts.push(filter.label + ': ' + filter.value);
  }

  return createExtractionOutput_(
    ss,
    tagsDict,
    range,
    matchedCount,
    shouldExportPdf,
    shouldExportSheet,
    titleParts.join(' - ')
  );
}

function getAccountsOptions_(ss) {
  return getDropdownValues_(ss, 'account');
}

function getWalletsOptions_(ss) {
  return getDropdownValues_(ss, 'wallet');
}

function buildTagsDict_(ss) {
  const dict = {};
  getDropdownValues_(ss, 'tag').forEach((tag) => {
    dict[tag] = 0;
  });

  if (!(CFG.NULL_TAG_LABEL in dict)) dict[CFG.NULL_TAG_LABEL] = 0;
  return dict;
}

function getDropdownValues_(ss, type) {
  const dropdownSheet = ss.getSheetByName(CFG.SHEET_DROPDOWN);
  if (!dropdownSheet) throw new Error('Sheet introuvable: ' + CFG.SHEET_DROPDOWN);

  const column = getDropdownColumnIndex_(dropdownSheet, type);
  const lastRow = dropdownSheet.getLastRow();
  if (lastRow <= CFG.HEADER_ROW) return [];

  const values = dropdownSheet
    .getRange(CFG.HEADER_ROW + 1, column, lastRow - CFG.HEADER_ROW, 1)
    .getValues();

  const seen = {};
  const result = [];

  values.forEach((row) => {
    const value = String(row[0] || '').trim();
    if (!value) return;

    const key = normalizeValue_(value);
    if (seen[key]) return;

    seen[key] = true;
    result.push(value);
  });

  return result;
}

function getTransactionColumns_(txSheet) {
  return {
    date: CFG.COL_DATE,
    amount: CFG.COL_AMOUNT,
    tag: getColumnIndex_(txSheet, CFG.HEADER_ALIASES.transactionTag, CFG.COL_TAG),
    account: getColumnIndex_(txSheet, CFG.HEADER_ALIASES.transactionAccount, CFG.COL_ACCOUNT),
    wallet: getColumnIndex_(txSheet, CFG.HEADER_ALIASES.transactionWallet, CFG.COL_WALLET)
  };
}

function getDropdownColumnIndex_(dropdownSheet, type) {
  if (type === 'tag') {
    return getColumnIndex_(dropdownSheet, CFG.HEADER_ALIASES.dropdownTag, CFG.COL_DROPDOWN_TAG);
  }
  if (type === 'account') {
    return getColumnIndex_(dropdownSheet, CFG.HEADER_ALIASES.dropdownAccount, CFG.COL_DROPDOWN_ACCOUNT);
  }
  if (type === 'wallet') {
    return getColumnIndex_(dropdownSheet, CFG.HEADER_ALIASES.dropdownWallet, CFG.COL_DROPDOWN_WALLET);
  }

  throw new Error('Type de dropdown non supporté: ' + type);
}

function getColumnIndex_(sheet, aliases, fallbackColumn) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn <= 0) {
    throw new Error('Aucune colonne trouvée dans ' + sheet.getName());
  }

  const headers = sheet.getRange(CFG.HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  const normalizedAliases = (aliases || []).map(normalizeValue_);

  for (let i = 0; i < headers.length; i++) {
    const header = normalizeValue_(headers[i]);
    if (header && normalizedAliases.indexOf(header) !== -1) {
      return i + 1;
    }
  }

  if (fallbackColumn) return fallbackColumn;

  throw new Error("Colonne introuvable dans '" + sheet.getName() + "' pour: " + (aliases || []).join(', '));
}

function buildFilterConfig_(extractionType, payload, txCols) {
  if (extractionType === 'tag') {
    return { column: null, value: '', label: '' };
  }

  if (!payload || !payload.selectionValue) {
    throw new Error('Sélection manquante.');
  }

  if (extractionType === 'account') {
    return {
      column: txCols.account,
      value: String(payload.selectionValue).trim(),
      label: 'Compte'
    };
  }

  if (extractionType === 'wallet') {
    return {
      column: txCols.wallet,
      value: String(payload.selectionValue).trim(),
      label: 'Wallet'
    };
  }

  throw new Error('Type d’extraction non supporté: ' + extractionType);
}

function buildDateRange_(payload) {
  if (!payload || !payload.mode) {
    throw new Error('Paramètres invalides.');
  }

  if (payload.mode === 'month') {
    if (!payload.month) throw new Error('Mois non renseigné.');

    const parts = payload.month.split('-').map(Number);
    const y = parts[0];
    const m = parts[1];
    if (!y || !m) throw new Error('Format de mois invalide.');

    const start = new Date(y, m - 1, 1, 0, 0, 0, 0);
    const end = new Date(y, m, 0, 23, 59, 59, 999);

    return {
      start: start,
      end: end,
      label: Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM yyyy')
    };
  }

  if (payload.mode === 'range') {
    if (!payload.startDate || !payload.endDate) {
      throw new Error('Plage de dates incomplète.');
    }

    const start = new Date(payload.startDate + 'T00:00:00');
    const end = new Date(payload.endDate + 'T23:59:59');

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('Dates invalides.');
    }
    if (start > end) {
      throw new Error('La date de début doit être <= date de fin.');
    }

    const tz = Session.getScriptTimeZone();
    return {
      start: start,
      end: end,
      label:
        Utilities.formatDate(start, tz, 'dd/MM/yyyy') +
        ' - ' +
        Utilities.formatDate(end, tz, 'dd/MM/yyyy')
    };
  }

  throw new Error('Mode non supporté: ' + payload.mode);
}

function isDateInRange_(d, start, end) {
  return d.getTime() >= start.getTime() && d.getTime() <= end.getTime();
}

function createExtractionOutput_(ss, dict, range, matchedCount, shouldExportPdf, shouldExportSheet, title) {
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const tempSheetName = CFG.SHEET_EXPORT_PREFIX + timestamp;
  const tempSheet = ss.insertSheet(tempSheetName);
  const generatedAt = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  const rows = Object.keys(dict).map(function(tag) {
    return [tag, dict[tag]];
  });

  tempSheet.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(14);
  tempSheet.getRange('A2').setValue('Periode: ' + range.label);
  tempSheet.getRange('A3').setValue('Genere le: ' + generatedAt);
  tempSheet.getRange('A4').setValue('Transactions depenses: ' + matchedCount);
  tempSheet.getRange('A6').setValue('Tag').setFontWeight('bold');
  tempSheet.getRange('B6').setValue('Total depenses').setFontWeight('bold');

  if (rows.length > 0) {
    tempSheet.getRange(7, 1, rows.length, 2).setValues(rows);
    tempSheet.getRange(7, 2, rows.length, 1).setNumberFormat(CFG.CURRENCY_FORMAT);
  }

  tempSheet.autoResizeColumns(1, 2);
  ss.setActiveSheet(tempSheet);
  SpreadsheetApp.flush();
  Utilities.sleep(700);

  const result = {
    exportPdf: false,
    exportSheet: shouldExportSheet,
    period: range.label,
    txCount: matchedCount,
    createdAt: generatedAt
  };

  if (shouldExportSheet) {
    result.sheetName = tempSheet.getName();
    result.sheetUrl = ss.getUrl() + '#gid=' + tempSheet.getSheetId();
  }

  try {
    if (shouldExportPdf) {
      const pdfBlob = exportSheetAsPdfBlob_(ss.getId(), tempSheet.getSheetId(), 'Extraction_' + timestamp + '.pdf');
      const file = savePdfBlob_(ss, pdfBlob);

      result.exportPdf = true;
      result.fileName = file.getName();
      result.fileUrl = file.getUrl();
    }

    return result;
  } finally {
    if (!shouldExportSheet) {
      ss.deleteSheet(tempSheet);
    }
  }
}

function exportSheetAsPdfBlob_(spreadsheetId, gid, fileName) {
  const token = ScriptApp.getOAuthToken();
  const baseUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
  const params = {
    format: 'pdf',
    portrait: true,
    size: 'A4',
    fitw: true,
    gridlines: false,
    printtitle: false,
    sheetnames: false,
    pagenumbers: true,
    fzr: false,
    gid: gid
  };

  const query = Object.keys(params)
    .map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    })
    .join('&');

  const res = UrlFetchApp.fetch(baseUrl + '?' + query, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Echec export PDF (HTTP ' + res.getResponseCode() + ').');
  }

  return res.getBlob().setName(fileName);
}

function savePdfBlob_(ss, blob) {
  if (!CFG.PDF_FOLDER_FALLBACK_TO_ROOT) {
    if (!CFG.PDF_FOLDER_ID) throw new Error('CFG.PDF_FOLDER_ID non défini.');
    return DriveApp.getFolderById(CFG.PDF_FOLDER_ID).createFile(blob);
  }

  const sourceFile = DriveApp.getFileById(ss.getId());
  const parents = sourceFile.getParents();

  if (parents.hasNext()) {
    return parents.next().createFile(blob);
  }

  return DriveApp.getRootFolder().createFile(blob);
}

function normalizeValue_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function matchesFilterValue_(rawValue, expectedValue) {
  return normalizeValue_(rawValue) === normalizeValue_(expectedValue);
}

function maxDefined_(max, value) {
  return value && value > max ? value : max;
}
