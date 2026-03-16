function runExtraction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extractionType = resolveExtractionType_(payload);
  const shouldExportPdf = !payload || payload.exportPdf !== false;
  const shouldExportSheet = !payload || payload.exportSheet !== false;

  if (!shouldExportPdf && !shouldExportSheet) {
    throw new Error('Choisis au moins un export: PDF ou Sheet.');
  }

  const range = buildDateRange_(payload);
  const txSheet = ss.getSheetByName(CFG.SHEET_TRANSACTIONS);
  if (!txSheet) throw new Error('Sheet introuvable: ' + CFG.SHEET_TRANSACTIONS);

  const txCols = getTransactionColumns_(txSheet);
  const filter = buildFilterConfig_(extractionType, payload, txCols);
  const lastRow = txSheet.getLastRow();

  if (lastRow <= CFG.HEADER_ROW) {
    throw new Error('Aucune transaction a lire dans ' + CFG.SHEET_TRANSACTIONS);
  }

  const dataWidth = [txCols.date, txCols.name, txCols.amount, txCols.tag, filter.column].reduce(maxDefined_, 0);
  const data = txSheet
    .getRange(CFG.HEADER_ROW + 1, 1, lastRow - CFG.HEADER_ROW, dataWidth)
    .getValues();

  const transactions = [];

  data.forEach(function(row) {
    const rawDate = row[txCols.date - 1];
    const rawName = row[txCols.name - 1];
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

    transactions.push({
      name: String(rawName || '').trim(),
      date: new Date(rawDate.getTime()),
      tag: normalizeTagValue_(rawTag),
      amount: amount
    });
  });

  if (extractionType === 'tag') {
    return buildTagsExtractionResult_(
      ss,
      transactions,
      range,
      shouldExportPdf,
      shouldExportSheet
    );
  }

  return buildTransactionsExtractionResult_(
    ss,
    transactions,
    range,
    shouldExportPdf,
    shouldExportSheet,
    filter
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
  getDropdownValues_(ss, 'tag').forEach(function(tag) {
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

  values.forEach(function(row) {
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
    name: CFG.COL_NAME,
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

  throw new Error('Type de dropdown non supporte: ' + type);
}

function getColumnIndex_(sheet, aliases, fallbackColumn) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn <= 0) {
    throw new Error('Aucune colonne trouvee dans ' + sheet.getName());
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
    throw new Error('Selection manquante.');
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

  throw new Error('Type d extraction non supporte: ' + extractionType);
}

function buildTagsExtractionResult_(ss, transactions, range, shouldExportPdf, shouldExportSheet) {
  const tagsDict = buildTagsDict_(ss);

  transactions.forEach(function(tx) {
    const tag = tx.tag in tagsDict ? tx.tag : CFG.NULL_TAG_LABEL;
    tagsDict[tag] += tx.amount;
  });

  return createTagSummaryOutput_(
    ss,
    tagsDict,
    range,
    transactions.length,
    shouldExportPdf,
    shouldExportSheet,
    'Extraction depenses par tag',
    'EXT-BY-TAG'
  );
}

function buildTransactionsExtractionResult_(ss, transactions, range, shouldExportPdf, shouldExportSheet, filter) {
  transactions.sort(function(a, b) {
    return a.date.getTime() - b.date.getTime();
  });

  return createTransactionsListingOutput_(
    ss,
    transactions,
    range,
    shouldExportPdf,
    shouldExportSheet,
    'Extraction transactions - ' + filter.label + ': ' + filter.value,
    filter.label === 'Compte' ? 'EXT-BY-ACC' : 'EXT-BY-WAL'
  );
}

function buildDateRange_(payload) {
  if (!payload || !payload.mode) {
    throw new Error('Parametres invalides.');
  }

  if (payload.mode === 'month') {
    if (!payload.month) throw new Error('Mois non renseigne.');

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
      throw new Error('Plage de dates incomplete.');
    }

    const start = new Date(payload.startDate + 'T00:00:00');
    const end = new Date(payload.endDate + 'T23:59:59');

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('Dates invalides.');
    }
    if (start > end) {
      throw new Error('La date de debut doit etre <= date de fin.');
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

  throw new Error('Mode non supporte: ' + payload.mode);
}

function isDateInRange_(d, start, end) {
  return d.getTime() >= start.getTime() && d.getTime() <= end.getTime();
}

function normalizeTagValue_(rawTag) {
  const tag = String(rawTag || '').trim();
  return tag || CFG.NULL_TAG_LABEL;
}
