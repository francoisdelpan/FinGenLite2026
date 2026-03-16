function createTagSummaryOutput_(ss, dict, range, matchedCount, shouldExportPdf, shouldExportSheet, title, exportPrefix) {
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const exportName = timestamp + '_' + exportPrefix;
  const tempSheetName = exportName;
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
      const pdfBlob = exportSheetAsPdfBlob_(ss.getId(), tempSheet.getSheetId(), exportName + '.pdf');
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

function createTransactionsListingOutput_(ss, transactions, range, shouldExportPdf, shouldExportSheet, title, exportPrefix) {
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const exportName = timestamp + '_' + exportPrefix;
  const tempSheetName = exportName;
  const tempSheet = ss.insertSheet(tempSheetName);
  const generatedAt = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  const rows = transactions.map(function(tx) {
    return [tx.name, tx.date, tx.tag, tx.amount];
  });

  tempSheet.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(14);
  tempSheet.getRange('A2').setValue('Periode: ' + range.label);
  tempSheet.getRange('A3').setValue('Genere le: ' + generatedAt);
  tempSheet.getRange('A4').setValue('Transactions depenses: ' + transactions.length);
  tempSheet.getRange('A6').setValue('Name').setFontWeight('bold');
  tempSheet.getRange('B6').setValue('Date').setFontWeight('bold');
  tempSheet.getRange('C6').setValue('Tag').setFontWeight('bold');
  tempSheet.getRange('D6').setValue('Amount').setFontWeight('bold');

  if (rows.length > 0) {
    tempSheet.getRange(7, 1, rows.length, 4).setValues(rows);
    tempSheet.getRange(7, 2, rows.length, 1).setNumberFormat('dd/MM/yyyy');
    tempSheet.getRange(7, 4, rows.length, 1).setNumberFormat(CFG.CURRENCY_FORMAT);
  }

  tempSheet.autoResizeColumns(1, 4);
  ss.setActiveSheet(tempSheet);
  SpreadsheetApp.flush();
  Utilities.sleep(700);

  const result = {
    exportPdf: false,
    exportSheet: shouldExportSheet,
    period: range.label,
    txCount: transactions.length,
    createdAt: generatedAt
  };

  if (shouldExportSheet) {
    result.sheetName = tempSheet.getName();
    result.sheetUrl = ss.getUrl() + '#gid=' + tempSheet.getSheetId();
  }

  try {
    if (shouldExportPdf) {
      const pdfBlob = exportSheetAsPdfBlob_(ss.getId(), tempSheet.getSheetId(), exportName + '.pdf');
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
    if (!CFG.PDF_FOLDER_ID) throw new Error('CFG.PDF_FOLDER_ID non defini.');
    return DriveApp.getFolderById(CFG.PDF_FOLDER_ID).createFile(blob);
  }

  const sourceFile = DriveApp.getFileById(ss.getId());
  const parents = sourceFile.getParents();

  if (parents.hasNext()) {
    return parents.next().createFile(blob);
  }

  return DriveApp.getRootFolder().createFile(blob);
}
