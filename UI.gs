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
    description: 'Genere un export de toutes les depenses de la periode, regroupees par tag.',
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
    description: 'Filtre les transactions sur un compte precis puis genere leur listing chronologique.',
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
    description: 'Filtre les transactions sur un wallet precis puis genere leur listing chronologique.',
    selectionLabel: 'Wallet',
    selectionPlaceholder: 'Choisis un wallet',
    selectionOptions: getWalletsOptions_(ss)
  });
}

function openExtractionModal_(config) {
  const template = HtmlService.createTemplateFromFile('ExtractionModal');
  template.extractionType = config.extractionType;
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
