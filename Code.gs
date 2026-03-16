/**
 * ====== CONFIG GLOBALE ======
 * Modifie uniquement cette section si tes noms/colonnes changent.
 */
const CFG = {
  MENU_NAME: 'Extractions',
  MENU_ITEM_TAGS: 'par Tags',
  MENU_ITEM_ACCOUNTS: 'par Compte',
  MENU_ITEM_WALLETS: 'par Wallet',

  SHEET_TRANSACTIONS: 'Transactions',
  SHEET_DROPDOWN: 'Dropdown',
  SHEET_EXPORT_PREFIX: 'Export_TMP_',

  // Colonnes 1-based
  COL_DATE: 1,              // Transactions ColA
  COL_NAME: 2,              // Transactions ColB
  COL_AMOUNT: 3,            // Transactions ColC
  COL_TAG: 6,               // Transactions ColF
  COL_ACCOUNT: 4,           // Transactions: colonne Compte
  COL_WALLET: 5,            // Transactions: colonne Wallet
  COL_DROPDOWN_TAG: 5,      // Dropdown: colonne Tag
  COL_DROPDOWN_ACCOUNT: 1,  // Dropdown: colonne Compte
  COL_DROPDOWN_WALLET: 3,   // Dropdown: colonne Wallet

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
