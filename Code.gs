const SPREADSHEET_ID = '1129_zCqs6UkxgV4gtApFg1aldv5kqnf9_Cij9t8FEZs';
const EXPENSE_SHEET = 'Expenses';
const CONFIG_SHEET = 'Sites_And_Heads';

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Interior Site Expense Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Automatic Sheets aur Headers setup karne ke liye
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let expSheet = ss.getSheetByName(EXPENSE_SHEET);
  if (!expSheet) {
    expSheet = ss.insertSheet(EXPENSE_SHEET);
    expSheet.appendRow(['ID', 'Date', 'Client Site', 'Major Head', 'Sub-Item / Description', 'Vendor / Labour Name', 'Payment Mode', 'Amount (₹)', 'Timestamp']);
    expSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    expSheet.setFrozenRows(1);
  }

  let cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(CONFIG_SHEET);
    cfgSheet.appendRow(['Client Sites', 'Major Expense Heads']);
    cfgSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    
    const defaultSites = ['Sharma Villa - 3BHK', 'Verma Penthouse', 'Greenwood Office Interior'];
    const defaultHeads = [
      'Carpentry Work',
      'Civil Work',
      'Tile & Flooring Work',
      'Painting & Polish Work',
      'Electrical Work',
      'Plumbing & Sanitary',
      'False Ceiling (POP/Gypsum)',
      'Modular Kitchen & Hardware',
      'Glass & Aluminium Work',
      'Site Misc & Transport'
    ];
    
    const maxLen = Math.max(defaultSites.length, defaultHeads.length);
    for (let i = 0; i < maxLen; i++) {
      cfgSheet.appendRow([defaultSites[i] || '', defaultHeads[i] || '']);
    }
  }
}

// Initial Data (Sites, Heads aur Purane Expenses) fetch karna
function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  const expSheet = ss.getSheetByName(EXPENSE_SHEET);

  const cfgData = cfgSheet.getDataRange().getValues();
  const sites = [];
  const heads = [];
  
  for (let i = 1; i < cfgData.length; i++) {
    if (cfgData[i][0]) sites.push(cfgData[i][0].toString().trim());
    if (cfgData[i][1]) heads.push(cfgData[i][1].toString().trim());
  }

  const expData = expSheet.getDataRange().getValues();
  const expenses = [];
  
  for (let i = 1; i < expData.length; i++) {
    if (expData[i][0]) {
      let rawDate = expData[i][1];
      let formattedDate = rawDate instanceof Date 
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
        : rawDate.toString();

      expenses.push({
        id: expData[i][0],
        date: formattedDate,
        site: expData[i][2],
        head: expData[i][3],
        desc: expData[i][4],
        vendor: expData[i][5],
        mode: expData[i][6],
        amount: Number(expData[i][7]) || 0
      });
    }
  }

  return {
    sites: sites,
    heads: heads,
    expenses: expenses.reverse() // Latest entry sabse upar
  };
}

// Naya Expense Save karna
function addExpense(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const expSheet = ss.getSheetByName(EXPENSE_SHEET);
  
  const id = 'EXP-' + Math.floor(100000 + Math.random() * 900000);
  const timestamp = new Date();
  
  expSheet.appendRow([
    id,
    data.date,
    data.site,
    data.head,
    data.desc,
    data.vendor,
    data.mode,
    Number(data.amount),
    timestamp
  ]);

  return getInitialData();
}

// Nayi Client Site add karna
function addNewSite(siteName) {
  if (!siteName) return getInitialData();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  const lastRow = cfgSheet.getLastRow();
  
  // Khali cell dhundhna Column A mein
  const colA = cfgSheet.getRange(1, 1, lastRow + 1, 1).getValues();
  for (let i = 1; i < colA.length; i++) {
    if (!colA[i][0]) {
      cfgSheet.getRange(i + 1, 1).setValue(siteName.trim());
      return getInitialData();
    }
  }
  cfgSheet.getRange(lastRow + 1, 1).setValue(siteName.trim());
  return getInitialData();
}