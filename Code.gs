
const SCRIPT_NAME = "BudgetTemplate"
const SCRIPT_VERSION = "v1.0"

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rose Money Manager')
    .addItem('Add current transaction', 'addCurrentTransaction')
    .addToUi()    
}

function addCurrentTransaction() {
  addTransaction_({
    name: 'Assets:Current Account',
  })
}

function addTransaction_(transaction, sheet, activate) {
  
  if (sheet === undefined) {
  
    sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Transactions')
  }
      
  if (activate === undefined) {
    activate = true
  }    
      
  var date = transaction.date || new Date()
  date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/YYYY')
  var name = transaction.name || ''
  var number = transaction.number || ''
  var description = transaction.description || ''
  var memo = transaction.memo || ''
  var category = transaction.category || ''
  var reconciled = (transaction.reconciled === undefined) ? 'N' : ''
  var amount = transaction.amount || 0
  
  sheet
    .appendRow([
      date,
      name,
      number,
      description,
      memo,
      category,
      reconciled,
      amount])
      
  if (activate) {   
    sheet
      .getRange(sheet.getLastRow(), 4) // Activate Description field
      .activate()
  }

} // addTransaction_()

/**
 * @customfunction
 * 
 * @param {array} array
 * @param {string} term
 * @return array
*/

function ARRAYFILTER(values, filterWord, numberOfWords = 2, excludeWords = "") {

  const extractedValues = [];
  const lowerCaseExcludeWords = excludeWords.split(",").map(word => word.trim().toLowerCase());

  for (const row of values) {

    const originalValue = row[0];
    const splitValue = originalValue.split(":");

    if (splitValue.length >= 2) {

      const extractedWord = splitValue.slice(0, numberOfWords).join(":")

      if (filterWord && 
          originalValue.toLowerCase().includes(filterWord.toLowerCase()) && 
          !containsExcludedWords(lowerCaseExcludeWords, originalValue)) {

        if (!extractedValues.includes(extractedWord)) {
          extractedValues.push(extractedWord);
        }

      } else if (!filterWord && !lowerCaseExcludeWords.some(word => originalValue.toLowerCase().includes(word))) {

        if (!extractedValues.includes(extractedWord)) {
          extractedValues.push(extractedWord);
        }
      }
    }
  }

  extractedValues.sort();
  return extractedValues;

  // Private Functions
  // -----------------

  function containsExcludedWords(lowerCaseExcludeWords, originalValue) {
    if (lowerCaseExcludeWords[0] === "") return false
    return lowerCaseExcludeWords.some(word => originalValue.toLowerCase().includes(word)) 
  }
}
