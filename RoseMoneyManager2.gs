
/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu('Rose Money Manager') 
    .addItem('💸 Add current transaction', 'addCurrentTransaction')
    .addItem('🕙 Add recurring transactions', 'addRecurringTransactions')
    .addSeparator()

  if (Triggers_.isTriggerCreated()) {
    menu.addItem('🔴 Disable monthly recurring transactions import', "onSetupTriggers")
  } else {
    menu.addItem('🟢 Enable monthly recurring transactions import', "onSetupTriggers")
  }
  menu.addToUi()
}

function addCurrentTransaction() {addTransaction_({name: 'Assets:Current Account',})}

function onSetupTriggers() {Triggers_.setup()}

/**
 * Run on the 25th of each month to add next months upcoming transactions to the 
 * Transactions tab
 */

function addRecurringTransactions() {

  const ss = Utils_.getSpreadsheet()
  const inputSheet = ss.getSheetByName('Recurring Transactions')
  if (!inputSheet) throw new Error('Can not find the "Recurring Transactions" tab')
  const inputData = inputSheet.getDataRange().getValues()
  inputData.shift() // Remove headers

  const outputSheet = ss.getSheetByName('Transactions')

  // Allow for next month being the year after this one
  const today = new Date()
  let nextMonth = today.getMonth() + 1
  const thisYear = (new Date(today.getFullYear(),nextMonth)).getFullYear()
  nextMonth = nextMonth % 12

  inputData.forEach(row => {

    let outputRows = []

    if (!row[0]) return
    const transactionDayOfMonth = row[0].getDate()
    const frequency = row[8]

    if (frequency === 'Once a week') {
      addWeeklyTransactions()
    } else if (frequency === 'Once a calendar month') {
      addMonthlyTransactions()
    } else if (frequency === 'Once a year') {
      addYearlyTransactions()
    } else {
      throw new Error(`Unexpected frequency ${frequency} in Recurring Transactions`)
    }

    return

    // Private Functions
    // -----------------

    function addYearlyTransactions() {
      const transactionMonth = row[0].getMonth()
      if (transactionMonth !== nextMonth) return 
      outputRows = row
      outputRows[0] = new Date(thisYear, nextMonth, transactionDayOfMonth)
      outputRows[6] = 'N' // Reconciled state
      delete outputRows[8] // Remove the freqency field
      outputSheet.appendRow(outputRows)
      log_(`Added Yearly ${JSON.stringify(outputRows)} to "Transactions"`)
    }

    function addWeeklyTransactions() {
      const endOfNextMonth = new Date(today.getFullYear(), nextMonth + 1, 0)
      if (transactionDayOfMonth > 7) throw new Error('The start date for weekly transactions has to be in the first week of the month')
      let nextDate = new Date(today.getFullYear(), nextMonth, transactionDayOfMonth)
      let weekCount = 0
      while (nextDate <= endOfNextMonth) {
        let nextRow = row.slice(0)
        nextRow[0] = nextDate
        nextRow[6] = 'N' // Reconciled state
        delete nextRow[8] // Remove the freqency field
        outputRows.push(nextRow)
        nextDate = new Date(thisYear, nextMonth, transactionDayOfMonth + (++weekCount * 7))
      } 
      outputSheet.getRange(outputSheet.getLastRow(), 1, outputRows.length, outputRows[0].length).setValues(outputRows)
      log_(`Added Weekly ${JSON.stringify(outputRows)} to "Transactions"`)
      debugger
    }

    function addMonthlyTransactions() {
      outputRows = row
      outputRows[0] = new Date(thisYear, nextMonth, transactionDayOfMonth)
      outputRows[6] = 'N' // Reconciled state
      delete outputRows[8] // Remove the freqency field
      outputSheet.appendRow(outputRows)
      log_(`Added Monthly ${JSON.stringify(outputRows)} to "Transactions"`)
    }

  }) // for each input row

} // addRecurringTransactions()

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
 * @param {array} values
 * @param {string} filterWord
 * @param {number} numberOfWords - total number of categories and sub-categories to include
 * @param {string} excludeWords - comma separated string of words to exclude
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

function log_(message) {Utils_.getSpreadsheet().getSheetByName('Log').appendRow([new Date(), message])}
