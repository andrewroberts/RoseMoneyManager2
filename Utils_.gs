// JSHint - TODO
/* jshint asi: true */
/* jshint esversion: 6 */

(function(){"use strict"})()

// Utils_.gs
// =========
//
// Object template

var Utils_ = (function(ns) {

  /**
   * Get the active spreadsheet, failing that the test one.
   *
   * @return {Spreadsheet} spreadsheet
   */
 
  ns.getSpreadsheet = function() {
  
    var spreadsheet = SpreadsheetApp.getActive()
    
    if (spreadsheet === null) {
      if (!PRODUCTION_VERSION_) {
        spreadsheet = SpreadsheetApp.openById(TEST_SHEET_ID_)
      } else {
        throw new Error('No active spreadsheet')
      }
    }
    
    return spreadsheet
  }
        
  ns.alert = function(message) {
    var spreadsheet = SpreadsheetApp.getActive()
    if (spreadsheet === null) return
    SpreadsheetApp.getUi().alert(message)
    log_(message)
  }

  return ns

})({})
