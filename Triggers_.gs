// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Triggers_.gs
// ==============
//
// Manage script's triggers

var Triggers_ = (function(ns) {
 
  ns.setup = function() {
    var trigger = Triggers_.getTrigger('TRIGGER_SCRIPT_NAME_')    
    if (trigger === null) {
      Triggers_.createTrigger()        
    } else {
      Triggers_.deleteTrigger()        
    }    
  } 

  ns.createTrigger = function() {
  
    var trigger = Triggers_.getTrigger()
    
    if (trigger !== null) {
      throw new Error('Trying to create a trigger when there is already one: ' + trigger.getUinqueId())
    }
  
    trigger = ScriptApp
      .newTrigger(TRIGGER_SCRIPT_NAME_)
      .timeBased()
      .onMonthDay(25)
      .create()
    
    var triggerId = trigger.getUniqueId()        
    
    PROPERTIES_.setProperty('AUTOMATIC_TRIGGER_ID', triggerId)
    Utils_.alert('New monthly trigger created.', 'Manage Trigger')
    onOpen() 
  }

  ns.getTrigger = function() {
  
    var trigger = null
    
    ScriptApp.getProjectTriggers().forEach(function(nextTrigger) {
      if (nextTrigger.getHandlerFunction() === TRIGGER_SCRIPT_NAME_) {
        if (trigger !== null) {throw new Error(`Multiple ${TRIGGER_SCRIPT_NAME_} triggers`)}
        trigger = nextTrigger
        // log_('Found trigger; ' + trigger.getUniqueId())
      }
    })

    if (trigger === null && !!PROPERTIES_.getProperty('AUTOMATIC_TRIGGER_ID')) {
      throw new Error('Trigger ID stored, but no trigger')
    }
    
    return trigger
  }

  ns.deleteTrigger = function() {  
  
    var trigger = Triggers_.getTrigger()
    
    if (trigger === null) {
      throw new Error('Trying to delete a trigger when there is not one')
    }    
    
    ScriptApp.deleteTrigger(trigger)
    PROPERTIES_.deleteProperty('AUTOMATIC_TRIGGER_ID')
    Utils_.alert('Clock trigger removed, it will no longer run monthly on the 25th.', 'Manage Trigger') 
    onOpen()    
  }

  ns.isTriggerCreated = function() {
    return !!PROPERTIES_.getProperty('AUTOMATIC_TRIGGER_ID')
  }

  return ns

})(Triggers_ || {})
