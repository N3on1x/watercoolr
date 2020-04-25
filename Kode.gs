/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('WaterCoolrApp')
      .addItem('Run', 'main')
      .addItem('Clear', 'removeWCList')
      .addToUi();
}

var ss = SpreadsheetApp.getActiveSpreadsheet();
var empList = ss.getSheetByName("Employees List");
var urlPrefix = ss.getSheetByName("Video Chat URL").getRange('A1').getValue();
var WCListSheetName = "WC List";
var WCList = ss.getSheetByName(WCListSheetName);
var defaultGroupSize = 4;
var employeesCount = getEmployeesCount();
var numOfGroups = calcNumOfGroups();

function getEmployeesCount() {
  var count = empList.getDataRange().getLastRow() - 1; // Subtract header row
  Logger.log("Numer of employees: " + count);
  return count;
}

function calcNumOfGroups() {
  var numOfGroups = parseInt(Math.ceil(employeesCount/defaultGroupSize));
  Logger.log("Number of groups: " + numOfGroups);
  return numOfGroups;
}

function removeWCList() {
  if (WCList == null) {
    Logger.log("WCList already removed");
    return;
  }
  Logger.log("Removing WC List");
  ss.deleteSheet(WCList);
}

function resetWCList() {
  // TODO Instead of delete; clear list if it exists
  removeWCList();
  WCList = ss.insertSheet(WCListSheetName, 3);
}
 

function main() {
  
  Logger.log(urlPrefix);
  
  resetWCList();
  
  // Copy employee list and shuffle the order
  empList.getRange(2,1,employeesCount,1).copyTo(WCList.getRange(2,2,employeesCount,1));
  WCList.getRange(2,2,employeesCount,1).randomize();
  
  // Set vertical alignment to 'middle' for group name and url columns
  WCList.getRange(1,1,WCList.getMaxRows(),1).setVerticalAlignment("middle");
  WCList.getRange(1,3,WCList.getMaxRows(),1).setVerticalAlignment("middle");
  
  var nextRow = 2;
  var i;
  for (i = 0; i < numOfGroups; i++) {
    var currentGroupNumber = i+1;
    WCList.getRange(nextRow,1,4,1).merge().setValue(`WC group ${currentGroupNumber}`);
    WCList.getRange(nextRow,3,4,1).merge().setValue(urlPrefix + `${currentGroupNumber}`);
    nextRow = nextRow + 4;
  }
  WCList.autoResizeColumns(1, 3)
  SpreadsheetApp.getUi()
     .alert('Congratulations! A new WC group mix has been created.');
  
  /* 
  This code shuffles the employee list.
  Commented out after the Range.copyTo(destination) and Range.randomize() was discovered
  
  var employeesList = empList.getRange(2,1,employeesCount,1).getValues(); // returns array of array Array[][]
  var shuffledEmployeesList = shuffle(employeesList);
  WCList.getRange(2,2,employeesCount,1).setValues(shuffledEmployeesList);
  */
}


// Not used anymore, but kept because it's neat

/** 
 * Fisher-Yates Shuffle
 * See: https://bost.ocks.org/mike/shuffle/
 *
 * @param {Array} array
 * @return {Array} shuffledEmployeesList
 */
function shuffle(array) {
  var numRemaining = array.length;
  var index, temp;
  
  while(numRemaining) {
    // Pick random item from unpicked items
    index = Math.floor(Math.random() * numRemaining);
    numRemaining = numRemaining - 1;
    
    temp = array[numRemaining];
    array[numRemaining] = array[index];
    array[index] = temp;
  }
  return array;
}

