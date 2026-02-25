
function BGHEX(row, column) {
  var background = SpreadsheetApp.getActive().getDataRange().getCell(row, column).getBackground();
  return background;
}

// Function to convert a 1D array to a 2D array and fill with empty values if needed
function prepareDataForRange(array, range) {
    var numRows = range.getNumRows();
    var data = array.slice(0, numRows).map(function(element) {
        return [element];
    });
    while (data.length < numRows) {
        data.push(['']); // Add an empty row
    }
    return data;
}

function getNonEmptyValuesFromRange(range) {
  var values = range.getValues();
  var nonEmptyValues = [];
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] !== '') {
      nonEmptyValues.push(values[i][0]);
    }
  }
  return nonEmptyValues;
}

// Colors for classes
var colorMapping = {
  'Tank': '#a66d36',
  'Warrior':'#a66d36', 
  'Rogue': '#fff601',
  'Hunter': '#93c47d',
  'Warlock': '#b4a7d6', 
  'Mage': '#6d9eeb',
  'Priest': '#ffffff',
  'Paladin': '#ff9fe7',
  'Druid': '#ffa000',
  'Shaman': '#0070DD',
  '': '#efefef',
};

// Spec to classes
var specToClass = {
  'Protection': 'Tank',
  'Fury': 'Warrior',
  'Combat': 'Rogue',
  'Marksmanship': 'Hunter',
  'Holy': 'Priest',
  'Discipline': 'Priest',
  'Restoration': 'Druid',
  'Holy1': 'Paladin',
  'Destruction': 'Warlock',
  'Fire': 'Mage',
  'Frost': 'Mage',
  'Retribution': 'Paladin',
  'Balance': 'Boomkin'
}

function ShowAllSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    
    allSheets.forEach(function(sheet) {
        if (sheet.isSheetHidden()) {
            sheet.showSheet();
        }
    });
}