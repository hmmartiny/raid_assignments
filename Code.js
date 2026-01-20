function postTwinMarks(input) {
  // Define the lookup table
  var lookupTable = {
    "skull": ["B4", "B13"],
    "cross": ["B5", "B10"],
    "square": ["B6", "B8"],
    "moon": ["B7", "B12"],
    "triangle": ["B9", "B11"],
  };

  // Get the current sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the cells to concatenate based on the input
  var cells = lookupTable[input];
  if (!cells) {
    return "That is not a mark used for tanking post twins.";
  }

  // Concatenate the values of the specified cells
  var result = cells.map(function(cell) {
    return sheet.getRange(cell).getValue();
  }).join(" / ");

  return result;
}

function postTwinMarkss(input) {
  // Lookup table
  var lookup = {
    "skull": '=IMAGE("https://i.imgur.com/EvHFvne.png")', // Replace FORMULA_1 with the actual formula
    "cross": '=IMAGE("https://i.imgur.com/6RVhiPo.png")', // Replace FORMULA_1 with the actual formula
    "square": '=IMAGE("https://i.imgur.com/7J91le2.png")', // Replace FORMULA_1 with the actual formula
    "moon": '=IMAGE("https://i.imgur.com/j5Xey4w.png")', // Replace FORMULA_1 with the actual formula
    "triangle": '=IMAGE("https://i.imgur.com/IKy7xxT.png")', // Replace FORMULA_1 with the actual formula
    // Add more mappings here
  };

  // Get the formula to match based on the input
  var formulaToMatch = lookup[input];
  if (!formulaToMatch) {
    return "Invalid input";
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getRange("H4:H22");
  var formulas = range.getFormulas();
  var correspondingCells = [];

  // Loop through the formulas
  for (var i = 0; i < formulas.length; i++) {
    // If the formula matches the formula to match
    if (formulas[i][0] === formulaToMatch) {
      var row = i + 4; // +4 because the range starts at row 4
      var cell = sheet.getRange("B" + row).getValue();
      correspondingCells.push(cell);
    }
  }

  // If there are corresponding cells, return them
  if (correspondingCells.length > 0) {
    return correspondingCells.join(" / ");
  } else {
    return "No matching formulas found";
  }
}

function BGHEX(row, column) {
  var background = SpreadsheetApp.getActive().getDataRange().getCell(row, column).getBackground();
  return background;
}


function ImportAQ40Roster() {
  
  // clear data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A7:H11").clearContent();
  sheet.getRange("A7:H11").setBackground(null); // This clears the background color

  sheet.getRange("A15:I40").clearContent();
  sheet.getRange("A15:I40").setBackground(null);


  // Get the value from cell E1
  var cellValue = SpreadsheetApp.getActiveSpreadsheet().getRange('G5').getValue();
  
  // Build the URL
  var url = 'https://raid-helper.dev/api/raidplan/' + cellValue;
  
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  
  // Initialize an empty array to store the extracted data
  var extractedData = [];

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
  
  // Check if the data has 'raidDrop' property
  if(data.hasOwnProperty('raidDrop')) {
    // Loop through each item in 'raidDrop'
    for(var i = 0; i < data.raidDrop.length; i++) {
      var item = data.raidDrop[i];
      
      // Check if the item has 'name', 'class', and 'spec' properties, and partyId is between 1 and 8
      if(item.hasOwnProperty('name') && item.hasOwnProperty('class') && item.hasOwnProperty('spec') && item.hasOwnProperty('partyId') && item.partyId >= 1 && item.partyId <= 8) {
        
        // Replace 'item.class' with the value from 'specToClass'
        var classValue = specToClass[item.spec] || item.class;
        
        // Push the extracted data to the array
        extractedData.push([item.name, classValue, item.spec, item.partyId, item.slotId]);
      }
    }
  }
  
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Write the extracted data to the sheet starting from cell A6
  //sheet.getRange(6, 1, extractedData.length, extractedData[0].length).setValues(extractedData);
  // Initialize an empty 2D array to represent the table
  var table = Array(5).fill().map(() => Array(8).fill(''));

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
      'Boomkin': '#ffa000',
      '': '#efefef',
  }; 

  // Loop through the extracted data
  for(var i = 0; i < extractedData.length; i++) {
    var item = extractedData[i];
    
    // Get the row and column indices from 'slotId' and 'partyId'
    // Subtract 1 because array indices start from 0
    var rowIndex = item[4] - 1;
    var columnIndex = item[3] - 1;
    
    // Check if the indices are within the table dimensions
    if(rowIndex >= 0 && rowIndex < 5 && columnIndex >= 0 && columnIndex < 8) {
      // Assign the item name to the corresponding cell in the table
      table[rowIndex][columnIndex] = item[0];

      // Get the color for the item's class and spec
      var color = item[1] ? colorMapping[item[1]] : '#efefef';
      
      // Set the background color of the cell
      if(color) {
        sheet.getRange(rowIndex + 7, columnIndex + 1).setBackground(color);
      }
    }
  }

  // Write the table data to the sheet starting from cell A7
  for(var i = 0; i < table.length; i++) {
    for(var j = 0; j < table[i].length; j++) {
      sheet.getRange(i + 7, j + 1).setValue(table[i][j]);
    }
  }
  
  var startColumn = 1;
  Object.keys(colorMapping).forEach(function(key, index) { 
    var classColor = colorMapping[key];
    var tankItems = extractedData.filter(item => item[1] === key);

    if(tankItems.length > 0){
      // Create a 2D array where each sub-array contains one item name
      var tankItemsArray = tankItems.map(item => [item[0]]);

      // Write key 
      sheet.getRange(14, startColumn + index).setValue(key);  
      sheet.getRange(14, startColumn + index).setFontWeight('bold');
      // Write the tank items to column A15 starting from row 15
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setValues(tankItemsArray);  

      // Color background
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setBackground(classColor);
    }
  });

  var aq40Range = sheet.getRange("A7:H11");
  var bwlRange = sheet.getRange("A44:H48");
  var mcRange = sheet.getRange("A52:H56");

  aq40Range.copyTo(bwlRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  aq40Range.copyTo(bwlRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  aq40Range.copyTo(mcRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  aq40Range.copyTo(mcRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function ImportNaxxRoster() {

  // clear data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("M7:T11").clearContent();
  sheet.getRange("M7:T11").setBackground(null); // This clears the background color

  sheet.getRange("M15:T40").clearContent();
  sheet.getRange("M15:T40").setBackground(null);

  // Get the value from cell E1
  var cellValue = SpreadsheetApp.getActiveSpreadsheet().getRange('S5').getValue();
  
  // Build the URL
  var url = 'https://raid-helper.dev/api/raidplan/' + cellValue;
  
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  
  // Initialize an empty array to store the extracted data
  var extractedData = [];

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
  }
  
  // Check if the data has 'raidDrop' property
  if(data.hasOwnProperty('raidDrop')) {
    // Loop through each item in 'raidDrop'
    for(var i = 0; i < data.raidDrop.length; i++) {
      var item = data.raidDrop[i];
      
      // Check if the item has 'name', 'class', and 'spec' properties, and partyId is between 1 and 8
      if(item.hasOwnProperty('name') && item.hasOwnProperty('class') && item.hasOwnProperty('spec') && item.hasOwnProperty('partyId') && item.partyId >= 1 && item.partyId <= 8) {
        // Push the extracted data to the array

        // Replace 'item.class' with the value from 'specToClass'
        var classValue = specToClass[item.spec] || item.class;

        extractedData.push([item.name, classValue, item.spec, item.partyId, item.slotId]);
      }
    }
  }
  
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Write the extracted data to the sheet starting from cell A6
  //sheet.getRange(6, 1, extractedData.length, extractedData[0].length).setValues(extractedData);
  // Initialize an empty 2D array to represent the table
  var table = Array(5).fill().map(() => Array(8).fill(''));

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
    '': '#efefef',
  }; 

  // Loop through the extracted data
  for(var i = 0; i < extractedData.length; i++) {
    var item = extractedData[i];
    
    // Get the row and column indices from 'slotId' and 'partyId'
    // Subtract 1 because array indices start from 0
    var rowIndex = item[4] - 1;
    var columnIndex = item[3] - 1;
    
    // Check if the indices are within the table dimensions
    if(rowIndex >= 0 && rowIndex < 5 && columnIndex >= 0 && columnIndex < 8) {
      // Assign the item name to the corresponding cell in the table
      table[rowIndex][columnIndex] = item[0];

      // Get the color for the item's class and spec
      var color = item[1] ? colorMapping[item[1]] : '#efefef';
      
      // Set the background color of the cell
      if(color) {
        sheet.getRange(rowIndex + 7, columnIndex + 13).setBackground(color);
      }
    }
  }

  // Write the table data to the sheet starting from cell A7
  for(var i = 0; i < table.length; i++) {
    for(var j = 0; j < table[i].length; j++) {
      sheet.getRange(i + 7, j + 13).setValue(table[i][j]);
    }
  }

  var startColumn = 13;
  Object.keys(colorMapping).forEach(function(key, index) { 
    var classColor = colorMapping[key];
    var tankItems = extractedData.filter(item => item[1] === key);

    if(tankItems.length > 0){
      // Create a 2D array where each sub-array contains one item name
      var tankItemsArray = tankItems.map(item => [item[0]]);

      // Write key 
      sheet.getRange(14, startColumn + index).setValue(key);  
      sheet.getRange(14, startColumn + index).setFontWeight('bold');
      // Write the tank items to column A15 starting from row 15
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setValues(tankItemsArray);  

      // Color background
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setBackground(classColor);
    }
  });
}

function AssignAQ40(){
    // Define the source range in Roster!A14:I28
    // Columns are classes
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster').getRange('A15:J35');

    var tankRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B4:B7');
    var warriorRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B8:B22');
    var rogueRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B24:B33');
    var hunterRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B35:B38');
    var warlockRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B40:B44');
    var mageRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B46:B56');
    var priestRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B58:B64');
    var paladinRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B66:B73');
    var druidRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B75:B77');
    var boomkinRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40').getRange('B78:B78');

    // Clear data in the ranges
    tankRange.clearContent();
    warriorRange.clearContent();
    rogueRange.clearContent();
    hunterRange.clearContent();
    warlockRange.clearContent();
    mageRange.clearContent();
    tankRange.clearContent();
    paladinRange.clearContent();
    druidRange.clearContent();
    boomkinRange.clearContent();

    var tanks = []; // column A
    var warriors = []; // column B
    var rogues = []; // column C
    var hunters = []; // column D
    var warlocks = []; // column E
    var mages = []; // column F
    var priests = []; // column G
    var paladins = []; // column H
    var druids = []; // column I
    var boomkins = []; // column J

    // Get all values from the source range at once
    var values = sourceRange.getValues();

    // Loop through each row in the source range
    for (var i = 0; i < values.length; i++) {
        // Add the value in each column to the corresponding list if it's not empty
        if (values[i][0] !== '') tanks.push(values[i][0]); // Column A
        if (values[i][1] !== '') warriors.push(values[i][1]); // Column B
        if (values[i][2] !== '') rogues.push(values[i][2]); // Column C
        if (values[i][3] !== '') hunters.push(values[i][3]); // Column D
        if (values[i][4] !== '') warlocks.push(values[i][4]); // Column E
        if (values[i][5] !== '') mages.push(values[i][5]); // Column F
        if (values[i][6] !== '') priests.push(values[i][6]); // Column G
        if (values[i][7] !== '') paladins.push(values[i][7]); // Column H
        if (values[i][8] !== '') druids.push(values[i][8]); // Column I
        if (values[i][9] !== '') boomkins.push(values[i][9]); // Column J
    }

    // Check if tanks has less than four elements
    while (tanks.length < 4 && warriors.length > 0) {
        // Move an element from warriors to tanks
        var warrior = warriors.shift();
        tanks.push(warrior);
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

    tankRange.setValues(prepareDataForRange(tanks, tankRange));
    warriorRange.setValues(prepareDataForRange(warriors, warriorRange));
    rogueRange.setValues(prepareDataForRange(rogues, rogueRange));
    hunterRange.setValues(prepareDataForRange(hunters, hunterRange));
    warlockRange.setValues(prepareDataForRange(warlocks, warlockRange));
    mageRange.setValues(prepareDataForRange(mages, mageRange));
    priestRange.setValues(prepareDataForRange(priests, priestRange));
    paladinRange.setValues(prepareDataForRange(paladins, paladinRange));
    druidRange.setValues(prepareDataForRange(druids, druidRange));
    boomkinRange.setValues(prepareDataForRange(boomkins, boomkinRange));

    // Update post-twin trash healing marks - since function doesn't automatically update itself.
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40');
    // Priests
    // sheet.getRange('H58').setFormula('=postTwinMarks("square")');
    // sheet.getRange('H59').setFormula('=postTwinMarks("cross")');
    // sheet.getRange('H60').setFormula('=postTwinMarks("square")');
    // sheet.getRange('H61').setFormula('=postTwinMarks("moon")');
    // sheet.getRange('H62').setFormula('=postTwinMarks("triangle")');

    // // Paladins
    // sheet.getRange('H66').setFormula('=postTwinMarks("skull")');
    // sheet.getRange('H67').setFormula('=postTwinMarks("cross")');
    // sheet.getRange('H68').setFormula('=postTwinMarks("square")');
    // sheet.getRange('H69').setFormula('=postTwinMarks("moon")');
    // sheet.getRange('H70').setFormula('=postTwinMarks("triangle")');

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

var tankRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B5:B8');
var warriorRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B9:B23');
var warriorDPSRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B11:B23');
var rogueRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B25:B34');
var hunterRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B36:B39');
var warlockRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B41:B43');
var mageRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B45:B54');
var priestRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B56:B62');
var paladinRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B64:B69');
var druidRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B71:B73');

function AssignNaxx(){
    // Define the source range in Roster!A14:I28
    // Columns are classes
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster').getRange('M15:U35');

    // Clear data in the ranges
    tankRange.clearContent();
    warriorRange.clearContent();
    rogueRange.clearContent();
    hunterRange.clearContent();
    warlockRange.clearContent();
    mageRange.clearContent();
    tankRange.clearContent();
    paladinRange.clearContent();
    druidRange.clearContent();

    var tanks = []; // column A
    var warriors = []; // column B
    var rogues = []; // column C
    var hunters = []; // column D
    var warlocks = []; // column E
    var mages = []; // column F
    var priests = []; // column G
    var paladins = []; // column H
    var druids = []; // column I

    // Get all values from the source range at once
    var values = sourceRange.getValues();

    // Loop through each row in the source range
    for (var i = 0; i < values.length; i++) {
        // Add the value in each column to the corresponding list if it's not empty
        if (values[i][0] !== '') tanks.push(values[i][0]); // Column A
        if (values[i][1] !== '') warriors.push(values[i][1]); // Column B
        if (values[i][2] !== '') rogues.push(values[i][2]); // Column C
        if (values[i][3] !== '') hunters.push(values[i][3]); // Column D
        if (values[i][4] !== '') warlocks.push(values[i][4]); // Column E
        if (values[i][5] !== '') mages.push(values[i][5]); // Column F
        if (values[i][6] !== '') priests.push(values[i][6]); // Column G
        if (values[i][7] !== '') paladins.push(values[i][7]); // Column H
        if (values[i][8] !== '') druids.push(values[i][8]); // Column I
    }

    // Check if tanks has less than four elements
    while (tanks.length < 3 && warriors.length > 0) {
        // Move an element from warriors to tanks
        var warrior = warriors.shift();
        tanks.push(warrior);
    }

    tankRange.setValues(prepareDataForRange(tanks, tankRange));
    warriorRange.setValues(prepareDataForRange(warriors, warriorRange));
    rogueRange.setValues(prepareDataForRange(rogues, rogueRange));
    hunterRange.setValues(prepareDataForRange(hunters, hunterRange));
    warlockRange.setValues(prepareDataForRange(warlocks, warlockRange));
    mageRange.setValues(prepareDataForRange(mages, mageRange));
    priestRange.setValues(prepareDataForRange(priests, priestRange));
    paladinRange.setValues(prepareDataForRange(paladins, paladinRange));
    druidRange.setValues(prepareDataForRange(druids, druidRange));

    // Update names for spore groups, 4HM healers, and KT positions
    GetSporeNames();
    Get4HMHealerNames();
    GetKTNames();
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

function priestPIpositions(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKT = ss.getSheetByName('Saph+KT');
  var sheetNaxx = ss.getSheetByName('Naxx');

  // Priest - PI pairs
  var cellPairs = [
    ['W25', 'W13'], // Naxx!B56
    ['W21', 'W5'],  // Naxx!B57
    ['W27', 'W17'], // Naxx!B58
    ['W26', 'W15'], // Naxx!B59
    ['W22', 'W7'],  // Naxx!B60
    ['W20', 'W19']  // Naxx!B61
  ];

  // 🔁 Clear all source and target cells first
  cellPairs.forEach(pair => {
    sheetKT.getRange(pair[0]).clearContent(); // source cell
    sheetKT.getRange(pair[1]).clearContent(); // target cell
  });

  // Get the relevant Naxx data
  var naxxColumnB = sheetNaxx.getRange('B56:B62').getValues();
  var naxxColumnF = sheetNaxx.getRange('F56:F62').getValues();

  for (var i = 0; i < cellPairs.length; i++) {
    var sourceCell = cellPairs[i][0];
    var targetCell = cellPairs[i][1];

    var sourceValue = naxxColumnB[i][0]; // B56 + i
    sheetKT.getRange(sourceCell).setValue(sourceValue);

    // Filter F56:F62 where B56:B62 == sourceValue
    var filtered = [];
    for (var j = 0; j < naxxColumnB.length; j++) {
      if (naxxColumnB[j][0] === sourceValue) {
        filtered.push([naxxColumnF[j][0]]);
        }
      }
    // Write filtered results to the target cell (vertically)
    if (filtered.length > 0) {
      sheetKT.getRange(targetCell).offset(0, 0, filtered.length, 1).setValues(filtered);
      } else {
        sheetKT.getRange(targetCell).setValue('');
        }
  }
}

function palaPositions(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKT = ss.getSheetByName('Saph+KT');

  var paladinValues = getNonEmptyValuesFromRange(paladinRange);

  // Define the target cells
  var targetCells = ['W23', 'W24', 'W10', 'W11'];

  // First, clear the content in all target cells 
  targetCells.forEach(cell => {
    sheetKT.getRange(cell).clearContent();
    });

  // Place paladins
  for (var i = 0; i < targetCells.length; i++) {
    var value = paladinValues[i] || ''; // Use empty string if not enough values
    sheetKT.getRange(targetCells[i]).setValue(value);
  }
}

function otherRangedPositions(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKT = ss.getSheetByName('Saph+KT');

  // Place hunter(s)
  var hunterValues = getNonEmptyValuesFromRange(hunterRange)
  var hunterCells = ['W6', 'W9'];

  for (var i = 0; i < hunterCells.length; i++){
    var value = hunterValues[i] || '';
    var cell = hunterCells[i];
    sheetKT.getRange(cell).clearContent();
    sheetKT.getRange(cell).setValue(value);
  }

  // Place druids
  var druidValues = getNonEmptyValuesFromRange(druidRange);
  var druidCells = ['W8', 'W16', 'W18'];

  for(var i = 0; i < druidCells.length; i++){
    var value = druidValues[i] || '';
    var cell = druidCells[i];

    sheetKT.getRange(cell).clearContent();
    sheetKT.getRange(cell).setValue(value);
  }

  // Get spots already filled
  var allPositionsRange = sheetKT.getRange('W4:W27')
  var allPositionValues = allPositionsRange.getValues(); // 2D array
  var filledPositions = [];
  for (var i = 0; i < allPositionValues.length; i++){
    if (allPositionValues[i][0] !== '') { // Only non-empty cells
      var cell = allPositionsRange.getCell(i + 1, 1).getA1Notation();
      filledPositions.push([cell, allPositionValues[i][0]]);
    }
  }

  // Place warlocks not already placed
  var warlockValues = getNonEmptyValuesFromRange(warlockRange);
  var warlockCells = ['W4', 'W14', 'W19'];
  var filledValues = filledPositions.map(function(pos) {return pos[1]; })
  for (var i = 0; i < warlockValues.length; i++){
    var warlock = warlockValues[i]

    // Skip if already placed
    if(filledValues.indexOf(warlock) !== -1) continue;

    // Find first empty warlock cell
    for (var j = 0; j < warlockCells.length; j++) {
      var cell = warlockCells[j]

      // Check if filled pos
      var isFilled = filledPositions.some(function(pos) { return pos[0] == cell});

      if (!isFilled) {
        sheetKT.getRange(cell).setValue(warlock);
        filledPositions.push([cell, warlock]); // Update filled list
        break; // Move to next warlock
      }
    }
  }

}

function tankKTPositions(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKT = ss.getSheetByName('Saph+KT');


  // Define the target cells AA17:AA20
  var tanktargetCells = ['AA17', 'AA18', 'AA19', 'AA20'];
  var tankValues = getNonEmptyValuesFromRange(tankRange)

  // Clear content in all target cells
  tanktargetCells.forEach(cell => {
    sheetKT.getRange(cell).clearContent();
  });

  // Place tanks
  for (var i = 0; i < tanktargetCells.length; i++){
    var value = tankValues[i] || ''; // Use empty string if not enough values
    sheetKT.getRange(tanktargetCells[i]).setValue(value);
  }

}

function meleeKTPositons(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKT = ss.getSheetByName('Saph+KT');

  var meleeKTrightCells = ['Y12', 'Y13', 'Y14', 'Y15', 'Y16', 'Y17', 'Y18'];
  var meleeKTleftCells = ['AC12', 'AC13', 'AC14', 'AC15', 'AC16', 'AC17', 'AC18'];
  var meleeKTtop = ['AA6', 'AA7', 'AA8', 'AA9', 'AA10', 'AA11'];

  // Get rogues and fury warriors
  var rogueValues = getNonEmptyValuesFromRange(rogueRange);
  var warriorValues = getNonEmptyValuesFromRange(warriorRange);

  // Clear content in all target cells
  meleeKTtop.forEach(cell => {
    sheetKT.getRange(cell).clearContent();
  });
  meleeKTrightCells.forEach(cell => {
    sheetKT.getRange(cell).clearContent();
  });
  meleeKTleftCells.forEach(cell => {
    sheetKT.getRange(cell).clearContent();
  });

  // Place rogues in meleeKTtop
  for (var i = 0; i < meleeKTtop.length; i++){
    var value = rogueValues[i] || '';
    sheetKT.getRange(meleeKTtop[i]).setValue(value);
  }

  // Find empty spots in meleeKTtop after placing rogues
  var emptyTopSpots = [];
  for (var i = 0; i < meleeKTtop.length; i++) {
    if (!rogueValues[i]) {
      emptyTopSpots.push(i);
    }
  }

  // Distribute warriors equally across the three ranges using round-robin
  var warriorIndex = 0;
  var topIndex = 0;
  var rightIndex = 0;
  var leftIndex = 0;
  
  // Cycle through ranges: right -> left -> top (if spots available)
  while (warriorIndex < warriorValues.length) {
    // Place in right
    if (rightIndex < meleeKTrightCells.length) {
      sheetKT.getRange(meleeKTrightCells[rightIndex]).setValue(warriorValues[warriorIndex]);
      rightIndex++;
      warriorIndex++;
      if (warriorIndex >= warriorValues.length) break;
    }
    
    // Place in left
    if (leftIndex < meleeKTleftCells.length) {
      sheetKT.getRange(meleeKTleftCells[leftIndex]).setValue(warriorValues[warriorIndex]);
      leftIndex++;
      warriorIndex++;
      if (warriorIndex >= warriorValues.length) break;
    }
    
    // Place in top (empty spots only)
    if (topIndex < emptyTopSpots.length) {
      sheetKT.getRange(meleeKTtop[emptyTopSpots[topIndex]]).setValue(warriorValues[warriorIndex]);
      topIndex++;
      warriorIndex++;
      if (warriorIndex >= warriorValues.length) break;
    }
    
    // If all ranges are full, break
    if (rightIndex >= meleeKTrightCells.length && 
        leftIndex >= meleeKTleftCells.length && 
        topIndex >= emptyTopSpots.length) {
      break;
    }
  }

}

function GetKTNames(){
  
  // Update names to put on KT positions
  var sheetKT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Saph+KT');
  var rangeKTMelee = sheetKT.getRange('AF5:AF32');
  var rangeKTRanged = sheetKT.getRange('AG5:AG27');
  var rangeKTHealer = sheetKT.getRange('AH5:AH27');

  // Clear content in those first
  rangeKTMelee.clearContent()
  rangeKTRanged.clearContent()
  rangeKTHealer.clearContent()

  var meleeList = getNonEmptyValuesFromRange(tankRange).concat(getNonEmptyValuesFromRange(warriorRange)).concat(getNonEmptyValuesFromRange(rogueRange));
  var rangedList = getNonEmptyValuesFromRange(hunterRange).concat(getNonEmptyValuesFromRange(warlockRange)).concat(getNonEmptyValuesFromRange(mageRange));
  var healerList = getNonEmptyValuesFromRange(priestRange).concat(getNonEmptyValuesFromRange(paladinRange)).concat(getNonEmptyValuesFromRange(druidRange));

  rangeKTMelee.setValues(prepareDataForRange(meleeList, rangeKTMelee));
  rangeKTRanged.setValues(prepareDataForRange(rangedList, rangeKTRanged));
  rangeKTHealer.setValues(prepareDataForRange(healerList, rangeKTHealer));

  // Clear ranged positions
  var rangedKT = sheetKT.getRange('W4:W27')
  rangedKT.clearContent()

  // Add priest locations + PI caster locations
  priestPIpositions();
  palaPositions();
  
  // Place other classes
  otherRangedPositions();

  // Clear melee positions
  var tankKT = sheetKT.getRange('AA17:AA21')
  var meleeKTright = sheetKT.getRange('Y12:Y18')
  var meleeKTleft = sheetKT.getRange('AC12:AC18')
  var meleeKTtop = sheetKT.getRange('AA6:AA11')

  tankKT.clearContent();
  meleeKTright.clearContent();
  meleeKTleft.clearContent();
  meleeKTtop.clearContent();

  // Place melees around KT
  tankKTPositions();
  meleeKTPositons();

}

function GetSporeNames(){
  // var tankRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B5:78');
  // var warriorRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B8:B22');
  // var rogueRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B24:B33');
  // var hunterRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B35:B38');
  // var warlockRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B40:B42');
  // var mageRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B44:B53');


  // Get values
  var mages = getNonEmptyValuesFromRange(mageRange);
  var warriors = getNonEmptyValuesFromRange(tankRange).concat(getNonEmptyValuesFromRange(warriorRange));
  var rogues = getNonEmptyValuesFromRange(rogueRange);
  var hunters = getNonEmptyValuesFromRange(hunterRange);
  var warlocks = getNonEmptyValuesFromRange(warlockRange);

  var sheetSpore = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SporeGrps');
  var rangeMage = sheetSpore.getRange('B16:B36');
  var rangeWarrior = sheetSpore.getRange('C16:C36');  
  var rangeRogue = sheetSpore.getRange('D16:D36');
  var rangeHunter = sheetSpore.getRange('E16:E36');
  var rangeWarlock = sheetSpore.getRange('F16:F36');

  //  Clear content
  rangeMage.clearContent()
  rangeWarrior.clearContent()
  rangeRogue.clearContent()
  rangeHunter.clearContent()
  rangeWarlock.clearContent()

  // Write
  rangeMage.setValues(prepareDataForRange(mages, rangeMage));
  rangeWarrior.setValues(prepareDataForRange(warriors, rangeWarrior));
  rangeRogue.setValues(prepareDataForRange(rogues, rangeRogue));
  rangeHunter.setValues(prepareDataForRange(hunters, rangeHunter));
  rangeWarlock.setValues(prepareDataForRange(warlocks, rangeWarlock));
}

function place4HMHealers(priests, paladins, druids){

  var sheet4HM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM');

  var thaneHealerCells = ['C21', 'C22', 'C23'];
  var mograineHealerCells = ['O21', 'O22', 'O23'];
  var blaumeuxHealerCells = ['C6', 'C7', 'C8'];
  var zeliakHealerCells = ['O6', 'O7', 'O8'];

  allHealerCells = thaneHealerCells.concat(mograineHealerCells).concat(blaumeuxHealerCells).concat(zeliakHealerCells);

  // Clear all healer cells first
  allHealerCells.forEach(function(cell){
    sheet4HM.getRange(cell).clearContent();
  });

  // Define positions for each class
  palaPositions = [thaneHealerCells[0], mograineHealerCells[1], blaumeuxHealerCells[2], zeliakHealerCells[1]];
  druidPositions = [blaumeuxHealerCells[1]];
  priestPositions = [
    blaumeuxHealerCells[1], // priest from grp 2
    zeliakHealerCells[2], // priest priest from grp 3
    thaneHealerCells[1], // priest from grp 4
    mograineHealerCells[1], // priest from grp 5
    mograineHealerCells[2], // priest from grp 6
    zeliakHealerCells[0], // priest from grp 7
  ];

  extraHealerPositions = [thaneHealerCells[2]]; // extra healer position

  // Place paladins
  for (var i = 0; i < palaPositions.length; i++){
    var value = paladins[i] || '';
    var cell = palaPositions[i];
    sheet4HM.getRange(cell).setValue(value);
  }

  // Place druids
  for (var i = 0; i < druidPositions.length; i++){
    var value = druids[i] || '';
    var cell = druidPositions[i];
    sheet4HM.getRange(cell).setValue(value);
  }

  // Place priests 
  for (var i = 0; i < priestPositions.length; i++){
    var value = priests[i] || '';
    var cell = priestPositions[i];
    sheet4HM.getRange(cell).setValue(value);
  }

  // Place extra healer if any (likely number 12)
  if (priests.length + paladins.length + druids.length > 11){
    var extraHealer = priests[6] || paladins[4] || druids[1] || '';
    var extraCell = extraHealerPositions[0];
    sheet4HM.getRange(extraCell).setValue(extraHealer);
  }
}

function Get4HMHealerNames(){
  // var priestRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B55:B61');
  // var paladinRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B63:B68');
  // var druidRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Naxx').getRange('B70:B72');

  var priests = getNonEmptyValuesFromRange(priestRange);
  var paladins = getNonEmptyValuesFromRange(paladinRange);
  var druids = getNonEmptyValuesFromRange(druidRange);

  // Write
  var rangePriests = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('R6:R12');
  var rangePaladins = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('S6:S12');
  var rangeDruids = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('T6:T12');

  rangePriests.clearContent();
  rangePaladins.clearContent();
  rangeDruids.clearContent();

  rangePriests.setValues(prepareDataForRange(priests, rangePriests));
  rangePaladins.setValues(prepareDataForRange(paladins, rangePaladins));
  rangeDruids.setValues(prepareDataForRange(druids, rangeDruids));

  // Place healers in 4HM healer positions
  place4HMHealers(priests, paladins, druids);
}

function Get4HMMeleeDPS(){

  var warriors = getNonEmptyValuesFromRange(warriorDPSRange).flat();
  var rogues = getNonEmptyValuesFromRange(rogueRange)

  // Write warriors
  var rangeWarriors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('W17:W31');
  rangeWarriors.clearContent();
  rangeWarriors.setValues(prepareDataForRange(warriors, rangeWarriors));

  // Write rogues
  var rangeRogues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('X17:X31');
  rangeRogues.clearContent();
  rangeRogues.setValues(prepareDataForRange(rogues, rangeRogues));
}