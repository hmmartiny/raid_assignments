
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
