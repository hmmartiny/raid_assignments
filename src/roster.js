/**
 * Imports and organizes an AQ40 raid roster from an external API.
 * 
 * This function fetches raid roster data from the raid-helper.dev API using a raid plan ID
 * stored in cell G5. It parses the response and extracts player information (name, class, spec, 
 * party ID, and slot ID). The function then:
 * 
 * 1. Clears existing roster data from the sheet
 * 2. Filters valid players (those in parties 1-8 with required properties)
 * 3. Maps player specs to classes using the specToClass lookup table
 * 4. Populates a 5x8 table (rows 7-11, columns A-H) with player names positioned by party and slot
 * 5. Colors cells based on class using the colorMapping lookup table
 * 6. Creates a class summary list (starting row 14) with players grouped by class
 * 7. Copies the organized roster to BWL and MC raid sections for easy viewing
 * 
 * @requires specToClass - A mapping object that converts spec names to class names
 * @requires colorMapping - A mapping object that associates class names with hex color codes
 * @requires SpreadsheetApp - Google Apps Script Spreadsheet service
 * @requires UrlFetchApp - Google Apps Script URL fetch service
 * @requires Logger - Google Apps Script Logger service
 * 
 * @returns {void}
 */
function ImportAQ40Roster() {

  Logger.log("Started AQ40 roster import")
  
  // clear data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A7:H11").clearContent();
  sheet.getRange("A7:H11").setBackground(null); // This clears the background color

  sheet.getRange("A15:J40").clearContent();
  sheet.getRange("A15:J40").setBackground(null);


  // Get the value from cell E1
  var cellValue = SpreadsheetApp.getActiveSpreadsheet().getRange('G5').getValue();
  
  // Build the URL
  var url = 'https://raid-helper.dev/api/raidplan/' + cellValue;
  
  Logger.log("Fetching data from URL: " + url);

  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  Logger.log("Data fetched: " + JSON.stringify(data).substring(0, 100) + "...");
  
  // Initialize an empty array to store the extracted data
  var extractedData = [];
  
  // Check if the data has 'slots' property (new API response)
  if(data.hasOwnProperty('slots')) {
    // Loop through each item in 'slots'
    for(var i = 0; i < data.slots.length; i++) {
      var item = data.slots[i];
      Logger.log("i:" + i);

      Logger.log("Processing item: " + JSON.stringify(item));
      
      // Check if the item has 'name', 'className', and 'specName' properties
      if(item.hasOwnProperty('name') && item.hasOwnProperty('className') && item.hasOwnProperty('specName')) {
        Logger.log("Valid item found: " + JSON.stringify(item));
        // Replace 'item.className' with the value from 'specToClass'
        var classValue = specToClass[item.specName] || item.className;

        // Prefer explicit partyId/slotId if present, otherwise derive from slotNumber (1-40)
        var partyId = item.hasOwnProperty('partyId') ? parseInt(item.partyId, 10) : null;
        var slotId = item.hasOwnProperty('slotId') ? parseInt(item.slotId, 10) : null;
        if(!partyId && item.hasOwnProperty('groupNumber')) {
          partyId = parseInt(item.groupNumber, 10);
        }
        if(!slotId && item.hasOwnProperty('slotNumber')) {
          slotId = parseInt(item.slotNumber, 10);
        }
        if(!partyId || !slotId) {
          var slotNumber = item.hasOwnProperty('slotNumber') ? parseInt(item.slotNumber, 10) : null;
          if(slotNumber) {
            partyId = Math.floor((slotNumber - 1) / 5) + 1;
            slotId = ((slotNumber - 1) % 5) + 1;
          }
        }

        // Push the extracted data to the array
        if(partyId && slotId) {
          extractedData.push([item.name, classValue, item.specName, partyId, slotId]);
        }
      }
    }
  }

  Logger.log("Extracted data length: " + extractedData.length);
  
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
  
  // Check if the data has 'slots' property (new API response)
  if(data.hasOwnProperty('slots')) {
    // Loop through each item in 'slots'
    for(var i = 0; i < data.slots.length; i++) {
      var item = data.slots[i];
      
      // Check if the item has 'name', 'className', and 'specName' properties
      if(item.hasOwnProperty('name') && item.hasOwnProperty('className') && item.hasOwnProperty('specName')) {
        // Replace 'item.className' with the value from 'specToClass'
        var classValue = specToClass[item.specName] || item.className;

        // Prefer explicit partyId/slotId if present, otherwise derive from slotNumber (1-40)
        var partyId = item.hasOwnProperty('partyId') ? parseInt(item.partyId, 10) : null;
        var slotId = item.hasOwnProperty('slotId') ? parseInt(item.slotId, 10) : null;
        if(!partyId && item.hasOwnProperty('groupNumber')) {
          partyId = parseInt(item.groupNumber, 10);
        }
        if(!slotId && item.hasOwnProperty('slotNumber')) {
          slotId = parseInt(item.slotNumber, 10);
        }
        if(!partyId || !slotId) {
          var slotNumber = item.hasOwnProperty('slotNumber') ? parseInt(item.slotNumber, 10) : null;
          if(slotNumber) {
            partyId = Math.floor((slotNumber - 1) / 5) + 1;
            slotId = ((slotNumber - 1) % 5) + 1;
          }
        }

        if(partyId && slotId) {
          extractedData.push([item.name, classValue, item.specName, partyId, slotId]);
        }
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
  var furyGroupOrder = [2, 4, 5, 3, 6, 7];
  var furyGroupPriority = {};
  for (var g = 0; g < furyGroupOrder.length; g++) {
    furyGroupPriority[furyGroupOrder[g]] = g;
  }

  Object.keys(colorMapping).forEach(function(key, index) { 
    var classColor = colorMapping[key];
    var tankItems = extractedData.filter(function(item) { return item[1] === key; });

    if (key === 'Warrior') {
      tankItems.sort(function(a, b) {
        var aGroup = a[3];
        var bGroup = b[3];

        var aPriority = furyGroupPriority.hasOwnProperty(aGroup) ? furyGroupPriority[aGroup] : 999;
        var bPriority = furyGroupPriority.hasOwnProperty(bGroup) ? furyGroupPriority[bGroup] : 999;

        if (aPriority !== bPriority) return aPriority - bPriority;

        // If both are outside priority list, keep deterministic group order
        if (aGroup !== bGroup) return aGroup - bGroup;

        // Within same group: slot 1..5
        return a[4] - b[4];
      });
    }

    if (tankItems.length > 0) {
      var tankItemsArray = tankItems.map(function(item) { return [item[0]]; });

      sheet.getRange(14, startColumn + index).setValue(key);
      sheet.getRange(14, startColumn + index).setFontWeight('bold');
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setValues(tankItemsArray);
      sheet.getRange(15, startColumn + index, tankItemsArray.length, tankItemsArray[0].length).setBackground(classColor);
    }
  });
}
