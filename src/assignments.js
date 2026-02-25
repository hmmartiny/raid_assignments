function AssignAQ40(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
    
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
    paladinRange.clearContent();
    druidRange.clearContent();
    boomkinRange.clearContent();

    var tanks = [];
    var warriors = [];
    var rogues = [];
    var hunters = [];
    var warlocks = [];
    var mages = [];
    var priests = [];
    var paladins = [];
    var druids = [];
    var boomkins = [];

    // Get tanks from Group 1 (A7:A9) for AQ40 roster
    var group1Tanks = sheet.getRange('A7:A9').getValues();
    for (var i = 0; i < group1Tanks.length; i++) {
        if (group1Tanks[i][0] !== '') tanks.push(group1Tanks[i][0]);
    }
    
    // Get 4th tank from Group 2 first warrior (B7)
    var group2FirstWarrior = sheet.getRange('B7').getValue();
    if (group2FirstWarrior !== '') {
        tanks.push(group2FirstWarrior);
    }

    // Get the rest of the roster data from the class columns
    var sourceRange = sheet.getRange('A15:J35');
    var values = sourceRange.getValues();

    // Loop through each row in the source range
    for (var i = 0; i < values.length; i++) {
        // Add the value in each column to the corresponding list if it's not empty
        // Skip column A (tanks) since we handled that from groups
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

    // Hide Naxx sheets if H3=TRUE
    ManageSheetVisibility('H3', 'AQ40');
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
    
    // Clear data in the ranges
    tankRange.clearContent();
    warriorRange.clearContent();
    rogueRange.clearContent();
    hunterRange.clearContent();
    warlockRange.clearContent();
    mageRange.clearContent();
    paladinRange.clearContent();
    druidRange.clearContent();

    var tanks = [];
    var warriors = [];
    var rogues = [];
    var hunters = [];
    var warlocks = [];
    var mages = [];
    var priests = [];
    var paladins = [];
    var druids = [];

    // Get tanks from Group 1 (M7:M9) for Naxx roster
    var group1Tanks = sheet.getRange('M7:M9').getValues();
    for (var i = 0; i < group1Tanks.length; i++) {
        if (group1Tanks[i][0] !== '') tanks.push(group1Tanks[i][0]);
    }
    
    // Get 4th tank from Group 2 first warrior (N7)
    var group2FirstWarrior = sheet.getRange('N7').getValue();
    if (group2FirstWarrior !== '') {
        tanks.push(group2FirstWarrior);
    }
    
    // Now get the rest of the warriors starting from B8 onwards
    // and collect all other classes from the appropriate roster range
    var sourceRange = sheet.getRange('M15:U35');
    var values = sourceRange.getValues();

    for (var i = 0; i < values.length; i++) {
        // Skip the first tank column since we handled that separately
        if (values[i][1] !== '') warriors.push(values[i][1]); // Column B (warriors/fury)
        if (values[i][2] !== '') rogues.push(values[i][2]); // Column C
        if (values[i][3] !== '') hunters.push(values[i][3]); // Column D
        if (values[i][4] !== '') warlocks.push(values[i][4]); // Column E
        if (values[i][5] !== '') mages.push(values[i][5]); // Column F
        if (values[i][6] !== '') priests.push(values[i][6]); // Column G
        if (values[i][7] !== '') paladins.push(values[i][7]); // Column H
        if (values[i][8] !== '') druids.push(values[i][8]); // Column I
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

    // Hide AQ40 sheets if T3=TRUE
    ManageSheetVisibility('T3', 'Naxx');
}

function ManageSheetVisibility(checkCell, raidType) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rosterSheet = ss.getSheetByName('Roster');
    var hideSheets = rosterSheet.getRange(checkCell).getValue();
    
    var aq40Sheets = ['AQ40'];
    var naxxSheets = ['Naxx', 'LIP/AOE taunts', 'SporeGrps', '4HM', 'Saph+KT'];
    
    if (hideSheets === true) {
        
        if (raidType === 'AQ40') {
            // Hide Naxx sheets, show AQ40 sheets
            naxxSheets.forEach(function(sheetName) {
                var sheetToHide = ss.getSheetByName(sheetName);
                if (sheetToHide) {
                    sheetToHide.hideSheet();
                }
            });
            
            aq40Sheets.forEach(function(sheetName) {
                var sheetToShow = ss.getSheetByName(sheetName);
                if (sheetToShow && sheetToShow.isSheetHidden()) {
                    sheetToShow.showSheet();
                }
            });
        } else if (raidType === 'Naxx') {
            // Hide AQ40 sheets, show Naxx sheets
            aq40Sheets.forEach(function(sheetName) {
                var sheetToHide = ss.getSheetByName(sheetName);
                if (sheetToHide) {
                    sheetToHide.hideSheet();
                }
            });
            
            naxxSheets.forEach(function(sheetName) {
                var sheetToShow = ss.getSheetByName(sheetName);
                if (sheetToShow && sheetToShow.isSheetHidden()) {
                    sheetToShow.showSheet();
                }
            });
        }
    } else {
        // Show all sheets if the checkbox is FALSE
        var allSheets = aq40Sheets.concat(naxxSheets);
        allSheets.forEach(function(sheetName) {
            var sheetToShow = ss.getSheetByName(sheetName);
            if (sheetToShow && sheetToShow.isSheetHidden()) {
                sheetToShow.showSheet();
            }
        });
    }
}