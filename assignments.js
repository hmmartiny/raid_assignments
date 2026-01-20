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
    // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AQ40');
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