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