function place4HMHealers(priests, paladins, druids){

  Logger.log('Priests: ' + priests);
  Logger.log('Paladins: ' + paladins);
  Logger.log('Druids: ' + druids);

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

  // Define positions for each class (order is mark 1, 2, and 3)
  // Thane: Paladin (MT1), Priest (Grp 3), extra healer
  // Mograine: Priest (Grp 6), Paladin (MT2), Priest (Grp 5)
  // Blaumeux: Druid, Priest (Grp 2), Paladin (MT3)
  // Zeliak: Paladin (MT4), Priest (Grp 7), Priest (Grp 4)

  palaPositions = [
    thaneHealerCells[0], // Paladin MT1 -> thane mark 1
    mograineHealerCells[1], // Paladin MT2 -> mograine mark 2
    blaumeuxHealerCells[2],  // Paladin MT3 -> blaumeux mark 3
    zeliakHealerCells[0]]; // Paladin MT4 -> zeliak mark 1
  
  priestPositions = [
    blaumeuxHealerCells[1], // priest from grp 2 -> blaumeux mark 2
    thaneHealerCells[1], // priest from grp 3 -> thane mark 2
    zeliakHealerCells[2], // // priest from grp 4 -> zeliak mark 3
    mograineHealerCells[2], // priest from grp 5 -> mograine mark 3
    mograineHealerCells[0], // priest from grp 6 -> mograine mark 1
    zeliakHealerCells[1], // // priest from grp 7 -> zeliak mark 2
  ];

  druidPositions = [blaumeuxHealerCells[0]];
  extraHealerPositions = [thaneHealerCells[2]]; // extra healer position

  // Place paladins
  for (var i = 0; i < palaPositions.length; i++){
    var value = paladins[i] || '';
    var cell = palaPositions[i];
    sheet4HM.getRange(cell).setValue(value);
  }

  // Place priests 
  for (var i = 0; i < priestPositions.length; i++){
    var value = priests[i] || '';
    var cell = priestPositions[i];
    sheet4HM.getRange(cell).setValue(value);
  }
  
  // Place druids
  for (var i = 0; i < druidPositions.length; i++){
    var value = druids[i] || '';
    var cell = druidPositions[i];
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
  Logger.log("Hello from Get4HMHealerNames");

  var rangePriests = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('R6:R12');
  var rangePaladins = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('S6:S12');
  var rangeDruids = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4HM').getRange('T6:T12');

  rangePriests.clearContent();
  rangePaladins.clearContent();
  rangeDruids.clearContent();

  var priests = getNonEmptyValuesFromRange(priestRange);
  var paladins = getNonEmptyValuesFromRange(paladinRange);
  var druids = getNonEmptyValuesFromRange(druidRange);

  Logger.log('Priests: ' + priests);
  Logger.log('Paladins: ' + paladins);
  Logger.log('Druids: ' + druids);

  // Write
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