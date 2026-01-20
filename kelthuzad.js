function priestKTPIpositions(){
  
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

function palaKTPositions(){
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
  priestKTPIpositions();
  palaKTPositions();
  
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