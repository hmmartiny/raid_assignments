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
