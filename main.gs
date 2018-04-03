function onOpen(){
  main();
}

function main(){
  SpreadsheetApp.getUi()
      .createMenu('Food for mood')
      .addItem('Import', 'openMenuLoadDialog')
      .addToUi();
}

function openMenuLoadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ImportMenu')
     .setWidth(700)
     .setHeight(800);
  SpreadsheetApp.getUi()
     .showModalDialog(html, "Import from text");
}

function fillTable(dishes){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet
    .clearContents()
    .clearFormats()
    .clearNotes();
  
  var hungryWorkers = ["Employee1", "Employee2"];
  var totalWidth = hungryWorkers.length + 4;
  
  createColumns(sheet, hungryWorkers);
  fillRows(sheet, dishes, hungryWorkers, totalWidth);
  fillTotalRow(sheet, dishes, hungryWorkers, totalWidth);
}

function fillRows(sheet, dishes, hungryWorkers, totalWidth){  
  for (var i = 0, dishesBlockRowCounter = 0; i < dishes.length; i++, dishesBlockRowCounter++){
    var currentRow = sheet.getRange(2 + i, 1, 1, totalWidth);
    var rowText = dishes[i];
    
    if (isDishCategoryRow(rowText)){
      fillDishCategoryRow(currentRow, rowText);
      dishesBlockRowCounter = 0;
    }
    else{
      var dishWithPrice = /(.*)\D(\d{2,3})\s*$/.exec(rowText);
      if (dishWithPrice){
        sheet
          .getRange(2 + i, 1)
          .setValue(dishWithPrice[1])
          .setWrap(true);
        sheet
          .getRange(2 + i, 2)
          .setValue(dishWithPrice[2])
      }
      else{
        sheet
          .getRange(2 + i, 1)
          .setValue("Can't parse \"" + rowText + "\"")
      }
      
      sheet
        .getRange(2 + i, 3 + hungryWorkers.length)
        .setFormulaR1C1("=SUM(R[0]C[-" + hungryWorkers.length + "]:R[0]C[-1])")
        .protect();
      sheet
        .getRange(2 + i, 4 + hungryWorkers.length)
        .setFormulaR1C1("=MULTIPLY(R[0]C[-1]; R[0]C[-" + (hungryWorkers.length + 2) + "])")
        .protect();
    }    
    paintRow(currentRow, dishesBlockRowCounter);
  }
}

function fillTotalRow(sheet, dishes, hungryWorkers, totalWidth){
  var totalRowIndex = 2 + dishes.length;
  sheet
    .getRange(totalRowIndex, 1, 1, totalWidth)
    .setBackground("yellow");
  
  for (var col = 0; col < hungryWorkers.length; col ++){
    sheet
      .getRange(totalRowIndex, 3 + col)
      .setFormulaR1C1("=SUMPRODUCT(R[-" + (dishes.length) + "]C[0]:R[-1]C[0];R[-" + (dishes.length) + "]C[-" + (col + 1) + "]:R[-1]C[-" + (col + 1) + "])")
      .setFontWeight("bold")
      .protect();
  }
  sheet
    .getRange(totalRowIndex, 3 + hungryWorkers.length)
    .setValue("Итог")
    .setFontWeight("bold")
    .protect();
  sheet
    .getRange(totalRowIndex, 4 + hungryWorkers.length)
    .setFormulaR1C1("=SUM(R[0]C[-" + (hungryWorkers.length + 2) + "]:R[0]C[-1])")
    .setFontWeight("bold")
    .protect();
}

function fillDishCategoryRow(row, rowText){
  row
  .merge()
  .setValue(rowText)
  .setWrap(true)
  .setFontWeight("bold")
  .setFontSize(14);
}

function paintRow(row, positionInDishBlock){
  var colour = "white";
  if (positionInDishBlock == 0){
    colour = "lightgreen";
  }
  else{
    if (positionInDishBlock % 2 == 0){
      colour = "lightgrey";
    }
  }
  row.setBackground(colour);  
}

function isDishCategoryRow(row){
  return !/\d/.test(row);
}

function createColumns(sheet, hungryWorkers) {
  sheet
    .getRange(1, 1)
    .setValue("Наименование")
    .setFontWeight("bold")
    .protect();
  sheet.setColumnWidth(1, 500);
  sheet
    .getRange(1, 2)
    .setValue("Цена")
    .setFontWeight("bold")
    .protect();
  sheet.setColumnWidth(2, 50);
  for (var i = 0; i < hungryWorkers.length; i++)
  {
    sheet
      .getRange(1, i + 3)
      .setValue(hungryWorkers[i])
      .setBackground("black")
      .setFontColor("white").setWrap(true)
      .protect();
    sheet.setColumnWidth(i + 3, 80);
  }
  sheet
    .getRange(1, 3 + hungryWorkers.length)
    .setValue("Количество")
    .setFontWeight("bold")
    .protect();
  sheet
    .getRange(1, 4 + hungryWorkers.length)
    .setValue("Стоимость")
    .setFontWeight("bold")
    .protect();  
}