function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('广告数据优化')
       .addItem('CPS_calculate', 'CPS_calculate')
       .addItem('CPO_plot', 'createPivotTable')
       .addItem('删除重复项', 'writeArrayToColumn')
       .addItem('I类', 'cat1Arr')
       .addItem('非I类', 'non1CatArr')
       .addItem('II类', 'cat2Arr')
       .addItem('III类', 'cat3Arr')
       .addItem('data_clean', 'data_clean')
       .addToUi();
 }

function CPS_calculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActive();
  var input = Browser.inputBox('search terms');
  var included = input.split(" ")[0];
  var excluded = input.split(" ")[1];
  
  var width = sheet.getLastColumn();
  var length = sheet.getLastRow();
  var range = 'A1:H' + length;
  
  // filter
  var filter_data = sheet.getDataRange().getValues();
  
  var included_list = included.split(',');
  for (var j=0;j<included_list.length;j++){
    filter_data = filter_data.filter(function (d){return d[4].indexOf(included_list[j]) !==-1});
  }
  
  if (excluded){
    var excluded_list = excluded.split(',');
    for (var i=0;i<excluded_list.length;i++){
      filter_data = filter_data.filter(function (d){return d[4].indexOf(excluded_list[i]) ===-1});
    }
  }
  
  var spend_sum = 0;
  var order_sum = 0;
  var clicks_sum = 0;
  for (a=0; a<filter_data.length; a++){
    spend_sum += filter_data[a][6]
    order_sum += filter_data[a][7]
    clicks_sum += filter_data[a][5]
  }
  var CPO = spend_sum / order_sum;
  var CR = order_sum / clicks_sum;
  
  // CPO expression
  sheet.getRange('K1').setValue('contains');
  sheet.getRange('L1').setValue('CPS');
  sheet.getRange('M1').setValue('clicks_sum');
  sheet.getRange('N1').setValue('CR');
  sheet.getRange('O1').setValue('excluded');
  sheet.getRange(1, 11, 1, 5).setBackground('Yellow').setFontWeight("bold");
  
  var Kvals = sheet.getRange("K1:K").getValues();
  var Klast = Kvals.filter(String).length + 1;
  var Kcell = 'K' + Klast;
  var Lcell = 'L' + Klast;
  var Mcell = 'M' + Klast;
  var Ncell = 'N' + Klast;
  var Ocell = 'O' + Klast;
  var Kcellnew = 'K' + (Klast + 1);
  
  sheet.getRange(Kcell).setValue(included)
  sheet.getRange(Lcell).setValue(CPO);
  sheet.getRange(Mcell).setValue(clicks_sum);
  sheet.getRange(Ncell).setValue(CR);
  sheet.getRange(Ocell).setValue(excluded);
  sheet.getRange('L:L').activate();
  sheet.getActiveRangeList().setNumberFormat('#,##0.00')
  sheet.getRange('N:N').activate();
  sheet.getActiveRangeList().setNumberFormat('#,##0.00%')
  sheet.getRange(Kcellnew).activate();
};

function data_clean(){
  data_clean1();
  data_clean2();
}

function data_clean1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:C').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1:T2000').activate();
  spreadsheet.getRange('D1:W2000').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('F:F').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H:I').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('F:F').activate();
  spreadsheet.getRange('G:G').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('G1:Q2000').activate();
  spreadsheet.getRange('J1:T2000').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('H:K').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('P:Q').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('L:M').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H1').activate();
  spreadsheet.getRange('N1:O2000').moveTo(spreadsheet.getActiveRange());
  testReplaceInSheet()
};


function data_clean2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J1').activate();
  spreadsheet.getCurrentCell().setValue('order');
  spreadsheet.getRange('L2').activate();
  spreadsheet.getCurrentCell().setFormula('=H2+I2');
  spreadsheet.getRange('J2').activate();
  spreadsheet.getCurrentCell().setFormula('=H2+I2');
  spreadsheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('J1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('L1').activate();
  spreadsheet.getRange('J:J').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('H:J').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H1').activate();
  spreadsheet.getRange('L:L').moveTo(spreadsheet.getActiveRange());
  // spreadsheet.getRange('I1').setValue('CPO');
  // spreadsheet.getRange('I2').setFormula('=G2/H2').autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
};

function testReplaceInSheet(){
    var sheet = SpreadsheetApp.getActiveSheet()
    replaceInSheet(sheet,'$ ','');
}

function replaceInSheet(sheet, to_replace, replace_with) {
  //get the current data range values as an array
  var values = sheet.getDataRange().getValues();

  //loop over the rows in the array
  for(var row in values){

    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value){
      return original_value.toString().replace(to_replace,replace_with);
    });

    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }

  //write the updated values to the sheet
  sheet.getDataRange().setValues(values);
}


function createPivotTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  // The name of the sheet containing the data you want to put in a table.
  var sheetName = "Data";
  var sheetId = sheet.getSheetId();
  
  var pivotTableParams = {};
  
  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = {
   sheetId: sheetId
  };
  
  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{
    sourceColumnOffset: 2,
    sortOrder: "ASCENDING"
  }];
  
  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [{
    summarizeFunction: "SUM",
    sourceColumnOffset: 5,
  },{
    summarizeFunction: "SUM",
    sourceColumnOffset: 6,
  },{
    summarizeFunction: "SUM",
    sourceColumnOffset: 7,
  }
  ];
    
  // Create a new sheet which will contain our Pivot Table
  if (ss.getSheetByName('plot')) {
    ss.deleteSheet(ss.getSheetByName('plot'));
  };
  var pivotTableSheet = ss.insertSheet('plot');
  var pivotTableSheetId = pivotTableSheet.getSheetId();
  
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": pivotTableSheetId,
        "rowIndex": 0,
        "columnIndex": 0
      },
      "fields": "pivotTable"
    }
  };

  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
  ss.getRange('E1').setValue('CPO');
  ss.getRange('E2').activate().setFormula('=if(D2>0,C2/D2,0)');
  ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getRange('E:E').setNumberFormat('#,##0.00');
  scatter_plot();
}

function scatter_plot() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var pivot = spreadsheet.getSheetByName('plot');
  var data = spreadsheet.getSheetByName('Data');
  var row = pivot.getLastRow();
  var max_cpo = pivot.getRange(row+1, 5).setFormula('=max(E2:E' + row + ')').getValue();
  
  chart = pivot.newChart()
  .asScatterChart()
  .addRange(spreadsheet.getRange('plot!B1:B' + row))
  .addRange(pivot.getRange(1, 5, row, 5))
  .addRange(pivot.getRange(1, 1, row ,1))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('curveType', 'none')
  .setOption('legend.position', 'right')
  .setOption('domainAxis.direction', 1)
  .setOption('title', 'CPS distribution')
  .setOption('treatLabelsAsText', false)
  .setXAxisTitle('Clicks')
  .setOption('series.0.hasAnnotations', true)
  .setOption('series.0.dataLabel', 'custom')
  .setOption('series.0.pointSize', 7)
  .setOption('series.0.labelInLegend', 'CPS')
  .setOption('vAxes.0.viewWindow.max', max_cpo * 1.2)
  .setPosition(1, 1, 57, 104)
  .build();
  data.insertChart(chart);
  data.activate();
};

function setFilter() {
  // get 1 D array of selection
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  var arr = selection.getActiveRange().getValues();
  var arr_result = [];
  for (var i=0; i<arr.length; i++){
    if (arr[i][0].length > 0){
      arr_result.push(arr[i][0]);
    }
  }
  return arr_result; 
}

function removeDuplicateUsingFilter(){
  // remove duplicates
  var arr = setFilter()
  var unique_array = arr.filter(function(elem, index, self) {
    return index == self.indexOf(elem);
  });
  return unique_array 
}

function writeArrayToColumn() {
  // turn 1D array to 2D array
  var input_column = Browser.inputBox('which column, please input integer.');
  var Sheet = SpreadsheetApp.getActiveSheet();
  var array = removeDuplicateUsingFilter()
  var arr = array.map(function (el) {return [el];});
  var range = Sheet.getRange(2, input_column, arr.length);
  // set values
  if (input_column == 2){
    Sheet.getRange('B1').setValue('nonDuplicate');
    Sheet.getRange('B1').setFontWeight('bold');
    Sheet.getRange('B1').setBackground('yellow');
   }
  range.setValues(arr);
};

function cat1Contains() {
  // K列数据， I类包含词
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('K2').activate();
  var arr1 = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  return arr1
};

function cat3Contains() {
  // M列数据， III类包含词
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('M2').activate();
  var arr3 = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  return arr3
};

function cat1Arr() {
  // I类词
  var arr1 = cat1Contains();
  var sheet = SpreadsheetApp.getActiveSheet();
  var arrCat1 = [['start']];
  
  sheet.getRange('C1').setValue('I类');
  sheet.getRange('C1').setFontWeight('bold');
  sheet.getRange('C1').setBackground('yellow');
  
  sheet.getRange('B2').activate();
  // arrND is arr-non-duplicate
  var arrND = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  arrND.sort();
  
  //for (var i=0; i<arr1.length; i++){
    //arrCat1 = arrND.filter(function(item) { return (item[0].indexOf(arr1[i][0]) != -1)})
  //}
  
  for (var i=0; i<arrND.length; i++){
    for (var j=0; j<arr1.length; j++){
      if (arrND[i][0].indexOf(arr1[j][0]) !== -1 && arrCat1[arrCat1.length-1][0] !== arrND[i][0]){
        arrCat1.push(arrND[i]);
      }
    }
  }
  
  arrCat1.shift();
  // paste
  var range = sheet.getRange(2, 3, arrCat1.length);
  range.setValues(arrCat1);
  return arrND
}

function non1CatArr(){
  // 非I类词
  var sheet = SpreadsheetApp.getActiveSheet();
  var arrND = cat1Arr();
  var arrCat1 = cat1Contains()
  
  sheet.getRange('D1').setValue('非I类');
  sheet.getRange('D1').setFontWeight('bold');
  sheet.getRange('D1').setBackground('yellow');

  for (var i=0; i<arrCat1.length; i++){
    arrND = arrND.filter(function(item) { return item[0].indexOf(arrCat1[i][0]) == -1 });
  }
  
  var range = sheet.getRange(2, 4, arrND.length);
  range.setValues(arrND);
  return arrND
}
   

function cat2Arr(){
  // II类词，非III类词
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('F2').activate();
  // arrND is arr-non-duplicate
  var arr3 = sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  // var arr3 = cat3Contains();
  var arr = non1CatArr();
  var arrNew = [['start']];
  
  sheet.getRange('E1').setValue('II类');
  sheet.getRange('E1').setFontWeight('bold');
  sheet.getRange('E1').setBackground('yellow');
  
  for (var i=0; i<arr3.length; i++){
    arr = arr.filter(function(item) { return item[0].indexOf(arr3[i][0]) == -1 })
  }
  
    // paste value
  var range = sheet.getRange(2, 5, arr.length);
  range.setValues(arr);
}

function cat3Arr(){
  // III类词
  var sheet = SpreadsheetApp.getActiveSheet();
  var arr3 = cat3Contains();
  var arr = non1CatArr();
  var arrNew = [['start']];
  
  /*
  for (var i=0; i<arr3.length; i++){
    arr = arr.filter(function(item) { return item[0].indexOf(arr3[i][0]) != -1 })
  }
  Logger.log(arr);
  */
  
  sheet.getRange('F1').setValue('III类');
  sheet.getRange('F1').setFontWeight('bold');
  sheet.getRange('F1').setBackground('yellow');
  
  for (var i=0; i<arr.length; i++){
    for (var j=0; j<arr3.length; j++){
      var input = arr3[j][0];
      if (input.indexOf(',') !== -1){
        var included_list = input.split(',');
        var keyword_1 = included_list[0];
        var keyword_2 = included_list[1];
        if (arr[i][0].indexOf(keyword_1) !== -1 && arr[i][0].indexOf(keyword_2) !== -1 && arrNew[arrNew.length-1][0] !== arr[i][0]){
          arrNew.push(arr[i]);
        }
      }
      else{
        if (arr[i][0].indexOf(arr3[j][0]) !== -1 && arrNew[arrNew.length-1][0] !== arr[i][0]){
          arrNew.push(arr[i]);
        }
      }
    }
  }
  arrNew.shift();
  var range = sheet.getRange(2, 6, arrNew.length);
  range.setValues(arrNew);
}


