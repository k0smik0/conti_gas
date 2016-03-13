function queryFromParseAndUpdateSheet() {
  var spreadsheet = SpreadsheetApp.openById(SpreadsheetConfig.id);
      
  // classExpense: what, howmuch, when, where
  var expenseResults = parseQuery(ParseConfig.expenseClass, {sheeted:false} , {"order": "when"} /*, {username: user}*/); // restore sheeted:false
  for (var r=0; r<expenseResults.length; r++) {
    var sheeted = expenseResults[r].sheeted;
    var expenseCosa = expenseResults[r].what;
    var expenseQuanto = expenseResults[r].howmuch;
    var expenseQuandoEpoch = expenseResults[r].when;
    var expenseDove = expenseResults[r].where;
    var expenseUsername = expenseResults[r].username;      
    var expenseObjectid = expenseResults[r].objectId;
    
    var usernameIndex = Users[expenseUsername];
    var columnStartExpenseUserCurrent = "columnStartExpenseUser"+usernameIndex;
    var columnEndExpenseUserCurrent = "columnEndExpenseUser"+usernameIndex;
    var columnToFind = SpreadsheetConfig[columnStartExpenseUserCurrent];
    var firstEmptyRow = getFirstEmptyRowInColumn_(columnToFind)+1;
    var rangeStart = SpreadsheetConfig[columnStartExpenseUserCurrent]+firstEmptyRow;
    var rangeEnd = SpreadsheetConfig[columnEndExpenseUserCurrent]+firstEmptyRow;
    var range = rangeStart+":"+rangeEnd;
      
    var expenseQuandoDate = new Date(expenseQuandoEpoch);
    var expenseQuando = expenseQuandoDate.getDate()+"/"+(expenseQuandoDate.getMonth()+1)+"/"+expenseQuandoDate.getFullYear();
    var tuple = [ [expenseCosa,expenseQuanto,expenseQuando,expenseDove,expenseObjectid] ];
    spreadsheet.getRange(range).setValues(tuple);
    var parseResult = 
//          parseUpdate(ParseConfig.expenseClass, expenseObjectid, {sheeted: true});
          "";
    Logger.log(" parseResult: "+JSON.stringify(parseResult));
  }
    
  // classCredit_user*: cosa, quanto, quando
  var creditResults = parseQuery(ParseConfig.creditClass, {sheeted: false});
  Logger.log(" creditResults: "+creditResults);
  for (var r=0; r<creditResults.length; r++) {
    var creditSheeted = creditResults[r].sheeted;
    
    var creditCosa = creditResults[r].what;
    var creditQuanto = creditResults[r].howmuch;
    //      var creditQuando = creditResults[r].when;
    var creditQuandoEpoch = creditResults[r].when;
    var creditObjectid = creditResults[r].objectId;
    
    var usernameIndex = Users[expenseUsername];
    var columnStartCreditUserCurrent = "columnStartCreditUser"+usernameIndex;
    var columnEndCreditUserCurrent = "columnEndCreditUser"+usernameIndex;
    var columnToFind = SpreadsheetConfig[columnStartCreditUserCurrent];
    //      var firstEmptyRow = getFirstEmptyRowInColumn_(SpreadsheetConfig[columnStartCreditUserCurrent]);
    var firstEmptyRow = getFirstEmptyRowInColumn_(columnToFind)+1;
    var firstEmptyRow = getFirstEmptyRowInColumn_(columnToFind)+1;
    var rangeStart = SpreadsheetConfig[columnStartCreditUserCurrent]+firstEmptyRow;
    var rangeEnd = SpreadsheetConfig[columnEndCreditUserCurrent]+firstEmptyRow;
    //      var range = columnStartCreditUserCurrent+firstEmptyRow+":"+columnEndCreditUserCurrent+i;
    var range = rangeStart+":"+rangeEnd;
    
    
    
    
    var creditQuandoDate = new Date(creditQuandoEpoch);
    var creditQuando = creditQuandoDate.getDate()+"/"+(creditQuandoDate.getMonth()+1)+"/"+creditQuandoDate.getFullYear();
    var values = [ [creditCosa, creditQuanto, creditQuando, creditObjectid] ];
    Logger.log(values);
    spreadsheet.getRange(range).setValues(values);
    //      var parseResult = parseUpdate(ParseConfig.creditPrefix+user, creditObjectid, {sheeted: true});
    //      Logger.log(user+" parseResult: "+parseResult);
  }

  Logger.log("\n");

  for (var i in SpreadsheetConfig.rangesToSort) {
    var rangeObject = SpreadsheetConfig.rangesToSort[i];
    var sheetToSortName = spreadsheet.getSheets()[0].getName();
    rangeSort_(spreadsheet,
               spreadsheet.getSheets()[0].getName(),
               SpreadsheetConfig.sheetTmp,
               rangeObject.rangeNotation,
               rangeObject.rangeTmp,
               rangeObject.sortingColumn);
  }
}

function findFirstEmptyRowInColumn_(columnName) {
  var spreadsheet = getUsingSpreadsheet_();
  var range = columnName+":"+columnName;
  var column = spreadsheet.getRange(range);
  var values = column.getValues(); // get all data in one call
  var rowIndex = 0;
  while ( values[rowIndex][0] != "" ) {
	  Logger.log("skipping: "+rowIndex+" "+values[rowIndex][0]);
    rowIndex++;
  }
  Logger.log("returning: "+rowIndex+"("+values[rowIndex][0]+")");
  return (rowIndex);
}