// application id: MZQdm0ZZ0EaTYYE8GnuWlqWu2fzgwscDVtIU5cWs
// rest id: XtI77oOBItZjMtCu0lTKVyPHe8vw0gbgGcgTpiVD
// client key: 3LWN9ZFWmoqEzzkzStLbJzrxLXmSUosz54xrMfoD

var Users = {
  index: {
    massimiliano: 1,
    giovanna: 2    
  }
}

var ColumnClasses = {
  expense: "expense",
  credit: "credit",
}

var SpreadsheetConfig = {
  id: "1-Cz8cQVpuMFNMmiMHLzjw-roPPi081NI7teOb7lBuX4",

  column: {
    expense: {
      start: {
        "1": "A",
        "2": "O"
      },
      end: {
        "1": "E",
        "2": "S"
      }      
    },
    credit: {
      start: {
        "1": "G",
        "2": "U"
      },
      end: {
        "1": "J",
        "2": "X"
      }
    },
    result: {
      user: {
        "1": "AC",
        "2": "AD"
      }
//      ,date: "AE"
    },
    parsed: "AI"
  },
  
  rangesToSort: [
    // PagatoMax, quando
    {rangeNotation:"A4:E",sortingColumn:3, rangeTmp:"A4:E"},
    // creditoMax, quando
    {rangeNotation:"G4:J",sortingColumn:3, rangeTmp:"A4:D"},
    // pagatoGio, quando
    {rangeNotation:"O4:S",sortingColumn:3, rangeTmp:"A4:E"},
    // creditoGio, quando
    {rangeNotation:"U4:X",sortingColumn:3, rangeTmp:"A4:D"}
  ],
  sheetTmp: "tmp",
  
  firstEmptyRow: {
    expense: {
      "1": "AJ4",
      "2": "AL4"
    },
    credit: {
      "1": "AK4",
      "2": "AM4"
    }
  }
}
function init() {
  var sheetTmp = getUsingSpreadsheet_().getSheetByName(SpreadsheetConfig.sheetTmp);
  Logger.log(sheetTmp);
  if (sheetTmp==null) {
    sheetTmp = getUsingSpreadsheet_().insertSheet();
    sheetTmp.setName(SpreadsheetConfig.sheetTmp);
  }
}

var TransformFunctions = {
	doNothing: function (value) {
		return value;
	},
	epochToFormattedStringDate: function(epochTime) {
		var date = new Date(epochTime);
		var s = date.getDate()+"/"+(date.getMonth()+1)+"/"+date.getFullYear();
		return s;
	}
}
var ParseConfig = {
  application_id: "MZQdm0ZZ0EaTYYE8GnuWlqWu2fzgwscDVtIU5cWs",
  rest_api_key: "XtI77oOBItZjMtCu0lTKVyPHe8vw0gbgGcgTpiVD",
  
  expenseClass: ColumnClasses.expense,
  creditClass: ColumnClasses.credit,
  
  classes: {
	  expense: {
		  name: "expense",
		  fields: {
			  what: TransformFunctions.doNothing,
			  howmuch: TransformFunctions.doNothing,
			  when: TransformFunctions.epochToFormattedStringDate,
			  where: TransformFunctions.doNothing,
			  username: TransformFunctions.doNothing,
			  sheeted: TransformFunctions.doNothing,
			  objectId: TransformFunctions.doNothing
		  }
	  },
	  credit: {
		  name: "credit",
		  fields: {
			  what: TransformFunctions.doNothing,
			  howmuch: TransformFunctions.doNothing,
			  when: TransformFunctions.epochToFormattedStringDate,
			  username: TransformFunctions.doNothing,
			  sheeted: TransformFunctions.doNothing,
			  objectId: TransformFunctions.doNothing
		  }
	  }
  }
}

function queryToParse() {
  var spreadsheet = SpreadsheetApp.openById(SpreadsheetConfig.id);

  var parseResults = [];
  var parseClassObjectKeys = Object.keys(ParseConfig.classes);
  for (var parseClassObjectKeyIndex in parseClassObjectKeys) {		
    var parseClassObjectKey = parseClassObjectKeys[parseClassObjectKeyIndex];
    var parseClassName = parseClassObjectKey;
    var results = parseQuery(parseClassName, {sheeted:false} , {"order": "when"} );
    for (var r=0; r<results.length; r++) {
      var valuesForTuple = [];
      var parseClassObject = ParseConfig.classes[parseClassObjectKey];
      var fieldKeys = Object.keys(parseClassObject.fields);
      var result = results[r];
      var username = "";
      var parseObjectId = "";
      for (var fieldKeyIndex in fieldKeys) {
        var fieldKey = fieldKeys[fieldKeyIndex];
        var fromParse = result[fieldKey];
        if (fieldKey=="sheeted") { continue; }
        if (fieldKey=="username") { username = fromParse; continue; }
        if (fieldKey == "objectId") { parseObjectId = fromParse; }
        var transformed = parseClassObject.fields[fieldKey](fromParse);
        valuesForTuple.push(transformed);
      }
      
      var userIndex = Users.index[username];
      var columnStartUserCurrent = SpreadsheetConfig.column[parseClassName].start[userIndex];
      var columnEndUserCurrent = SpreadsheetConfig.column[parseClassName].end[userIndex];
      
      
      var cellStoringFirstEmptyRow = SpreadsheetConfig.firstEmptyRow[parseClassName][userIndex];
      var firstEmptyRow = retrieveFirstEmptyRowInColumn_(cellStoringFirstEmptyRow);
      var rangeStart = columnStartUserCurrent+firstEmptyRow;
      var rangeEnd = columnEndUserCurrent+firstEmptyRow;
      var range = rangeStart+":"+rangeEnd;
      
      var tuple = [ valuesForTuple ];
      
      spreadsheet.getRange(range).setValues(tuple);
      storeFirstEmptyRowInColumn_(cellStoringFirstEmptyRow,firstEmptyRow+1);
      Logger.log("storing "+(firstEmptyRow+1)+" in "+cellStoringFirstEmptyRow);
      var parseResult = parseUpdate(parseClassName, parseObjectId, {sheeted: true});
      parseResults.push(parseResult);
      Logger.log(" parseResult: "+JSON.stringify(parseResult));
    }
  }
  sortRanges_(spreadsheet);
  return parseResults;
}


function testInsertToParse() {  
  parseInsert("expense", {username: "massimiliano", what: "spesa", how: "13,45", when: "01/12/2015", where: "pam", sheeted: false});
}


function getUsingSpreadsheet_() {
  return SpreadsheetApp.openById(SpreadsheetConfig.id);
}

function retrieveFirstEmptyRowInColumn_(cellName) {
  var spreadsheet = getUsingSpreadsheet_();
  var value = spreadsheet.getRange(cellName).getValue();
  return value;
}
function storeFirstEmptyRowInColumn_(cellName, value) {
  var spreadsheet = getUsingSpreadsheet_();
  spreadsheet.getRange(cellName).setValue(value);
}



function getFirstEmptyRow_() {
  var ec = getFirstEmptyRowInColumn_(SpreadsheetConfig.column_parsed);
//  Logger.log("ec: "+ec);
  return ec;
}

function sortRanges_(spreadsheet) {
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

function rangeSort_(spreadSheet, sheetNameOnSortTo, tmpSheetName, rangeToSort, rangeTmp, sortingColumn) {
  // This function sorts the Nth column in range "xxN:yyM" in sheet
  // Needed to create a temporary sheet 
  // This is only a temporary solution until the range.sort works for ranges not including column A
  
  var source_sheet = spreadSheet.getSheetByName(sheetNameOnSortTo);
  var target_sheet = spreadSheet.getSheetByName(tmpSheetName);

  target_sheet.clear();
  
  var source_range = source_sheet.getRange(rangeToSort);
  var target_range = target_sheet.getRange(rangeTmp);
  
  //source_range.copyTo(target_range, {contentsOnly: true}); **** use this if format and formulas are not to be copied
  source_range.copyTo(target_range);
  
  target_range.sort(sortingColumn);
  
  //target_range.copyTo(source_range, {contentsOnly: true}); **** use this if format and formulas are not to be copied
  target_range.copyTo(source_range); 
}