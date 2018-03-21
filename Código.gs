function verifyEmailAddressInDoc() {
// Abre unos documentos Doc que contienen tablas con nombres y direcciones de correo y verifica que estas direcciones existan en una hoja de cálculo  
 // var document = docId ?
 //DocumentApp.openById(docId) :
  var docUrl= ""; 
  var SPREADSHEET_NAME = "";
 document = DocumentApp.openByUrl(docUrl);
 //document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var search = null;
  var tables = [];
  
  // Extract all the tables inside the Google Document
  while (search = body.findElement(DocumentApp.ElementType.TABLE, search)) {
    tables.push(search.getElement().asTable());
  }
  var ssUrl= 'https://docs.google.com/spreadsheets/d/1pJrEe1DJgYuQCx9SFPU7MWVM0c5VrHX-7XT6_E7m-9A/edit'  
  var SPREADSHEET_NAME = "Correo UNA-UNIEDPA-UNAD 2018";
  var SEARCH_COL_IDX = 3;
  var RETURN_COL_IDX = 0;
  var regex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  var ss = SpreadsheetApp.openByUrl(ssUrl); 
  var values = ss.getSheetByName(SPREADSHEET_NAME).getDataRange().getValues();
  tables.forEach(function (table) {
    var rows = table.getNumRows();
    // Iterate through each row of the table
    for (var r = rows - 1; r >= 0; r--) {
      var found=0;
      var string=table.getCell(r,2).getText().toLowerCase().trim(); //Ver la diposicion de la columna de email
      if (regex.test(string)) {
        for (var i = 0; i < values.length; i++) {
          var email = values[i][SEARCH_COL_IDX].toString().toLowerCase().trim();
          if (email == string.trim())
          { 
            found += 1;
          } 
        }
      }  Logger.log(r+" "+string+" "+found);
    } 
  }
                )}
