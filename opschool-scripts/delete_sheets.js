function deleteSheets(e) {
    ss = SpreadsheetApp.getActive();
    var sheets = ss.getSheets();
    for (var i = 1; i < sheets.length; i++) {
        var formVar = sheets[i].getName().split(" ")[0];

        if (formVar == "Form") {
            var sheet_name = ss.getSheetByName(sheets[i].getName());

            var formUrl = sheets[i].getFormUrl();
            if (formUrl)
                FormApp.openByUrl(formUrl).removeDestination();
            ss.deleteSheet(sheets[i]);
        }
    }

}