// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
    // You can add logic here that runs when the add-in starts
};

function insertText() {
    // Check if we are in Excel or Word and run the appropriate API
    if (Office.context.host.get_name() === 'Excel') {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.getRange("A1").values = [["Hello from my Add-in!"]];
            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    } else if (Office.context.host.get_name() === 'Word') {
        Word.run(function (context) {
            context.document.body.insertText("Hello from my Add-in!", Word.InsertLocation.start);
            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    }
}
