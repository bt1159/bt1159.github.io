Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
        console.log("Office is ready in: " + info.host);
        // You can also enable buttons here if you started them as disabled
    }
});

async function insertText() {
    try {
        if (Office.context.host === Office.HostType.Excel) {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange("A1");
                range.values = [["Hello from my Add-in!"]];
                
                await context.sync();
                console.log("Text inserted into Excel");
            });
        } 
        else if (Office.context.host === Office.HostType.Word) {
            await Word.run(async (context) => {
                context.document.body.insertText("Hello from my Add-in!", "Start");
                await context.sync();
                console.log("Text inserted into Word");
            });
        }
    } catch (error) {
        console.error("Error: " + error);
    }
}
