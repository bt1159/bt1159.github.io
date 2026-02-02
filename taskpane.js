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

async function createAndInsertImage() {
    try {
        await Excel.run(async (context) => {
            // 1. Get a value from cell A1 to use in our image
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1");
            range.load("values");
            await context.sync();
            
            const cellValue = range.values[0][0];

            // 2. Draw on the Canvas
            const canvas = document.getElementById("myCanvas");
            const ctx = canvas.getContext("2d");
            
            // Background
            ctx.fillStyle = "#4CAF50"; // Excel Green
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            
            // Text
            ctx.fillStyle = "white";
            ctx.font = "20px Arial";
            ctx.fillText("Value: " + cellValue, 10, 50);

            // 3. Convert Canvas to Base64 Image String
            // We strip the header "data:image/png;base64," because Excel just wants the raw code
            const fullDataUrl = canvas.toDataURL("image/png");
            const base64Image = fullDataUrl.replace(/^data:image\/(png|jpg);base64,/, "");

            // 4. Insert the Image into the Sheet
            sheet.shapes.addImage(base64Image);
            
            await context.sync();
            console.log("Image inserted!");
        });
    } catch (error) {
        console.error(error);
    }
}

async function getTableData() {
    try {

        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            let dateTable = sheet.tables.getItemAt(0);

            // Get data from the header row.
            let headerRange = dateTable.getHeaderRowRange().load("values");

            // Get data from the table.
            let bodyRange = dateTable.getDataBodyRange().load("values");

            // Get data from a single column.
            let columnRange = dateTable.columns.getItem("Merchant").getDataBodyRange().load("values");

            // Get data from a single row.
            let rowRange = dateTable.rows.getItemAt(1).load("values");

            // Sync to populate proxy objects with data from Excel.
            await context.sync();

            let headerValues = headerRange.values;
            let bodyValues = bodyRange.values;
            let merchantColumnValues = columnRange.values;
            let secondRowValues = rowRange.values;

            console.log('headerValues: ' + headerValues);
        });
    } catch (error) {
        console.error(error);
    }
}