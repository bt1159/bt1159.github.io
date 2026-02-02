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

function excelDateToJS(excelDate) {
    // Excel's date system starts 70 years before JS, 
    // and there's a "leap year bug" in Excel's 1900 logic to account for.
    let excelDateFixed = excelDate ?? 0;
    const date = new Date(Math.round((excelDateFixed - 25569) * 86400 * 1000));
    return date;
}

function drawDiamond(ctx, topLX, topLY, size, color) {
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(topLX + (size / 2), topLY);      // Top
    ctx.lineTo(topLX, topLY + (size / 2));      // Right
    ctx.lineTo(topLX + (size / 2), topLY + size);      // Bottom
    ctx.lineTo(topLX + size, topLY + (size / 2));      // Left
    ctx.closePath();
    ctx.fill();
}

function drawHLine(ctx, y, color) {
    ctx.strokeStyle = color;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0,y);
    ctx.lineTo(ctx.canvas.width,y);
    ctx.closePath();
    ctx.stroke();
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

            // // Get data from a single column.
            // let columnRange = dateTable.columns.getItem("Merchant").getDataBodyRange().load("values");

            // Get data from a single row.
            let rowRange = dateTable.rows.getItemAt(1).load("values");

            //Canvas setup
            const canvas = document.getElementById("myCanvas");
            const ctx = canvas.getContext("2d");

            // Sync to populate proxy objects with data from Excel.
            await context.sync();

            // let headerValues = headerRange.values;
            // let bodyValues = bodyRange.values;
            // // let merchantColumnValues = columnRange.values;
            // let secondRowValues = rowRange.values;

            // 1. Get the 1D array of headers from the 2D array
            const headers = headerRange.values[0]; 

            // 2. Use the standard JS indexOf method
            const typeIndex = headers.indexOf("Type");
            const startIndex = headers.indexOf("Start date");
            const endIndex = headers.indexOf("End date");
            const titleIndex = headers.indexOf("Title");
            
            let data = bodyRange.values;

            // Canvas math
            
            const size0 = 20;
            const buffer = size0;
            const lineBuffer = 5;
            const templateHeight = size0 + lineBuffer;

            // Simple Math Setup
            const types = data.map(row => row[typeIndex]);
            const startDates = data.map(row => new Date(excelDateToJS(row[startIndex])));
            const endDates = data.map(row => new Date(excelDateToJS(row[endIndex])));
            const titles = data.map(row => row[titleIndex]);
            const projectStart = new Date(Math.min(...startDates));
            const projectEnd = new Date(Math.max(...endDates));
            const totalDays = (projectEnd - projectStart) / (1000 * 60 * 60 * 24);
            ctx.font = "14px Arial";
            let maxTimestamps = data.map((row, index) => Math.max(
                (startDates[index] instanceof Date && !isNaN(startDates[index])) ? startDates[index].getTime() : 0,
                (endDates[index] instanceof Date && !isNaN(endDates[index])) ? endDates[index].getTime() : 0
            ));
            const availablePixels = data.map((row, index) => {
                if (types[index] == "Activity") {
                    return (canvas.width - ctx.measureText(row[titleIndex]).width - 2 * buffer - 5);
                } else if (types[index] == "Milestone") {
                    return (canvas.width - ctx.measureText(row[titleIndex]).width - 2 * buffer - 5 + size0 / 2);
                } else {
                    return 0;
                }
            });
            const requiredDayWidth = data.map((row, index) => (maxTimestamps[index] - projectStart) / (1000 * 60 * 60 * 24));
            const theoreticalPxPerDay = availablePixels.map((row, index) => row / requiredDayWidth[index]);
            // const pxPerDay = (canvas.width - 2 * buffer) / totalDays;
            const pxPerDay = Math.min(...theoreticalPxPerDay);


            // Draw each task
            data.forEach((row, index) => {
                if (types[index] == "Activity") {
                    const taskStart = new Date(excelDateToJS(row[startIndex]));
                    const taskEnd = new Date(excelDateToJS(row[endIndex]));
                    const duration = (taskEnd - taskStart) / (1000 * 60 * 60 * 24);

                    
                    const x = (taskStart - projectStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                    const y = index * templateHeight; // 30px height per row
                    const width = duration * pxPerDay;

                    drawHLine(ctx,y,"red");
                    // Draw the bar
                    ctx.fillStyle = "#217346"; // Excel Green
                    ctx.fillRect(x, y, width, size0);
                    
                    // Draw the label
                    ctx.fillStyle = "black";
                    ctx.font = "14px Arial";
                    ctx.textBaseline = "middle";
                    const title = row[titleIndex];
                    const metrics = ctx.measureText(title);
                    const textWidth = metrics.width;
                    if (x + width + 5 + textWidth < ctx.canvas.width) {
                        ctx.fillText(title, x + width + 5, y + size0 / 2);
                    } else if (textWidth + 5 < x) {
                        ctx.fillText(title, x - textWidth - 5, y + size0 / 2);
                    } else {
                        ctx.fillText(title, ctx.canvas.width - textWidth, y + size0 / 2);
                    }
                    
                    drawHLine(ctx,y + size0,"blue");

                } else if (types[index] == "Milestone") {
                    const taskStart = new Date(excelDateToJS(row[startIndex]));
                    const size = size0;
                    const x = (taskStart - projectStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer - (size / 2);
                    const y = index * templateHeight; // 30px height per row
                    drawDiamond(ctx,x,y,size,"orange");
                    drawHLine(ctx,y,"red");
                    
                    // Draw the label
                    ctx.fillStyle = "black";
                    ctx.font = "14px Arial";
                    ctx.textBaseline = "middle";
                    const title = row[titleIndex];
                    const metrics = ctx.measureText(title);
                    const textWidth = metrics.width;
                    if (x + size + 5 + textWidth < ctx.canvas.width) {
                        ctx.fillText(title, x + size + 5, y + size0 / 2);
                    } else if (textWidth + 5 < x) {
                        ctx.fillText(title, x - textWidth - 5, y + size0 / 2);
                    } else {
                        ctx.fillText(title, ctx.canvas.width - textWidth, y + size0 / 2);
                    }
                    drawHLine(ctx,y + size0,"blue");

                }
            });

            // Convert to image and push to Excel
            const image = canvas.toDataURL("image/png").replace(/^data:image\/(png|jpg);base64,/, "");
            sheet.shapes.addImage(image);
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}