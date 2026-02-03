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

function drawVLine(ctx, x, y0, y1, color, width) {
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.beginPath();
    ctx.moveTo(x,y0);
    ctx.lineTo(x,y1);
    ctx.closePath();
    ctx.stroke();
}
function drawCenteredLabel(ctx, text, centerX, centerY, rectWidth, rectHeight, rectColor, strokeColor) {
    // 1. Draw the Rectangle (centered)
    ctx.fillStyle = rectColor;
    ctx.fillRect(centerX - rectWidth / 2, centerY - rectHeight / 2, rectWidth, rectHeight);

    // 2. Draw the Border (Stroke)
    ctx.strokeStyle = strokeColor;
    ctx.lineWidth = 2;
    ctx.strokeRect(centerX - rectWidth / 2, centerY - rectHeight / 2, rectWidth, rectHeight);

    // 2. Configure Text Alignment
    ctx.fillStyle = "white";
    let fontSize = 14;
    ctx.font = fontSize + "px Arial";
    let success = ctx.measureText(text).width < rectWidth;
    while (!success) {
        if (fontSize <= 8) {
            success = true;
        } else {
            fontSize--;
            ctx.font = fontSize + "px Arial";
            success = ctx.measureText(text).width < rectWidth
        }
    }
    if (fontSize > 8) {
        fontSize--;
        ctx.font = fontSize + "px Arial";
    }

    console.log('fontSize: ' + fontSize);
    console.log('ctx.measureText(text).width: ' + ctx.measureText(text).width);
    console.log('rectWidth: ' + rectWidth);
    
    // These two lines are the "magic" for centering
    ctx.textAlign = "center";     // Horizontal center
    ctx.textBaseline = "middle";  // Vertical center

    // 3. Draw the text at the EXACT same center point
    ctx.fillText(text, centerX, centerY);
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
            const size0 = 20;
            const buffer = size0;
            const lineBuffer = 5;
            const templateHeight = size0 + lineBuffer;
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
            const ganttBottom = templateHeight * data.length;

            // Canvas math
            

            // Simple Math Setup
            const types = data.map(row => row[typeIndex]);
            const startDates = data.map(row => new Date(excelDateToJS(row[startIndex])));
            const endDates = data.map(row => new Date(excelDateToJS(row[endIndex])));
            const titles = data.map(row => row[titleIndex]);
            const projectStart = new Date(Math.min(...startDates));
            const projectEnd = new Date(Math.max(...endDates));
            const rangeStart = new Date(projectStart.getFullYear(),projectStart.getMonth(),1);
            // rangeEnd will always be the last day of the month in which projectEnd falls
            const rangeEnd = new Date(projectEnd.getFullYear(), projectEnd.getMonth() + 1, 0);
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
            const requiredDayWidth = data.map((row, index) => (maxTimestamps[index] - rangeStart) / (1000 * 60 * 60 * 24));
            const theoreticalPxPerDay = availablePixels.map((row, index) => row / requiredDayWidth[index]);
            const pxPerDay = Math.min(...theoreticalPxPerDay);

            //I think there is a problem where the pxPerDay gets based on the furthest point, which uses date and title.  The problem is 
            // that this bases things off a date that is not the last of the month.  Won't this make the scale different for the labels
            // at the bottom?

            console.log('projectEnd: ' + projectEnd);
            const yearDiff = rangeEnd.getFullYear() - rangeStart.getFullYear();
            console.log('yearDiff: ' + yearDiff);
            const monthDiff = rangeEnd.getMonth() - rangeStart.getMonth() + 1;
            console.log('monthDiff: ' + monthDiff);
            const noMonths = (yearDiff * 12) + monthDiff;
            console.log('noMonths: ' + noMonths);

            // Draw monthly gridlines and titles at bottom
            for (let m = 0; m <= noMonths; m++) {
                //Draw monthly gridlines
                const month = rangeStart.getMonth() + m;
                const thisDate = new Date(rangeStart.getFullYear(),month,1);
                const x = (thisDate - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                const width = month == 12 ? 2 : 1;
                const color = month == 12 ? "rgb(0, 0, 0)" : "rgb(180, 180, 180)";
                drawVLine(ctx,x,0,ganttBottom, color, width);

                // Draw label at bottom
                if (m < noMonths) {
                    const monthAbbr = thisDate.toLocaleString('en-US', { month: 'short' });
                    const nextMDate = new Date(rangeStart.getFullYear(),month + 1,1);
                    const nextX = (nextMDate - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                    const rectWidth = nextX - x;
                    const rectHeight = size0;
                    const centerX = (nextX + x) / 2;
                    const centerY = ganttBottom + rectHeight / 2;
                    drawCenteredLabel(ctx,monthAbbr, centerX, centerY, rectWidth, rectHeight, "rgb(180,180,180)","rgb(90,90,90)");
                }
            }

            // Draw yearly titles at bottom
            for (let y = 0; y < yearDiff + 1; y++) {
                const yearLabel = projectStart.getFullYear() + y;
                let thisDate;
                let nextDate;
                if (y == 0) {
                    thisDate = rangeStart;
                } else {
                    thisDate = new Date(yearLabel,0,1);
                }
                if (y < yearDiff) {
                    nextDate = new Date(yearLabel + 1,0,1);
                } else {
                    nextDate = new Date(rangeEnd);
                }
                const x = (thisDate - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                const nextX = (nextDate - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                const rectWidth = nextX - x;
                const rectHeight = size0;
                const centerX = (nextX + x) / 2;
                const centerY = ganttBottom + rectHeight + rectHeight / 2;
                drawCenteredLabel(ctx,yearLabel, centerX, centerY, rectWidth, rectHeight, "rgb(180,180,180)","rgb(90,90,90)");
             }
            
            // Draw red line for today
            if ( new Date() > rangeStart && new Date() < rangeEnd) {
                const thisDate = new Date();
                console.log('thisDate: ' + thisDate);
                console.log('thisDate > rangeStart: ' + (thisDate > rangeStart));
                const x = (thisDate - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                const width = 2;
                const color = "rgb(255, 0, 0)";
                drawVLine(ctx,x,0,ganttBottom, color, width);
            } else {
                const thisDate = new Date();
                console.log('thisDate: ' + thisDate);
                console.log('thisDate > rangeStart: ' + (thisDate > rangeStart));
                console.log('rangeStart: ' + rangeStart);
            }


            // Draw each task
            data.forEach((row, index) => {
                if (types[index] == "Activity") {
                    const taskStart = new Date(excelDateToJS(row[startIndex]));
                    const taskEnd = new Date(excelDateToJS(row[endIndex]));
                    const duration = (taskEnd - taskStart) / (1000 * 60 * 60 * 24);

                    
                    const x = (taskStart - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer;
                    const y = index * templateHeight; // 30px height per row
                    const width = duration * pxPerDay;

                    // drawHLine(ctx,y,"red");
                    // Draw the bar
                    ctx.fillStyle = "#217346"; // Excel Green
                    ctx.fillRect(x, y, width, size0);
                    
                    // Draw the label
                    ctx.fillStyle = "black";
                    ctx.font = "14px Arial";
                    ctx.textBaseline = "middle";
                    ctx.textAlign = "left";
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

                } else if (types[index] == "Milestone") {
                    const taskStart = new Date(excelDateToJS(row[startIndex]));
                    const size = size0;
                    const x = (taskStart - rangeStart) / (1000 * 60 * 60 * 24) * pxPerDay + buffer - (size / 2);
                    const y = index * templateHeight; // 30px height per row
                    drawDiamond(ctx,x,y,size,"orange");
                    // drawHLine(ctx,y,"red");
                    
                    // Draw the label
                    ctx.font = "14px Arial";
                    ctx.textBaseline = "middle";
                    ctx.textAlign = "left";
                    const title = row[titleIndex];
                    const metrics = ctx.measureText(title);
                    const textWidth = metrics.width;
                    const textHeight = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
                    const textY = y + size0 / 2;
                    let textX;
                    if (x + size + 5 + textWidth < ctx.canvas.width) {
                        textX = x + size + 5;
                    } else if (textWidth + 5 < x) {
                        textX = x - textWidth - 5;
                    } else {
                        textX = ctx.canvas.width - textWidth;
                    }
                    ctx.fillStyle = "rgba(255,255,255,0.5)";
                    ctx.fillRect(textX, textY - textHeight / 2, textWidth, textHeight);
                    ctx.fillStyle = "rgb(0,0,0)";
                    ctx.fillText(title, textX, textY);

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
