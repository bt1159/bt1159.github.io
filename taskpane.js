// Fixes needed
// - Set font sizes to be programmatic and related to templateHeight.
// - Listeners
//  - The listener for changes to the table contents is working!   basically...
//    One thing is that I need to tweak the chart creation function so that, if
//    it is being re-created, it stays in the same place and size.
// - Currently, if a "Milestone" row has something in the End date column, it is
//   just ignored.  I could add some sort of note (not an error) that informs
//   users about it.  That way, they are not confused.
// - Add a test to even check if there is a table with the right name.  Or, if I
//   am ignoring table name, make sure that there is more than zero tables.  If
//   more than 1, give some sort of a warning.
// - Make sure to close listeners somewhere.
// - What happens if the user deletes the table.  I don't want an error.
// - I should try to speed things up.  I could: ask Gemini what probably takes
//   the most time, or ask Gemini how to tell how long things take.
// - I have an order problem with checking if Excel is in edit mode and
//   listening for changes.  If I make a change and then quickly begin to edit
//   another cell, the edited data causes a fire of the function.  It then sees
//   that Excel is in edit mode and hits that logic/loop.  When I tested it,
//   after finishing the second change, it corrected itself, but I still got
//   multiple errors.  I don't think it is robust at the moment.  2) This time,
//   I edited a cell and immediately went into edit mode in a different cell,
//   BUT, I didn't edit the second cell.  While in edit mode, my code was
//   looping and waiting in 1 second increments waiting for the user to leave
//   edit mode.  But then, after I did, it failed saying "No existing Gantt
//   group found to reformat."   That makes no sense.  If I then click "Reformat
//   chart to new size", it gives that same error.  If I click "Create chart"
//   everything is ok again.  For some reason, when doing what I did, it
//   resulted in all the shapes being there but not grouped.
// - Make sure the overwrite of currentHash, or whatever it's called, happens
//   after the final await context.sync().  This should minimize the chance that
//   the hash is updated but the chart is not (due to an error).
// - Close toast any time the script finishes successfully.
// - Put a global lock in place. Also, I will add a variant of the killing the
//   wait loop. Specifically, I will use a variable (i.e., a tag) to record
//   whether the script has been triggered automatically or by the user clicking
//   that original "create chart" button. If the former, edit mode will kill the
//   script. If the latter, edit mode should show a toast so the user knows that
//   their chart will not be created until the hit enter/tab and click the
//   button again.  Also, I need to change the listener to the entire worksheet
//   instead of just the table.  That way, if the user edits the table and then
//   immediately goes into edit mode in a cell outside the table, it is safe to
//   kill the first run.  When the user hits enter, even though they are outside
//   the table, it will still trigger a hash check.
// - BEFORE ANYTHING ELSE, I NEED TO RESTRUCTURE THE FLOW OF THIS SCRIPT.  IT IS
//   A MESS AND IT IS TERRIBLY CONFUSING.

// Notes
// - I am going to implement a listening approach that checks for changes to the
//   data.  But, when the data has changed, it simply reruns the entire fuction.
//   It's possible that a faster approach would be to see what has changed and
//   devise a way to only update the specific shapes that need it (changing the
//   text inside a textbox if the new length results in the same labelLayout,
//   changing just the length of a bar if the end date changed, moving some
//   shapes down and creating some new ones if a new activity was added in the
//   middle).  That would require, however, a much more robust way of keeping
//   track of the shapes and tying them back to the specific data they come
//   from.  Currently, looking up those shapes would be nearly impossible and
//   certainly take far too long. For simplicity, I am going to make a smart
//   process to look only for meaningful changes, but if one is found, I am
//   going to recreate the entire table.
// - Since you're redrawing a chart, if your "start date" column contains a mix
//   of actual Excel dates and just text strings, JSON.stringify might be
//   slightly inconsistent. If you notice the hash changing when nothing
//   actually changed, you may want to wrap the value in String(val).trim()
//   inside the normalizedData step to ensure the hash is seeing identical
//   strings.

/**
 * Called by Excel when the add-in is loaded (i.e., when the taskbar is opened).
 * I don't know how this would change if I started using the shared runtime (I
 * think that's what it's called).  Put things here that need to run at the very
 * beginning.
 *
 * One thing this is often used for is interacting with taskbar html elements
 * to, for instance, add listeners, add onclick callbacks, etc.
 *
 * @param {OnReadyResult} - An object with two properties: host: HostType and platform: PlatformType
 */
Office.onReady(async function (info) {
  if (info.host === Office.HostType.Excel) {
    console.log("Office is ready in: " + info.host);

    const createChartButton = document.getElementById("create-chart-btn");
    const resetButton = document.getElementById("reset-btn");
    const inputColorButton = document.getElementById("input-color");
    const widthInput = document.getElementById("input-width");
    const heightInput = document.getElementById("input-height");
    const hexColorLabel = document.getElementById("color-hex-label");
    const includeFYInput = document.getElementById("includeFY");

    createChartButton.onclick = () => {
      // Grab values and enforce defaults if blank
      let valWidth =
        widthInput.value === "" ? helpers.chartConfiguration.defaultChartWidth : Number(widthInput.value);
      let valHeight =
        heightInput.value === ""
          ? helpers.chartConfiguration.defaultTemplateHeight
          : Number(heightInput.value);

      // Enforce Integers and "0 or Higher" rule Math.round handles decimals;
      // Math.max(0, ...) handles negatives
      const finalWidth = Math.max(0, Math.round(valWidth));
      const finalHeight = Math.max(0, Math.round(valHeight));

      // Update the UI fields to show the "cleaned" numbers to the user
      widthInput.value = finalWidth;
      heightInput.value = finalHeight;

      const selectedColor = document.getElementById("input-color").value;
      const includeFY = includeFYInput.checked;

      createGanttChart({
        chartWidth: finalWidth,
        templateHeight: finalHeight,
        includeFY: includeFY,
        selectedColor: selectedColor,
      });
    };

    resetButton.onclick = () => {
      widthInput.value = helpers.chartConfiguration.defaultChartWidth;
      heightInput.value = helpers.chartConfiguration.defaultTemplateHeight;
      inputColorButton.value = "#ff7777";
      hexColorLabel.innerText = "#ff7777";
    };

    inputColorButton.oninput = (e) => {
      hexColorLabel.innerText = e.target.value;
    };

    try {
      await Excel.run(async (context) => {
        const table = context.workbook.tables.getItemAt(0);
        table.onChanged.add(debouncedUpdate);
        await context.sync();
        debouncedUpdate();
      });
    } catch (error) {
      console.warn("Could not find a table. Listener not attached.", error);
    }
  }
});

const debouncedUpdate = (function (func, delay) {
  let timeoutId;
  return (...args) => {
    if (timeoutId) clearTimeout(timeoutId);
    timeoutId = setTimeout(() => {
      func.apply(null, args);
    }, delay);
  };
})(checkTableAndRedraw, 500);

async function checkTableAndRedraw() {
  try {
    await Excel.run(async (context) => {
      console.log("got into await Excel.run() in checkTableAndRedraw");
      const table = context.workbook.tables.getItemAt(0);
      const range = table.getRange().load("values");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();

      const length = shapes.items.length;

      console.log(
        "passed first context.sync() in checkTableAndRedraw, and length is: ",
        length,
      );
      const newHash = await helpers.getTableFingerprint(range.values);
      const oldHash = localStorage.getItem("tableHash");

      if (newHash !== oldHash) {
        await handleReformatButtonClick();

        console.log(
          "passed await handleReformatButtonClick() in checkTableAndRedraw",
        );
        localStorage.setItem("tableHash", newHash);
      } else {
        console.log(
          "hashes are equal.  This means table data is equal, i.e., no data has changed.  Script will now exit",
        );
      }
    });
  } catch (error) {
    console.log("Something went wrong and was caught by checkTableAndRedraw");
    console.error("FULL ERROR OBJECT:", JSON.stringify(error));
    console.error("ERROR CODE:", error.code);
    if (error.code === "InvalidOperationInCellEditMode") {
      console.log("Excel is busy (cell edit). Retrying in 1 second...");
      // If busy, wait 1 second and try one more time automatically
      setTimeout(checkTableAndRedraw, 1000);
    } else {
      GanttErrorHandler.handle(error);
    }
  }
}

async function handleReformatButtonClick() {
  const geo = await getExistingChartGeometry();

  console.log(
    "passed await getExistingChartGeometry() in handleReformatButtonClick",
  );
  if (geo) {
    await createGanttChart({
      chartWidth: geo.width,
      chartHeight: geo.height,
      chartTop: geo.top,
      chartLeft: geo.left,
    });

    console.log("passed await createGanttChart() in handleReformatButtonClick");
  }
}

async function getExistingChartGeometry() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes.load(
      "items/name, items/top, items/left, items/width, items/height",
    );

    await context.sync();

    const ganttShapeGroup = shapes.items.find((s) =>
      s.name.startsWith("GanttShapeGroup"),
    );
    // GanttShapeGroup

    if (!ganttShapeGroup) {
      console.error("No existing Gantt group found to reformat.");
      return null;
    }

    // Return the live dimensions the user manually set
    return {
      top: ganttShapeGroup.top,
      left: ganttShapeGroup.left,
      width: ganttShapeGroup.width,
      height: ganttShapeGroup.height,
    };
  });
}

/**
 * The function called by the taskbar button.  This function removes all
 * previously created GanttChart shapes, creates new shapes, and groups them.
 * * @param {Object} options - The options to provide the function
 * @param {number} [options.chartWidth] - width given by taskbar as the desired
 * width of the created chart, which will mean the width of the background
 * rectangle
 * @param {number} [options.chartHeight] - Total height of the chart.  This
 * particularly used when replacing an already made chart
 * @param {number} [options.chartLeft] - Global X value for chart.  This
 * particularly used when replacing an already made chart
 * @param {number} [options.chartTop] - Global Y value for chart.  This
 * particularly used when replacing an already made chart
 * @param {number} [options.templateHeight] - height given by taskbar for each
 * Gantt Chart bar.  This is particularly used when creating a chart from
 * scratch.
 * @param {boolean} [options.includeFY] - color given by taskbar for Gantt
 * Chart bars
 * @param {string} [options.selectedColor] - color given by taskbar for Gantt
 * Chart bars
 */
async function createGanttChart({
  chartWidth,
  chartHeight,
  chartLeft,
  chartTop,
  templateHeight,
  includeFY,
  selectedColor,
} = {}) {
  if (!chartHeight && !templateHeight)
    throw "improper height info sent to chart function";
  /**
   * Rounds the input chartWidth in order to make further rounding logic simpler.
   */
  const safeChartWidth = Math.round(chartWidth);
  /**
   * Nearly the same as chartWidth except that it is reduced by the portion of
   * the border around the chart that overlaps the shape itself AND BY the
   * padding set manually in this file.  That makes this the maximum width that
   * any chart contents should have if they span the entire chart.
   * @type {number}
   */
  const chartInternalSafeWidth = safeChartWidth - 2 * helpers.chartConfiguration.localXOffset;
  const barLabelMarginPtLeft = 0;
  const barLabelMarginPtRight = 0;
  const barLabelMarginPtTop = 0;
  const barLabelMarginPtBottom = 0;
  /** A width added to the calculated width of text elements in the create label
   * functions.  This makes it less likely than the width calculated using the
   * html canvas is too small for the actual text box Excel.Shape created. */
  const safetyPxLabel = 8;
  /**
   * Controls whether Gantt bar labels are bold or not
   * @type {boolean}
   */
  const barLabelBold = false;
  /**
   * Controls whether axis labels are bold or not
   * @type {boolean}
   */
  const axisLabelBold = false;
  /**
   * Font used for all labels
   * @type {string}
   */
  const fontName = "Arial";

  if (!selectedColor) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items/name, items/type");
      await context.sync();

      const ganttGroup = shapes.items.find(
        (shape) => shape.name === "GanttShapeGroup",
      );
      if (!ganttGroup) {
        console.log("No shapes found in group. Default color will be used.");
        selectedColor = null;
      } else if (ganttGroup.type !== Excel.ShapeType.group) {
        console.log(
          "A shape was found with the right name, but it wasn't a group.  It was: ",
          ganttGroup.type,
        );
        selectedColor = null;
      } else {
        const groupShapes = ganttGroup.group.shapes;
        groupShapes.load("items/name");

        await context.sync();

        const firstBar = groupShapes.items.find((shape) =>
          shape.name.startsWith("GanttShape_Bar"),
        );

        if (firstBar) {
          firstBar.load("fill/foregroundColor");
          await context.sync();
          selectedColor = firstBar.fill.foregroundColor;
        } else {
          selectedColor = null;
        }
      }
    });
  }

  /**
   * The color to make Gantt bars and milestones.  The default here is
   * overwritten by what the user selects in the taskbar.
   * @type {string}
   */
  const defaultColor = selectedColor || "#ff7f7f";
  /**
   * Defines the global x value for chart's localX = 0.  Used in the functions
   * that actually create the shapes in Excel to convert locally calculated x
   * positions to actual x values in Excel.
   * @type {number}
   */
  let anchorLeft;
  /**
   * Defines the global y value for chart's localY = 0.  Used in the functions
   * that actually create the shapes in Excel to convert locally calculated y
   * positions to actual y values in Excel.
   * @type {number}
   */
  let anchorTop;

  const overlay = document.getElementById("loading-overlay");
  overlay.style.display = "flex";

  /**
   * Creates a simple rectangle shape without any text.  Enter locations
   * relative to top-left anchor point.  Leave fillColor or borderSize null for
   * no fill or border, respectively.
   * @param {Excel.ShapeCollection} shapes - ShapeCollection from sheet, where
   * this shaped will be added
   * @param {number} left - in localX
   * @param {number} top - in localY
   * @param {number} width
   * @param {number} height
   * @param {string} fillColor - string converted to color object.  Leave null
   * for no fill.
   * @param {number} fillTransparency - From 0 to 1, inclusive.  0 is opaque.
   * @param {string} nameSuffix - string to be added between "GanttShape_" and
   * the index
   * @param {string} borderColor - string converted to color object.  Ignored if
   * borderSize is null.
   * @param {number} borderSize - leave null for no border
   * @returns {Excel.Shape}
   */
  function createShapeRect(
    shapes,
    left,
    top,
    width,
    height,
    fillColor,
    fillTransparency,
    nameSuffix,
    borderColor,
    borderSize,
  ) {
    // Ensure we are sending absolute numbers, not strings or undefined
    const safeLeft = (Number(left) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;
    const safeTop = (Number(top) || 0) + helpers.chartConfiguration.localYOffset + anchorTop;
    const safeWidth = Math.max(Number(width), 1); // Width can't be 0
    const safeHeight = Number(height) || 1;
    const safeTransparency = Math.min(
      Math.max(Number(fillTransparency) || 0, 0),
      1,
    );

    const shape = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    shape.left = safeLeft;
    shape.top = safeTop;
    shape.width = safeWidth;
    shape.height = safeHeight;
    if (fillColor) {
      shape.fill.setSolidColor(fillColor);
      shape.fill.transparency = safeTransparency;
    } else {
      shape.fill.clear();
    }

    if (Number(borderSize)) {
      shape.lineFormat.weight = borderSize;
      shape.lineFormat.color = borderColor || "black";
    } else {
      shape.lineFormat.visible = false;
    }
    helpers.ganttManager.nameAndTrack(shape, nameSuffix);
    return shape;
  }

  /**
   * Creates a simple diamond shape without any text.  Enter locations relative
   * to top-left anchor point.  height is applied to both height & width, and
   * are measured point-to-point.  Leave fillColor or borderSize null for no
   * fill or border, respectively.
   * @param {Excel.ShapeCollection} shapes - ShapeCollection from sheet, where
   * this shaped will be added
   * @param {number} left - localX
   * @param {number} top - localY
   * @param {number} height
   * @param {string} fillColor - string converted to color object.  Leave null
   * for no fill.
   * @param {number} fillTransparency - From 0 to 1, inclusive.  0 is opaque.
   * @param {string} nameSuffix - string to be added between "GanttShape_" and
   * the index
   * @param {string} borderColor - string converted to color object.  Ignored if
   * borderSize is null.
   * @param {number} borderSize - leave null for no border
   * @returns {Excel.Shape}
   */
  function createShapeDiamond(
    shapes,
    left,
    top,
    height,
    fillColor,
    fillTransparency,
    nameSuffix,
    borderColor,
    borderSize,
  ) {
    // Native diamond fills the bounding box tip-to-tip
    const shape = shapes.addGeometricShape(Excel.GeometricShapeType.diamond);

    // To keep it proportional (a "square" diamond), set width = height
    const safeSize = Math.max(Math.round(height) || 0, 1);
    const safeLeft =
      (Number(left) || 0) - safeSize / 2 + helpers.chartConfiguration.localXOffset + anchorLeft;
    const safeTop = (Number(top) || 0) + helpers.chartConfiguration.localYOffset + anchorTop;

    shape.left = safeLeft;
    shape.top = safeTop;
    shape.width = safeSize;
    shape.height = safeSize;
    shape.lockAspectRatio = true;

    // Style
    if (fillColor) {
      shape.fill.setSolidColor(fillColor);
      shape.fill.transparency = Math.min(
        Math.max(Number(fillTransparency) || 0, 0),
        1,
      );
    } else {
      shape.fill.clear();
    }

    if (Number(borderSize)) {
      shape.lineFormat.weight = borderSize;
      shape.lineFormat.color = borderColor || "black";
    } else {
      shape.lineFormat.visible = false;
    }

    helpers.ganttManager.nameAndTrack(shape, nameSuffix);
    return shape;
  }

  /**
   * Creates a basic label.  Primarily intended for axis categories or other
   * labels that will be positioned on their own.  For bar labels, use
   * createBarLabel().
   * @param {Excel.ShapeCollection} shapes - ShapeCollection from sheet, where
   * this shaped will be added
   * @param {string} text - text to put inside label
   * @param {number} left - localX
   * @param {number} top - localY
   * @param {number} width
   * @param {number} height
   * @param {number} fontSizePx - font size, in pixels
   * @param {boolean} bold - whether the label should be bold or standard
   * @param {string} fillColor - string converted to color object.  Leave null
   * for no fill.
   * @param {number} fillTransparency - From 0 to 1, inclusive.  0 is opaque.
   * @param {string} nameSuffix - string to be added between "GanttShape_" and
   * the index
   * @param {string} borderColor - string converted to color object.  Ignored if
   * borderSize is null.
   * @param {number} borderSize - leave null for no border
   * @returns {Excel.Shape}
   */
  function createSimpleLabel(
    shapes,
    text,
    left,
    top,
    width,
    height,
    fontSizePx,
    bold,
    fillColor,
    fillTransparency,
    nameSuffix,
    borderColor,
    borderSize,
  ) {
    const safeLeft = (Number(left) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;
    const safeTop = (Number(top) || 0) + helpers.chartConfiguration.localYOffset + anchorTop;
    const safeWidth = Number(width) || 0;
    const safeHeight = Number(height) || 0;
    const safeTransparency = Math.min(
      Math.max(Number(fillTransparency) || 0, 0),
      1,
    );

    const shape = shapes.addTextBox("");
    shape.left = safeLeft;
    shape.top = safeTop;
    shape.width = safeWidth;
    shape.height = safeHeight;
    shape.textFrame.horizontalAlignment =
      Excel.ShapeTextHorizontalAlignment.center;
    shape.textFrame.leftMargin = 0;
    shape.textFrame.rightMargin = 0;
    shape.textFrame.topMargin = 0;
    shape.textFrame.bottomMargin = 0;
    shape.textFrame.autoSizeSetting = Excel.ShapeAutoSize.autoSizeNone;
    shape.textFrame.textRange.font.name = fontName;
    shape.textFrame.textRange.font.size = fontSizePx;
    shape.textFrame.textRange.font.bold = bold;

    // Now set the text
    shape.textFrame.textRange.text = text;
    shape.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;

    if (fillColor) {
      shape.fill.setSolidColor(fillColor);
      shape.fill.transparency = safeTransparency;
    } else {
      shape.fill.clear();
    }

    if (Number(borderSize)) {
      shape.lineFormat.weight = borderSize;
      shape.lineFormat.color = borderColor || "black";
    } else {
      shape.lineFormat.visible = false;
    }

    helpers.ganttManager.nameAndTrack(shape, nameSuffix);
    return shape;
  }

  /**
   * Create a label designed to fit correctly with a Gantt bar.  barTop,
   * barLeft, and barHeight are used to size & locate bar for various label
   * layout options.
   * @param {Excel.ShapeCollection} shapes - ShapeCollection from sheet, where
   * this shaped will be added
   * @param {string} text - text to put inside label
   * @param {number} left - localX position of the left side of the bar label IF
   * the label uses left layout.
   * @param {number} barTop - in localY.  Pass the same value as was passed to
   * createShapeRect() as the top for the related bar.
   * @param {number} barLeft - in localX.  Pass the same value as was passed to
   * createShapeRect() as the left for the related bar.
   * @param {number} barHeight - Pass the same value as was passed to
   * createShapeRect() as the height for the related bar.
   * @param {boolean} insideAllowed - is the label allowed to be inside the
   * shape.  Essentially, this is asking whether this row is an activity.
   * @param {number} fontSizePx - font size, in pixels
   * @param {string} fillColor - string converted to color object.  Leave null
   * for no fill.
   * @param {number} fillTransparency - From 0 to 1, inclusive.  0 is opaque.
   * @param {string} nameSuffix - string to be added between "GanttShape_" and
   * the index
   * @returns {Excel.Shape}
   */
  function createBarLabel(
    shapes,
    text,
    left,
    barTop,
    barLeft,
    barHeight,
    insideAllowed,
    fontSizePx,
    fillColor,
    fillTransparency,
    nameSuffix,
  ) {
    const safeLeft = (Number(left) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;
    const cssFont = `${barLabelBold ? "bold " : ""}${fontSizePx}px ${fontName}`;
    const safeTransparency = Math.min(
      Math.max(Number(fillTransparency) || 0, 0),
      1,
    );
    const safeBarLeft = (Number(barLeft) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;

    // Measure the text size in pixels
    const textPxSize = helpers.measureTextPx(text, cssFont);
    const textPx = textPxSize.width;
    const textPxH = textPxSize.height;

    // Convert margins to pixels
    const marginPxW = helpers.ptToPx(
      barLabelMarginPtLeft + barLabelMarginPtRight,
    );
    const marginPxH = helpers.ptToPx(
      barLabelMarginPtTop + barLabelMarginPtBottom,
    );

    const targetWidthPx = Math.ceil(textPx + marginPxW + safetyPxLabel);
    const targetHeightPx =
      textPxH == 0
        ? helpers.ptToPx(fontSizePx) +
          helpers.ptToPx(topMarginPt + bottomMarginPt) +
          10
        : Math.ceil(textPxH + marginPxH + safetyPxLabel);

    let labelLayout;

    if (
      safeLeft + targetWidthPx - helpers.chartConfiguration.localXOffset - anchorLeft <=
      chartInternalSafeWidth
    ) {
      labelLayout = helpers.LabelLayout.RIGHT;
    } else if (
      targetWidthPx <
      safeBarLeft - helpers.chartConfiguration.barLabelBuffer - helpers.chartConfiguration.localXOffset - anchorLeft
    ) {
      labelLayout = helpers.LabelLayout.LEFT;
    } else {
      labelLayout = helpers.LabelLayout.INSIDE;
    }

    // Create the textbox with empty text first (prevents early layout issues)
    const shape = shapes.addTextBox("");

    shape.textFrame.leftMargin = barLabelMarginPtLeft;
    shape.textFrame.rightMargin = barLabelMarginPtRight;
    shape.textFrame.topMargin = barLabelMarginPtTop;
    shape.textFrame.bottomMargin = barLabelMarginPtBottom;

    const barTopCenterLine = barTop + barHeight / 2 + anchorTop + helpers.chartConfiguration.localYOffset;
    shape.top = Math.max(barTopCenterLine - targetHeightPx / 2, 0);
    switch (labelLayout) {
      case helpers.LabelLayout.RIGHT:
        shape.left = safeLeft;
        shape.textFrame.horizontalAlignment =
          Excel.ShapeTextHorizontalAlignment.left;
        break;
      case helpers.LabelLayout.LEFT:
        shape.left = safeBarLeft - targetWidthPx - helpers.chartConfiguration.barLabelBuffer;
        shape.textFrame.horizontalAlignment =
          Excel.ShapeTextHorizontalAlignment.right;
        break;
      case helpers.LabelLayout.INSIDE:
        shape.left = safeBarLeft;
        shape.textFrame.horizontalAlignment =
          Excel.ShapeTextHorizontalAlignment.center;
        break;
    }

    shape.textFrame.horizontalOverflow = Excel.ShapeTextHorizontalOverflow.clip;
    shape.textFrame.verticalOverflow = Excel.ShapeTextVerticalOverflow.ellipsis;

    // IMPORTANT: set width BEFORE setting text
    shape.width =
      labelLayout == helpers.LabelLayout.INSIDE
        ? safeLeft - helpers.chartConfiguration.barLabelBuffer - safeBarLeft
        : targetWidthPx;

    // Set a reasonable height (Excel will expand if needed, but we aim for one
    // line)
    shape.height = targetHeightPx;

    // Disable autosize since we are controlling width deterministically
    shape.textFrame.autoSizeSetting = Excel.ShapeAutoSize.autoSizeNone;

    shape.textFrame.textRange.font.name = fontName;
    shape.textFrame.textRange.font.size = fontSizePx;

    // Now set the text
    shape.textFrame.textRange.text = text;
    shape.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;

    const labelOverBackgroundColor = helpers.getContrastColor(
      fillColor ?? "#FFFFFF",
      safeTransparency,
    );
    const labelOverActivityColor = helpers.getContrastColor(defaultColor, 0);
    const labelColorsSame = labelOverBackgroundColor == labelOverActivityColor;

    switch (labelLayout) {
      case helpers.LabelLayout.RIGHT:
      case helpers.LabelLayout.LEFT:
        if (fillColor) {
          shape.fill.setSolidColor(fillColor);
          shape.fill.transparency = safeTransparency;
          shape.textFrame.textRange.font.color = labelOverBackgroundColor;
        } else {
          shape.fill.clear();
        }
        break;
      case helpers.LabelLayout.INSIDE:
        shape.fill.clear();
        labelTextColor = helpers.getContrastColor(defaultColor, 0);
        shape.textFrame.textRange.font.color = labelTextColor;
        break;
    }

    shape.lineFormat.visible = false;

    helpers.ganttManager.nameAndTrack(shape, nameSuffix);
    return shape;
  }

  /**
   * Create a simple line.  Location defined relative to top-left anchor point.
   * @param {Excel.ShapeCollection} shapes - ShapeCollection from sheet, where
   * this shaped will be added
   * @param {number} startLeft - starting localX
   * @param {number} startTop - starting localY
   * @param {number} endLeft - ending localX
   * @param {number} endTop - ending localY
   * @param {number} weight - stroke weight.  If null or erroneous, stroke wight
   * will be 1.
   * @param {string} color - string converted to color object.  If null, stroke
   * will be black
   * @param {string} nameSuffix - string to be added between "GanttShape_" and
   * the index
   * @returns {Excel.Shape}
   */
  function createLine(
    shapes,
    startLeft,
    startTop,
    endLeft,
    endTop,
    weight,
    color,
    nameSuffix,
  ) {
    const sL = (Number(startLeft) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;
    const sT = (Number(startTop) || 0) + helpers.chartConfiguration.localYOffset + anchorTop;
    const eL = (Number(endLeft) || 0) + helpers.chartConfiguration.localXOffset + anchorLeft;
    const eT = (Number(endTop) || 0) + helpers.chartConfiguration.localYOffset + anchorTop;

    const shape = shapes.addLine(sL, sT, eL, eT);

    const wt = Math.round(Number(weight)) || 1;

    shape.lineFormat.weight = wt;
    shape.lineFormat.color = color || "black";

    helpers.ganttManager.nameAndTrack(shape, nameSuffix);
    return shape;
  }

  try {
    // First Excel.run call creates all the individual shapes and then closes
    // the session.
    await Excel.run(async (context) => {
      // File setup
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      await helpers.ganttManager.clearAll(context, sheet);
      const shapes = sheet.shapes;
      shapes.load("items/name");
      await context.sync();

      if (chartTop && chartLeft) {
        anchorLeft = chartLeft;
        anchorTop = chartTop;
      } else {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("left, top");
        await context.sync();

        anchorLeft = Math.round(Math.max(20, selectedRange.left));
        anchorTop = Math.round(Math.max(20, selectedRange.top));
      }

      // Data processing
      const dateTable = sheet.tables.getItemAt(0);
      const headerRange = dateTable.getHeaderRowRange().load("values");
      const bodyRange = dateTable.getDataBodyRange().load("values");
      await context.sync();

      helpers.testData(headerRange.values, bodyRange.values);

      // Process table data into usable arrays
      const headers = headerRange.values[0].map((header) =>
        header.toLowerCase(),
      );
      const typeIndex = headers.indexOf("type");
      const startIndex = headers.indexOf("start date");
      const endIndex = headers.indexOf("end date");
      const titleIndex = headers.indexOf("title");
      let data = bodyRange.values;

      let initialSafeTemplateHeight;
      if (chartHeight) {
        initialSafeTemplateHeight = Math.round(
          (chartHeight -
            (2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding + 2 * helpers.chartConfiguration.scaleAxisHeight)) /
            data.length,
        );
      } else if (templateHeight) {
        initialSafeTemplateHeight = Math.round(templateHeight);
      } else {
        throw "No good height given to chart function";
      }

      /**
       * Rounds the input templateHeight in order to make further rounding logic simpler.
       */
      const safeTemplateHeight = initialSafeTemplateHeight;

      /**
       * Defines the height for Gantt bars & milestones.  templateHeight controls the
       * spacing from top to top of these bars, but this controls the height directly.
       * @type {number}
       */
      const size0 = Math.round((3 / 4) * safeTemplateHeight);

      const startDates = data.map(
        (row) => new Date(helpers.excelDateToJS(row[startIndex])),
      );
      const endDates = data.map(
        (row) => new Date(helpers.excelDateToJS(row[endIndex])),
      );
      // This returns for each row, whichever is more, the startDate plus 1 or
      // the endDate.
      const maxTimestamps = data.map((row, index) =>
        Math.max(
          startDates[index] instanceof Date && !isNaN(startDates[index])
            ? new Date(startDates[index]).setDate(
                startDates[index].getDate() + 1,
              )
            : 0,
          endDates[index] instanceof Date && !isNaN(endDates[index])
            ? endDates[index].getTime()
            : 0,
        ),
      );

      // Calculate scope of data
      const projectStart = new Date(Math.min(...startDates));
      const projectEnd = new Date(Math.max(...maxTimestamps));
      // rangeStart will always be the first of the month that contains
      // projectStart.  They could be equal.
      const rangeStart = new Date(
        projectStart.getFullYear(),
        projectStart.getMonth(),
        1,
      );
      const rangeStartM = rangeStart.getTime();
      // rangeEnd will always be the first day of the month after the month in
      // which projectEnd falls unless projectEnd is the first of the month.
      const rangeEnd = new Date(
        projectEnd.getFullYear(),
        projectEnd.getDate() == 1
          ? projectEnd.getMonth()
          : projectEnd.getMonth() + 1,
        1,
      );
      const totalMs = rangeEnd - rangeStart;
      const ganttWidth = chartInternalSafeWidth;
      const pxPerM = ganttWidth / totalMs;

      const monthTicks = helpers.calculateMonthTicks(
        rangeStart,
        rangeEnd,
        ganttWidth,
      );
      const yearTicks = helpers.calculateYearTicks(
        rangeStart,
        rangeEnd,
        ganttWidth,
      );
      const fiscalYearTicks = includeFY
        ? helpers.calculateFiscalYearTicks(rangeStart, rangeEnd, ganttWidth)
        : null;

      // Map raw Excel rows into "Widget-ready" objects
      const processedData = data.map((row, index) => {
        const startM = helpers.excelDateToJS(row[startIndex]).getTime();
        const endM = helpers.excelDateToJS(row[endIndex]).getTime();

        const xVal = Math.round((startM - rangeStartM) * pxPerM);
        let wVal = Math.round((endM - rangeStartM) * pxPerM) - xVal;

        // If it's a milestone or the dates are the same, give it a fixed width
        if (row[typeIndex] === "Milestone" || wVal <= 0) {
          wVal = 10;
        }

        return {
          taskName: String(row[titleIndex]) || "Unnamed Task",
          type: row[typeIndex],
          localX: isNaN(xVal) ? 0 : xVal,
          width: wVal,
          localY: index * safeTemplateHeight +  helpers.chartConfiguration.ganttTopBottomInternalPadding,
        };
      });

      // Create background
      const backgroundRect = createShapeRect(
        shapes,
        0,
        0,
        safeChartWidth,
        safeTemplateHeight * processedData.length +
          2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding +
          2 * helpers.chartConfiguration.scaleAxisHeight +
          (includeFY ? helpers.chartConfiguration.scaleAxisHeight : 0),
        "#FFFFFF",
        0,
        "BackgroundRect",
      );

      // Create monthly tick lines
      monthTicks.forEach((tick) => {
        createLine(
          shapes,
          tick.localX,
          0,
          tick.localX,
          safeTemplateHeight * processedData.length +
            2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding,
          helpers.chartConfiguration.innerBorderWidth,
          "#DDDDDD",
          0,
          "MonthlyTickLine",
        );
      });

      // Create Month axis labels
      monthTicks.forEach((tick, index) => {
        if (index < monthTicks.length - 1) {
          // 1. Pre-round the start and end X positions
          const currentX = Math.round(tick.localX);
          const nextX = Math.round(monthTicks[index + 1].localX);

          // 2. Calculate width based on the rounded integers
          const labelWidth = nextX - currentX;

          createSimpleLabel(
            shapes,
            tick.label,
            currentX,
            safeTemplateHeight * processedData.length +
              2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding,
            labelWidth,
            helpers.chartConfiguration.scaleAxisHeight,
            8,
            axisLabelBold,
            "#FFFFFF",
            0,
            "AxisMonthLabel",
            "#000000",
            helpers.chartConfiguration.innerBorderWidth,
          );
        }
      });

      // Create yearly tick lines
      yearTicks.forEach((tick) => {
        createLine(
          shapes,
          tick.localX,
          0,
          tick.localX,
          safeTemplateHeight * processedData.length +
            2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding +
            helpers.chartConfiguration.scaleAxisHeight,
          helpers.chartConfiguration.innerBorderWidth,
          "#000000",
          0,
          "YearlyTickLine",
        );
      });

      // Create Year axis labels
      yearTicks.forEach((tick, index) => {
        if (index < yearTicks.length - 1) {
          // 1. Pre-round the start and end X positions
          const currentX = Math.round(tick.localX);
          const nextX = Math.round(yearTicks[index + 1].localX);

          // 2. Calculate width based on the rounded integers
          const labelWidth = nextX - currentX;

          createSimpleLabel(
            shapes,
            tick.label,
            currentX,
            safeTemplateHeight * processedData.length +
              2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding +
              helpers.chartConfiguration.scaleAxisHeight,
            labelWidth,
            helpers.chartConfiguration.scaleAxisHeight,
            8,
            axisLabelBold,
            "#FFFFFF",
            0,
            "AxisYearLabel",
            "#000000",
            helpers.chartConfiguration.innerBorderWidth,
          );
        }
      });

      if (includeFY) {
        // Create Fiscal Year axis labels
        fiscalYearTicks.forEach((tick, index) => {
          if (index < fiscalYearTicks.length - 1) {
            // 1. Pre-round the start and end X positions
            const currentX = Math.round(tick.localX);
            const nextX = Math.round(fiscalYearTicks[index + 1].localX);

            // 2. Calculate width based on the rounded integers
            const labelWidth = nextX - currentX;

            createSimpleLabel(
              shapes,
              tick.label,
              currentX,
              safeTemplateHeight * processedData.length +
                2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding +
                2 * helpers.chartConfiguration.scaleAxisHeight,
              labelWidth,
              helpers.chartConfiguration.scaleAxisHeight,
              8,
              axisLabelBold,
              "#FFFFFF",
              0,
              "AxisYearLabel",
              "#000000",
              helpers.chartConfiguration.innerBorderWidth,
            );
          }
        });
      }

      // Get current local time
      const today = new Date();
      const todayMs = today.getTime();

      // Calculate X position relative to your project start rangeStartM and
      // pxPerM are already defined in your script
      if (todayMs >= rangeStartM && todayMs <= rangeEnd.getTime()) {
        const todayX = Math.round((todayMs - rangeStartM) * pxPerM);

        // Create a vertical "Today" line
        const todayLine = createLine(
          shapes,
          todayX,
          0,
          todayX,
          safeTemplateHeight * processedData.length +
            2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding,
          2, // Slightly thicker
          "#FF0000", // Red
          "TodayLine",
        );

        // Optional: Make it dashed
        todayLine.lineFormat.dashStyle = Excel.ShapeLineDashStyle.dash;
      }

      // Create bars and labels
      processedData.forEach((row) => {
        const activityBar =
          row.type == "Activity"
            ? createShapeRect(
                shapes,
                row.localX,
                row.localY,
                row.width,
                size0,
                defaultColor,
                0,
                "Bar",
              )
            : createShapeDiamond(
                shapes,
                row.localX,
                row.localY,
                size0,
                defaultColor,
                0,
                "Diamond",
              );
        //   actualRect.load("top,width,left,height");
        const activityLabel = createBarLabel(
          shapes,
          row.taskName,
          row.type == "Activity"
            ? row.localX + row.width + helpers.chartConfiguration.barLabelBuffer
            : row.localX + Math.ceil(size0 / 2) + helpers.chartConfiguration.barLabelBuffer,
          row.localY,
          row.type == "Activity"
            ? row.localX
            : row.localX - Math.ceil(size0 / 2),
          size0,
          row.type == "Activity",
          10,
          "#FFFFFF",
          0.2,
          "Label",
        );
      });

      const outerBox = createShapeRect(
        shapes,
        0,
        0,
        safeChartWidth,
        safeTemplateHeight * processedData.length +
          2 *  helpers.chartConfiguration.ganttTopBottomInternalPadding +
          2 * helpers.chartConfiguration.scaleAxisHeight +
          (includeFY ? helpers.chartConfiguration.scaleAxisHeight : 0),
        "#FFFFFF",
        1,
        "OuterBox",
        "#000000",
         helpers.chartConfiguration.outerBorderWidth,
      );

      shapes.load("items/name, items/visible, items/zOrderPosition");

      await context.sync();

      // Move LabelBackground shapes behind Bars
      const bars = shapes.items.filter((item) =>
        item.name.startsWith("GanttShape_Bar"),
      );
      const minBarZ = Math.min(...bars.map((bar) => bar.zOrderPosition));

      const labelBackgrounds = shapes.items.filter((item) =>
        item.name.startsWith("GanttShape_LabelBackground"),
      );
      labelBackgrounds.forEach((labelBackground) => {
        if (labelBackground.zOrderPosition > minBarZ) {
          const n = labelBackground.zOrderPosition - minBarZ;
          [...Array(n)].forEach(() =>
            labelBackground.setZOrder(Excel.ShapeZOrder.sendBackward),
          );
        }
      });

      // Group shapes
      helpers.ganttManager.groupAll(shapes);

      shapes.load("items/name, items/type");
      await context.sync();

      const ghosts = shapes.items.filter(
        (shape) =>
          shape.name.startsWith("GanttShape") && shape.type !== "Group",
      );
      ghosts.forEach((shape) => {
        shape.delete();
      });

      if (ghosts.length > 0) {
        console.log("found erroneous shapes and deleted them");
      } else {
        console.log("no erroneous shapes found");
      }
    });

    console.log("Finished block 3");
  } catch (error) {
    console.error("FULL ERROR OBJECT:", JSON.stringify(error));
    console.error("ERROR CODE:", error.code);

    GanttErrorHandler.handle(error);
  } finally {
    // Hides the loader when everything (success or failure) is done
    overlay.style.display = "none";
  }
}

async function cleanUpUngroupedShapes() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes;
    shapes.load("items/name, items/type");
    await context.sync();

    const ghosts = shapes.items.filter(
      (shape) => shape.name.startsWith("GanttShape") && shape.type !== "Group",
    );
    ghosts.forEach((shape) => {
      shape.delete();
    });

    if (ghosts.length > 0) {
      console.log("found erroneous shapes and deleted them");
    } else {
      console.log("no erroneous shapes found");
    }
  });
}

async function runGanttDiagnostic() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    // Load top level
    const shapes = sheet.shapes;
    shapes.load("items/name, items/type, items/visible, items/width");
    await context.sync();

    console.log("--- GANTT SHAPE DIAGNOSTIC REPORT ---");

    for (let shape of shapes.items) {
      if (shape.name.startsWith("Gantt")) {
        console.log(
          `[Top Level] Name: ${shape.name} | Type: ${shape.type} | Width: ${shape.width.toFixed(2)}`,
        );

        if (shape.type === "Group") {
          const groupShapes = shape.group.shapes;

          // CRITICAL FIX: Load the nested properties directly into the items
          // collection We use the slash path to ensure 'wordWrap' is attached
          // to each item's textFrame
          groupShapes.load(
            "items/name, items/type, items/visible, items/width",
          );

          await context.sync(); // This "hydrates" all children at once

          groupShapes.items.forEach((child) => {
            const status = child.visible ? "VISIBLE" : "HIDDEN";

            console.log(`   └── [In Group] Name: ${child.name}`);
            console.log(
              `          > Status: ${status} | Type: ${child.type} | Width: ${child.width.toFixed(2)}`,
            );
          });
        }
      }
    }
    console.log("---------------------------------------");
  });
}
