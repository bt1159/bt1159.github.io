const helpers = {
  /**
   * Convert an Excel serial date number into a JavaScript Date.
   *
   * Excel’s default (1900) date system stores dates as the number of days since
   * 1899‑12‑30 (because of the historical “1900 is a leap year” quirk).
   *
   * @param {number} excelDate - Days since 1899‑12‑30 in the Excel 1900 system.
   * @returns {Date} A JavaScript Date representing the same calendar day.
   */
  excelDateToJS: function (excelDate) {
    let excelDateFixed = excelDate ?? 0;
    const date = new Date(Math.round((excelDateFixed - 25569) * 86400 * 1000));
    return date;
  },

  /**
   * @callback NameAndTrack
   * @param {any} shape
   * @param {string} suffix
   * @returns {string}
   */

  /**
   * @callback ClearAll
   * @param {any} context
   * @param {any} sheet
   * @returns {Promise<void>}
   */

  /**
   * @callback GroupAll
   * @param {any} shapesCollection
   * @returns {any|null}
   */

  /**
   * @typedef {Object} GanttManager
   * @property {number} sI
   * @property {Excel.Shape[]} shapesArray
   * @property {NameAndTrack} nameAndTrack
   * @property {ClearAll} clearAll
   * @property {GroupAll} groupAll
   */

  /**
   * This object handles the naming of shapes, including the tracking of the sI
   * index, and adding them to the shapesArray group.
   * @type {GanttManager}
   */
  ganttManager: {
    sI: 0,
    shapesArray: [],

    // Your "extension-like" method
    nameAndTrack: function (shape, suffix) {
      const name = `GanttShape_${suffix}_${this.sI}`;
      shape.name = name;

      // Automatically track and increment
      this.shapesArray.push(shape);
      this.sI++;

      return name;
    },

    // A single command to wipe the slate clean
    clearAll: async function (context, sheet) {
      const shapes = sheet.shapes;
      shapes.load("items/name");
      await context.sync();

      // Collect ghosts and previous chart pieces in one go
      const targets = shapes.items.filter((s) =>
        s.name.startsWith("GanttShape"),
      );
      targets.forEach((s) => s.delete());

      // Reset state for a new run
      this.sI = 0;
      this.shapesArray = [];

      await context.sync();
      console.log(`Cleanup complete. Removed ${targets.length} shapes.`);
    },

    groupAll: function (shapesCollection) {
      if (this.shapesArray.length > 1) {
        const group = shapesCollection.addGroup(this.shapesArray);
        group.name = "GanttShapeGroup";

        return group;
      } else {
        console.log("Not enough shapes to group.");
        return null;
      }
    },
  },

  /**
   * Convert point size to pixels
   * @param {number} pt - A point size for fonts, shapes, etc.
   * @returns {number} - A size in pixels
   */
  ptToPx: function (pt) {
    // 1pt = 1/72 inch; CSS px assume 96 DPI => 96/72 px per pt
    return pt * (96 / 72);
  },

  /**
   * Creates an html canvas element with the string and font given and returns the
   * width and height of the created text element.  This will then get used by
   * other functions to create Excel shapes that have the appropriate width &
   * height, in lieu of an appropriate auto-sizing ability.
   * @param {string} text
   * @param {string} cssFont
   * @returns {{width: number, height: number}} A size object in pixels. If
   * measuring fails, returns {width:0, height:0}.
   */
  measureTextPx: function (text, cssFont) {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    if (!ctx)
      return {
        width: 0,
        height: 0,
      };
    ctx.font = cssFont;
    const m = ctx.measureText(text);
    const preciseW =
      (m.actualBoundingBoxLeft ?? 0) + (m.actualBoundingBoxRight ?? m.width);
    const preciseH =
      (m.actualBoundingBoxAscent ?? 0) + (m.actualBoundingBoxDescent ?? 0);
    const output = {
      width: Math.max(m.width, preciseW),
      height: preciseH,
    };
    return output;
  },

  /**
   * Determines whether black or white text is more readable based on background
   * color.  If this is not great for very dark or very saturated colors, consider
   * using the more complex WCAG 2.1 "gamma correction" standard
   * @param {string} hexColor - The background color in Hex format (e.g.
   * "#FFFFFF")
   * @param {number} transparency - 0 (opaque) to 1 (transparent)
   * @returns {string} - the ideal text color in hex code (e.g. "#FFFFFF")
   */
  getContrastColor: function (hexColor, transparency) {
    // 1. Remove the # and parse the hex values
    const hex = hexColor.replace("#", "");
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);

    // 2. Calculate Alpha (Excel transparency 0 = Opaque, so Alpha = 1 - transparency)
    const alpha = 1 - (Number(transparency) || 0);

    // 3. Blend with White Background (RGB 255, 255, 255)
    // Formula: (Color * Alpha) + (255 * (1 - Alpha))
    const blendR = r * alpha + 255 * (1 - alpha);
    const blendG = g * alpha + 255 * (1 - alpha);
    const blendB = b * alpha + 255 * (1 - alpha);

    // 2. Calculate relative luminance
    // Values are divided by 255 to get them into the 0-1 range
    const luminance =
      (0.2126 * blendR + 0.7152 * blendG + 0.0722 * blendB) / 255;

    // 3. Return black for light backgrounds, white for dark
    // 0.5 is the standard midpoint, but 0.6 is often better for "Office" aesthetics
    return luminance > 0.5 ? "#000000" : "#FFFFFF";
  },

  /**
   * @typedef {Object} Tick
   * @property {number} localX - the x position of the tick, from localX = 0,
   * which already takes into account anchorLeft and localXOffset.  SHOULD BE
   * ROUNDED.
   * @property {string} label - the label that should be used for that tick in the
   * axis.  For final tick, this is null.
   */

  /**
   * Function that takes the bounding dates and the width of the chart and returns
   * an array with the data needed to create the vertical monthly bars and the
   * corresponding monthly axis labels
   * @param {Date} start
   * @param {Date} end
   * @param {number} ganttWidth - the actual width of the Gantt chart, used to
   * calculate the spacing between ticks
   * @returns {Tick[]} - the information required for each tick: localX position
   * and label for the axis
   */
  calculateMonthTicks: function (start, end, ganttWidth) {
    const ticks = [];
    const startTime = start.getTime();
    const totalDuration = end.getTime() - startTime;
    const pixelsPerMs = ganttWidth / (totalDuration || 1); // Avoid div by zero

    let current = new Date(start);

    while (current <= end) {
      const elapsed = current.getTime() - startTime;
      const xPos = Math.round(elapsed * pixelsPerMs);

      ticks.push({
        localX: isNaN(xPos) ? 0 : xPos,
        label:
          current == end
            ? null
            : current.toLocaleDateString("en-US", { month: "short" }),
      });

      current.setMonth(current.getMonth() + 1);
    }
    return ticks;
  },

  /**
   * Function that takes the bounding dates and the width of the chart and returns
   * an array with the data needed to create the vertical yearly bars and the
   * corresponding yearly axis labels
   * @param {Date} start
   * @param {Date} end
   * @param {number} ganttWidth - the actual width of the Gantt chart, used to
   * calculate the spacing between ticks
   * @returns {Tick[]} - the information required for each tick: localX position
   * and label for the axis
   */
  calculateYearTicks: function (start, end, ganttWidth) {
    const ticks = [];
    const startTime = start.getTime();
    const totalDuration = end.getTime() - startTime;
    const pixelsPerMs = ganttWidth / (totalDuration || 1); // Avoid div by zero

    let current = new Date(start);
    let keepLooping = true;
    const rangeWithinOneYear = start.getFullYear() == end.getFullYear();
    let i = 0;

    // What if all the same year?
    while (keepLooping) {
      if (i >= 10) keepLooping = false;
      if (current > end) {
        keepLooping = false;
      } else if (
        current.getFullYear() == end.getFullYear() &&
        current.getMonth() == end.getMonth()
      ) {
        keepLooping = false;
      } else if (rangeWithinOneYear) {
        keepLooping = false;
      }
      const elapsed = Math.min(current.getTime(), end.getTime()) - startTime;
      const xPos = Math.round(elapsed * pixelsPerMs);

      ticks.push({
        localX: isNaN(xPos) ? 0 : xPos,
        label: !keepLooping
          ? null
          : current.toLocaleDateString("en-US", { year: "numeric" }),
      });
      const currentYear = current.getFullYear();
      current = new Date(currentYear + 1, 0, 1);
      i++;
    }

    return ticks;
  },

  /**
   * Function that takes the bounding dates and the width of the chart and returns
   * an array with the data needed to create the fiscal year axis labels.  Since
   * CTC fiscal year begins on July 1 of that calendar year, we can essentially
   * move certain dates six months later.
   * @param {Date} start
   * @param {Date} end
   * @param {number} ganttWidth - the actual width of the Gantt chart, used to
   * calculate the spacing between ticks
   * @returns {Tick[]} - the information required for each tick: localX position
   * and label for the axis
   */
  calculateFiscalYearTicks: function (startTrue, endTrue, ganttWidth) {
    const ticks = [];
    const start = new Date(startTrue.setMonth(startTrue.getMonth() + 6));
    const end = new Date(endTrue.setMonth(endTrue.getMonth() + 6));
    const startTime = start.getTime();
    const totalDuration = end.getTime() - startTime;
    const pixelsPerMs = ganttWidth / (totalDuration || 1); // Avoid div by zero

    let current = new Date(start);
    let keepLooping = true;
    const rangeWithinOneYear = start.getFullYear() == end.getFullYear();
    let i = 0;

    // What if all the same year?
    while (keepLooping) {
      if (i >= 10) keepLooping = false;
      if (current > end) {
        keepLooping = false;
      } else if (
        current.getFullYear() == end.getFullYear() &&
        current.getMonth() == end.getMonth()
      ) {
        keepLooping = false;
      } else if (rangeWithinOneYear) {
        keepLooping = false;
      }
      const elapsed = Math.min(current.getTime(), end.getTime()) - startTime;
      const xPos = Math.round(elapsed * pixelsPerMs);

      let labelYear = null;
      if (keepLooping) {
        labelYearCode =
          current.toLocaleDateString("en-US", { year: "numeric" }) % 100;
      }

      const label = !keepLooping ? null : "FY" + labelYearCode;

      ticks.push({
        localX: isNaN(xPos) ? 0 : xPos,
        label: label,
      });
      const currentYear = current.getFullYear();
      current = new Date(currentYear + 1, 0, 1);
      i++;
    }

    return ticks;
  },

  /**
   * Multi-step validation function to test the input data and make sure it has
   * the correct types, columns, etc.
   * @param {[string[]]} headerValues - 2D array with one item.  That item is an
   * array of strings.  Those strings are the column headers.
   * @param {any[][]} bodyValues - 2D array of table values.  Each top-level
   * array item is an array that represents a row from the table.
   * @returns {string} - true means data is valid
   */
  testData: function (headerValues, bodyValues) {
    const headers = headerValues[0].map((header) => header.toLowerCase());

    // Test that the correct headers are there.
    const typeIndex = headers.indexOf("type");
    if (typeIndex == -1) throw new GanttHeaderError("type");
    const startIndex = headers.indexOf("start date");
    if (startIndex == -1) throw new GanttHeaderError("start date");
    const endIndex = headers.indexOf("end date");
    if (endIndex == -1) throw new GanttHeaderError("end date");
    if (!headers.includes("title")) throw new GanttHeaderError("title");
    // NOTE: there is no need to test the title column since, in processedData, I
    // am defining the label with: String(row[titleIndex]) || "Unnamed task". That
    // way, even if the input cannot be converted to a string, it still will not
    // fail.  It will just be replaced with "Unnamed task".

    // Test the values in Type column
    let firstProblem = bodyValues.findIndex(
      (row) => !["Activity", "Milestone"].includes(row[typeIndex]),
    );
    if (firstProblem > -1)
      throw new GanttDataTypeError(
        firstProblem + 1,
        "Type",
        "'Activity' or 'Milestone'",
      );

    // Test the values in Start date column
    firstProblem = bodyValues.findIndex(
      (row) => typeof row[startIndex] !== "number",
    );
    if (firstProblem > -1)
      throw new GanttDataTypeError(firstProblem + 1, "Start date", "a number");

    // Test the values in End date column
    firstProblem = bodyValues.findIndex(
      (row) =>
        row[typeIndex] == "Activity" && typeof row[endIndex] !== "number",
    );
    if (firstProblem > -1)
      throw new GanttDataTypeError(firstProblem + 1, "End date", "a number");
    return true;
  },

  /** Acts like an enum with the various options for the spacing of each Gantt bar
   * label */
  LabelLayout: Object.freeze({
    LEFT: { align: "right" },
    RIGHT: { align: "left" },
    INSIDE: { align: "center" },
  }),

  /**
   * Generates a fingerprint for the table based ONLY on 4 specific columns.
   * @param {Array<Array>} rawData - The 2D array from Excel (including headers).
   * @returns {Promise<string>} A SHA-256 hash string.
   */
  getTableFingerprint: async function (rawData) {
    const targetHeaders = ["type", "start date", "end date", "title"];
    const headers = rawData[0].map((h) => h?.toString().trim().toLowerCase());

    // 1. Find indices of our targets
    const indices = targetHeaders.map((target) => headers.indexOf(target));

    // 2. Build a normalized 2D array (skipping headers, picking only our 4 columns)
    const normalizedData = rawData.slice(1).map((row) => {
      return indices.map((i) => {
        const val = row[i];
        // We return the value exactly as is (case-sensitive)
        // but we ensure null/undefined are treated as empty strings
        return val !== undefined && val !== null ? val.toString() : "";
      });
    });

    // 3. Convert to JSON and Hash
    const jsonString = JSON.stringify(normalizedData);
    const msgUint8 = new TextEncoder().encode(jsonString);
    const hashBuffer = await crypto.subtle.digest("SHA-256", msgUint8);

    // Convert buffer to hex string
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
  },

  chartConfiguration: (function() {
    const chartBorderWidth = 1;
    const chartBorderWidthInternalOnly = Math.ceil(chartBorderWidth/2);
    return {
      defaultChartWidth: 400,
      defaultTemplateHeight: 20,
      chartBorderWidth: chartBorderWidth,
      chartBorderWidthInternalOnly: chartBorderWidthInternalOnly,
      localXOffset: chartBorderWidthInternalOnly,
      localYOffset: chartBorderWidthInternalOnly,
      barLabelBuffer: 5,
      scaleAxisHeight: 20,
      innerBorderWidth: 1,
      outerBorderWidth: 2,
      ganttTopBottomInternalPadding: 5,
    };
  })(),
  
  // chartConfiguration2: {
  //   /**
  //    * The default width to make the entire image.  This is the value that the input
  //    * field in the taskbar uses as its default, and it supercedes a blank submitted
  //    * value
  //    * @type {number}
  //    */
  //   defaultChartWidth: 400,
  //   /**
  //    * The default height to make each Gantt chart bar or milestone.  This is the
  //    * value that the input field in the taskbar uses as its default, and it
  //    * supercedes a blank submitted value
  //    * @type {number}
  //    */
  //   defaultTemplateHeight: 20,
  //   /**
  //    * The width of the border around the entire Gantt chart.
  //    * @type {number}
  //    */
  //   chartBorderWidth: 1,
  //   /**
  //    * The portion of the width of the border that overlaps the internal area of the
  //    * background rectangle.  Moving to the right this many pixels gets you to the
  //    * chart local x = 0 point.  For decimal/integer precision, this is rounded.  It
  //    * is rounded up to ensure it is actually safe.
  //    * @type {number}
  //    */
  //   chartBorderWidthInternalOnly: Math.ceil(helpers.chartConfiguration.chartBorderWidth / 2),
  //   /**
  //    * The horizontal offset between leftAnchor and localX = 0
  //    * @type {number}
  //    */
  //   localXOffset: helpers.chartConfiguration.chartBorderWidthInternalOnly,
  //   /**
  //    * The vertical offset between topAnchor and localY = 0
  //    * @type {number}
  //    */
  //   localYOffset: helpers.chartConfiguration.chartBorderWidthInternalOnly,
  //   /**
  //    * The horizontal gap between a Gantt bar and its label, if laid out that way.
  //    */
  //   barLabelBuffer: 5,
  //   /**
  //    * Height of the axis labels
  //    * @type {number}
  //    */
  //   scaleAxisHeight: 20,
  //   /**
  //    * Stroke weight for borders of chart shapes excluding the border around the entire chart.
  //    * @type {number}
  //    */
  //   innerBorderWidth: 1,
  //   /**
  //    * Stroke weight for border around the entire chart.
  //    * @type {number}
  //    */
  //   outerBorderWidth: 2,
  //   /**
  //    * Vertical padding inside the Gantt chart above the first bar and below
  //    * the last bar
  //    */
  //   ganttTopBottomInternalPadding: 5,
  // },
};
