
/**
 * Base class for all Gantt-related exceptions.
 * @extends Error
 */
class GanttError extends Error {
  /**
   * @param {string} message - The error message.
   */
  constructor(message) {
    super(message);
    this.name = this.constructor.name;
    this.helpUrl = "https://bt1159.github.io/HelpPage";
  }
}

/**
 * Error thrown when a required column header is missing.
 * @extends GanttError
 */
class GanttHeaderError extends GanttError {
  /**
   * @param {string} correctColumnName - The exact string name the column should have.
   */
  constructor(correctColumnName) {
    super(
      "Your table is missing a column called '" +
        correctColumnName +
        "'.  Case does not matter.",
    );
    this.name = this.constructor.name;
    /** @type {string} */
    this.correctColumnName = correctColumnName;
  }
}

/**
 * Error thrown when a cell contains a value of the wrong type (e.g., text in a date column).
 * @extends GanttError
 */
class GanttDataTypeError extends GanttError {
  /**
   * @param {number} row - The 1-based index of the row in the table.
   * @param {string} columnName - The name of the column where the error occurred.
   * @param {string} columnName - The name of the column where the error occurred.
   */
  constructor(row, columnName, correctType) {
    super(
      "Your table has data of the wrong type.  In the '" +
        columnName +
        "' column, in row " +
        row +
        " (1 is the top row).  It should be: " +
        correctType +
        ".",
    );
    this.name = this.constructor.name;
    /** @type {number} */
    this.row = row;
    /** @type {string} */
    this.columnName = columnName;
    /** @type {string} */
    this.correctType = correctType;
  }
}

class GanttErrorHandler {
  /**
   * Centralized method to handle any error in the add-in
   * @param {Error} error
   */
  static handle(error) {
    // 1. Always log the full stack trace for the developer
    console.error(`[DEBUG] Origin: ${error.name}\nStack:`, error.stack);
    // Check for the "Cell Editing" error

    if (error instanceof OfficeExtension.Error) {
      // Now you are safe to check error.code
      if (
        error.code === "InvalidOperationInCellEditMode" ||
        error.code === "HostIsBusy"
      ) {
        this.showToast(
          "⚠️ Excel is Busy: Please press Enter or Esc to finish editing your cell, then try again.",
          "https://support.microsoft.com/en-us/office/edit-cell-content-877ad3c5-950c-4df4-942b-58673f32488a",
        );
        return;
      }
    }

    // 2. Decide how to notify the user
    if (error instanceof GanttHeaderError) {
      this.showDetailedAlert(error);
    } else if (error instanceof GanttDataTypeError) {
      this.showDetailedAlert(error);
    } else if (error.name === "OfficeExtension.Error") {
      this.handleExcelInternalError(error);
    } else {
      this.showGenericCrash(error);
    }
  }

  static showToast(msg, url) {
    const toast = document.getElementById("error-toast");
    const messageEl = document.getElementById("toast-message");
    const linkEl = document.getElementById("toast-link");

    if (!toast || !messageEl) {
      // Fallback if the HTML isn't ready or elements are missing
      console.error("Toast elements not found in HTML:", msg);
      return;
    }

    // Set the message
    messageEl.innerText = msg;

    // Set the link if provided, otherwise hide it
    if (url) {
      linkEl.href = url;
      linkEl.style.display = "inline";
    } else {
      linkEl.style.display = "none";
    }

    // Make it visible
    toast.style.display = "block";
  }

  static showDetailedAlert(error) {
    const message = `📊 Data Issue: ${error.message}`;
    this.showToast(message, error.helpUrl);
  }

  static highlightUIError(error) {
    const el = document.getElementById(error.elementId);
    if (el) {
      el.style.border = "2px solid red";
      setTimeout(() => (el.style.border = ""), 3000);
    }
    alert(`Input Error: ${error.message}`);
  }

  static handleExcelInternalError(error) {
    console.error("Excel Error Code:", error.code);
    // alert("Excel had trouble processing the request. Please try again.");
  }

  static showGenericCrash(error) {
    console.error("Generic excel error", error.message);
    console.error("code: ", error.code);
    // alert(
    //   "An unexpected error occurred. Please check the console for debugging info.",
    // );
  }
}