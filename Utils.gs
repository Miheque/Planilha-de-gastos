function transformPtoCheckmark(sheet, range, col, row, value) {
  if (sheet.getName() !== "SALDO" && col >= 4 && col <= 16 && value && typeof value === "string") {
    if (/^[Pp]/.test(value)) {
      range.setValue("✓ " + value.substring(1).toUpperCase());
      return true; // indica que transformou
    }
  }
  return false;
}

function updateDateColumn(sheet, row, targetCol, editedCol, value) {
  if (editedCol === targetCol && row > 0) {
    var cellA = sheet.getRange(row, 1);
    if (!value) {
      cellA.clearContent();
    } else if (cellA.isBlank()) {
      cellA.setValue(getFormattedDate());
    }
  }
}

function uppercaseText(sheet, range, value) {
  if (value && typeof value === "string" && value !== value.toUpperCase()) {
    range.setValue(value.toUpperCase());
  }
}

function getFormattedDate() {
  var today = new Date();
  return Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM/dd");
}

function handleDefault(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var value = e.value;

  if (value && typeof value === "string" && value !== value.toUpperCase()) {
    range.setValue(value.toUpperCase());
  }
}
