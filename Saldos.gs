function handleSaldos(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();
  var value = e.value;

  // Atualizar data na coluna A se edição na coluna 3 (C)
  updateDateColumn(sheet, row, 3, col, value);

  // Converter texto para maiúscula
  if (value !== undefined) {
    uppercaseText(sheet, range, value);
  }

  // Aplicar cor na linha se edição for na coluna C
  if (col === 3) {
    applyRowColor(sheet, row, value);
  }
}

function updateDateColumn(sheet, row, targetCol, editedCol, value) {
  if (editedCol === targetCol && value !== undefined && value !== '') {
    var now = new Date();
    var dateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // zerar hora
    var dateCell = sheet.getRange(row, 1); // Coluna A
    dateCell.setValue(dateOnly);
    dateCell.setNumberFormat("MMM/dd"); // Exibe como Jun/13
  }
}

function uppercaseText(sheet, range, value) {
  if (typeof value === 'string') {
    range.setValue(value.toUpperCase());
  }
}

function applyRowColor(sheet, row, value) {
  var lastColumn = sheet.getLastColumn();
  var rangeToColor = sheet.getRange(row, 1, 1, lastColumn);

  // Se valor for vazio ou apagado
  if (value === undefined || value === '') {
    rangeToColor.setBackground('black');
    rangeToColor.setFontColor('white');
    return;
  }

  var numericValue = parseFloat(value);

  if (!isNaN(numericValue)) {
    if (numericValue > 0) {
      rangeToColor.setBackground('#4ea72e'); // Verde
    } else if (numericValue < 0) {
      rangeToColor.setBackground('#ff0000'); // Vermelho
    } else {
      rangeToColor.setBackground('#0000ff'); // Azul
    }
    rangeToColor.setFontColor('white');
  } else {
    rangeToColor.setBackground('#0000ff'); // Azul (texto)
    rangeToColor.setFontColor('white');
  }
}
