class MockRange {
  constructor(sheet, startRow, startCol, numRows, numCols) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.startCol = startCol;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  _ensureSize(row, col) {
    while (this.sheet.data.length < row) {
      this.sheet.data.push(new Array(this.sheet.maxColumns || col).fill(''));
    }
    var targetRow = this.sheet.data[row - 1];
    while (targetRow.length < col) {
      targetRow.push('');
    }
    this.sheet.maxColumns = Math.max(this.sheet.maxColumns, col);
  }

  getValues() {
    var values = [];
    for (var r = 0; r < this.numRows; r++) {
      var rowIndex = this.startRow - 1 + r;
      var row = this.sheet.data[rowIndex] || [];
      var rowValues = [];
      for (var c = 0; c < this.numCols; c++) {
        var colIndex = this.startCol - 1 + c;
        rowValues.push(row[colIndex]);
      }
      values.push(rowValues);
    }
    return values;
  }

  setValues(values) {
    for (var r = 0; r < this.numRows; r++) {
      var rowIndex = this.startRow - 1 + r;
      this._ensureSize(rowIndex + 1, this.startCol + this.numCols - 1);
      var targetRow = this.sheet.data[rowIndex];
      for (var c = 0; c < this.numCols; c++) {
        var colIndex = this.startCol - 1 + c;
        targetRow[colIndex] = values[r][c];
      }
    }
  }

  setValue(value) {
    this.setValues([[value]]);
  }

  getValue() {
    return this.getValues()[0][0];
  }

  setFontWeight() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
  setNumberFormat() { return this; }
  setDataValidation() { return this; }
}

class MockSheet {
  constructor(name, rows) {
    this.name = name;
    this.data = rows && rows.length ? rows.map(function(row) { return row.slice(); }) : [['']];
    var initialCols = this.data.length && this.data[0].length ? this.data[0].length : 1;
    this.maxColumns = this.data.reduce(function(max, row) {
      return Math.max(max, row.length);
    }, initialCols);
    var self = this;
    this.data = this.data.map(function(row) {
      var copy = row.slice();
      while (copy.length < self.maxColumns) copy.push('');
      return copy;
    });
  }

  getName() {
    return this.name;
  }

  getDataRange() {
    return new MockRange(this, 1, 1, this.getLastRow(), Math.max(this.maxColumns, 1));
  }

  getRange(row, col, numRows, numCols) {
    numRows = numRows || 1;
    numCols = numCols || 1;
    return new MockRange(this, row, col, numRows, numCols);
  }

  appendRow(row) {
    var copy = row.slice();
    this.maxColumns = Math.max(this.maxColumns, copy.length);
    while (copy.length < this.maxColumns) copy.push('');
    this.data.push(copy);
  }

  setColumnWidth() {}
  getLastRow() {
    return this.data.length;
  }

  getLastColumn() {
    return this.maxColumns;
  }

  setConditionalFormatRules() {}
  setFrozenRows() {}
}

class MockSpreadsheet {
  constructor(initialSheets) {
    this.sheets = {};
    var names = Object.keys(initialSheets || {});
    if (names.length === 0) {
      this.sheets['Sheet1'] = new MockSheet('Sheet1', [['']]);
    } else {
      names.forEach((name) => {
        var value = initialSheets[name];
        var rows = Array.isArray(value) ? value : value.rows;
        this.sheets[name] = new MockSheet(name, rows);
      });
    }
  }

  getSheetByName(name) {
    return this.sheets[name] || null;
  }

  insertSheet(name) {
    var sheetName = name || ('Sheet' + (Object.keys(this.sheets).length + 1));
    var sheet = new MockSheet(sheetName, [['']]);
    this.sheets[sheetName] = sheet;
    return sheet;
  }

  getSheets() {
    return Object.values(this.sheets);
  }
}

function createMockSpreadsheet(initialSheets) {
  return new MockSpreadsheet(initialSheets);
}

function getSheetValues(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) return [];
  return sheet.data.map(function(row) { return row.slice(); });
}

module.exports = {
  createMockSpreadsheet: createMockSpreadsheet,
  getSheetValues: getSheetValues
};
