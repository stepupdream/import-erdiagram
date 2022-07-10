const driveFolderId = "************************";
const separationKey = "ColumnName";
const groupKey = "";

////////////////////////////////////////////
// Menu Method
////////////////////////////////////////////
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("Import ER Diagram");
  menu.addItem("Import only the current sheet", "loadSingle");
  menu.addItem("All import", "loadAll");
  menu.addToUi();
}

////////////////////////////////////////////
// Main Method
////////////////////////////////////////////
/**
 * Read all ER diagrams
 */
function loadAll() {
  syncUML(true);
}

/**
 * Read only the ER diagram for the currently open sheet name
 */
function loadSingle() {
  syncUML(false);
}

/**
 * Reflect ER diagram in table definition
 *
 * @param isAll
 */
function syncUML(isAll) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = activeSpreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  const folder = DriveApp.getFolderById(driveFolderId);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    let fileName = file.getName();
    fileName = fileName.replace('.puml', '');
    const contents = file.getBlob().getDataAsString("utf-8").split(/[\n]+/);
    if (isAll == false && pascalCase(activeSheetName) != pascalCase(fileName)) {
      continue;
    }

    // If there is no target sheet, insert the initial value and generate it
    let targetSheet = activeSpreadsheet.getSheetByName(pascalCase(fileName));
    if (targetSheet == null) {
      targetSheet = createNewSheet(activeSpreadsheet, fileName);
    }

    const readERDiagramValues = convertERDiagram(activeSpreadsheet.getName(), contents);
    sync(targetSheet, readERDiagramValues);
  }
}

/**
 * Synchronize the contents of the ER diagram to a spreadsheet
 *
 * @param activeSheet
 * @param readERDiagramValues
 */
function sync(activeSheet, readERDiagramValues) {
  const lastRow = activeSheet.getLastRow();
  const lastColumn = activeSheet.getLastColumn();
  const range = activeSheet.getRange(1, 1, lastRow, lastColumn);
  const sheetArrayData = getValuesAndFormulas(range);
  const headersMain = getParentAttributeKeyName(sheetArrayData);
  const headersSub = getattributeKeyName(sheetArrayData);
  const rowsData = convertSheet(sheetArrayData);
  const readSheetValues = readGoogleSpreadsheet(headersMain, headersSub, rowsData);
  const lastRowStart = activeSheet.getLastRow();

  if (lastRowStart != 1) {
    activeSheet.getRange(2, 1, lastRowStart - 1, lastColumn).setBackground(null);
    activeSheet.getRange(2, 1, lastRowStart - 1, lastColumn).clearContent();
  }

  createTableDefinition(activeSheet, readERDiagramValues, readSheetValues, headersMain, headersSub);
}

////////////////////////////////////////////
// Google Spreadsheet Method
////////////////////////////////////////////
/**
 * Get an array of spreadsheet information
 *
 * @param sheet
 * @param readERDiagramValues
 * @param readSheetValues
 * @param headersMain
 * @param headersSub
 */
function createTableDefinition(sheet, readERDiagramValues, readSheetValues, headersMain, headersSub) {
  let setRowsValues = [];
  const requiredHeaderMain = { "TableName": "TableName", "TableDescription": "TableDescription", "ConnectionName": "ConnectionName" };
  const requiredHeaderSub = { "ColumnName": "ColumnName", "ColumnDescription": "ColumnDescription", "DataType": "DataType", "IsNullable": "IsNullable" };
  const headersMainAndSub = headersMain.concat(headersSub);
  const readTableNamesByERDiagram = Object.keys(readERDiagramValues);
  let readTableNamesBySheetValue;
  let readTableNames;
  let rowNumber = 1;
  let colorStart = 2;

  // Get the list of required table names to be output as a table definition document. If there is no table data, only ER diagram is targeted.
  // If it is on the table definition side before reading and it is not in the ER diagram, it shall not be deleted without permission.
  if (Object.keys(readSheetValues).length) {
    readTableNamesBySheetValue = Object.keys(readSheetValues);
    readTableNames = readTableNamesByERDiagram.concat(readTableNamesBySheetValue);
    let tableKyesSet = new Set(readTableNames);
    readTableNames = Array.from(tableKyesSet);
  } else {
    readTableNames = readTableNamesByERDiagram;
  }

  // Output the definition for each table
  const lastKeyName = Object.values(readTableNames).pop();
  for (let tableName of readTableNames) {
    let isFirstRow = true;

    // Branch the output contents depending on whether the table exists in the ER diagram
    if (readERDiagramValues[tableName]) {
      let readColumnNamesByERDiagram = Object.keys(readERDiagramValues[tableName]['attributes']);
      let readColumnNamesBySheetValue
      let readColumnNames;

      // Get the list of required column names to output as a table definition document. If there is no table data, only ER diagram is targeted.
      // If it is on the table definition side before output and it is not in the ER diagram, it will not be deleted without permission.
      if (readSheetValues[tableName]) {
        readColumnNamesBySheetValue = Object.keys(readSheetValues[tableName]["attributes"]);
        readColumnNames = readColumnNamesByERDiagram.concat(readColumnNamesBySheetValue);
        let set = new Set(readColumnNames);
        readColumnNames = Array.from(set);
      } else {
        readColumnNames = readColumnNamesByERDiagram;
      }

      // Generate output contents line by line
      for (let columnName of readColumnNames) {
        let row = [];

        // Insert the data of the parent element only in the first row
        if (isFirstRow === true) {
          for (let headerMain of headersMain) {
            if (headerMain in requiredHeaderMain && readERDiagramValues[tableName]) {
              row.push(readERDiagramValues[tableName][headerMain]);
            } else if (readSheetValues[tableName] && readSheetValues[tableName][headerMain]) {
              // If it does not exist in the ER diagram but exists in the table definition
              row.push(readSheetValues[tableName][headerMain]);
            } else {
              // It does not exist in the ER diagram. Moreover, even if it is data, it does not exist in the table definition, but only the header exists (error avoidance).
              row.push("");
            }
          }
          isFirstRow = false
        } else {
          // The parent element is blank after the second line
          for (let columnIndex = 0; columnIndex < headersMain.length; columnIndex++) {
            row.push("");
          }
        }

        for (let headerSub of headersSub) {
          if (headerSub in requiredHeaderSub && readERDiagramValues[tableName]['attributes'][columnName]) {
            row.push(readERDiagramValues[tableName]['attributes'][columnName][headerSub]);
          } else if (readSheetValues[tableName] && readSheetValues[tableName]['attributes'][columnName] && readSheetValues[tableName]['attributes'][columnName][headerSub]) {
            // If it does not exist in the ER diagram but exists in the table definition
            row.push(readSheetValues[tableName]['attributes'][columnName][headerSub]);
          } else {
            // It does not exist in the ER diagram. Moreover, even if it is data, it does not exist in the table definition, but only the header exists (error avoidance).
            row.push("");
          }
        }

        rowNumber++;
        setRowsValues.push(row);
      }
    } else {
      // Consider only the table definition because it is a table that does not exist in the ER diagram
      let columnNames = Object.keys(readSheetValues[tableName]['attributes']);

      for (let columnName of columnNames) {
        let row = [];
        if (isFirstRow === true) {
          for (let headerMain of headersMain) {
            row.push(readSheetValues[tableName][headerMain]);
          }
          isFirstRow = false;
        } else {
          for (let columnIndex = 0; columnIndex < headersMain.length; columnIndex++) {
            row.push("");
          }
        }

        for (let headerSub of headersSub) {
          if (readSheetValues[tableName]['attributes'][columnName][headerSub]) {
            row.push(readSheetValues[tableName]['attributes'][columnName][headerSub]);
          } else {
            row.push("");
          }
        }

        rowNumber++;
        setRowsValues.push(row);
      }
    }

    // coloring
    rowNumber++;
    sheet.getRange(colorStart, 1, rowNumber - colorStart, headersMainAndSub.length).setBackground("rgba(255,235,148,0.76)");
    colorStart = rowNumber + 1;

    // Insert blank rows below except for the last table
    if (tableName !== lastKeyName) {
      let row = [];
      for (let columnIndex = 0; columnIndex < headersMainAndSub.length; columnIndex++) {
        row.push("");
      }
      setRowsValues.push(row);
    }
  }

  if (setRowsValues.length >= 1) {
    sheet.getRange(2, 1, setRowsValues.length, setRowsValues[0].length).setBorder(true, true, true, true, true, true);
    sheet.getRange(2, 1, setRowsValues.length, setRowsValues[0].length).setValues(setRowsValues);
  }
}

/**
 * Get an array of spreadsheet information
 *
 * @param headersMain
 * @param headersSub
 * @param rowsData
 * @return
 */
function readGoogleSpreadsheet(headersMain, headersSub, rowsData) {
  let rowNumber = 0;
  let parentAttributes = {};

  // Loop until out of range
  while (rowsData[rowNumber] !== undefined) {
    if (isAllEmpty(rowsData[rowNumber])) {
      rowNumber++;
      continue;
    }

    let parentMainKeyName = rowsData[rowNumber][headersMain[0]];

    [rowNumber, parentAttribute] = createParentAttribute(headersMain, headersSub, rowsData, rowNumber);
    parentAttributes[parentMainKeyName] = parentAttribute;
  }

  return parentAttributes;
};

/**
 * Get one data group from a spreadsheet (ex: 1 table of data)
 *
 * @param headersMain
 * @param headersSub
 * @param rowsData
 * @param rowNumber
 * @return
 */
function createParentAttribute(headersMain, headersSub, rowsData, rowNumber) {
  let parentAttribute = {};
  for (let i = 0; i < headersMain.length; i++) {
    let headerMain = headersMain[i];
    parentAttribute[headerMain] = rowsData[rowNumber][headerMain];
  }

  while (rowsData[rowNumber] !== undefined && !isAllEmpty(rowsData[rowNumber])) {
    if (groupKey !== "") {
      let groupKeyName = rowsData[rowNumber][groupKey];
      [rowNumber, parentAttribute[groupKeyName]] = createAttributeGroup(headersSub, rowsData, rowNumber, 1);
    } else {
      [rowNumber, parentAttribute["attributes"]] = createAttributeGroup(headersSub, rowsData, rowNumber, 0);
    }
  }

  return [rowNumber, parentAttribute];
}

/**
 * A parent can have multiple child groups. Generate a group of one child
 *
 * @param headersSub
 * @param rowsData
 * @param rowNumber
 * @param keyIndex
 * @return
 */
function createAttributeGroup(headersSub, rowsData, rowNumber, keyIndex) {
  let attributes = {};
  const beforeKeyValue = rowsData[rowNumber][headersSub[0]];

  while (true) {
    let attribute = {};
    for (let j = 0; j < headersSub.length; j++) {
      let headerSub = headersSub[j];
      attribute[headerSub] = rowsData[rowNumber][headerSub];
    }

    if (!isAllEmpty(rowsData[rowNumber])) {
      attributes[rowsData[rowNumber][headersSub[keyIndex]]] = attribute;
    }
    rowNumber++;

    if (!rowsData[rowNumber] || isAllEmpty(rowsData[rowNumber])) {
      break;
    }

    if (groupKey != "" && rowsData[rowNumber][headersSub[0]] != "" && rowsData[rowNumber][headersSub[0]] != beforeKeyValue) {
      break;
    }
  }

  return [rowNumber, attributes];
}

/**
 * Get both values in the specified range and formulas
 *
 * @param range
 * @return
 */
function getValuesAndFormulas(range) {
  // Get only the value
  let valuesAndFomulas = range.getValues();

  // Get formula
  let tempFormulas = range.getFormulas();

  // Combine each
  for (let column = 0; column < valuesAndFomulas[0].length; column++) {
    for (let row = 0; row < valuesAndFomulas.length; row++) {
      if (tempFormulas[row][column].length != 0) {
        valuesAndFomulas[row][column] = tempFormulas[row][column];
      }
    }
  }

  return valuesAndFomulas;
};

/**
 * Generate a new sheet with default data
 *
 * @param activeSpreadsheet
 * @param fileName
 * @return
 */
function createNewSheet(activeSpreadsheet, fileName) {
  const sheet = activeSpreadsheet.insertSheet(pascalCase(fileName));

  sheet.getRange(1, 1, 1, 9).setBackgroundRGB(182, 215, 168);
  sheet.getRange(1, 1).setValue('TableName');
  sheet.getRange(1, 2).setValue('TableDescription');
  sheet.getRange(1, 3).setValue('ConnectionName');
  sheet.getRange(1, 4).setValue('ColumnName');
  sheet.getRange(1, 5).setValue('ColumnDescription');
  sheet.getRange(1, 6).setValue('DataType');
  sheet.getRange(1, 7).setValue('IsNullable');
  sheet.getRange(1, 8).setValue('IsUnsigned');
  sheet.getRange(1, 9).setValue('Comment');

  return sheet;
}

/**
 * Get the first row in an array (parent side)
 *
 * @param values
 * @return
 */
function getParentAttributeKeyName(values) {
  const keys = values[0];
  let names = [];
  for (let i = 0; i < keys.length; i++) {
    let key = keys[i];
    if (key == separationKey) {
      break;
    }
    names.push(key);
  }

  return names;
};

/**
 * Get the first row as an array
 *
 * @param values
 * @return
 */
function getattributeKeyName(values) {
  const keys = values[0];
  let names = [];
  let subKeyStart = false;
  for (let i = 0; i < keys.length; i++) {
    let key = keys[i];
    if (key == separationKey) {
      subKeyStart = true;
    }

    if (subKeyStart != false) {
      names.push(key);
    }
  }

  return names;
};

/**
 * Whether the entire line is empty
 *
 * @param obj
 * @return
 */
function isAllEmpty(obj) {
  for (let key in obj) {
    if (obj.hasOwnProperty(key)) {
      if (obj[key] != '') {
        return false;
      }
    }
  }
  return true;
};

////////////////////////////////////////////
// ERDiagram Method
////////////////////////////////////////////
/**
 * Format the data read from the ER diagram
 *
 * @param activeSpreadsheetName
 * @param contents
 * @return
 */
function convertERDiagram(activeSpreadsheetName, contents) {
  let readERDiagrams = {};

  for (let rowIndex = 0; rowIndex < contents.length; rowIndex++) {
    // When the row with the description of entity is reached, it is judged as the contents of the table definition until} appears after that (it is managed by the isLoading flag).
    if (contents[rowIndex].match(/entity \"/)) {
      let databaseCategoryName = getDatabaseCategoryName(contents[rowIndex]);

      // Do not output if the currently executing sheet and table category are different (Example: Master data is not output to User type table)
      if (activeSpreadsheetName != databaseCategoryName) {
        continue;
      }

      const [tableName, tableDescription] = getTableName(contents[rowIndex]);
      const connectionName = getConnectionName(contents[rowIndex]);

      // Insert what you read from the ER diagram
      parentAttribute = {};
      parentAttribute["DatabaseCategoryName"] = databaseCategoryName;
      parentAttribute["TableName"] = tableName;
      parentAttribute["TableDescription"] = tableDescription;
      parentAttribute["ConnectionName"] = connectionName;
      rowIndex++;

      attributes = {};
      while (!contents[rowIndex].match(/}/)) {
        // Those with-are assumed to be nullable columns
        let isNullable = 'TRUE';
        if (contents[rowIndex].match(/\-/)) {
          isNullable = 'FALSE';
        }
        // Remove unnecessary symbols with regular expressions and extract column information
        // ex1) + id : bigInteger [ID]
        // ex2) - id : string[] [name]
        let attributeText = contents[rowIndex].replace(/[-#+~]/, '').trim();
        attributeText = deleteByTargetNumber(attributeText, attributeText.lastIndexOf(']'));
        attributeText = replaceByTargetNumber(attributeText, '+', attributeText.lastIndexOf('['));
        const columnDescription = attributeText.split("+")[1];
        const tempText = attributeText.split("+")[0].split(":");
        const columnName = tempText[0].trim();
        const dataType = tempText[1].trim();

        let attribute = {};
        attribute["ColumnName"] = columnName;
        attribute["ColumnDescription"] = columnDescription;
        attribute["DataType"] = dataType;
        attribute["IsNullable"] = isNullable;
        attributes[columnName] = attribute;
        rowIndex++;
      }
      parentAttribute["attributes"] = attributes;
      readERDiagrams[tableName] = parentAttribute;
    }
  }

  return readERDiagrams;
}

/**
 * Get the database category (used to determine if it matches the title part of Spreadsheet)
 *
 * @param content
 * @return
 */
function getDatabaseCategoryName(content) {
  content = content.split(">>")[0];

  return content.split(",")[1].trim();
};

/**
 * Get the connection name
 *
 * @param content
 * @return
 */
function getConnectionName(content) {
  const str = content.split(">>")[1];

  return str.split("{")[0].trim();
};

/**
 * Get table name and table details
 *
 * @param content
 * @return
 */
function getTableName(content) {
  const replaced = content.replace('entity "', '');
  const str = replaced.split("]")[0].split("[");

  return [str[0].trim(), str[1].trim()];
};

////////////////////////////////////////////
// Util Method
////////////////////////////////////////////
/**
 * Convert to camel case
 *
 * @param text
 * @return
 */
function camelCase(text) {
  if (!isString(text)) {
    return text;
  }
  text = text.charAt(0).toLowerCase() + text.slice(1);
  return text.replace(/[-_](.)/g, function (match, string) {
    return string.toUpperCase();
  });
};

/**
 * Convert to snake case
 *
 * @param text
 * @return
 */
function snakeCase(text) {
  if (!isString(text)) {
    return text;
  }
  const camelText = camelCase(text);

  return camelText.replace(/[A-Z]/g, function (string) {
    return "_" + string.charAt(0).toLowerCase();
  });
};

/**
 * Convert to Pascal case
 *
 * @param text
 * @return
 */
function pascalCase(text) {
  if (!isString(text)) {
    return text;
  }
  const camelText = this.camelCase(text);

  return camelText.charAt(0).toUpperCase() + camelText.slice(1);
};

/**
 * Whether the specified variable is a string
 *
 * @param obj
 * @return
 */
function isString(obj) {
  return typeof (obj) == "string" || obj instanceof String;
};

/**
 * Delete the character string at the specified position
 *
 * @param text
 * @param targetNumber
 * @return
 */
function deleteByTargetNumber(text, targetNumber) {
  if (targetNumber <= 0) {
    return text.slice(1);
  }
  const before = text.slice(0, targetNumber);
  const after = text.slice(targetNumber + 1);

  return before + after;
};

/**
 * Replace the character string at the specified position
 *
 * @param text
 * @param replace
 * @param targetNumber
 * @return
 */
function replaceByTargetNumber(text, replace, targetNumber) {
  if (targetNumber <= 0) {
    return replace + text.slice(1);
  }
  const before = text.slice(0, targetNumber);
  const after = text.slice(targetNumber + 1);

  return before + replace + after;
};

/**
 * Convert a 2D array to an associative array (key to the first row)
 *
 * @param values
 * @return
 */
function convertSheet(values) {
  const keys = values.splice(0, 1)[0];
  return values.map(function (row) {
    let object = [];
    row.map(function (column, index) {
      object[keys[index]] = column;
    });

    return object;
  });
};
