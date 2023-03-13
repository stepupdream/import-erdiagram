const driveFolderId = "**************************";
const separationKey = "ColumnName";

////////////////////////////////////////////
// Menu Method
////////////////////////////////////////////
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("ER図を取り込む");
  menu.addItem("現在のシートのみ取り込む", "loadSingle");
  menu.addItem("全データを取り込む", "loadAll");
  menu.addToUi();
}

////////////////////////////////////////////
// Main Method
////////////////////////////////////////////
/**
 * ER図をすべて読み取る
 */
function loadAll() {
  syncUML(true);
}

/**
 * 現在開いているシート名に関するER図のみを読み取る
 */
function loadSingle() {
  syncUML(false);
}

/**
 * ER図をテーブル定義書に反映する
 *
 * @param isAll すべてのファイルを対象とするかどうか
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

    // 対象のシートがなければ初期値を挿入して生成する
    let targetSheet = activeSpreadsheet.getSheetByName(pascalCase(fileName));
    if (targetSheet == null) {
      targetSheet = createNewSheet(activeSpreadsheet, fileName);
    }

    const readERDiagramValues = convertERDiagram(activeSpreadsheet.getName(), contents);
    sync(targetSheet, readERDiagramValues);
  }
}

/**
 * ER図の内容をスプレッドシートに同期する
 *
 * @param activeSpreadsheet 現在開いているGoogleSpreadsheet
 * @param readERDiagramValues ER図から読み取ったデータ
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
 * スプレッドシートの情報をオブジェクト配列化したものを取得
 *
 * @param sheet シートデータ
 * @param readERDiagramValues ER図から読み取ったデータ
 * @param readSheetValues 既存のシート上に記載されていたテーブル情報から読み取ったデータ
 * @param headersMain 親側のキー一覧
 * @param headersSub キー一覧
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

  // テーブル定義書として出力必要なテーブル名一覧を取得する。なお、テーブルデータがない場合はER図のみを対象する。
  // 読み取り前にテーブル定義側にあり、ER図にないといったケースの場合、勝手に削除は行わないものとする。
  if (Object.keys(readSheetValues).length) {
    readTableNamesBySheetValue = Object.keys(readSheetValues);
    readTableNames = readTableNamesByERDiagram.concat(readTableNamesBySheetValue);
    let tableKyesSet = new Set(readTableNames);
    readTableNames = Array.from(tableKyesSet);
  } else {
    readTableNames = readTableNamesByERDiagram;
  }

  // テーブルごとで定義を出力していく
  const lastKeyName = Object.values(readTableNames).pop();
  for (let tableName of readTableNames) {
    let isFirstRow = true;

    // ER図に存在するテーブルかどうかで出力する内容を分岐する
    if (readERDiagramValues[tableName]) {
      let readColumnNamesByERDiagram = Object.keys(readERDiagramValues[tableName]['attributes']);
      let readColumnNamesBySheetValue
      let readColumnNames;

      // テーブル定義書として出力必要なカラム名一覧を取得する。なお、テーブルデータがない場合はER図のみを対象する。
      // 出力前からテーブル定義側にあり、ER図にないといったケースの場合、勝手に削除は行わないものとする。
      if (readSheetValues[tableName]) {
        readColumnNamesBySheetValue = Object.keys(readSheetValues[tableName]["attributes"]);
        readColumnNames = readColumnNamesByERDiagram.concat(readColumnNamesBySheetValue);
        let set = new Set(readColumnNames);
        readColumnNames = Array.from(set);
      } else {
        readColumnNames = readColumnNamesByERDiagram;
      }

      // 1行ごとに出力内容を生成していく
      for (let columnName of readColumnNames) {
        let row = [];

        // 最初の行のみ親要素のデータを入れ込む
        if (isFirstRow === true) {
          for (let headerMain of headersMain) {
            if (headerMain in requiredHeaderMain && readERDiagramValues[tableName]) {
              row.push(readERDiagramValues[tableName][headerMain]);
            } else if (readSheetValues[tableName] && readSheetValues[tableName][headerMain]) {
              // ER図には存在しないがテーブル定義には存在していた場合
              row.push(readSheetValues[tableName][headerMain]);
            } else {
              // ER図には存在しない。なおかつデータとしてもテーブル定義には存在しないがヘッダーだけは存在している場合（エラー回避）
              row.push("");
            }
          }
          isFirstRow = false
        } else {
          // 親要素は二行目以降は空白とする
          for (let columnIndex = 0; columnIndex < headersMain.length; columnIndex++) {
            row.push("");
          }
        }

        for (let headerSub of headersSub) {
          if (headerSub in requiredHeaderSub && readERDiagramValues[tableName]['attributes'][columnName]) {
            row.push(readERDiagramValues[tableName]['attributes'][columnName][headerSub]);
          } else if (readSheetValues[tableName] && readSheetValues[tableName]['attributes'][columnName] && readSheetValues[tableName]['attributes'][columnName][headerSub]) {
            // ER図には存在しないがテーブル定義には存在していた場合
            row.push(readSheetValues[tableName]['attributes'][columnName][headerSub]);
          } else {
            // ER図には存在しない。なおかつデータとしてもテーブル定義には存在しないがヘッダーだけは存在している場合（エラー回避）
            row.push("");
          }
        }

        rowNumber++;
        setRowsValues.push(row);
      }
    } else {
      // ER図には存在していないテーブルであるため、テーブル定義だけを考慮する
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

    // セルの着色
    rowNumber++;
    sheet.getRange(colorStart, 1, rowNumber - colorStart, headersMainAndSub.length).setBackground("rgba(255,235,148,0.76)");
    colorStart = rowNumber + 1;

    // 最後のテーブル以外は空白行を下に挟む
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
 * スプレッドシートの情報をオブジェクト配列化したものを取得
 *
 * @param headersMain 親側のキー一覧
 * @param headersSub キー一覧
 * @param rowsData スプレッドシートの各行ごとのデータ
 * @return スプレッドシートの内容をテーブルごとに配列でまとめた情報
 */
function readGoogleSpreadsheet(headersMain, headersSub, rowsData) {
  let rowNumber = 0;
  let parentAttributes = {};

  // 範囲外になるまでループする
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
 * スプレッドシートから1つ分のデータ群を取得する（ex:1テーブル分のデータ）
 *
 * @param headersMain 親側のキー一覧
 * @param headersSub キー一覧
 * @param rowsData スプレッドシートの各行ごとのデータ
 * @param rowNumber 読み取り中の行番号
 * @return [読み取り中の行番号, 1つ分のデータ群]
 */
function createParentAttribute(headersMain, headersSub, rowsData, rowNumber) {
  let parentAttribute = {};
  for (let i = 0; i < headersMain.length; i++) {
    let headerMain = headersMain[i];
    parentAttribute[headerMain] = rowsData[rowNumber][headerMain];
  }

  while (rowsData[rowNumber] !== undefined && !isAllEmpty(rowsData[rowNumber])) {
    [rowNumber, parentAttribute["attributes"]] = createAttributeGroup(headersSub, rowsData, rowNumber, 0);
  }

  return [rowNumber, parentAttribute];
}

/**
 * 親は複数行の子グループを持つことができる。1つ分の子のグループを生成する
 *
 * @param headersSub キー一覧
 * @param rowsData スプレッドシートの各行ごとのデータ
 * @param rowNumber 読み取り中の行番号
 * @param keyIndex キーとするindex番号
 * @return [読み取り中の行番号, 1つ分の子グループ]
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
  }

  return [rowNumber, attributes];
}

/**
 * 指定範囲の値と数式の両方を取得する
 *
 * @param range Spreadsheetの範囲指定
 * @return string[] 指定範囲のデータを取得したもの
 */
function getValuesAndFormulas(range) {
  // 値だけを取得
  let valuesAndFomulas = range.getValues();

  // 数式を取得
  let tempFormulas = range.getFormulas();

  // それぞれを結合
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
 * デフォルトデータを入れた新しいシートを生成する
 *
 * @param activeSpreadsheet 現在開いているSpreadsheet
 * @param fileName umlファイルのファイル名
 * @return 作成したシートデータ
 */
function createNewSheet(activeSpreadsheet, fileName) {
  const sheet = activeSpreadsheet.insertSheet(pascalCase(fileName));
  sheet.getRange(1, 1, 1, 11).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(1, 1, 1, 11).setBackgroundRGB(182, 215, 168);
  sheet.getRange(1, 1).setValue('TableName');
  sheet.getRange(1, 2).setValue('TableDescription');
  sheet.getRange(1, 3).setValue('ConnectionName');
  sheet.getRange(1, 4).setValue('ColumnName');
  sheet.getRange(1, 5).setValue('ColumnDescription');
  sheet.getRange(1, 6).setValue('DataType');
  sheet.getRange(1, 7).setValue('MigrationDataType');
  sheet.getRange(1, 8).setValue('IsNullable');
  sheet.getRange(1, 9).setValue('IsUnsigned');
  sheet.getRange(1, 10).setValue('Version');
  sheet.getRange(1, 11).setValue('Comment');

  return sheet;
}

/**
 * 最初の行を配列化して取得する（親側）
 *
 * @param values Spreadsheetの内容を二次元配列したもの
 * @return Sheetのヘッダー一覧
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
 * 最初の行を配列化して取得する
 *
 * @param values Spreadsheetの内容を二次元配列したもの
 * @return Sheetのヘッダー一覧
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
 * 行全体がすべて空かどうか
 *
 * @param obj 行の情報が入っているオブジェクト情報
 * @return trueであればすべて空
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
 * ER図から読み取ったデータを整形する
 *
 * @param activeSpreadsheetName スプレッドシートのタイトル
 * @param contents ER図から読み取った内容
 * @return ER図の内容をテーブルごとに配列でまとめた情報
 */
function convertERDiagram(activeSpreadsheetName, contents) {
  let readERDiagrams = {};

  for (let rowIndex = 0; rowIndex < contents.length; rowIndex++) {
    // entityの記述がある行に到達した場合はそれ以降は } が現れるまではテーブル定義の内容と判断する（isLoadingフラグでそのことを管理）
    if (contents[rowIndex].match(/entity \"/)) {
      let databaseCategoryName = getDatabaseCategoryName(contents[rowIndex]);

      // 現在実行中のシートとテーブルのカテゴリが異なる場合は出力しない（例：User系のテーブルにマスターデータは出力しない）
      if (activeSpreadsheetName != databaseCategoryName) {
        continue;
      }

      const [tableName, tableDescription] = getTableName(contents[rowIndex]);
      const connectionName = getConnectionName(contents[rowIndex]);

      // ER図から読み取るものを挿入する
      parentAttribute = {};
      parentAttribute["DatabaseCategoryName"] = databaseCategoryName;
      parentAttribute["TableName"] = tableName;
      parentAttribute["TableDescription"] = tableDescription;
      parentAttribute["ConnectionName"] = connectionName;
      rowIndex++;

      attributes = {};
      while (!contents[rowIndex].match(/}/)) {
        // -がついているものはnull許容カラムであるとする
        let isNullable = 'FALSE';
        if (contents[rowIndex].match(/\-/)) {
          isNullable = 'TRUE';
        }
        // 正規表現で不要な記号を削除し、カラム情報を抽出する
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
 * データベースカテゴリを取得する（Spreadsheetのタイトル部分と一致しているかどうかの判定に使用する）
 *
 * @param content ER図から読み取った内容
 * @return データベースカテゴリ名
 */
function getDatabaseCategoryName(content) {
  content = content.split(">>")[0];

  return content.split(",")[1].trim();
};

/**
 * コネクション名を取得する
 *
 * @param content ER図から読み取った内容
 * @return コネクション名
 */
function getConnectionName(content) {
  const str = content.split(">>")[1];

  return str.split("{")[0].trim();
};

/**
 * テーブル名、テーブル詳細を取得する
 *
 * @param content ER図から読み取った内容
 * @return [テーブル名、テーブル詳細]
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
 * キャメルケースへ変換
 *
 * @param text 対象文字列
 * @return 変換後の文字列
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
 * スネークケースへ変換
 *
 * @param text 対象文字列
 * @return 変換後の文字列
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
 * パスカルケースへ変換
 *
 * @param text 対象文字列
 * @return 変換後の文字列
 */
function pascalCase(text) {
  if (!isString(text)) {
    return text;
  }
  const camelText = this.camelCase(text);

  return camelText.charAt(0).toUpperCase() + camelText.slice(1);
};

/**
 * 指定変数が文字列かどうか
 *
 * @param obj 検証したい対象
 * @return 文字列であればtrue
 */
function isString(obj) {
  return typeof (obj) == "string" || obj instanceof String;
};

/**
 * 指定位置の文字列を削除する
 *
 * @param text 削除前の文字列
 * @param targetNumber 削除したい位置
 * @return 削除後の文字列
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
 * 指定位置の文字列を置換する
 *
 * @param text 置換前の文字列
 * @param replace 置換文字
 * @param targetNumber 置換したい位置
 * @return 置換後の文字列
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
 * 2次元配列を連想配列に変換する（最初の行をキーとする）
 *
 * @param values 二次元配列
 * @return 連想配列化したオブジェクト
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


