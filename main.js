var rowCount = 20;

function setup() { 
  // 初期設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bodySheet = ss.getSheetByName('BodyText');
  let varSheet = ss.getSheetByName('Variable');
  let resultSheet = ss.getSheetByName('Result');
  let scriptSheet = ss.getSheetByName('Script');

  if (bodySheet === null) {
    bodySheet = ss.insertSheet();
    bodySheet.setName('BodyText');
    bodySheet.getRange(1, 1).setValue('BodyTextSheet');
    bodySheet.getRange(2, 1).setValue('A3以下に段落ごとに文章を挿入').setBorder(true, true, true, true, false, false);
    bodySheet.getRange(3, 1, rowCount).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    bodySheet.setColumnWidth(1, 500);
  }

  if (varSheet === null) {
    varSheet = ss.insertSheet();
    varSheet.setName('Variable');
    varSheet.getRange(1, 1).setValue('VariableSheet');
    varSheet.getRange(2, 1, 1, 2).setValues([['識別名', '挿入内容']]).setBorder(true, true, true, true, true, false);
  }

  if (resultSheet === null) {
    resultSheet = ss.insertSheet();
    resultSheet.setName('Result');
    resultSheet.getRange(1, 1).setValue('ResultSheet');
    resultSheet.getRange(2, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    resultSheet.setColumnWidth(1, 500);
  }

  if (scriptSheet === null) {
    scriptSheet = ss.insertSheet();
    scriptSheet.setName('Script');
    scriptSheet.getRange(2, 2).setValue('Coded by chishige1217200');
    scriptSheet.getRange(3, 2).setValue('https://github.com/chishige1217200/FixedPhraserGAS');
  }

  console.log('BodyTextシートとVariableシートを記入してください．');
}

function createFixedPhrase() {
  // シートが準備されているか確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bodySheet = ss.getSheetByName('BodyText');
  let varSheet = ss.getSheetByName('Variable');
  let resultSheet = ss.getSheetByName('Result');
  if (bodySheet === null || varSheet === null || resultSheet === null) {
    setup();
    return;
  }

  // 情報をシートから読み出す
  let body = [];
  if (bodySheet.getLastRow() >= 3) {
    body = bodySheet.getRange(3, 1, bodySheet.getLastRow() - 2, 1).getValues();
  }
  if (body.length === 0) {
    console.log('文章情報が存在しません．処理は中止されました．');
    return;
  }

  let variable = [];
  if (varSheet.getLastRow() >= 3) {
    variable = varSheet.getRange(3, 1, varSheet.getLastRow() - 2, 2).getValues();
  }

  //console.log(body);
  //console.log(variable);

  // 文字列置換を行う
  if (variable.length !== 0) {
    for (let i = 0; i < variable.length; i++) {
      for (let j = 0; j < body.length; j++) {
        let regexp = new RegExp('%_' + variable[i][0] + '_%', 'ig'); // RegExpオブジェクトでreplaceを行わないと先頭要素しか置換されない
        body[j][0] = body[j][0].replace(regexp, variable[i][1]);
      }
    }
  }
  else {
    console.log('変数情報が存在しません．文字列置換はスキップされました．');
  }

  // 文字列連結を行う
  let fixedPhrase = '';

  for (let i = 0; i < body.length; i++) {
    fixedPhrase += (body[i][0] + '\n\n');
  }

  resultSheet.getRange(2, 1).setValue(fixedPhrase);
  resultSheet.autoResizeRows(2, 1);

  console.log('Resultシートに文章を出力しました．');
}