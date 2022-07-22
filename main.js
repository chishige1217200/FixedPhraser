function setup() { // 初期設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let bodySheet = ss.getSheetByName('BodyText');
    let varSheet = ss.getSheetByName('Variable');
    let resultSheet = ss.getSheetByName('Result');
    let scriptSheet = ss.getSheetByName('Script');

    if (bodySheet === null) {
        bodySheet = ss.insertSheet();
        bodySheet.setName('BodyText');
        bodySheet.getRange(1, 1).setValue('BodyTextSheet');
        bodySheet.getRange(2, 1).setValue('C1以下に段落ごとに文章を挿入');
    }

    if (varSheet === null) {
        varSheet = ss.insertSheet();
        varSheet.setName('Variable');
        varSheet.getRange(1, 1).setValue('VariableSheet');
        varSheet.getRange(2, 1, 1, 2).setValues([['識別名', '挿入内容']]);
    }

    if (resultSheet === null) {
        resultSheet = ss.insertSheet();
        resultSheet.setName('Result');
        resultSheet.getRange(1, 1).setValue('ResultSheet');
        resultSheet.getRange(2, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }

    if (scriptSheet === null) {
        scriptSheet = ss.insertSheet();
        scriptSheet.setName('Script');
        scriptSheet.getRange(2, 2, 1, 1).setValue('Coded by chishige1217200');
        scriptSheet.getRange(3, 2, 1, 1).setValue('https://github.com/chishige1217200/AttendManageGAS');
    }

    console.log('Setup Complete.');
}

function createFixedPhrase() {
    // シートが準備されているか確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let bodySheet = ss.getSheetByName('BodyText');
    let varSheet = ss.getSheetByName('Variable');
    let resultSheet = ss.getSheetByName('Result');
    if (bodySheet === null) {
        setup();
        return;
    }
    if (varSheet === null) {
        setup();
        return;
    }
    if (resultSheet === null) {
        setup();
        return;
    }

    // 情報をシートから読み出す
    let body = bodySheet.getRange(3, 1, bodySheet.getLastRow() - 2, 1).getValues();
    let variable = varSheet.getRange(3, 1, varSheet.getLastRow() - 2, 2).getValues();

    console.log(body);
    console.log(variable);

    //resultSheet.autoResizeColumn(1);

}