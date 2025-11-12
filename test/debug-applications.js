/**
 * 申請データのデバッグスクリプト
 * GASエディタで実行して、申請シートの全データを確認します
 */

function debugApplicationSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('申請');

  if (!sheet) {
    Logger.log('❌ 申請シートが見つかりません');
    return;
  }

  Logger.log('✅ 申請シートが見つかりました');

  // データ範囲を取得
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  Logger.log('最終行: ' + lastRow);
  Logger.log('最終列: ' + lastCol);

  if (lastRow <= 1) {
    Logger.log('❌ データが存在しません（ヘッダー行のみ）');
    return;
  }

  // 全データを取得
  var values = sheet.getDataRange().getValues();

  Logger.log('\n========== ヘッダー行 ==========');
  Logger.log(values[0]);

  Logger.log('\n========== データ行（最大10行表示） ==========');
  for (var i = 1; i < Math.min(values.length, 11); i++) {
    var row = values[i];
    Logger.log('\n--- 行 ' + (i + 1) + ' ---');
    Logger.log('A列(利用者番号): "' + row[0] + '" (型: ' + typeof row[0] + ', 真偽: ' + !!row[0] + ')');
    Logger.log('B列(氏名): "' + row[1] + '"');
    Logger.log('C列(残日数): "' + row[2] + '"');
    Logger.log('D列(申請日): "' + row[3] + '"');
    Logger.log('E列(申請日時): "' + row[4] + '"');
    Logger.log('F列(ステータス): "' + row[5] + '"');
    Logger.log('G列(コメント): "' + row[6] + '"');
    Logger.log('H列(申請日数): "' + row[7] + '"');
    Logger.log('全体: ' + JSON.stringify(row));
  }

  Logger.log('\n========== 利用者番号123456のレコード検索 ==========');
  var found123456 = [];
  for (var i = 1; i < values.length; i++) {
    var userId = String(values[i][0]);
    if (userId === '123456' || userId.indexOf('123456') !== -1) {
      found123456.push({
        row: i + 1,
        data: values[i]
      });
    }
  }

  if (found123456.length === 0) {
    Logger.log('❌ 利用者番号123456のレコードが見つかりません');
  } else {
    Logger.log('✅ 利用者番号123456のレコードが ' + found123456.length + ' 件見つかりました');
    found123456.forEach(function(item) {
      Logger.log('\n行番号: ' + item.row);
      Logger.log('データ: ' + JSON.stringify(item.data));
    });
  }

  Logger.log('\n========== getApplications関数の実行テスト ==========');
  var result = getApplications({});
  Logger.log('取得されたレコード数: ' + result.records.length);
  Logger.log('総レコード数: ' + result.pagination.total);

  if (result.records.length > 0) {
    Logger.log('\n最初の3件:');
    for (var i = 0; i < Math.min(result.records.length, 3); i++) {
      Logger.log('レコード' + (i + 1) + ': ' + JSON.stringify(result.records[i]));
    }
  }

  // 123456のレコードが含まれているか確認
  var found123456InResult = result.records.filter(function(r) {
    return r.userId === '123456' || r.userId.indexOf('123456') !== -1;
  });

  if (found123456InResult.length > 0) {
    Logger.log('\n✅ getApplicationsの結果に123456が含まれています (' + found123456InResult.length + '件)');
    found123456InResult.forEach(function(r) {
      Logger.log(JSON.stringify(r));
    });
  } else {
    Logger.log('\n❌ getApplicationsの結果に123456が含まれていません');
  }
}

/**
 * 特定の利用者番号のレコードを検索
 */
function searchByUserId(targetUserId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('申請');

  if (!sheet) {
    Logger.log('❌ 申請シートが見つかりません');
    return;
  }

  var values = sheet.getDataRange().getValues();

  Logger.log('検索対象利用者番号: "' + targetUserId + '"');
  Logger.log('総行数（ヘッダー含む）: ' + values.length);

  var matches = [];
  for (var i = 1; i < values.length; i++) {
    var userId = String(values[i][0]).trim();
    if (userId === String(targetUserId)) {
      matches.push({
        row: i + 1,
        userId: values[i][0],
        userName: values[i][1],
        remaining: values[i][2],
        applyDate: values[i][3],
        timestamp: values[i][4],
        status: values[i][5],
        comment: values[i][6],
        applyDays: values[i][7]
      });
    }
  }

  Logger.log('\n検索結果: ' + matches.length + '件');
  matches.forEach(function(m) {
    Logger.log('\n行番号: ' + m.row);
    Logger.log(JSON.stringify(m, null, 2));
  });

  return matches;
}
