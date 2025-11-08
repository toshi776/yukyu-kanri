/**
 * テストデータ生成のテスト
 * GASエディタから直接実行してください
 */

/**
 * テストデータ生成を実行してログに詳細を出力
 */
function testGenerateTestData() {
  console.log('=== テストデータ生成テスト開始 ===');

  try {
    // 1. スプレッドシート存在確認
    console.log('1. スプレッドシート確認中...');
    var ss = getSpreadsheet();
    console.log('✓ スプレッドシート取得成功:', ss.getName());

    // 2. マスターシート確認
    console.log('2. マスターシート確認中...');
    var masterSheet = ss.getSheetByName('マスター');
    if (!masterSheet) {
      console.error('✗ マスターシートが見つかりません');
      return;
    }
    console.log('✓ マスターシート確認成功');

    // 3. 現在のデータ数を確認
    var beforeMasterRows = masterSheet.getLastRow();
    console.log('現在のマスターシート行数:', beforeMasterRows);

    // 4. テストデータ生成実行
    console.log('3. テストデータ生成実行中...');
    var result = generateTestData();

    // 5. 結果確認
    console.log('4. 生成結果確認:');
    console.log(JSON.stringify(result, null, 2));

    if (result.success) {
      console.log('✓ テストデータ生成成功');

      // 6. データが実際に追加されたか確認
      var afterMasterRows = masterSheet.getLastRow();
      console.log('生成後のマスターシート行数:', afterMasterRows);
      console.log('追加された行数:', (afterMasterRows - beforeMasterRows));

      // 7. 付与履歴シートを確認
      var grantHistorySheet = ss.getSheetByName('付与履歴');
      if (grantHistorySheet) {
        var grantRows = grantHistorySheet.getLastRow();
        console.log('付与履歴シート行数:', grantRows);
      }

      console.log('\n=== 生成されたデータ ===');
      console.log('テストユーザー数:', result.details.masterUsers);
      console.log('付与履歴レコード数:', result.details.grantRecords);

      if (result.details.scenarios) {
        console.log('\nテストシナリオ:');
        result.details.scenarios.forEach(function(scenario, index) {
          console.log((index + 1) + '. ' + scenario);
        });
      }

      console.log('\n✅ テストデータ生成テスト完了');
    } else {
      console.error('✗ テストデータ生成失敗:', result.message);
      if (result.error) {
        console.error('エラー詳細:', result.error);
      }
    }

  } catch (error) {
    console.error('✗ テスト実行エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * マスターシートの内容を確認
 */
function checkMasterSheet() {
  console.log('=== マスターシート内容確認 ===');

  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');

    if (!masterSheet) {
      console.error('マスターシートが見つかりません');
      return;
    }

    var lastRow = masterSheet.getLastRow();
    console.log('総行数:', lastRow);

    // ヘッダー行を取得
    var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    console.log('ヘッダー:', headers);

    // 最後の10行を表示
    var displayRows = Math.min(10, lastRow - 1);
    if (displayRows > 0) {
      console.log('\n最後の' + displayRows + '行:');
      var data = masterSheet.getRange(lastRow - displayRows + 1, 1, displayRows, masterSheet.getLastColumn()).getValues();

      data.forEach(function(row, index) {
        console.log((lastRow - displayRows + index + 1) + ':', row);
      });
    }

    // TEST01-TEST06を検索
    console.log('\nテストユーザー検索:');
    var allData = masterSheet.getDataRange().getValues();
    var testUsers = ['TEST01', 'TEST02', 'TEST03', 'TEST04', 'TEST05', 'TEST06'];

    testUsers.forEach(function(testId) {
      var found = false;
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === testId) {
          console.log('✓ ' + testId + ' が見つかりました:', allData[i][1]);
          found = true;
          break;
        }
      }
      if (!found) {
        console.log('✗ ' + testId + ' が見つかりません');
      }
    });

  } catch (error) {
    console.error('確認エラー:', error);
  }
}

/**
 * 付与履歴シートの内容を確認
 */
function checkGrantHistorySheet() {
  console.log('=== 付与履歴シート内容確認 ===');

  try {
    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!grantHistorySheet) {
      console.error('付与履歴シートが見つかりません');
      return;
    }

    var lastRow = grantHistorySheet.getLastRow();
    console.log('総行数:', lastRow);

    // ヘッダー行を取得
    var headers = grantHistorySheet.getRange(1, 1, 1, grantHistorySheet.getLastColumn()).getValues()[0];
    console.log('ヘッダー:', headers);

    // テストユーザーのデータを検索
    console.log('\nテストユーザーの付与履歴:');
    var allData = grantHistorySheet.getDataRange().getValues();
    var testUsers = ['TEST01', 'TEST02', 'TEST03', 'TEST04', 'TEST05', 'TEST06'];

    testUsers.forEach(function(testId) {
      var count = 0;
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === testId) {
          count++;
        }
      }
      console.log(testId + ': ' + count + '件');
    });

  } catch (error) {
    console.error('確認エラー:', error);
  }
}

/**
 * テストデータをクリア（削除）
 */
function clearTestData() {
  if (!Browser.msgBox('テストデータを削除しますか？', Browser.Buttons.YES_NO) === 'yes') {
    console.log('キャンセルされました');
    return;
  }

  console.log('=== テストデータクリア開始 ===');

  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    var deletedMaster = 0;
    var deletedGrant = 0;

    // マスターシートからTEST01-TEST06を削除
    if (masterSheet) {
      var data = masterSheet.getDataRange().getValues();
      var testUsers = ['TEST01', 'TEST02', 'TEST03', 'TEST04', 'TEST05', 'TEST06'];

      for (var i = data.length - 1; i >= 1; i--) {
        if (testUsers.indexOf(String(data[i][0])) !== -1) {
          masterSheet.deleteRow(i + 1);
          deletedMaster++;
        }
      }
    }

    // 付与履歴シートからTEST01-TEST06のデータを削除
    if (grantHistorySheet) {
      var grantData = grantHistorySheet.getDataRange().getValues();
      var testUsers = ['TEST01', 'TEST02', 'TEST03', 'TEST04', 'TEST05', 'TEST06'];

      for (var i = grantData.length - 1; i >= 1; i--) {
        if (testUsers.indexOf(String(grantData[i][0])) !== -1) {
          grantHistorySheet.deleteRow(i + 1);
          deletedGrant++;
        }
      }
    }

    console.log('✓ マスターシートから削除:', deletedMaster + '行');
    console.log('✓ 付与履歴シートから削除:', deletedGrant + '行');
    console.log('✅ テストデータクリア完了');

  } catch (error) {
    console.error('✗ クリアエラー:', error);
  }
}
