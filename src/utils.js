/** スプレッドシートID */
var SPREADSHEET_ID = '1ENEljNya5MPY2Z7FPZR0LAguFPeCunwbJ3B9kgSjyK0';

/**
 * 安全にスプレッドシートを取得
 * @returns {Spreadsheet}
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    console.error('スプレッドシートの取得に失敗:', e);
    throw new Error('スプレッドシートにアクセスできません');
  }
}

/**
 * マスターシートから利用者情報を取得
 * @param {string} id
 * @returns {{name:string, remaining:number}|null}
 */
function getMasterRecord(id) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('マスター');
    if (!sheet) {
      console.error('マスターシートが見つかりません');
      return null;
    }
    
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        return { name: rows[i][1], remaining: rows[i][2] };
      }
    }
    return null;
  } catch (e) {
    console.error('マスターレコード取得エラー:', e);
    return null;
  }
}

/**
 * 申請データを追加し、マスターシートの残日数を１日減算する
 * @param {Object} params e.parameter をそのまま渡す
 */
function appendApplication(params) {
  try {
    var ss = getSpreadsheet();
    
    // 1) 申請シートに行を追加
    var appSheet = ss.getSheetByName('申請');
    if (!appSheet) {
      throw new Error('申請シートが見つかりません');
    }
    
    appSheet.appendRow([
      params.userId,
      params.userName,
      Number(params.remaining),
      params.applyDate,
      new Date(),
      'Pending',
      ''
    ]);

    // 2) マスターシートの残日数を１日減算
    var masterSheet = ss.getSheetByName('マスター');
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var finder = masterSheet.createTextFinder(params.userId).matchEntireCell(true);
    var foundCell = finder.findNext();
    if (foundCell) {
      var rowIndex = foundCell.getRow();
      var remCell = masterSheet.getRange(rowIndex, 3);
      var oldVal = Number(remCell.getValue());
      var newVal = oldVal - 1;
      remCell.setValue(newVal);
    }
  } catch (e) {
    console.error('申請追加エラー:', e);
    throw e;
  }
}

/**
 * 申請一覧取得（管理画面用）
 * @param {Object} params フィルター条件
 * @returns {Object} 申請データとページネーション情報
 */
function getApplications(params) {
  try {
    console.log('getApplications呼び出し - params:', params);
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('申請'); // シート名を'申請'に修正
    
    if (!sheet) {
      console.error('申請シートが見つかりません');
      return { records: [], pagination: { page: 1, totalPages: 0, total: 0 } };
    }
    
    var values = sheet.getDataRange().getValues();
    console.log('取得した行数:', values.length);
    
    if (values.length <= 1) {
      console.log('データが存在しません');
      return { records: [], pagination: { page: 1, totalPages: 0, total: 0 } };
    }
    
    values.shift(); // ヘッダー行を削除
    var data = [];
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (row[0]) { // 利用者番号が存在する行のみ処理
        data.push({
          recordId: i + 2, // スプレッドシートの行番号（ヘッダー分+1、0-based index分+1）
          userId: String(row[0]),
          userName: String(row[1] || ''),
          remaining: Number(row[2] || 0),
          applyDate: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          timestamp: row[4] ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
          status: String(row[5] || 'Pending'),
          applyDays: Number(row[7] || 1)  // 申請日数（H列、デフォルト1日）
        });
      }
    }
    
    console.log('処理後のデータ件数:', data.length);

    // フィルター処理
    if (params && params.filter) {
      var f = params.filter.toLowerCase();
      data = data.filter(function(r) {
        return r.userId.toLowerCase().indexOf(f) !== -1 || r.userName.toLowerCase().indexOf(f) !== -1;
      });
    }

    if (params && params.status) {
      data = data.filter(function(r) {
        return r.status === params.status;
      });
    }

    // 有給希望日の昇順でソート（古い日付が上に）
    data.sort(function(a, b) {
      var dateA = new Date(a.applyDate);
      var dateB = new Date(b.applyDate);
      return dateA - dateB; // 昇順
    });

    // ページング（100件に増量）
    var pageSize = 100;
    var page = (params && params.page) ? Number(params.page) : 1;
    var total = data.length;
    var totalPages = Math.ceil(total / pageSize);
    var start = (page - 1) * pageSize;
    var pageData = data.slice(start, start + pageSize);

    console.log('最終的な返却データ件数:', pageData.length);
    
    return { 
      records: pageData, 
      pagination: { page: page, totalPages: totalPages, total: total } 
    };
    
  } catch (e) {
    console.error('申請一覧取得エラー:', e);
    return { records: [], pagination: { page: 1, totalPages: 0, total: 0 } };
  }
}

/**
 * 古い申請データを削除（2025/06/17以前）
 * @returns {number} 削除された行数
 */
function deleteOldApplications() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('申請');
    
    if (!sheet) {
      console.error('申請シートが見つかりません');
      return 0;
    }
    
    var data = sheet.getDataRange().getValues();
    var cutoffDate = new Date('2025-06-17');
    cutoffDate.setHours(0, 0, 0, 0);
    
    var rowsToDelete = [];
    
    // データをチェック（ヘッダー行をスキップ）
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[3]) { // 申請日がある場合
        var applyDate = new Date(row[3]);
        applyDate.setHours(0, 0, 0, 0);
        
        if (applyDate < cutoffDate) {
          rowsToDelete.push(i + 1); // スプレッドシートの行番号（1ベース）
        }
      }
    }
    
    console.log('削除対象行:', rowsToDelete.length + '行');
    
    // 後ろから削除（行番号が変わらないようにするため）
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    return rowsToDelete.length;
  } catch (e) {
    console.error('古いデータ削除エラー:', e);
    return 0;
  }
}

/**
 * レコード更新
 * @param {Object} obj 更新情報
 * @returns {boolean} 成功/失敗
 */
function updateRecord(obj) {
  try {
    var recordId = obj.recordId;
    var field = obj.field;
    var value = obj.value;
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('申請'); // シート名を'申請'に修正
    
    if (!sheet) {
      console.error('申請シートが見つかりません');
      return false;
    }
    
    var col;
    if (field === 'remaining') col = 3;
    else if (field === 'applyDate') col = 4;
    else return false;
    
    if (field === 'applyDate') {
      sheet.getRange(recordId, col).setValue(new Date(value));
    } else {
      sheet.getRange(recordId, col).setValue(value);
    }
    
    return true;
  } catch (e) {
    console.error('レコード更新エラー:', e);
    return false;
  }
}

/**
 * 申請承認
 * @param {Object} obj レコード情報
 * @returns {boolean} 成功/失敗
 */
function approveRecord(obj) {
  var lock = LockService.getScriptLock();
  try {
    // 最大10秒間ロックを試行
    lock.waitLock(10000);
    
    var rowNumber = obj.rowNumber;
    var status = obj.status || 'Approved';
    var rejectionReason = obj.reason || '';
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('申請'); // シート名を'申請'に修正
    
    if (!sheet) {
      console.error('申請シートが見つかりません');
      return false;
    }
    
    console.log('承認処理:', rowNumber, status);
    
    // 現在のステータスを確認（既に処理済みかチェック）
    var currentStatus = sheet.getRange(rowNumber, 6).getValue();
    if (currentStatus !== 'Pending') {
      console.log('既に処理済みの申請です:', currentStatus);
      throw new Error('この申請は既に処理されています（現在のステータス: ' + currentStatus + '）');
    }
    
    var userId = sheet.getRange(rowNumber, 1).getValue(); // 利用者番号取得
    var applyDays = sheet.getRange(rowNumber, 8).getValue() || 1; // 申請日数取得（H列）
    
    // 却下の場合は有給残日数を復旧
    if (status === 'Rejected') {
      var masterSheet = ss.getSheetByName('マスター');
      
      if (masterSheet) {
        var masterData = masterSheet.getDataRange().getValues();
        for (var i = 1; i < masterData.length; i++) {
          if (String(masterData[i][0]) === String(userId)) {
            var currentRemaining = masterData[i][2];
            masterSheet.getRange(i + 1, 3).setValue(currentRemaining + applyDays); // 申請日数分を復旧
            console.log('有給残日数を復旧:', userId, currentRemaining + applyDays, '（申請日数:', applyDays, '）');
            break;
          }
        }
      }
    }
    
    sheet.getRange(rowNumber, 6).setValue(status); // ステータス列（F列）
    sheet.getRange(rowNumber, 7).setValue(new Date()); // 更新ログ列（G列）
    
    // 承認結果通知を送信（エラーがあっても承認処理は継続）
    try {
      if (typeof sendApprovalResultNotification === 'function') {
        var applicantData = {
          userId: userId,
          userName: sheet.getRange(rowNumber, 2).getValue(), // B列から名前取得
          applyDate: Utilities.formatDate(sheet.getRange(rowNumber, 4).getValue(), 'JST', 'yyyy/MM/dd'),
          applyDays: applyDays
        };
        var notifyResult = sendApprovalResultNotification(applicantData, status, rejectionReason);
        console.log('承認結果通知結果:', notifyResult);
      }
    } catch (notifyError) {
      console.error('承認結果通知エラー(無視):', notifyError);
    }
    
    console.log('承認処理完了');
    return true;
  } catch (e) {
    console.error('承認処理エラー:', e);
    throw e; // エラーを再スローして呼び出し元でキャッチ
  } finally {
    // ロックを解放
    lock.releaseLock();
  }
}

/**
 * 申請取消
 * @param {Object} obj レコード情報
 * @returns {boolean} 成功/失敗
 */
function cancelRecord(obj) {
  try {
    var recordId = obj.recordId;
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('申請'); // シート名を'申請'に修正
    
    if (!sheet) {
      console.error('申請シートが見つかりません');
      return false;
    }
    
    sheet.getRange(recordId, 6).setValue('Canceled'); // ステータス列
    sheet.getRange(recordId, 5).setValue(new Date()); // タイムスタンプ列
    return true;
  } catch (e) {
    console.error('取消処理エラー:', e);
    return false;
  }
}

/**
 * 利用者マスター取得
 * @returns {Array} マスターデータ配列
 */
function getMasterData() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('マスター');
    
    if (!sheet) {
      console.error('マスターシートが見つかりません');
      return [];
    }
    
    var values = sheet.getDataRange().getValues();
    values.shift(); // ヘッダー削除
    
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (row[0]) { // 利用者番号が存在する行のみ
        result.push({ 
          userId: String(row[0]), 
          userName: String(row[1] || ''), 
          remaining: Number(row[2] || 0) 
        });
      }
    }
    return result;
  } catch (e) {
    console.error('マスターデータ取得エラー:', e);
    return [];
  }
}

/**
 * 利用者一覧を取得（備考欄対応版）
 * @param {Object} params - フィルター条件
 * @return {Object} 利用者データと件数
 */
function getEmployees(params) {
  console.log('getEmployees called with params:', params);
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('マスター');
    
    if (!sheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = [];
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 空行はスキップ
      if (!row[0] && !row[1]) continue;
      
      var record = {
        rowNumber: i + 1,  // 行番号（編集時に使用）
        userId: String(row[0] || ''),  // 文字列に変換
        userName: String(row[1] || ''), // 文字列に変換
        remaining: Number(row[2] || 0),
        remarks: String(row[3] || '')   // 備考欄を追加（D列）
      };
      
      // フィルター処理
      if (params && params.filter !== undefined && params.filter !== '') {
        var filterText = String(params.filter).toLowerCase();
        
        var userIdStr = String(record.userId).toLowerCase();
        var userNameStr = String(record.userName).toLowerCase();
        var remarksStr = String(record.remarks).toLowerCase();
        
        var userIdMatch = userIdStr.indexOf(filterText) !== -1;
        var userNameMatch = userNameStr.indexOf(filterText) !== -1;
        var remarksMatch = remarksStr.indexOf(filterText) !== -1;
        
        if (!userIdMatch && !userNameMatch && !remarksMatch) {
          continue;
        }
      }
      
      // 残日数フィルター（追加機能）
      if (params && params.remainingFilter) {
        if (params.remainingFilter === 'low' && record.remaining > 5) continue;
        if (params.remainingFilter === 'zero' && record.remaining > 0) continue;
      }
      
      records.push(record);
    }
    
    // ソート処理（追加機能）
    if (params && params.sortBy) {
      records.sort(function(a, b) {
        if (params.sortBy === 'remaining') {
          return a.remaining - b.remaining; // 残日数昇順
        } else if (params.sortBy === 'userName') {
          return a.userName.localeCompare(b.userName);
        }
        return 0;
      });
    }
    
    console.log('getEmployees found:', records.length, 'records');
    
    // 統計情報も計算
    var stats = {
      total: records.length,
      zeroCount: records.filter(function(r) { return r.remaining === 0; }).length,
      lowCount: records.filter(function(r) { return r.remaining <= 5; }).length
    };
    
    return {
      records: records,
      total: records.length,
      stats: stats
    };
    
  } catch (error) {
    console.error('getEmployees error:', error);
    throw error;
  }
}

/**
 * 利用者の備考を更新
 * @param {Object} params - {rowNumber, newRemarks}
 * @return {Object} 成功/失敗
 */
function updateEmployeeRemarks(params) {
  console.log('updateEmployeeRemarks called with params:', params);
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('マスター');
    
    if (!sheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var rowNumber = parseInt(params.rowNumber);
    var newRemarks = String(params.newRemarks || '');
    
    // 備考を更新（D列 = 4列目）
    sheet.getRange(rowNumber, 4).setValue(newRemarks);
    
    console.log('updateEmployeeRemarks success - row:', rowNumber);
    
    return {
      success: true,
      message: '備考を更新しました'
    };
    
  } catch (error) {
    console.error('updateEmployeeRemarks error:', error);
    throw error;
  }
}

/**
 * 利用者の有給残日数を更新
 * @param {Object} params - {rowNumber, newRemaining}
 * @return {Object} 成功/失敗
 */
function updateEmployeeRemaining(params) {
  console.log('updateEmployeeRemaining called with params:', params);
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('マスター');
    
    if (!sheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    // 行番号と新しい残日数を取得
    var rowNumber = parseInt(params.rowNumber);
    var newRemaining = parseFloat(params.newRemaining);
    
    if (isNaN(rowNumber) || isNaN(newRemaining)) {
      throw new Error('無効な数値です');
    }
    
    if (newRemaining < 0) {
      throw new Error('残日数は0以上である必要があります');
    }
    
    // 残日数を更新（C列 = 3列目）
    sheet.getRange(rowNumber, 3).setValue(newRemaining);
    
    console.log('updateEmployeeRemaining success - row:', rowNumber, 'new value:', newRemaining);
    
    return {
      success: true,
      message: '更新しました'
    };
    
  } catch (error) {
    console.error('updateEmployeeRemaining error:', error);
    throw error;
  }
}

// ============================
// 申請フォーム拡張用の関数
// ============================

/**
 * 申請を取り消す
 * @param {Object} params - {userId, rowNumber}
 * @return {Object} 処理結果
 */
function cancelUserApplication(params) {
  console.log('cancelUserApplication called with params:', params);
  
  var lock = LockService.getScriptLock();
  try {
    // 最大10秒間ロックを試行
    lock.waitLock(10000);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var applySheet = ss.getSheetByName('申請');
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!applySheet || !masterSheet) {
      throw new Error('必要なシートが見つかりません');
    }
    
    var rowNumber = parseInt(params.rowNumber);
    
    // 申請データを取得（H列まで取得して申請日数も含める）
    var applyRow = applySheet.getRange(rowNumber, 1, 1, 8).getValues()[0];
    var userId = applyRow[0];
    var applyDate = new Date(applyRow[3]);
    var currentStatus = applyRow[5];
    var applyDays = applyRow[7] || 1; // 申請日数（H列、デフォルト1日）
    
    // 既に処理済みかチェック
    if (currentStatus === 'Canceled') {
      throw new Error('この申請は既に取り消されています');
    }
    if (currentStatus === 'Approved') {
      throw new Error('承認済みの申請は取り消せません');
    }
    if (currentStatus === 'Rejected') {
      throw new Error('却下済みの申請は取り消せません');
    }
    if (currentStatus !== 'Pending') {
      throw new Error('この申請は処理できない状態です（現在のステータス: ' + currentStatus + '）');
    }
    
    // 申請日が今日以降かチェック（前日まで取消可能）
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    applyDate.setHours(0, 0, 0, 0);
    
    if (applyDate <= today) {
      throw new Error('申請日の前日までしか取り消しできません');
    }
    
    // ユーザーIDが一致するか確認
    if (String(userId) !== String(params.userId)) {
      throw new Error('他のユーザーの申請は取り消せません');
    }
    
    // マスターシートでユーザーを検索
    var masterData = masterSheet.getDataRange().getValues();
    var userRowIndex = -1;
    
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === String(userId)) {
        userRowIndex = i;
        break;
      }
    }
    
    if (userRowIndex === -1) {
      throw new Error('ユーザーが見つかりません');
    }
    
    // 現在の残日数を取得して申請日数分戻す
    var currentRemaining = masterData[userRowIndex][2];
    var newRemaining = currentRemaining + applyDays;
    
    // マスターシートの残日数を更新
    masterSheet.getRange(userRowIndex + 1, 3).setValue(newRemaining);
    
    // 申請シートのステータスを「Canceled」に更新
    applySheet.getRange(rowNumber, 6).setValue('Canceled');
    
    // 更新ログに取消情報を追加
    var updateLog = applyRow[6] || '';
    var cancelLog = '\n[' + Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss') + '] 取消';
    applySheet.getRange(rowNumber, 7).setValue(updateLog + cancelLog);
    
    console.log('cancelUserApplication success - userId:', userId, 'new remaining:', newRemaining);
    
    return {
      success: true,
      message: '申請を取り消しました。有給残日数: ' + newRemaining + '日',
      newRemaining: newRemaining
    };
    
  } catch (error) {
    console.error('cancelUserApplication error:', error);
    throw error;
  } finally {
    // ロックを解放
    lock.releaseLock();
  }
}


/**
 * 利用者番号からユーザー情報を検索
 * @param {string} userId - 利用者番号
 * @return {Object} {found: boolean, user: Object}
 */
function findUser(userId) {
  console.log('findUser called with userId:', userId);
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('マスター');
    
    if (!sheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップして検索
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 利用者番号が一致するか確認（文字列として比較）
      if (String(row[0]) === String(userId)) {
        console.log('User found:', row);
        return {
          found: true,
          user: {
            userId: row[0],
            name: row[1],
            remaining: row[2] || 0
          }
        };
      }
    }
    
    console.log('User not found');
    return {
      found: false,
      user: null
    };
    
  } catch (error) {
    console.error('findUser error:', error);
    throw error;
  }
}

/**
 * 有給申請を処理（強化版）
 * @param {Object} params - {userId, applyDate, applyDays}
 * @return {Object} {success, newRemaining, message}
 */
function submitRequest(params) {
  console.log('submitRequest called with params:', params);
  
  // 高度なバリデーションを実行
  var validation = validateLeaveApplication(params);
  if (!validation.isValid) {
    throw new Error(validation.errors.join(', '));
  }
  
  // 警告がある場合はログに記録
  if (validation.hasWarnings) {
    console.warn('申請時の警告:', validation.warnings.join(', '));
  }
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName('マスター');
    var applySheet = ss.getSheetByName('申請');
    
    if (!masterSheet || !applySheet) {
      throw new Error('必要なシートが見つかりません');
    }
    
    // ユーザー情報を取得
    var masterData = masterSheet.getDataRange().getValues();
    var userRowIndex = -1;
    var userName = '';
    var currentRemaining = 0;
    
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === String(params.userId)) {
        userRowIndex = i;
        userName = masterData[i][1];
        currentRemaining = masterData[i][2] || 0;
        break;
      }
    }
    
    if (userRowIndex === -1) {
      throw new Error('ユーザーが見つかりません');
    }
    
    // 申請日数を取得（デフォルトは1日）
    var applyDays = Number(params.applyDays) || 1;
    
    // 申請日数の再検証
    if (applyDays !== 1 && applyDays !== 0.5) {
      throw new Error('申請日数が不正です: ' + applyDays);
    }
    
    // 残日数チェック
    if (currentRemaining < applyDays) {
      throw new Error('有給残日数が不足しています（申請: ' + applyDays + '日, 残: ' + currentRemaining + '日）');
    }
    
    // 祝日・会社休日チェック（サーバーサイド検証）
    var holidayCheck = validateHolidaysAndWeekends(new Date(params.applyDate));
    if (!holidayCheck.isValid && !holidayCheck.allowOverride) {
      throw new Error(holidayCheck.message);
    }

    // 重複日付チェック（サーバーサイド）- 合計日数が1日を超えないかチェック
    var applySheetData = applySheet.getDataRange().getValues();
    var requestDate = params.applyDate;
    var totalDaysForDate = 0;

    for (var i = 1; i < applySheetData.length; i++) {
      var row = applySheetData[i];
      // 同じユーザーで、同じ申請日で、ステータスがPendingまたはApprovedの申請の日数を合計
      if (String(row[0]) === String(params.userId)) {
        var existingDate = row[3];
        var existingStatus = String(row[5] || 'Pending');
        var existingDays = Number(row[7] || 1); // H列の申請日数

        // 日付を文字列化して比較
        var existingDateStr = '';
        if (existingDate instanceof Date) {
          existingDateStr = Utilities.formatDate(existingDate, 'JST', 'yyyy-MM-dd');
        } else if (typeof existingDate === 'string') {
          existingDateStr = existingDate;
        }

        var requestDateStr = requestDate;

        if (existingDateStr === requestDateStr && (existingStatus === 'Pending' || existingStatus === 'Approved')) {
          totalDaysForDate += existingDays;
        }
      }
    }

    // 今回の申請と合わせて1日を超える場合はエラー
    if (totalDaysForDate + applyDays > 1) {
      if (totalDaysForDate >= 1) {
        throw new Error('この日付（' + requestDate + '）はすでに1日分の申請があります');
      } else if (applyDays === 1) {
        throw new Error('この日付（' + requestDate + '）には既に半日申請があります。1日申請はできません');
      } else {
        throw new Error('この日付（' + requestDate + '）はすでに半日申請が2回あります');
      }
    }

    // 新しい残日数を計算
    var newRemaining = currentRemaining - applyDays;
    
    // マスターシートの残日数を更新
    masterSheet.getRange(userRowIndex + 1, 3).setValue(newRemaining);
    
    // 申請シートに新しい申請を追加
    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss');
    
    var newRow = [
      params.userId,                    // A列: 利用者番号
      userName,                         // B列: 利用者名
      currentRemaining,                 // C列: 申請時の残日数
      params.applyDate,                 // D列: 申請日
      timestamp,                        // E列: 申請日時
      'Pending',                        // F列: ステータス
      '[' + timestamp + '] 申請',       // G列: 更新ログ
      applyDays                         // H列: 申請日数
    ];
    
    applySheet.appendRow(newRow);
    
    console.log('submitRequest success - userId:', params.userId, 'new remaining:', newRemaining);
    
    // 申請通知を送信（エラーがあっても申請処理は継続）
    try {
      if (typeof sendApplicationNotification === 'function') {
        var notificationData = {
          userId: params.userId,
          userName: userName,
          applyDate: params.applyDate,
          applyDays: applyDays,
          timestamp: timestamp
        };
        var notifyResult = sendApplicationNotification(notificationData);
        console.log('申請通知結果:', notifyResult);
      }
    } catch (notifyError) {
      console.error('申請通知エラー(無視):', notifyError);
    }
    
    return {
      success: true,
      newRemaining: newRemaining,
      message: '申請が完了しました'
    };
    
  } catch (error) {
    console.error('submitRequest error:', error);
    throw error;
  }
}

/**
 * 祝日・会社休日のサーバーサイド検証
 * @param {Date} dateObj - 申請日
 * @return {Object} 検証結果
 */
function validateHolidaysAndWeekends(dateObj) {
  try {
    var dayOfWeek = dateObj.getDay();
    var month = dateObj.getMonth() + 1;
    var date = dateObj.getDate();
    var year = dateObj.getFullYear();
    
    // 土日チェック（警告レベル）
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      var dayName = dayOfWeek === 0 ? '日曜日' : '土曜日';
      return {
        isValid: false,
        message: dayName + 'は休日です。有給申請が必要か確認してください。',
        allowOverride: true // 上書き許可（警告レベル）
      };
    }
    
    // 固定祝日チェック
    var nationalHolidays = {
      '1-1': '元日',
      '2-11': '建国記念日',
      '2-23': '天皇誕生日',
      '4-29': '昭和の日',
      '5-3': '憲法記念日',
      '5-4': 'みどりの日',
      '5-5': 'こどもの日',
      '8-11': '山の日',
      '11-3': '文化の日',
      '11-23': '勤労感謝の日',
      '12-29': '年末休暩',
      '12-30': '年末休暩',
      '12-31': '年末休暩'
    };
    
    var dateKey = month + '-' + date;
    if (nationalHolidays[dateKey]) {
      return {
        isValid: false,
        message: '選択された日付は' + nationalHolidays[dateKey] + 'です。祝日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    // 移動祝日の簡易チェック（成人の日、海の日、敬老の日、スポーツの日）
    if (month === 1 && dayOfWeek === 1 && date >= 8 && date <= 14) {
      return {
        isValid: false,
        message: '成人の日です。祝日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    if (month === 7 && dayOfWeek === 1 && date >= 15 && date <= 21) {
      return {
        isValid: false,
        message: '海の日です。祝日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    if (month === 9 && dayOfWeek === 1 && date >= 15 && date <= 21) {
      return {
        isValid: false,
        message: '敬老の日です。祝日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    if (month === 10 && dayOfWeek === 1 && date >= 8 && date <= 14) {
      return {
        isValid: false,
        message: 'スポーツの日です。祝日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    // 会社休日（お盆休み）
    if (month === 8 && date >= 13 && date <= 16) {
      return {
        isValid: false,
        message: 'お盆休み期間中です。会社休日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    // 年末年始休暩
    if ((month === 12 && date >= 29) || (month === 1 && date <= 3)) {
      return {
        isValid: false,
        message: '年末年始休暩中です。会社休日の有給申請は通常不要です。',
        allowOverride: true
      };
    }
    
    return { isValid: true };
    
  } catch (error) {
    console.error('祝日チェックエラー:', error);
    return { isValid: true }; // エラー時はチェックをスキップ
  }
}



/**
 * 特定ユーザーの申請履歴を取得
 * @param {string} userId - 利用者番号
 * @return {Array} 申請履歴の配列
 */
function getUserApplications(userId) {
  console.log('getUserApplications called for userId:', userId);
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('申請');
    
    if (!sheet) {
      console.error('申請シートが見つかりません。シート名を確認してください。');
      // シート一覧を確認
      var sheets = ss.getSheets();
      console.log('利用可能なシート:', sheets.map(function(s) { return s.getName(); }));
      // エラーでも空の配列を返す
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var applications = [];
    
    // 今日の日付を取得（時刻を0:00:00にリセット）
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 該当ユーザーのデータのみ抽出（取消済みを除外、承認済みは含む）
      if (String(row[0]) === String(userId) && String(row[5] || 'Pending') !== 'Canceled') {
        var applyDate = new Date(row[3]); // 申請日
        applyDate.setHours(0, 0, 0, 0);
        
        // 申請日が今日以降かチェック（前日まで取消可能）
        var canCancel = applyDate > today;
        
        applications.push({
          rowNumber: i + 1,  // 行番号（取消時に使用）
          userId: String(row[0]),  // 文字列に変換
          userName: String(row[1] || ''),
          remaining: Number(row[2] || 0),  // 数値に変換、空の場合は0
          applyDate: Utilities.formatDate(applyDate, 'JST', 'yyyy/MM/dd'),
          timestamp: row[4] ? Utilities.formatDate(new Date(row[4]), 'JST', 'yyyy/MM/dd HH:mm:ss') : '',
          status: String(row[5] || 'Pending'),
          canCancel: canCancel,  // 取消可能フラグ
          applyDays: Number(row[7] || 1)  // 申請日数（H列、デフォルト1日）
        });
      }
    }
    
    // 申請日の降順でソート（新しい申請が上）
    applications.sort(function(a, b) {
      return new Date(b.applyDate) - new Date(a.applyDate);
    });
    
    console.log('getUserApplications found:', applications.length, 'records');
    console.log('getUserApplications returning:', applications); // 返却値を確認
    
    return applications;
    
  } catch (error) {
    console.error('getUserApplications error:', error);
    console.error('Error stack:', error.stack);
    return [];
  }
}

/**
 * 有給申請の高度なバリデーション
 * @param {Object} params - 申請パラメーター
 * @return {Object} 検証結果
 */
function validateLeaveApplication(params) {
  try {
    var validationErrors = [];
    var validationWarnings = [];
    
    // 基本パラメーターチェック
    if (!params.userId) {
      validationErrors.push('利用者IDが指定されていません');
    }
    
    if (!params.applyDate) {
      validationErrors.push('申請日が指定されていません');
    }
    
    var applyDays = params.applyDays || 1;
    if (applyDays !== 1 && applyDays !== 0.5) {
      validationErrors.push('申請日数が不正です');
    }
    
    // 日付バリデーション
    if (params.applyDate) {
      var applyDate = new Date(params.applyDate);
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      applyDate.setHours(0, 0, 0, 0);
      
      if (applyDate <= today) {
        validationErrors.push('申請日は明日以降でなければなりません');
      }
      
      // 3ヶ月先までの制限
      var maxDate = new Date();
      maxDate.setMonth(maxDate.getMonth() + 3);
      if (applyDate > maxDate) {
        validationErrors.push('申請日は3ヶ月以内でなければなりません');
      }
      
      // 祝日・休日チェック
      var holidayCheck = validateHolidaysAndWeekends(applyDate);
      if (!holidayCheck.isValid) {
        if (holidayCheck.allowOverride) {
          validationWarnings.push(holidayCheck.message);
        } else {
          validationErrors.push(holidayCheck.message);
        }
      }
    }
    
    return {
      isValid: validationErrors.length === 0,
      errors: validationErrors,
      warnings: validationWarnings,
      hasWarnings: validationWarnings.length > 0
    };
    
  } catch (error) {
    console.error('バリデーションエラー:', error);
    return {
      isValid: false,
      errors: ['バリデーション中にエラーが発生しました'],
      warnings: []
    };
  }
}

// =============================
// URL管理システム
// =============================

/**
 * ランダムなURL キーを生成
 * @return {string} 32文字のランダムキー
 */
function generateUrlKey() {
  var chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var result = '';
  for (var i = 0; i < 32; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * 利用者のURLキーを取得または生成
 * @param {string} userId - 利用者番号
 * @return {string} URLキー
 */
function getUserUrlKey(userId) {
  try {
    var ss = getSpreadsheet();
    var urlSheet = getOrCreateUrlSheet(ss);
    
    var data = urlSheet.getDataRange().getValues();
    
    // 既存のURLキーを検索
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId) && data[i][4] === 'active') {
        // 最終アクセス日時を更新
        urlSheet.getRange(i + 1, 4).setValue(new Date());
        return String(data[i][1]);
      }
    }
    
    // 新しいURLキーを生成
    var newKey = generateUrlKey();
    urlSheet.appendRow([
      userId,
      newKey,
      new Date(), // 作成日
      new Date(), // 最終アクセス
      'active'    // ステータス
    ]);
    
    console.log('新しいURLキーを生成:', userId, newKey);
    return newKey;
    
  } catch (error) {
    console.error('URLキー取得エラー:', error);
    throw new Error('URLキーの取得に失敗しました');
  }
}

/**
 * URLキーから利用者番号を取得
 * @param {string} urlKey - URLキー
 * @return {string|null} 利用者番号
 */
function getUserIdFromUrlKey(urlKey) {
  try {
    if (!urlKey) return null;
    
    var ss = getSpreadsheet();
    var urlSheet = getOrCreateUrlSheet(ss);
    
    var data = urlSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(urlKey) && data[i][4] === 'active') {
        // 最終アクセス日時を更新
        urlSheet.getRange(i + 1, 4).setValue(new Date());
        
        // アクセスログを記録
        logAccess(data[i][0], urlKey);
        
        return String(data[i][0]);
      }
    }
    
    console.log('無効なURLキー:', urlKey);
    return null;
    
  } catch (error) {
    console.error('利用者番号取得エラー:', error);
    return null;
  }
}

/**
 * URL管理シートを取得または作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @return {Sheet} URL管理シート
 */
function getOrCreateUrlSheet(ss) {
  var sheetName = 'URL管理';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log('URL管理シートを新規作成します');
    sheet = ss.insertSheet(sheetName);
    
    // ヘッダー行を追加
    sheet.getRange(1, 1, 1, 5).setValues([
      ['利用者番号', 'URLキー', '作成日', '最終アクセス', 'ステータス']
    ]);
    
    // ヘッダー行の書式設定
    var headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // 列幅調整
    sheet.setColumnWidth(1, 100); // 利用者番号
    sheet.setColumnWidth(2, 300); // URLキー
    sheet.setColumnWidth(3, 150); // 作成日
    sheet.setColumnWidth(4, 150); // 最終アクセス
    sheet.setColumnWidth(5, 100); // ステータス
  }
  
  return sheet;
}

/**
 * アクセスログを記録
 * @param {string} userId - 利用者番号
 * @param {string} urlKey - URLキー
 */
function logAccess(userId, urlKey) {
  try {
    // PropertiesServiceにアクセスログを保存（簡易版）
    var logKey = 'access_log_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
    var existingLog = PropertiesService.getScriptProperties().getProperty(logKey) || '';
    
    var newLogEntry = [
      Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm:ss'),
      userId,
      urlKey.substring(0, 8) + '...', // セキュリティのため一部のみ
      Session.getActiveUser().getEmail() || 'unknown'
    ].join(',');
    
    var updatedLog = existingLog ? existingLog + '\n' + newLogEntry : newLogEntry;
    PropertiesService.getScriptProperties().setProperty(logKey, updatedLog);
    
  } catch (error) {
    console.error('アクセスログ記録エラー:', error);
    // ログ記録の失敗はシステム全体の動作を止めない
  }
}

/**
 * 全利用者のURLキーを一括生成
 * @return {Object} 生成結果
 */
function generateAllUserUrls() {
  try {
    console.log('全利用者のURLキー一括生成を開始します');
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var masterData = masterSheet.getDataRange().getValues();
    var results = [];
    var successCount = 0;
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      var userName = String(masterData[i][1]);
      
      if (userId) {
        try {
          var urlKey = getUserUrlKey(userId);
          results.push({
            userId: userId,
            userName: userName,
            urlKey: urlKey,
            url: getWebAppUrl() + '?key=' + urlKey,
            success: true
          });
          successCount++;
        } catch (error) {
          results.push({
            userId: userId,
            userName: userName,
            error: error.message,
            success: false
          });
        }
      }
    }
    
    console.log('URLキー一括生成完了:', successCount + '/' + (masterData.length - 1));
    
    return {
      success: true,
      results: results,
      total: masterData.length - 1,
      successCount: successCount,
      message: successCount + '/' + (masterData.length - 1) + ' 名のURLキーを生成しました'
    };
    
  } catch (error) {
    console.error('URLキー一括生成エラー:', error);
    return {
      success: false,
      message: 'URLキー生成中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * WebアプリのURLを取得
 * @return {string} WebアプリのURL
 */
function getWebAppUrl() {
  try {
    // ScriptApp.getService().getUrl() で動的に取得を試みる
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf('YOUR_SCRIPT_ID') === -1) {
      return url;
    }
  } catch (e) {
    console.log('ScriptApp.getService().getUrl()でエラー:', e);
  }

  // フォールバック: 最新のデプロイメントID（バージョン13）
  return 'https://script.google.com/macros/s/AKfycbxWVjrZ1NfAiAXv0NnO8atjpHLmgKcFqGoBgn8OnfvfWCs3mmN_lwrb_we0w7A10P1R/exec';
}

/**
 * URLキーを無効化
 * @param {string} userId - 利用者番号
 * @return {boolean} 成功/失敗
 */
function deactivateUserUrl(userId) {
  try {
    var ss = getSpreadsheet();
    var urlSheet = getOrCreateUrlSheet(ss);
    
    var data = urlSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId) && data[i][4] === 'active') {
        urlSheet.getRange(i + 1, 5).setValue('inactive');
        console.log('URLキーを無効化:', userId);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('URLキー無効化エラー:', error);
    return false;
  }
}

/**
 * URL管理のテスト関数
 */
function testUrlManagement() {
  try {
    console.log('=== URL管理テスト開始 ===');
    
    var testUserId = 'R01';
    var results = [];
    
    // 1. URLキー生成テスト
    console.log('1. URLキー生成テスト');
    var urlKey1 = getUserUrlKey(testUserId);
    results.push({
      test: 'URLキー生成',
      result: urlKey1 ? 'SUCCESS' : 'FAILED',
      detail: 'URLキー: ' + urlKey1
    });
    
    // 2. 同じユーザーで再取得（同じキーが返されるか）
    console.log('2. URLキー再取得テスト');
    var urlKey2 = getUserUrlKey(testUserId);
    results.push({
      test: 'URLキー再取得',
      result: urlKey1 === urlKey2 ? 'SUCCESS' : 'FAILED',
      detail: '一致: ' + (urlKey1 === urlKey2)
    });
    
    // 3. 逆引きテスト
    console.log('3. 逆引きテスト');
    var retrievedUserId = getUserIdFromUrlKey(urlKey1);
    results.push({
      test: '逆引き',
      result: retrievedUserId === testUserId ? 'SUCCESS' : 'FAILED',
      detail: '取得した利用者番号: ' + retrievedUserId
    });
    
    // 4. 無効なキーのテスト
    console.log('4. 無効キーテスト');
    var invalidResult = getUserIdFromUrlKey('invalid_key_123');
    results.push({
      test: '無効キー',
      result: invalidResult === null ? 'SUCCESS' : 'FAILED',
      detail: '結果: ' + invalidResult
    });
    
    console.log('=== テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('URL管理テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}
