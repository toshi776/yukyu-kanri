const fs = require('fs');
const path = require('path');
const vm = require('vm');

const leaveGrantCore = require('../../src/leave-grant-core.js');
const { createMockSpreadsheet } = require('../unit/helpers/mock-spreadsheet');

const projectRoot = path.join(__dirname, '..', '..');
const dayMs = 24 * 60 * 60 * 1000;

function loadSource(relativePath) {
  return fs.readFileSync(path.join(projectRoot, relativePath), 'utf8');
}

const sources = {
  leaveGrant: loadSource('src/leave-grant.js'),
  notification: loadSource('src/notification.js'),
  statistics: loadSource('src/statistics-report.js'),
  triggerManager: loadSource('src/trigger-manager.js'),
  grantExpirySuite: loadSource('test/test-grant-expiry-management.js'),
  notificationSuite: loadSource('test/test-notification-production.js')
};

function formatDateValue(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function formatJapaneseDate(date) {
  const d = new Date(date);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

function createUtilities() {
  return {
    formatDate(date, _tz, format) {
      if (!date) return '';
      const d = new Date(date);
      if (isNaN(d.getTime())) return '';
      if (format === 'yyyy/MM/dd') {
        return formatDateValue(d);
      }
      if (format === 'yyyy/MM/dd HH:mm') {
        const time = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
        return `${formatDateValue(d)} ${time}`;
      }
      if (format === 'yyyy年M月d日') {
        return formatJapaneseDate(d);
      }
      return d.toISOString();
    }
  };
}

function createIntegrationContext(initialSheets) {
  const mockSpreadsheet = createMockSpreadsheet(initialSheets);
  const sentEmails = [];
  const scriptProperties = {};

  const context = {
    console,
    Logger: console,
    SPREADSHEET_ID: 'TEST_SPREADSHEET',
    SpreadsheetApp: {
      openById: () => mockSpreadsheet
    },
    GmailApp: {
      sendEmail: function(to, subject, body, options) {
        sentEmails.push({ to, subject, body, options: options || {} });
      }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        setProperty: (key, value) => { scriptProperties[key] = value; },
        getProperty: (key) => scriptProperties[key] || null
      })
    },
    Session: {
      getScriptTimeZone: () => 'Asia/Tokyo'
    },
    Utilities: createUtilities(),
    UrlFetchApp: {
      fetch: () => ({ getContentText: () => '' })
    },
    ScriptApp: {
      getProjectTriggers: () => []
    }
  };

  context.getSpreadsheet = () => mockSpreadsheet;
  context.global = context;

  Object.keys(leaveGrantCore).forEach((key) => {
    context[key] = leaveGrantCore[key];
  });

  vm.runInNewContext(sources.leaveGrant, context, { filename: 'leave-grant.js' });
  vm.runInNewContext(sources.notification, context, { filename: 'notification.js' });
  vm.runInNewContext(sources.statistics, context, { filename: 'statistics-report.js' });
  vm.runInNewContext(sources.triggerManager, context, { filename: 'trigger-manager.js' });
  vm.runInNewContext(sources.grantExpirySuite, context, { filename: 'test-grant-expiry-management.js' });
  vm.runInNewContext(sources.notificationSuite, context, { filename: 'test-notification-production.js' });

  return { context, mockSpreadsheet, sentEmails };
}

function buildBaseSheets() {
  const now = new Date();
  const thirtyDaysLater = new Date(now.getTime() + 30 * dayMs);
  const ninetyDaysLater = new Date(now.getTime() + 90 * dayMs);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['R01', 'ライズ太郎', 4, '', new Date(now.getFullYear() - 1, now.getMonth() - 7, 1), 5, '', ''],
    ['P01', 'パロン花子', 12, '', new Date(now.getFullYear() - 3, 3, 1), 5, new Date(now.getFullYear() - 3, 9, 1), new Date(now.getFullYear() - 1, 3, 1)],
    ['S01', 'シエル次郎', 2, '', new Date(now.getFullYear() - 2, 5, 15), 4, '', '']
  ];

  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['R01', new Date(now.getFullYear() - 1, now.getMonth(), 1), 10, ninetyDaysLater, 4, '初回', 0.5, new Date(now.getFullYear() - 1, now.getMonth(), 1)],
    ['P01', new Date(now.getFullYear(), now.getMonth() - 6, 1), 12, new Date(now.getFullYear() + 1, now.getMonth() - 6, 1), 12, '年次', 2, new Date(now.getFullYear(), now.getMonth() - 6, 1)],
    ['S01', new Date(now.getFullYear() - 1, now.getMonth() - 3, 10), 8, thirtyDaysLater, 2, '年次', 1.5, new Date(now.getFullYear() - 1, now.getMonth() - 3, 10)]
  ];

  const applications = [
    ['利用者番号', '氏名', '残日数', '申請日', '申請日時', '状態', '備考', '日数'],
    ['R01', 'ライズ太郎', 4, new Date(now.getFullYear(), now.getMonth() - 1, 15), new Date(now.getFullYear(), now.getMonth() - 1, 10), 'Approved', '', 1],
    ['P01', 'パロン花子', 12, new Date(now.getFullYear(), now.getMonth() - 2, 5), new Date(now.getFullYear(), now.getMonth() - 2, 1), 'Approved', '', 2]
  ];

  return {
    'マスター': masterSheet,
    '付与履歴': grantHistory,
    '申請': applications
  };
}

function runSuite(name, runner) {
  console.log(`\n=== ${name} ===`);
  const initialSheets = buildBaseSheets();
  const { context, sentEmails } = createIntegrationContext(initialSheets);
  const result = runner(context, sentEmails);
  if (!result || typeof result.success === 'undefined') {
    throw new Error(`${name} が結果を返しませんでした`);
  }
  console.log(`結果: ${result.success ? 'PASS' : 'FAIL'} (${result.summary || ''})`);
  if (sentEmails.length > 0) {
    console.log('送信メール件数:', sentEmails.length);
  }
  return result.success;
}

function main() {
  console.log('=== 結合テスト（ローカルモック環境）開始 ===');

  const suites = [
    {
      name: '付与・失効管理結合テスト',
      run: (context) => context.runGrantExpiryManagementTests()
    },
    {
      name: '通知機能結合テスト',
      run: (context) => context.runNotificationProductionTests()
    }
  ];

  let allPassed = true;
  suites.forEach((suite) => {
    const passed = runSuite(suite.name, suite.run);
    if (!passed) {
      allPassed = false;
    }
  });

  console.log('\n=== 結合テスト完了 ===');
  if (!allPassed) {
    console.error('結合テストで失敗が発生しました');
    process.exitCode = 1;
  } else {
    console.log('すべての結合テストが成功しました');
  }
}

main();
