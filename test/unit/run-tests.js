const assert = require('assert');
const path = require('path');
const fs = require('fs');
const vm = require('vm');

const leaveGrantCore = require(path.join(__dirname, '..', '..', 'src', 'leave-grant-core.js'));
const {
  createMockSpreadsheet,
  getSheetValues
} = require('./helpers/mock-spreadsheet');

const leaveGrantSource = fs.readFileSync(
  path.join(__dirname, '..', '..', 'src', 'leave-grant.js'),
  'utf8'
);
const notificationSource = fs.readFileSync(
  path.join(__dirname, '..', '..', 'src', 'notification.js'),
  'utf8'
);
const statisticsSource = fs.readFileSync(
  path.join(__dirname, '..', '..', 'src', 'statistics-report.js'),
  'utf8'
);

const tests = [];

function test(name, fn) {
  tests.push({ name, fn });
}

function run() {
  let passed = 0;

  tests.forEach(({ name, fn }, index) => {
    try {
      fn();
      passed++;
      console.log(`✅ (${index + 1}) ${name}`);
    } catch (error) {
      console.error(`❌ (${index + 1}) ${name}`);
      console.error(`   ${error.message}`);
      if (error.stack) {
        console.error(error.stack.split('\n').slice(1).join('\n'));
      }
    }
  });

  const failed = tests.length - passed;
  console.log('\n=== 単体テスト結果 ===');
  console.log(`成功: ${passed}/${tests.length}`);

  if (failed > 0) {
    process.exitCode = 1;
  }
}

function createLeaveGrantContext(initialSheets) {
  const mockSpreadsheet = createMockSpreadsheet(initialSheets);
  const context = {
    console: console,
    SpreadsheetApp: {
      openById: () => mockSpreadsheet
    },
    Utilities: {
      formatDate: (date) => {
        const d = new Date(date);
        if (isNaN(d.getTime())) return '';
        return (
          d.getFullYear() +
          '/' +
          String(d.getMonth() + 1).padStart(2, '0') +
          '/' +
          String(d.getDate()).padStart(2, '0')
        );
      }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        setProperty: () => {},
        getProperty: () => null
      })
    },
    Logger: console
  };
  Object.keys(leaveGrantCore).forEach((key) => {
    context[key] = leaveGrantCore[key];
  });
  context.getSpreadsheet = () => mockSpreadsheet;
  context.global = context;
  vm.runInNewContext(leaveGrantSource, context, { filename: 'leave-grant.js' });
  return { context, mockSpreadsheet };
}

function createNotificationContext(initialSheets) {
  const mockSpreadsheet = createMockSpreadsheet(initialSheets);
  const sentEmails = [];
  const scriptProperties = {};

  const context = {
    console: console,
    SPREADSHEET_ID: 'TEST_SPREADSHEET',
    SpreadsheetApp: {
      openById: () => mockSpreadsheet
    },
    GmailApp: {
      sendEmail: function(to, subject, body, options) {
        sentEmails.push({
          to: to,
          subject: subject,
          body: body,
          options: options || {}
        });
      }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        setProperty: (key, value) => { scriptProperties[key] = value; },
        getProperty: (key) => scriptProperties[key] || null
      })
    },
    Utilities: {
      formatDate: function(date, _tz, format) {
        const d = new Date(date);
        if (isNaN(d.getTime())) return '';
        if (format === 'yyyy/MM/dd') {
          return (
            d.getFullYear() +
            '/' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '/' +
            String(d.getDate()).padStart(2, '0')
          );
        }
        return d.toISOString();
      }
    }
  };
  context.global = context;
  vm.runInNewContext(notificationSource, context, { filename: 'notification.js' });
  context.NOTIFICATION_CONFIG.TEST_MODE = false;
  context.NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS = true;
  return { context, sentEmails, mockSpreadsheet };
}

function createStatisticsContext(initialSheets) {
  const mockSpreadsheet = createMockSpreadsheet(initialSheets);
  const context = {
    console: console,
    getSpreadsheet: () => mockSpreadsheet,
    SpreadsheetApp: {
      openById: () => mockSpreadsheet
    },
    Utilities: {
      formatDate: function(date, _tz, format) {
        const d = new Date(date);
        if (isNaN(d.getTime())) return '';
        if (format === 'yyyy/MM/dd HH:mm') {
          return (
            d.getFullYear() +
            '/' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '/' +
            String(d.getDate()).padStart(2, '0') +
            ' ' +
            String(d.getHours()).padStart(2, '0') +
            ':' +
            String(d.getMinutes()).padStart(2, '0')
          );
        }
        if (format === 'yyyy/MM/dd') {
          return (
            d.getFullYear() +
            '/' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '/' +
            String(d.getDate()).padStart(2, '0')
          );
        }
        return d.toISOString();
      }
    }
  };
  context.global = context;
  vm.runInNewContext(statisticsSource, context, { filename: 'statistics-report.js' });
  return { context, mockSpreadsheet };
}

// ---- Test Cases ----

test('calculateLeaveDays: 初回付与テーブル', () => {
  const table = [
    { week: 5, expected: 10 },
    { week: 4, expected: 7 },
    { week: 3, expected: 5 },
    { week: 2, expected: 3 },
    { week: 1, expected: 1 },
    { week: 0, expected: 0 }
  ];

  table.forEach(({ week, expected }) => {
    assert.strictEqual(
      leaveGrantCore.calculateLeaveDays(0.5, week, true),
      expected,
      `週${week}日勤務の初回付与が期待値と一致しません`
    );
  });
});

test('calculateLeaveDays: 週5日勤務の年次増加', () => {
  const table = [
    { years: 0.9, expected: 10 },
    { years: 1.1, expected: 11 },
    { years: 2.1, expected: 12 },
    { years: 3.1, expected: 14 },
    { years: 4.1, expected: 16 },
    { years: 5.1, expected: 18 },
    { years: 6.1, expected: 20 }
  ];

  table.forEach(({ years, expected }) => {
    assert.strictEqual(
      leaveGrantCore.calculateLeaveDays(years, 5, false),
      expected,
      `勤続${years}年の付与日数が期待値と異なります`
    );
  });
});

test('calculateAnnualLeaveDays: 勤務日数別の境界', () => {
  const result4 = leaveGrantCore.calculateAnnualLeaveDays(1.9, 4);
  const result3 = leaveGrantCore.calculateAnnualLeaveDays(3.9, 3);
  const result2 = leaveGrantCore.calculateAnnualLeaveDays(5.9, 2);
  assert.strictEqual(result4, 8, '週4日勤務の境界値が想定外');
  assert.strictEqual(result3, 8, '週3日勤務の境界値が想定外');
  assert.strictEqual(result2, 6, '週2日勤務の境界値が想定外');
});

test('calculateWorkYears: 日付差の計算', () => {
  const hire = new Date('2020-04-01');
  const base = new Date('2023-04-01');
  const years = leaveGrantCore.calculateWorkYears(hire, base);
  assert.ok(Math.abs(years - 3) < 0.01, '勤続年数計算が誤っています');
});

test('calculateWorkYears: 不正入力ハンドリング', () => {
  assert.strictEqual(leaveGrantCore.calculateWorkYears(null, new Date()), 0);
  assert.strictEqual(leaveGrantCore.calculateWorkYears(new Date(), null), 0);
});

test('isFiveDayObligationTarget: 閾値判定', () => {
  assert.strictEqual(leaveGrantCore.isFiveDayObligationTarget(9), false);
  assert.strictEqual(leaveGrantCore.isFiveDayObligationTarget(10), true);
  assert.strictEqual(leaveGrantCore.isFiveDayObligationTarget(11), true);
});

test('formatDate: 正常系と異常系', () => {
  assert.strictEqual(leaveGrantCore.formatDate('2024-01-15'), '2024/01/15');
  assert.strictEqual(leaveGrantCore.formatDate('invalid'), '-');
  assert.strictEqual(leaveGrantCore.formatDate(null), '-');
});

test('grantLeave: 履歴追加と残数加算', () => {
  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['U001', '山田太郎', 5, '', '2023/04/01', 5, '', '']
  ];
  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時']
  ];

  const { context, mockSpreadsheet } = createLeaveGrantContext({
    'マスター': masterSheet,
    '付与履歴': grantHistory
  });

  const grantDate = new Date('2024-04-01');
  const result = context.grantLeave('U001', grantDate, 10, '年次', 1.5);
  assert.ok(result.success, '付与処理が成功しませんでした');

  const updatedMaster = getSheetValues(mockSpreadsheet, 'マスター');
  assert.strictEqual(updatedMaster[1][2], 15, '残日数が正しく更新されていません');

  const history = getSheetValues(mockSpreadsheet, '付与履歴');
  assert.strictEqual(history.length, 2, '履歴行が追加されていません');
  assert.strictEqual(history[1][0], 'U001', '付与履歴の利用者番号が不正です');
  assert.strictEqual(history[1][2], 10, '付与日数が記録されていません');
});

test('consumeLeave: FIFOで複数付与を消費', () => {
  const today = new Date();
  const future = new Date(today);
  future.setFullYear(future.getFullYear() + 1);

  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['U900', new Date('2022-04-01'), 10, future, 5, '初回', 0.5, new Date('2022-04-01')],
    ['U900', new Date('2023-04-01'), 11, future, 7, '年次', 1.5, new Date('2023-04-01')]
  ];

  const { context, mockSpreadsheet } = createLeaveGrantContext({
    '付与履歴': grantHistory
  });

  const result = context.consumeLeave('U900', 6);
  assert.ok(result.success, '有給消費が失敗しました');
  assert.strictEqual(result.consumedFromGrants.length, 2, '複数付与からの消費が記録されていません');
  assert.strictEqual(result.consumedFromGrants[0].consumed, 5, '最古の付与から消費されていません');
  assert.strictEqual(result.consumedFromGrants[1].consumed, 1, '残りが新しい付与から消費されていません');

  const history = getSheetValues(mockSpreadsheet, '付与履歴');
  assert.strictEqual(history[1][4], 0, '古い付与の残日数がゼロになっていません');
  assert.strictEqual(history[2][4], 6, '新しい付与の日数が減算されていません');
});

test('calculateEffectiveRemainingDays: 失効済み分を除外', () => {
  const today = new Date();
  const past = new Date(today);
  past.setFullYear(past.getFullYear() - 1);
  const future = new Date(today);
  future.setFullYear(future.getFullYear() + 1);

  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['U777', new Date('2022-04-01'), 10, past, 4, '初回', 0.5, new Date('2022-04-01')],
    ['U777', new Date('2023-04-01'), 11, future, 6, '年次', 1.5, new Date('2023-04-01')]
  ];

  const { context } = createLeaveGrantContext({
    '付与履歴': grantHistory
  });

  const remaining = context.calculateEffectiveRemainingDays('U777');
  assert.strictEqual(remaining, 6, '失効済み日数が除外されていません');
});

test('processExpiredLeaves: 失効分をクリアしマスターを更新', () => {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const future = new Date(today);
  future.setFullYear(future.getFullYear() + 1);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['U500', '斎藤花子', 8, '', '2022/04/01', 5, '', '']
  ];
  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['U500', new Date('2022-04-01'), 10, yesterday, 3, '初回', 0.5, new Date('2022-04-01')],
    ['U500', new Date('2023-04-01'), 11, future, 5, '年次', 1.5, new Date('2023-04-01')]
  ];

  const { context, mockSpreadsheet } = createLeaveGrantContext({
    'マスター': masterSheet,
    '付与履歴': grantHistory
  });

  const result = context.processExpiredLeaves();
  assert.ok(result.success, '失効処理が失敗しました');
  assert.strictEqual(result.expiredCount, 3, '失効した日数が合致しません');

  const history = getSheetValues(mockSpreadsheet, '付与履歴');
  assert.strictEqual(history[1][4], 0, '失効対象の残日数が0になっていません');

  const updatedMaster = getSheetValues(mockSpreadsheet, 'マスター');
  assert.strictEqual(updatedMaster[1][2], 5, 'マスター残日数が再計算されていません');
});

test('getExpiringLeaves: 指定期間内のみ取得', () => {
  const today = new Date();
  const within10 = new Date(today);
  within10.setDate(today.getDate() + 10);
  const beyond30 = new Date(today);
  beyond30.setDate(today.getDate() + 40);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['U600', '田中進', 6, '', '2022/04/01', 5, '', '']
  ];
  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['U600', new Date('2022-04-01'), 10, within10, 3, '初回', 0.5, new Date('2022-04-01')],
    ['U600', new Date('2023-04-01'), 11, beyond30, 5, '年次', 1.5, new Date('2023-04-01')]
  ];

  const { context } = createLeaveGrantContext({
    'マスター': masterSheet,
    '付与履歴': grantHistory
  });

  const list = context.getExpiringLeaves(30);
  assert.strictEqual(list.length, 1, '期間外の失効データが混入しています');
  assert.strictEqual(list[0].remainingDays, 3, '取得した残日数が一致しません');
});

test('getSixMonthGrantTargets: 初回未付与者のみ抽出', () => {
  const today = new Date();
  const sixMonthsAgo = new Date(today);
  sixMonthsAgo.setMonth(today.getMonth() - 6);
  const fiveMonthsAgo = new Date(today);
  fiveMonthsAgo.setMonth(today.getMonth() - 5);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['M001', '入社6ヶ月経過', 0, '', sixMonthsAgo, 5, '', ''],
    ['M002', '入社5ヶ月', 0, '', fiveMonthsAgo, 5, '', ''],
    ['M003', '既に初回付与済み', 0, '', sixMonthsAgo, 5, today, '']
  ];

  const { context } = createLeaveGrantContext({
    'マスター': masterSheet
  });

  const targets = context.getSixMonthGrantTargets();
  assert.strictEqual(targets.length, 1, '初回未付与者だけが抽出されていません');
  assert.strictEqual(targets[0].userId, 'M001', '対象利用者が誤っています');
});

test('processSixMonthGrants: 出勤率チェックと初回付与', () => {
  const today = new Date();
  const sixMonthsAgo = new Date(today);
  sixMonthsAgo.setMonth(today.getMonth() - 6);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['S001', '高出勤率', 0, '', sixMonthsAgo, 5, '', ''],
    ['S002', '低出勤率', 0, '', sixMonthsAgo, 5, '', '']
  ];

  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時']
  ];

  const applySheet = [
    ['利用者番号', '氏名', '残', '申請日', '申請日時', '状態', '備考', '日数'],
    ['S002', '低出勤率', 0, sixMonthsAgo, sixMonthsAgo, 'Approved', '', 40]
  ];

  const { context, mockSpreadsheet } = createLeaveGrantContext({
    'マスター': masterSheet,
    '付与履歴': grantHistory,
    '申請': applySheet
  });

  const origCheckAttendanceRate = context.checkAttendanceRate;
  context.checkAttendanceRate = (userId) => {
    if (userId === 'S001') return 0.9;
    if (userId === 'S002') return 0.5;
    return origCheckAttendanceRate(userId);
  };

  const result = context.processSixMonthGrants();
  assert.ok(result.success, '6ヶ月付与処理が失敗しました');
  assert.strictEqual(result.processedCount, 1, '成功件数が一致しません');

  const masterValues = getSheetValues(mockSpreadsheet, 'マスター');
  assert.strictEqual(masterValues[1][2], 10, '付与後の残日数が反映されていません');
  assert.ok(masterValues[1][6], '初回付与日が記録されていません');
  assert.strictEqual(masterValues[2][2], 0, '出勤率不足の利用者に付与されています');
});

test('getAnnualGrantTargets: 次回付与日が到来した利用者のみ抽出', () => {
  const today = new Date();
  const twoYearsAgo = new Date(today);
  twoYearsAgo.setFullYear(today.getFullYear() - 2);
  const oneYearAgo = new Date(today);
  oneYearAgo.setFullYear(today.getFullYear() - 1);
  const futureDate = new Date(today);
  futureDate.setMonth(today.getMonth() + 6);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['A001', '年次対象', 10, '', twoYearsAgo, 5, twoYearsAgo, null],
    ['A002', '最新記録あり', 8, '', twoYearsAgo, 4, twoYearsAgo, oneYearAgo],
    ['A003', '初回未付与', 5, '', twoYearsAgo, 5, null, null],
    ['A004', '次回未来', 5, '', twoYearsAgo, 5, futureDate, null]
  ];

  const { context } = createLeaveGrantContext({
    'マスター': masterSheet
  });

  const targets = context.getAnnualGrantTargets();
  const userIds = Array.from(targets, (t) => t.userId).sort();
  assert.strictEqual(userIds.length, 2, '年次付与対象者数が想定外です: ' + JSON.stringify(userIds));
  assert.deepStrictEqual(userIds, ['A001', 'A002'], '年次付与対象が期待通りではありません');
});

test('processAnnualGrants: 出勤率に応じて付与可否を切り替え', () => {
  const today = new Date();
  const twoYearsAgo = new Date(today);
  twoYearsAgo.setFullYear(today.getFullYear() - 2);

  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['AN01', '高勤続', 12, '', twoYearsAgo, 5, twoYearsAgo, null],
    ['AN02', '低勤続', 8, '', twoYearsAgo, 3, twoYearsAgo, null]
  ];

  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時']
  ];

  const applySheet = [
    ['利用者番号', '氏名', '残', '申請日', '申請日時', '状態', '備考', '日数'],
    ['AN02', '低勤続', 0, twoYearsAgo, twoYearsAgo, 'Approved', '', 60]
  ];

  const { context, mockSpreadsheet } = createLeaveGrantContext({
    'マスター': masterSheet,
    '付与履歴': grantHistory,
    '申請': applySheet
  });

  const origCheck = context.checkAnnualAttendanceRate;
  context.checkAnnualAttendanceRate = (userId) => {
    if (userId === 'AN01') return 0.95;
    if (userId === 'AN02') return 0.5;
    return origCheck(userId);
  };

  const result = context.processAnnualGrants();
  assert.ok(result.success, '年次付与処理が失敗しました');
  assert.strictEqual(result.processedCount, 1, '年次付与成功件数が一致しません');

  const masterValues = getSheetValues(mockSpreadsheet, 'マスター');
  assert.strictEqual(masterValues[1][2], 24, '年次付与後の残日数が誤っています');
  assert.ok(masterValues[1][7], '年次付与日が記録されていません');
  assert.strictEqual(masterValues[2][2], 8, '出勤率不足の利用者に付与されています');
});

test('sendLowRemainingDaysAlert: 通常閾値では本人のみ通知', () => {
  const { context, sentEmails } = createNotificationContext({});
  const result = context.sendLowRemainingDaysAlert({
    userId: 'R10',
    userName: '通常閾値',
    remaining: 4
  });
  assert.ok(result.success, '通常アラート送信に失敗しました');
  assert.strictEqual(sentEmails.length, 1, '上長への通知が不要に送信されています');
  assert.strictEqual(sentEmails[0].to, 'r10@company.com', '宛先が本人ではありません');
});

test('sendLowRemainingDaysAlert: 臨界閾値で上長にも通知', () => {
  const { context, sentEmails } = createNotificationContext({});
  const result = context.sendLowRemainingDaysAlert({
    userId: 'R20',
    userName: '臨界閾値',
    remaining: 2
  });
  assert.ok(result.success, '臨界アラート送信に失敗しました');
  assert.strictEqual(sentEmails.length, 2, '上長への通知が送信されていません');
  assert.strictEqual(sentEmails[0].to, 'r20@company.com', '1通目が本人宛ではありません');
  assert.strictEqual(sentEmails[1].to, 'rise-manager@company.com', '上長宛先が想定と異なります');
});

test('sendApplicationNotification: 承認リンクと人事CC', () => {
  const { context, sentEmails } = createNotificationContext({});
  context.NOTIFICATION_CONFIG.APPROVAL_BASE_URL = 'https://example.com/approve';
  context.NOTIFICATION_CONFIG.HR_EMAIL = 'hr@example.com';
  context.NOTIFICATION_CONFIG.CC_HR_ON_APPLICATION = true;
  const result = context.sendApplicationNotification({
    userId: 'R30',
    userName: 'リンク太郎',
    applyDate: '2025/04/10',
    applyDays: 1,
    timestamp: '2025/04/01 09:00:00'
  });
  assert.ok(result.success, '申請通知送信に失敗しました');
  assert.strictEqual(sentEmails[0].to, 'rise-manager@company.com', '承認者宛先が不正です');
  assert.strictEqual(sentEmails[0].options.cc, 'hr@example.com', '人事CCが付与されていません');
  assert.ok(sentEmails[0].options.htmlBody.indexOf('https://example.com/approve') !== -1, '承認リンクが本文に含まれていません');
});

test('sendApprovalResultNotification: 却下理由と上長CC', () => {
  const { context, sentEmails } = createNotificationContext({});
  context.NOTIFICATION_CONFIG.CC_MANAGER_ON_APPROVAL = true;
  const result = context.sendApprovalResultNotification({
    userId: 'R40',
    userName: '却下花子',
    applyDate: '2025/04/12',
    applyDays: 0.5
  }, 'Rejected', '業務都合のため');
  assert.ok(result.success, '承認結果通知送信に失敗しました');
  assert.strictEqual(sentEmails[0].to, 'r40@company.com', '申請者宛先が不正です');
  assert.strictEqual(sentEmails[0].options.cc, 'rise-manager@company.com', '上長CCが付与されていません');
  assert.ok(sentEmails[0].options.htmlBody.indexOf('業務都合のため') !== -1, '理由が本文に含まれていません');
});

test('sendGrantNotification: 付与通知で上長にも共有', () => {
  const { context, sentEmails } = createNotificationContext({});
  context.NOTIFICATION_CONFIG.CC_MANAGER_ON_GRANT = true;
  const result = context.sendGrantNotification({
    userId: 'R50',
    userName: '付与太郎',
    grantDays: 11,
    grantDate: '2025/04/01',
    grantType: '年次付与'
  });
  assert.ok(result.success, '付与通知送信に失敗しました');
  assert.strictEqual(sentEmails[0].to, 'r50@company.com', '付与通知宛先が不正です');
  assert.strictEqual(sentEmails[0].options.cc, 'rise-manager@company.com', '上長共有が付与されていません');
});

test('calculateBasicStatistics: 従業員と残日数を集計', () => {
  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['R01', 'ライズ太郎', 5, '', new Date('2023-04-01'), 5, new Date('2023-10-01'), new Date('2024-04-01')],
    ['P01', 'パロン花子', 0, '', new Date('2022-04-01'), 4, new Date('2022-10-01'), new Date('2023-04-01')],
    ['S01', 'シエル次郎', 8, '', new Date('2023-07-01'), 3, new Date('2024-01-01'), null]
  ];

  const { context } = createStatisticsContext({
    'マスター': masterSheet
  });

  const stats = context.calculateBasicStatistics(new Date('2025-04-01'));
  assert.strictEqual(stats.totalEmployees, 3, '従業員数の集計が誤っています');
  assert.strictEqual(stats.totalRemainingDays, 13, '残日数合計が誤っています');
  assert.strictEqual(stats.lowRemainingCount, 1, '残5日以下の人数が誤っています');
  assert.strictEqual(stats.zeroRemainingCount, 1, '残0日人数が誤っています');
  assert.strictEqual(stats.divisionCounts.R, 1, '事業所別従業員数が誤っています');
});

test('generateMonthlyReport: 各統計を統合して出力', () => {
  const masterSheet = [
    ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日'],
    ['R01', 'ライズ太郎', 5, '', new Date('2023-04-01'), 5, new Date('2023-10-01'), new Date('2024-04-01')],
    ['P01', 'パロン花子', 0, '', new Date('2022-04-01'), 4, new Date('2022-10-01'), new Date('2023-04-01')]
  ];
  const applySheet = [
    ['利用者番号', '氏名', '残', '申請日', '申請日時', '状態', '備考', '日数'],
    ['R01', 'ライズ太郎', 5, new Date('2025-04-10'), new Date('2025-04-01'), 'Approved', '', 1],
    ['P01', 'パロン花子', 0, new Date('2025-03-20'), new Date('2025-03-10'), 'Rejected', '', 2]
  ];
  const grantHistory = [
    ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時'],
    ['R01', new Date('2025-04-01'), 11, new Date('2027-04-01'), 11, '年次', 2, new Date('2025-04-01')],
    ['P01', new Date('2025-04-05'), 10, new Date('2027-04-05'), 10, '初回', 0.5, new Date('2025-04-05')]
  ];

  const { context } = createStatisticsContext({
    'マスター': masterSheet,
    '申請': applySheet,
    '付与履歴': grantHistory
  });

  const report = context.generateMonthlyReport(2025, 4);
  assert.ok(report.success, '月次レポート生成に失敗しました');
  assert.strictEqual(report.reportData.applicationStatistics.totalApplications, 1, '期間内申請数が誤っています');
  assert.strictEqual(report.reportData.grantStatistics.totalGrants, 2, '期間内付与数が誤っています');
  assert.strictEqual(report.reportData.divisionStatistics.R.employees, 1, '事業所別集計が誤っています');
});

run();
