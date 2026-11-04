const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const backendSource = fs.readFileSync(path.join(projectRoot, 'code.gs'), 'utf8');
const frontendSource = fs.readFileSync(path.join(projectRoot, 'index.html'), 'utf8');
const TODAY = '2026-10-30';

function makeSheet(rows, sheetName, rangeSpecs) {
  return {
    getDataRange() {
      return { getValues: () => rows };
    },
    getRange(row, column, rowCount, columnCount) {
      rangeSpecs.push({ sheetName, row, column, rowCount, columnCount });
      return {
        getValues: () => rows.slice(row - 1, row - 1 + rowCount)
          .map(values => values.slice(column - 1, column - 1 + columnCount))
      };
    },
    getLastRow() {
      return rows.length;
    }
  };
}

function createHarness(liveRows, historyRows) {
  const properties = new Map();
  const rangeSpecs = [];
  let archiveCallCount = 0;
  const context = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: () => null })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => properties.get(key) || null,
        setProperty: (key, value) => properties.set(key, value)
      })
    },
    LockService: {
      getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} })
    },
    Utilities: {
      formatDate: value => value instanceof Date ? value.toISOString() : String(value)
    }
  };

  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });

  const liveSheet = makeSheet(liveRows, 'User_Mapping', rangeSpecs);
  const historySheet = makeSheet(historyRows, 'User_Mapping_History', rangeSpecs);
  context.requireSession_ = token => ({
    success: true,
    session: token === 'admin-token'
      ? { userName: 'Admin', userType: 'ADMIN' }
      : { userName: 'Kim', userType: 'VJ' }
  });
  const database = { getSheetByName: name => name === 'User_Mapping' ? liveSheet : (name === 'User_Mapping_History' ? historySheet : null) };
  context.getUserMappingDatabase_ = () => database;
  context.getOrCreateUserMappingSheet_ = () => liveSheet;
  context.getOrCreateUserMappingHistorySheet_ = () => historySheet;
  context.userMappingSnapshotDate_ = () => TODAY;
  context.archiveUserMappingDailySnapshot_ = (_database, _rows, existingHistory) => {
    archiveCallCount++;
    return existingHistory;
  };

  return {
    context,
    rangeSpecs,
    properties,
    get archiveCallCount() { return archiveCallCount; }
  };
}

const liveHeader = ['ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'Coins Today', 'Coins Month', 'Last Updated', 'Uploaded By', 'Total Days'];
const historyHeader = ['Snapshot Date', 'ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'Coins Today', 'Coins Month', 'Last Updated', 'Uploaded By', 'Total Days'];

test('VJ response includes only that VJ records in live and history data', () => {
  const live = [
    liveHeader,
    ['@kim-user', 'Kim User', '', 'Kim', 'Team A', 10, 10, TODAY + ' 09:00:00', 'Admin', 1],
    ['@lee-user', 'Lee User', '', 'Lee', 'Team B', 20, 20, TODAY + ' 09:00:00', 'Admin', 1]
  ];
  const history = [
    historyHeader,
    [TODAY, '@kim-user', 'Kim User', '', 'Kim', 'Team A', 10, 10, TODAY + ' 09:00:00', 'Admin', 1],
    [TODAY, '@lee-user', 'Lee User', '', 'Lee', 'Team B', 20, 20, TODAY + ' 09:00:00', 'Admin', 1]
  ];
  const harness = createHarness(live, history);
  const result = harness.context.getUserMappingData('ignored', 'ADMIN', 'vj-token', '127.0.0.1');

  assert.equal(result.success, true);
  assert.deepEqual(Array.from(result.list, row => row.assignedVJ), ['Kim']);
  assert.deepEqual(Array.from(result.historyList, row => row.assignedVJ), ['Kim']);
  assert.equal('userAvatar' in result.list[0], false);
  assert.equal('vjAvatar' in result.list[0], false);
  assert.deepEqual(harness.rangeSpecs, [
    { sheetName: 'User_Mapping', row: 1, column: 1, rowCount: live.length, columnCount: 10 },
    { sheetName: 'User_Mapping_History', row: 1, column: 1, rowCount: history.length, columnCount: 11 }
  ]);
});

test('ADMIN response includes all live records and first history page', () => {
  const live = [liveHeader];
  const history = [historyHeader];
  for (let index = 1; index <= 3; index++) {
    live.push(['@user' + index, 'User ' + index, '', 'VJ ' + index, 'Team ' + index, index, index, TODAY, 'Admin', 1]);
    history.push([TODAY, '@user' + index, 'User ' + index, '', 'VJ ' + index, 'Team ' + index, index, index, TODAY, 'Admin', 1]);
  }
  const harness = createHarness(live, history);
  const result = harness.context.getUserMappingData('', 'VJ', 'admin-token', '127.0.0.1');

  assert.equal(result.success, true);
  assert.equal(result.userType, 'ADMIN');
  assert.equal(result.list.length, 3);
  assert.equal(result.historyList.length, 3);
  assert.equal(result.historyHasMore, false);
});

test('history is returned newest-first in pages of 100 and exposes older rows', () => {
  const history = [historyHeader];
  for (let index = 1; index <= 235; index++) {
    history.push([TODAY, '@user' + index, 'User ' + index, '', 'Kim', 'Team A', index, index, TODAY, 'Admin', index]);
  }
  const harness = createHarness([liveHeader], history);
  const first = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');
  const second = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1', first.historyNextOffset);
  const third = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1', second.historyNextOffset);

  assert.equal(first.historyList.length, 100);
  assert.equal(first.historyHasMore, true);
  assert.equal(second.historyList.length, 100);
  assert.equal(second.historyHasMore, true);
  assert.equal(third.historyList.length, 35);
  assert.equal(third.historyHasMore, false);
  const allIds = [...first.historyList, ...second.historyList, ...third.historyList].map(row => row.tiktokId);
  assert.equal(new Set(allIds).size, 235);
  assert.equal(first.historyList[0].tiktokId, '@user235');
});

test('daily snapshot fallback is attempted at most once for the same day', () => {
  const harness = createHarness([liveHeader, ['@kim-user', 'Kim User', '', 'Kim', 'Team A', 1, 1, TODAY, 'Admin', 1]], [historyHeader]);
  harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');
  harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');

  assert.equal(harness.archiveCallCount, 1);
});

test('daily snapshot lookup reads only the current-date tail of the history sheet', () => {
  const history = [historyHeader];
  for (let index = 1; index <= 1400; index++) {
    history.push(['2026-10-29', '@old' + index, 'Old User ' + index, '', 'Kim', 'Team A', 0, 0, '2026-10-29', 'Admin', 1]);
  }
  for (let index = 1; index <= 10; index++) {
    history.push([TODAY, '@today' + index, 'Today User ' + index, '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 1]);
  }
  const harness = createHarness([liveHeader], history);
  const result = harness.context.getUserMappingTodayHistoryRows_(makeSheet(history, 'User_Mapping_History', harness.rangeSpecs), TODAY);

  assert.equal(result.length, 11);
  assert.equal(result[1][1], '@today10');
  assert.ok(harness.rangeSpecs[0].rowCount < history.length);
});

test('User Mapping report returns only the requested 50-row page and computes whole-filter summaries', () => {
  const live = [liveHeader];
  for (let index = 1; index <= 125; index++) {
    live.push(['@user' + index, 'User ' + index, '', index % 2 ? 'Kim' : 'Lee', index % 2 ? 'Team A' : 'Team B', index, index * 2, TODAY, 'Admin', index]);
  }
  const harness = createHarness(live, [historyHeader]);
  harness.properties.set('ACE_USER_MAPPING_ARCHIVED_DATE', TODAY);
  const first = harness.context.getUserMappingData('', 'ADMIN', 'admin-token', '127.0.0.1', 0, { reportMode: true, page: 1, pageSize: 50 });
  const second = harness.context.getUserMappingData('', 'ADMIN', 'admin-token', '127.0.0.1', 0, { reportMode: true, page: 2, pageSize: 50 });

  assert.equal(first.success, true);
  assert.equal(first.list.length, 50);
  assert.equal(first.totalCount, 125);
  assert.equal(first.totalPages, 3);
  assert.equal(first.report.uniqueUserCount, 125);
  assert.equal(first.report.todayCoins, 7875);
  assert.equal(second.list.length, 50);
  assert.equal(second.list[0].tiktokId, '@user51');
  assert.equal(first.historyList.length, 0);
  assert.equal(first.report.availableTeams.length, 2);
  assert.ok(harness.rangeSpecs.every(range => range.sheetName === 'User_Mapping' && range.columnCount === 10));
});

test('historyOnly returns one newest-first 100-row page without scanning the full archive', () => {
  const history = [historyHeader];
  for (let index = 1; index <= 1400; index++) {
    history.push([TODAY, '@user' + index, 'User ' + index, '', 'Kim', 'Team A', index, index, TODAY, 'Admin', index]);
  }
  const harness = createHarness([liveHeader], history);
  harness.properties.set('ACE_USER_MAPPING_ARCHIVED_DATE', TODAY);
  const result = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1', 0, { historyOnly: true });

  assert.equal(result.success, true);
  assert.equal(result.list.length, 0);
  assert.equal(result.historyList.length, 100);
  assert.equal(result.historyList[0].tiktokId, '@user1400');
  assert.equal(result.historyHasMore, true);
  assert.equal(result.historyTotalCount, null);
  assert.ok(harness.rangeSpecs.some(range => range.sheetName === 'User_Mapping_History' && range.rowCount < history.length));
});

test('Super Admin password lookup is role-gated, disambiguates accounts, audits safely, and never returns hashes', () => {
  const auditEntries = [];
  const adminRows = [
    ['Name', 'PIN', 'Role'],
    ['Admin', '001234', 'SUPER_ADMIN'],
    ['Staff', '009999', 'ADMIN']
  ];
  const hashValue = '$2b$12$fake-hash-value-that-must-never-be-returned';
  const vjRows = [
    ['Full Name', 'Stage Name', 'Team', 'Role', 'x', 'Shift', 'x', 'TikTok ID', 'Password'],
    ['Kim One', 'Kim', 'Team A', 'VJ', '', '', '', '@kim.one', '9876'],
    ['Kim Two', 'Kim', 'Team B', 'MC', '', '', '', '@kim.two', '2222'],
    ['Hash User', 'Hashed', 'Team A', 'HR', '', '', '', '@hashed', hashValue]
  ];
  const logRows = [['Timestamp', 'User Type', 'User Name', 'Action', 'Details', 'IP Address', 'Device Info', 'Month_Year']];
  function sheet(rows) {
    return {
      getDataRange: () => ({ getDisplayValues: () => rows }),
      getLastRow: () => rows.length,
      getLastColumn: () => rows[0].length,
      getRange: (row, column, rowCount, columnCount) => ({
        getDisplayValues: () => rows.slice(row - 1, row - 1 + rowCount).map(values => values.slice(column - 1, column - 1 + columnCount))
      })
    };
  }
  const sheets = { Admins: sheet(adminRows), 'vj mc all': sheet(vjRows), Logs: sheet(logRows) };
  const context = {
    console,
    Logger: { log: value => auditEntries.push(String(value)) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] || null }) }
  };
  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });
  context.requireSession_ = token => ({ success: true, session: { userName: token === 'staff-token' ? 'Staff' : 'Admin', userType: 'ADMIN' } });
  context.isLoginRateLimited_ = () => false;
  context.recordLoginFailure_ = () => {};
  context.clearLoginFailures_ = () => {};
  context.writeLog_ = (...args) => auditEntries.push(args);

  const accountPage = context.getPasswordAccountPage('', 'ALL', 1, 50, 'admin-token', '127.0.0.1');
  assert.equal(accountPage.success, true);
  assert.equal(accountPage.total, 3);
  assert.equal(accountPage.accounts[2].credentialStatus, 'HASHED');
  assert.equal(JSON.stringify(accountPage).includes('9876'), false);
  assert.equal(JSON.stringify(accountPage).includes(hashValue), false);

  const duplicateKim = accountPage.accounts.find(account => account.username === '@kim.two');
  const badPin = context.getVjPasswordForAdmin(duplicateKim, 'bad-pin', 'admin-token', '127.0.0.1', 'test');
  assert.equal(badPin.success, false);
  assert.equal(Object.hasOwn(badPin, 'password'), false);
  const wrongIdentity = context.getVjPasswordForAdmin({ ...duplicateKim, team: 'Team A' }, '001234', 'admin-token', '127.0.0.1', 'test');
  assert.equal(wrongIdentity.success, false);

  const staffPage = context.getPasswordAccountPage('', 'ALL', 1, 50, 'staff-token', '127.0.0.1');
  const staffReveal = context.getVjPasswordForAdmin(duplicateKim, '009999', 'staff-token', '127.0.0.1', 'test');
  assert.equal(staffPage.success, false);
  assert.equal(staffReveal.success, false);
  assert.equal(Object.hasOwn(staffReveal, 'password'), false);

  const allowed = context.getVjPasswordForAdmin(duplicateKim, '001234', 'admin-token', '127.0.0.1', 'test');
  assert.equal(allowed.success, true);
  assert.equal(allowed.stageName, 'Kim');
  assert.equal(allowed.username, '@kim.two');
  assert.equal(allowed.team, 'Team B');
  assert.equal(allowed.role, 'MC');
  assert.equal(allowed.password, '2222');
  const hashed = context.getVjPasswordForAdmin(accountPage.accounts[2], '001234', 'admin-token', '127.0.0.1', 'test');
  assert.equal(hashed.success, false);
  assert.match(hashed.message, /hash/);
  assert.equal(JSON.stringify(hashed).includes(hashValue), false);

  const realLogRows = auditEntries.filter(Array.isArray).map((entry, index) => [
    '2026-11-02 12:00:' + String(index).padStart(2, '0'), entry[0], entry[1], entry[2], entry[3], entry[4] || '-', entry[5] || '-', 'NOVEMBER 2026'
  ]);
  sheets.Logs = sheet([logRows[0], ...realLogRows]);
  const history = context.getPasswordAuditHistory({ query: '@kim.two', team: 'Team B', role: 'MC', activity: 'VIEWED' }, 1, 50, 'admin-token', '127.0.0.1');
  assert.equal(history.success, true);
  assert.equal(history.total, 1);
  assert.equal(history.records[0].stageName, 'Kim');
  assert.equal(history.records[0].username, '@kim.two');
  assert.equal(JSON.stringify(history).includes('2222'), false);
  assert.equal(context.isStoredPasswordHash_('$pbkdf2-sha256$600000$salt$digest', false), true);
  assert.equal(context.isStoredPasswordHash_('pbkdf2_sha256$600000$salt$digest', false), true);

  adminRows.push(['Admin', '008888', 'ADMIN']);
  assert.equal(context.getAdminSecurityContext_('Admin').role, 'ADMIN', 'duplicate Admin names must fail closed');
  assert.equal(context.getPasswordAccountPage('', 'ALL', 1, 50, 'admin-token', '127.0.0.1').success, false);
  assert.equal(context.getVjPasswordForAdmin(duplicateKim, '001234', 'admin-token', '127.0.0.1', 'test').success, false);
  const allLogs = JSON.stringify(auditEntries);
  assert.equal(allLogs.includes('001234'), false);
  assert.equal(allLogs.includes('009999'), false);
  assert.equal(allLogs.includes('2222'), false);
  assert.equal(allLogs.includes(hashValue), false);
});

test('frontend includes transient retry, bridge guard, and history continuation control', () => {
  assert.ok(/HTTP\s*-1/.test(frontendSource));
  assert.ok(frontendSource.includes('if (!hasAppsScriptBridge())'));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('VJ')"));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('ADMIN')"));
});

test('frontend pages User Mapping and requests one selected account only after Super Admin re-authentication', () => {
  assert.ok(frontendSource.includes("reportMode: true"));
  assert.ok(frontendSource.includes('pageSize: USER_MAPPING_PAGE_SIZE'));
  assert.ok(frontendSource.includes("{ historyOnly: true }"));
  assert.ok(frontendSource.includes(".getUserMappingData(userName, state.userType, token, clientIpAddress, 0"));
  assert.ok(frontendSource.includes('.getVjPasswordForAdmin(selected, adminPin, currentAdminToken'));
  assert.ok(frontendSource.includes('.getPasswordAccountPage('));
  assert.ok(frontendSource.includes('.getPasswordAuditHistory('));
  assert.match(frontendSource, /setTimeout\(function\(\)\s*\{\s*clearAdminPasswordReveal\(\);/);
  assert.ok(frontendSource.includes('}, 15000);'));
  assert.ok(frontendSource.includes('type="password" autocomplete="current-password"'));
});

test('legacy login still compares the stored password directly and does not force a reset', () => {
  const login = backendSource.match(/function verifyAndGetPersonalData\([\s\S]*?(?=function getAdminDashboardData)/);
  assert.ok(login);
  assert.ok(login[0].includes('storedPin !== targetPass'));
  assert.equal(/force.{0,20}reset|password.{0,20}reset/i.test(login[0]), false);
});

test('login accepts the unchanged stored password and rejects a different password', () => {
  const accounts = [
    ['Full Name', 'Stage Name', 'Team', 'Role', 'Shift', 'TikTok ID', 'Password'],
    ['Kim Full', 'Kim', 'Team A', 'MC', '09:00-18:00', '@kim', '0042']
  ];
  const logs = [];
  const context = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'vj mc all' ? {
      getDataRange: () => ({ getValues: () => accounts })
    } : null }) }
  };
  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });
  context.isLoginRateLimited_ = () => false;
  context.checkAccessPermission_ = () => ({ allowed: true });
  context.recordLoginFailure_ = () => {};
  context.clearLoginFailures_ = () => {};
  context.getProfilePicUrl_ = () => '';
  context.getAllMonthData_ = () => ({ list: [{ vjName: 'Kim', currentDiamond: 100, targetDiamond: 200, team: 'Team A' }] });
  context.getAttendanceRecords_ = () => ({ success: true, records: [], sourceWarning: '' });
  context.getActiveAnnouncements_ = () => ({ success: true, list: [] });
  context.getSystemVersionInfo = () => ({ success: true, version: 'test' });
  context.issueSession_ = () => 'test-session';
  context.writeLog_ = (...entry) => logs.push(entry);

  const accepted = context.verifyAndGetPersonalData('Kim', '0042', 'NOVEMBER 2026', '127.0.0.1', 'test-device', '');
  const rejected = context.verifyAndGetPersonalData('Kim', '0043', 'NOVEMBER 2026', '127.0.0.1', 'test-device', '');
  assert.equal(accepted.success, true);
  assert.equal(accepted.sessionToken, 'test-session');
  assert.equal(rejected.success, false);
  assert.equal(JSON.stringify(logs).includes('0042'), false);
});

test('table group builder groups by team/date and exposes native keyboard-operable controls', () => {
  const inlineScripts = [...frontendSource.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)].map(match => match[2]);
  const source = inlineScripts.find(script => script.includes('function buildAccessibleTableGroups'));
  assert.ok(source);
  const start = source.indexOf('function buildAccessibleTableGroups');
  const end = source.indexOf('\n    function applyAccessibleGroupingToBody', start);
  const groupBuilderSource = source.slice(start, end) + '\nthis.buildAccessibleTableGroups = buildAccessibleTableGroups;';
  const context = {};
  vm.createContext(context);
  vm.runInContext(groupBuilderSource, context);
  const groups = context.buildAccessibleTableGroups(['VJ', 'ทีม', 'ตำแหน่ง'], [['A', 'Team 1', 'VJ'], ['B', 'Team 2', 'MC'], ['C', 'Team 1', 'HR']], 'team');
  assert.equal(groups.groupColumn, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(groups.groups.map(group => [group.label, group.rowIndexes.length]))), [['Team 1', 2], ['Team 2', 1]]);
  const byDate = context.buildAccessibleTableGroups(['วัน / เวลา', 'VJ'], [['2026-11-01 09:00', 'A'], ['2026-11-01 15:30', 'B'], ['2026-11-02 09:00', 'C']], 'date');
  assert.deepEqual(JSON.parse(JSON.stringify(byDate.groups.map(group => [group.label, group.rowIndexes.length]))), [['2026-11-01', 2], ['2026-11-02', 1]]);
  const byTimeColumn = context.buildAccessibleTableGroups(['เวลา', 'ผู้ส่ง'], [['2026-11-01 09:00', 'A'], ['2026-11-01 15:30', 'B']], 'date');
  assert.deepEqual(JSON.parse(JSON.stringify(byTimeColumn.groups.map(group => [group.label, group.rowIndexes.length]))), [['2026-11-01', 2]]);
  const byRank = context.buildAccessibleTableGroups(['#', 'ผู้ส่ง'], [['1', 'A'], ['12', 'B']], 'rank');
  assert.deepEqual(JSON.parse(JSON.stringify(byRank.groups.map(group => [group.label, group.rowIndexes.length]))), [['1–10 อันดับ', 1], ['11–20 อันดับ', 1]]);
  assert.ok(frontendSource.includes("toggle.type = 'button'"));
  assert.ok(frontendSource.includes("toggle.setAttribute('aria-expanded'"));
  assert.ok(frontendSource.includes('rows[index].hidden = !expanded'));
  assert.ok(frontendSource.includes('aceGroupingModes'));
});

test('sidebar categories collapse to an icon rail on desktop and remain dismissible on mobile', () => {
  assert.match(frontendSource, /body\.ace-shell\.ace-sidebar-collapsed \{ padding-left:\s*78px/);
  assert.ok(frontendSource.includes("localStorage.setItem('ace_sidebar_collapsed'"));
  assert.ok(frontendSource.includes("matchMedia('(max-width: 767px)')"));
  assert.ok(frontendSource.includes("side.classList.toggle('open', open)"));
  assert.ok(frontendSource.includes("backdrop.classList.toggle('open', open)"));
  assert.ok(frontendSource.includes("event.key==='Escape'"));
  assert.ok(frontendSource.includes("{ title:'ภาพรวม'"));
  assert.ok(frontendSource.includes("{ title:'จัดการข้อมูล'"));
  assert.ok(frontendSource.includes("toggle.setAttribute('aria-label', collapsed ? 'ขยายเมนู' : 'ย่อเมนู')"));
  assert.ok(frontendSource.includes('ace-mobile-user-nav'));
});

test('inline frontend JavaScript parses without syntax errors', () => {
  const scriptTag = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let match;
  let parsedCount = 0;
  while ((match = scriptTag.exec(frontendSource))) {
    if (/\bsrc\s*=/.test(match[1]) || !match[2].trim()) continue;
    parsedCount++;
    assert.doesNotThrow(() => new vm.Script(match[2], { filename: 'index-inline-' + parsedCount + '.js' }));
  }
  assert.ok(parsedCount > 0, 'expected at least one inline application script');
});

test('both User Mapping views provide report filters, summary cards, and history tabs', () => {
  for (const id of [
    'vjUserMapTeamFilter', 'vjUserMapDateFrom', 'vjUserMapDateTo', 'vjUserMapOverviewPanel', 'vjUserMapInsightGrid', 'vjUserMapTabHistory', 'vjUserMapTabVjTeams',
    'admUserMapTeamFilter', 'admUserMapDateFrom', 'admUserMapDateTo', 'admUserMapOverviewPanel', 'admUserMapInsightGrid', 'admUserMapTabHistory', 'admUserMapTabVjTeams',
    'userMapVjDetailModal', 'userMapVjDetailTabGifters', 'userMapVjDetailTabGifts'
  ]) {
    assert.ok(frontendSource.includes('id="' + id + '"'), 'missing report UI element: ' + id);
  }
  assert.ok(frontendSource.includes("setUserMapQuickRange('VJ','week')"));
  assert.ok(frontendSource.includes("setUserMapQuickRange('ADMIN','month')"));
  assert.ok(frontendSource.includes('function renderUserMapTeamSummary'));
  assert.ok(frontendSource.includes('function renderUserMapOverview'));
  assert.ok(frontendSource.includes('function renderUserMapVjTeams'));
  assert.ok(frontendSource.includes('function openUserMapVjDetail'));
  assert.ok(frontendSource.includes('.ace-um-detail-search { width:min(280px,35%)'));
  assert.ok(frontendSource.includes('function renderUserMapVjDetailIdentity(vjTikTokId)'));
  assert.ok(frontendSource.includes("loadMappingAvatarCache([{ vjTikTokId:state.vjTikTokId }])"));
  assert.ok(frontendSource.includes('.ace-um-detail-tabs-list { width:100%; gap:.85rem; }'));
  assert.ok(frontendSource.includes('var USER_MAPPING_PAGE_SIZE = 50;'));
  assert.ok(frontendSource.includes('id="vjUserMapPageText"'));
  assert.ok(frontendSource.includes('id="admUserMapPageText"'));
  assert.ok(frontendSource.includes('function changeUserMapCurrentPage'));
});

test('VJ mobile navigation uses a safe-area-aware bottom bar and keeps all user views reachable', () => {
  for (const view of ['perf', 'userMap', 'att', 'ann']) {
    assert.ok(frontendSource.includes("onclick=\"switchVJTab('" + view + "')\""), 'missing mobile navigation action: ' + view);
  }
  assert.ok(frontendSource.includes('id="aceMobileUserNav"'));
  assert.ok(frontendSource.includes("document.body.classList.toggle('ace-vj-session', !admin && vj)"));
  assert.ok(frontendSource.includes('function updateAceVjNavigation'));
  assert.ok(frontendSource.includes('env(safe-area-inset-bottom)'));
  assert.ok(frontendSource.includes('body.ace-vj-session #aceMobileUserNav'));
});

function createGiftReportHarness(giftRows, mappingRows, sessionType, sessionName) {
  const reads = { databaseOpen: 0, giftRead: 0, mappingRead: 0, ranges: [] };
  function sheet(rows, type) {
    return {
      getLastRow: () => rows.length,
      getLastColumn: () => rows[0] ? rows[0].length : 0,
      getRange: (row, column, rowCount, columnCount) => ({
        getValues: () => {
          reads[type]++;
          reads.ranges.push({ type, row, column, rowCount, columnCount });
          return rows.slice(row - 1, row - 1 + rowCount).map(values => values.slice(column - 1, column - 1 + columnCount));
        }
      })
    };
  }
  const giftSheet = sheet(giftRows, 'giftRead');
  const mappingSheet = sheet(mappingRows, 'mappingRead');
  const database = {
    getSheetByName: name => name === 'Gift_Details' ? giftSheet : (name === 'User_Mapping' ? mappingSheet : null)
  };
  const context = { console, Logger: { log: () => {} } };
  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });
  context.requireSession_ = () => ({ success: true, session: { userName: sessionName, userType: sessionType } });
  context.getUserMappingDatabase_ = () => { reads.databaseOpen++; return database; };
  return { context, reads };
}

const giftHeader = ['Record Key', 'Event Time', 'Round', 'VJ', 'Gifter Name', 'TikTok ID', 'Bound', 'Gift Name', 'Quantity', 'Unit Price', 'Total Coins', 'Collection ID', 'Uploaded By'];

test('Gift_Details report joins TikTok ID to User_Mapping without changing the sheet schema', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 2, 1, 2, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Lee', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader'],
    ['r3', '2026-10-30 09:00:00', 'R2', 'Kim', 'Gifter C', 'giver-c', 'Bound', 'Heart', 4, 1, 4, 'c3', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-a', 'Mapped A', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0],
    ['@giver-c', 'Mapped C', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'ADMIN', 'Admin');
  const result = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { team: 'Team A' });

  assert.equal(result.success, true);
  assert.equal(result.totalRecords, 2);
  assert.equal(result.summary.totalCoins, 6);
  assert.equal(result.summary.mappedRecords, 2);
  assert.equal(result.list.length, 2);
  assert.equal(result.list[0].mappedUserName, 'Mapped A');
  assert.equal(result.summaryByMapping[0].tiktokId, 'giver-c');
  assert.equal(harness.reads.databaseOpen, 1);
  assert.equal(harness.reads.giftRead, 1);
  assert.equal(harness.reads.mappingRead, 1);
  assert.deepEqual(harness.reads.ranges, [
    { type: 'giftRead', row: 1, column: 2, rowCount: gifts.length, columnCount: 10 },
    { type: 'mappingRead', row: 1, column: 1, rowCount: mappings.length, columnCount: 5 }
  ]);
});

test('VJ Gift_Details report only returns own VJ rows and own mapped rows with blank VJ', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 2, 1, 2, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Lee', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader'],
    ['r3', '2026-10-30 09:00:00', 'R2', '', 'Gifter C', '@giver-c', 'Bound', 'Heart', 4, 1, 4, 'c3', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-a', 'Mapped A', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0],
    ['@giver-b', 'Mapped B', '', 'Lee', 'Team B', 0, 0, TODAY, 'Admin', 0],
    ['@giver-c', 'Mapped C', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'VJ', 'Kim');
  const result = harness.context.getGiftDetailsData('vj-token', '127.0.0.1', {});

  assert.equal(result.success, true);
  assert.equal(result.totalRecords, 2);
  assert.deepEqual(Array.from(result.list, row => row.tiktokId), ['@giver-a', '@giver-c']);
  assert.equal(result.summary.totalCoins, 6);
  assert.equal(result.list.some(row => row.tiktokId === '@giver-b'), false);
});

test('VJ detail report filters Gift_Details by VJ and returns aggregated top gifters', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 2, 1, 2, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Lee', 'Gifter A', '@giver-a', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader'],
    ['r3', '2026-10-31 12:00:00', 'R2', '', 'Gifter B', '@mapped-user', 'Bound', 'Heart', 4, 1, 4, 'c3', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@mapped-user', 'Mapped User', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'ADMIN', 'Admin');
  const result = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { vj: 'Kim' });

  assert.equal(result.totalRecords, 2);
  assert.deepEqual(Array.from(result.list, row => row.vj || row.mappedVJ), ['Kim', 'Kim']);
  assert.equal(result.summaryByGifter.length, 2);
  assert.equal(result.summaryByGifter[0].tiktokId, '@mapped-user');
  assert.equal(result.summaryByGifter[0].totalCoins, 4);
});

test('VJ Gift_Details rows do not reveal User Mapping details owned by a different VJ', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-b', 'Mapped B', '', 'Lee', 'Team B', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'VJ', 'Kim');
  const result = harness.context.getGiftDetailsData('vj-token', '127.0.0.1', {});

  assert.equal(result.totalRecords, 1);
  assert.equal(result.list[0].isMapped, false);
  assert.equal(result.list[0].mappedUserName, '');
  assert.equal(result.list[0].team, '');
  assert.equal(result.summaryByMapping.length, 0);
});

test('Gift_Details report pages rows and keeps whole-filter totals', () => {
  const gifts = [giftHeader];
  for (let index = 1; index <= 51; index++) {
    gifts.push(['r' + index, '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter ' + index, 'gifter-' + index, 'Bound', 'Rose', 1, 1, 1, 'c' + index, 'uploader']);
  }
  const harness = createGiftReportHarness(gifts, [liveHeader], 'ADMIN', 'Admin');
  const firstPage = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { page: 1 });
  const secondPage = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { page: 2 });

  assert.equal(firstPage.list.length, 50);
  assert.equal(firstPage.totalPages, 2);
  assert.equal(firstPage.totalRecords, 51);
  assert.equal(firstPage.summary.totalCoins, 51);
  assert.equal(secondPage.list.length, 1);
  assert.equal(secondPage.page, 2);
  assert.equal(Object.hasOwn(firstPage.list[0], 'sourceRow'), false);
});

test('Gift_Details column H gift name builds a direct BeetGames image URL', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 1, 1, 1, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Kim', 'Gifter B', '@giver-b', 'Bound', 'Cow-napping', 1, 1, 1, 'c2', 'uploader'],
    ['r3', '2026-10-31 12:00:00', 'R1', 'Kim', 'Gifter C', '@giver-c', 'Bound', '10/10', 1, 1, 1, 'c3', 'uploader']
  ];
  const harness = createGiftReportHarness(gifts, [liveHeader], 'ADMIN', 'Admin');
  let fetchCount = 0;
  harness.context.UrlFetchApp = { fetch: () => { fetchCount++; throw new Error('direct URL generation must not fetch the catalog'); } };

  const first = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', {});
  const byName = Object.fromEntries(Array.from(first.list, item => [item.giftName, item.giftImageUrl]));

  assert.equal(byName.Rose, 'https://beetgames.com/images/gifts/rose.webp');
  assert.equal(byName['Cow-napping'], 'https://beetgames.com/images/gifts/cow_napping.webp');
  assert.equal(byName['10/10'], 'https://beetgames.com/images/gifts/10_10.webp');
  assert.equal(fetchCount, 0);
  assert.ok(backendSource.includes("giftName: giftDetailsColumn_(headers, ['Gift Name', 'GiftName', 'Gift'], 6)"), 'Gift Name resolves to column H within the B:K read range');
});

test('both User Mapping views expose Gift_Details report tabs and filtered pages', () => {
  for (const id of [
    'vjUserMapTabGifts', 'vjUserMapGiftsPanel', 'vjGiftDateFrom', 'vjGiftTeamFilter', 'vjGiftReportTbody',
    'admUserMapTabGifts', 'admUserMapGiftsPanel', 'admGiftDateFrom', 'admGiftTeamFilter', 'admGiftReportTbody'
  ]) {
    assert.ok(frontendSource.includes('id="' + id + '"'), 'missing Gift_Details UI element: ' + id);
  }
  assert.ok(frontendSource.includes('.getGiftDetailsData(sessionToken, clientIpAddress, options)'));
  assert.ok(frontendSource.includes('function renderGiftDetailsReport'));
  assert.ok(frontendSource.includes('function giftReportProfileCell'));
  assert.ok(frontendSource.includes('function giftReportGiftCell'));
  assert.ok(frontendSource.includes('var cachedMappingAvatarUrls'));
  assert.ok(frontendSource.includes('loadMappingAvatarCache(pageRows)'));
  assert.ok(frontendSource.includes('loadMappingAvatarCache(giftAvatarRows)'));
  assert.ok(!frontendSource.includes('loadMappingAvatarCache(cachedUserMapData)'));
  assert.ok(frontendSource.includes('cachedMappingAvatarLookups[lookupKey]'));
  assert.ok(frontendSource.includes('pendingMappingAvatarLookups[lookupKey]'));
  assert.ok(backendSource.includes('function giftImageUrlFromName_(giftName)'));
  assert.ok(frontendSource.includes('beetgames\\.com\\/images\\/gifts'));
});

test('gift images load lazily with a fallback in both report and VJ gift detail tables', () => {
  assert.ok(frontendSource.includes("image.loading = 'lazy'"));
  assert.ok(frontendSource.includes("image.decoding = 'async'"));
  assert.ok(frontendSource.includes("image.onerror = function() { if (image.parentNode) image.parentNode.replaceChild(placeholder, image); }"));
  assert.match(frontendSource, /function renderUserMapVjDetailGifts\(result\)[\s\S]*?giftReportGiftCell\(row,item\)/);
  assert.ok(backendSource.includes('item.giftImageUrl = giftImageUrlFromName_(item.giftName)'));
});
