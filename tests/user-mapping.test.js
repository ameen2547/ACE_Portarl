const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const backendSource = fs.readFileSync(path.join(projectRoot, 'code.gs'), 'utf8');
const frontendSource = fs.readFileSync(path.join(projectRoot, 'index.html'), 'utf8');
const TODAY = '2026-10-30';

function makeSheet(rows) {
  return {
    getDataRange() {
      return { getValues: () => rows };
    },
    getLastRow() {
      return rows.length;
    }
  };
}

function createHarness(liveRows, historyRows) {
  const properties = new Map();
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

  const liveSheet = makeSheet(liveRows);
  const historySheet = makeSheet(historyRows);
  context.requireSession_ = token => ({
    success: true,
    session: token === 'admin-token'
      ? { userName: 'Admin', userType: 'ADMIN' }
      : { userName: 'Kim', userType: 'VJ' }
  });
  context.getUserMappingDatabase_ = () => ({});
  context.getOrCreateUserMappingSheet_ = () => liveSheet;
  context.getOrCreateUserMappingHistorySheet_ = () => historySheet;
  context.userMappingSnapshotDate_ = () => TODAY;
  context.archiveUserMappingDailySnapshot_ = (_database, _rows, existingHistory) => {
    archiveCallCount++;
    return existingHistory;
  };

  return {
    context,
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

test('frontend includes transient retry, bridge guard, and history continuation control', () => {
  assert.ok(/HTTP\s*-1/.test(frontendSource));
  assert.ok(frontendSource.includes('if (!hasAppsScriptBridge())'));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('VJ')"));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('ADMIN')"));
});
